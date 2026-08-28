"""Handing a downloaded attachment to the caller, and taking it off disk again.

``outlook_get_attachment`` writes what the mailbox holds to the filesystem of
the machine the *server* runs on. Over stdio that is also the caller's machine,
so the path in the answer is a file they can open and there is nothing more to
arrange. Over HTTP it is not: the file lands under the account the service runs
as, on a host the agent may never reach, and a path in a tool result is useless
to whoever asked for it.

This module is the bridge, and it is four decisions:

* **One route, /attachments/<token>.** A caller never names a path, only an
  opaque token this server minted, so there is nothing to traverse out of and no
  way to ask for a file the server did not offer.
* **The token is the whole credential.** The route asks for nothing else: no
  identity header, and no credential of the deployment's own. That is not a gap,
  it is the requirement. The party that has to fetch the file is the agent that
  just called the tool, and it reaches this server through an MCP client that
  holds the reverse proxy's key without ever showing it to the agent. A link the
  agent cannot follow is not a link. So the proxy has to let /attachments/
  through unauthenticated, and the token carries the weight: 256 unguessable
  bits naming one file, minted per call.
* **Fetching it destroys it.** The first successful GET pops the ticket and
  deletes the file from the server, so the link in the transcript is dead by the
  time the answer is written, and a second reader gets a 404. What nobody
  fetches is deleted anyway once [attachments].retention_minutes runs out, by
  ``reap_expired_downloads``. Both halves matter: the first is what makes an
  unauthenticated link safe, the second is what stops a host from slowly filling
  with other people's mail. The ticket table lives in memory, so a restart
  invalidates every outstanding link, which costs one repeated tool call.
* **Every file is filed under its user and its message**, at
  ``<download_path>/<user digest>/msg-<message digest>/<filename>``. The first
  level keeps one person's attachments out of another's directory listing, the
  same isolation the token caches have and for the same reason. The second is
  what makes deletion possible without an index to keep: the files downloaded
  from a message are exactly the contents of one directory, so
  ``outlook_delete_attachment_files`` can empty it without ever being told a
  path by the caller.

Over stdio the user level is absent (there is only one person, and the download
directory is their own), but the message level is not: the same tool deletes the
same way on both transports. Nothing else here applies over stdio, and the
expiry least of all: there the file the tool wrote *is* the answer, sitting in
the caller's own download directory, and a server that deleted it an hour later
would be deleting the user's file out from under them.
"""

import hashlib
import logging
import secrets
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Set

import anyio
from starlette.background import BackgroundTask
from starlette.requests import Request
from starlette.responses import FileResponse, PlainTextResponse, Response

from .app import get_config, mcp
from .auth import user_digest
from .config import DOWNLOAD_ROUTE

logger = logging.getLogger("outlook_mcp")

# A link is meant to be fetched by the agent that just asked for the file. Long
# enough to survive a slow client or a person clicking it by hand, short enough
# that a transcript kept afterwards holds nothing usable.
TICKET_TTL_SECONDS = 15 * 60

# Every tool call can mint one, so the table is bounded like the enrollment one.
MAX_PENDING_TICKETS = 256

# How often the reaper looks for downloads whose retention has run out. Short
# enough that "an hour" means about an hour, cheap enough to ignore: one walk of
# one directory tree.
SWEEP_INTERVAL_SECONDS = 5 * 60

# Directory name standing for one message. Shortened from the full digest
# because it only has to be unique within one mailbox's download directory, and
# a path a person may have to look at should stay readable.
MESSAGE_DIR_PREFIX = "msg-"
MESSAGE_DIGEST_CHARS = 16


@dataclass(frozen=True)
class _Ticket:
    """One file, offered once, to one user."""

    path: Path
    name: str
    content_type: str
    user: str
    minted_at: float


# token -> ticket.
_pending: Dict[str, _Ticket] = {}


def _prune(now: float) -> None:
    for token, ticket in list(_pending.items()):
        if now - ticket.minted_at > TICKET_TTL_SECONDS:
            del _pending[token]


def _mint(path: Path, name: str, content_type: str, user: str) -> str:
    now = time.monotonic()
    _prune(now)
    if len(_pending) >= MAX_PENDING_TICKETS:
        # Drop the oldest rather than refuse: a full table would otherwise let
        # one abandoned link stop everybody else from downloading anything.
        oldest = min(_pending, key=lambda t: _pending[t].minted_at)
        del _pending[oldest]
    token = secrets.token_urlsafe(32)
    _pending[token] = _Ticket(path, name, content_type, user, now)
    return token


def _take(token: str) -> Optional[_Ticket]:
    """The ticket named by `token`, consumed so the link cannot be replayed."""
    ticket = _pending.pop(token, None)
    if ticket is None:
        return None
    if time.monotonic() - ticket.minted_at > TICKET_TTL_SECONDS:
        return None
    return ticket


def _consume_file(path: Path) -> None:
    """Delete a file that has just been served, and forget any link left to it.

    Called after the response has been sent and never before: the body is
    streamed off this very file. A second ticket can point at the same path when
    the tool was called twice for one attachment, so the revoke is not
    bookkeeping, it is what stops the other link from later serving nothing.
    """
    try:
        path.unlink()
    except OSError:
        # Already gone, or a permission problem worth seeing in the log. Either
        # way the download itself succeeded, so there is nothing to fail here.
        logger.warning("Could not delete %s after serving it", path)
        return
    _revoke({path})
    _remove_if_empty(path.parent)


def _revoke(paths: Set[Path]) -> None:
    """Drop the outstanding links to files that no longer exist.

    Both sides are resolved before comparing: a ticket holds the path the tool
    reported, which is absolute, while a deletion walks the directory, which is
    only absolute if [attachments].download_path was.
    """
    gone = {path.resolve() for path in paths}
    for token, ticket in list(_pending.items()):
        if ticket.path.resolve() in gone:
            del _pending[token]


def download_root(user: Optional[str], base: Optional[Path] = None) -> Path:
    """The directory one caller's downloads live under."""
    root = base if base is not None else get_config().attachment_dir
    if user is None:
        return root
    return root / user_digest(user)


def message_dir(user: Optional[str], message_id: str, base: Optional[Path] = None) -> Path:
    """The directory holding what was downloaded from one message.

    Derived from the message id rather than recorded anywhere: a Graph message
    id is neither short nor a legal filename (it carries '/' and '+'), and a
    directory that can be recomputed cannot drift out of step with an index.
    """
    digest = hashlib.sha256(message_id.encode("utf-8")).hexdigest()[:MESSAGE_DIGEST_CHARS]
    return download_root(user, base) / f"{MESSAGE_DIR_PREFIX}{digest}"


def offer(path: Path, name: str, content_type: str, user: Optional[str]) -> Optional[str]:
    """A URL this file can be fetched at, or None when the caller needs none.

    None over stdio, where the caller already has the file, and None over HTTP
    without [auth].public_url, where the server cannot say what URL it is
    reachable on.
    """
    config = get_config()
    if user is None or not config.downloads_enabled:
        return None
    return config.download_url(_mint(path, name, content_type, user))


def delete_message_downloads(
    user: Optional[str], message_id: str, filename: Optional[str] = None
) -> List[str]:
    """Remove the files downloaded from one message. Returns the names removed.

    The caller chooses a message, never a path: the directory comes from their
    own identity and the message id, and `filename` is reduced to a bare name
    inside it, so nothing outside that one directory is reachable from here.
    """
    directory = message_dir(user, message_id)
    if not directory.is_dir():
        return []

    if filename is not None:
        safe = Path(filename).name
        targets = [directory / safe] if safe and (directory / safe).is_file() else []
    else:
        targets = sorted(child for child in directory.iterdir() if child.is_file())

    removed: List[str] = []
    for target in targets:
        target.unlink()
        removed.append(target.name)

    _revoke(set(targets))
    _remove_if_empty(directory)
    return removed


def _remove_if_empty(directory: Path) -> None:
    try:
        directory.rmdir()
    except OSError:
        # Still holds files, or is gone already. Either way there is nothing to
        # clean up, and an empty directory left behind is not worth an error.
        pass


def sweep() -> List[Path]:
    """Delete the downloads nobody fetched before their retention ran out.

    Walks the filesystem rather than the ticket table, because the table is the
    smaller of the two: a link expires after TICKET_TTL_SECONDS while the file
    it named lives on, a restart empties the table entirely, and a tool call the
    caller never followed up on leaves a file with no ticket at all. The
    directory tree is the only complete record of what is on this disk.

    Age is the file's mtime against the wall clock, not the monotonic clock the
    tickets use: a file outlives the process that wrote it.
    """
    config = get_config()
    retention = config.retention_seconds
    root = config.attachment_dir
    if retention is None or not root.is_dir():
        return []

    cutoff = time.time() - retention
    removed: List[Path] = []
    for path in root.rglob("*"):
        try:
            if not path.is_file() or path.stat().st_mtime > cutoff:
                continue
            path.unlink()
        except OSError:
            logger.warning("Could not delete the expired download %s", path)
            continue
        removed.append(path)

    if removed:
        _revoke(set(removed))
        _prune_empty_dirs(root)
        logger.info("Deleted %d expired download(s)", len(removed))
    return removed


def _prune_empty_dirs(root: Path) -> None:
    """Remove the per-message and per-user directories left empty. Never root.

    Deepest first, so emptying a message directory lets its user directory go in
    the same pass. `root` itself stays: it is configuration, not something this
    module created, and the next download needs it anyway.
    """
    for path in sorted(root.rglob("*"), key=lambda p: len(p.parts), reverse=True):
        if path.is_dir():
            _remove_if_empty(path)


async def reap_expired_downloads() -> None:
    """Sweep on a timer, for as long as the server runs.

    A sweep on the next tool call would be simpler and would be wrong: nothing
    says there will be a next call, and the file nobody came back for is exactly
    the one that has to go. Started by the lifespan and cancelled with it.
    """
    while True:
        try:
            sweep()
        except Exception:
            # A failed sweep must not take the timer down with it: the next one
            # would have cleaned up the same files anyway.
            logger.exception("Download sweep failed")
        await anyio.sleep(SWEEP_INTERVAL_SECONDS)


def _refuse(message: str, status: int) -> PlainTextResponse:
    return PlainTextResponse(message, status_code=status)


@mcp.custom_route(DOWNLOAD_ROUTE, methods=["GET"])
async def download_attachment(request: Request) -> Response:
    """Serve one attachment once, and take it off the server on the way out."""
    config = get_config()
    if not config.downloads_enabled:
        return _refuse("Not found", 404)
    if request.method != "GET":
        # Starlette answers HEAD on a GET route by itself; refuse it rather than
        # let a probe burn the one fetch the link is good for.
        return _refuse("Use GET to fetch this file.", 405)

    ticket = _take(request.path_params.get("token", ""))
    if ticket is None:
        return _refuse(
            "This download link is not valid: it has already been used, or it "
            "has expired. Ask for the attachment again to get a new one.",
            404,
        )
    if not ticket.path.is_file():
        return _refuse(
            "This attachment is no longer on the server. Ask for it again.", 410
        )

    # The user is logged rather than checked: the token was minted for them and
    # the proxy asserts nothing on this route, so it is the record of whose
    # mailbox a served file came out of, not a second gate.
    logger.info("Serving %s, downloaded for %s", ticket.name, ticket.user)
    return FileResponse(
        ticket.path,
        media_type=ticket.content_type,
        filename=ticket.name,
        background=BackgroundTask(_consume_file, ticket.path),
    )
