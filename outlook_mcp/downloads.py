"""Handing a downloaded attachment to the caller, and taking it off disk again.

``outlook_get_attachment`` writes what the mailbox holds to the filesystem of
the machine the *server* runs on. Over stdio that is also the caller's machine,
so the path in the answer is a file they can open and there is nothing more to
arrange. Over HTTP it is not: the file lands under the account the service runs
as, on a host the agent may never reach, and a path in a tool result is useless
to whoever asked for it.

This module is the bridge, and it is three decisions:

* **One route, /attachments/<token>.** A caller never names a path, only an
  opaque token this server minted, so there is nothing to traverse out of and no
  way to ask for a file the server did not offer.
* **A token belongs to one user and burns on use.** The route runs the same
  proxy identity policy the tools do and refuses a token minted for anybody
  else, so a link that ends up in a shared transcript is worth nothing to a
  second reader and nothing at all once it has been fetched. The table lives in
  memory: a restart costs one repeated tool call, while persisting it would keep
  live download links on disk.
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
same way on both transports.
"""

import hashlib
import logging
import secrets
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Set

from starlette.requests import Request
from starlette.responses import FileResponse, PlainTextResponse, Response

from .app import get_config, mcp, proxy_auth_policy
from .auth import CredentialsError, user_digest
from .config import DOWNLOAD_ROUTE

logger = logging.getLogger("outlook_mcp")

# A link is meant to be fetched by the agent that just asked for the file. Long
# enough to survive a slow client or a person clicking it by hand, short enough
# that a transcript kept afterwards holds nothing usable.
TICKET_TTL_SECONDS = 15 * 60

# Every tool call can mint one, so the table is bounded like the enrollment one.
MAX_PENDING_TICKETS = 256

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


def _take(token: str, user: str) -> Optional[_Ticket]:
    """The ticket this user was given under `token`, consumed so it cannot replay."""
    ticket = _pending.pop(token, None)
    if ticket is None:
        return None
    if time.monotonic() - ticket.minted_at > TICKET_TTL_SECONDS:
        return None
    # Consumed above before this check on purpose: the realistic way another
    # user arrives holding a token is that the link leaked, and a leaked link
    # should be dead rather than still redeemable by the person it was for.
    if ticket.user.strip().casefold() != user.strip().casefold():
        logger.warning("Download token offered by a user it was not minted for")
        return None
    return ticket


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


def _refuse(message: str, status: int) -> PlainTextResponse:
    return PlainTextResponse(message, status_code=status)


@mcp.custom_route(DOWNLOAD_ROUTE, methods=["GET"])
async def download_attachment(request: Request) -> Response:
    """Serve one attachment to the user the link was minted for, once."""
    config = get_config()
    policy = proxy_auth_policy(config)
    if not config.downloads_enabled or policy is None:
        return _refuse("Not found", 404)
    if request.method != "GET":
        # Starlette answers HEAD on a GET route by itself; refuse it rather than
        # let a probe burn the one fetch the link is good for.
        return _refuse("Use GET to fetch this file.", 405)

    try:
        user = policy.user_from_headers(request.headers)
    except CredentialsError as e:
        return _refuse(str(e), 403)

    ticket = _take(request.path_params.get("token", ""), user)
    if ticket is None:
        return _refuse(
            "This download link is not valid: it has already been used, it has "
            "expired, or it was issued for somebody else. Ask for the attachment "
            "again to get a new one.",
            404,
        )
    if not ticket.path.is_file():
        return _refuse(
            "This attachment is no longer on the server. Ask for it again.", 410
        )

    logger.info("Serving %s to %s", ticket.name, user)
    return FileResponse(
        ticket.path, media_type=ticket.content_type, filename=ticket.name
    )
