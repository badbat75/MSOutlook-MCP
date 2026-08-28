"""Unit tests for outlook_mcp.downloads: getting a file to a remote caller.

Four things have to hold, and none of them needs a network to check. A file one
user downloaded must not appear in another's directory or be deletable by them.
A download link must work exactly once, and serving it must take the file with
it, since the link itself asks for no credential. A file nobody fetched must go
by itself. And a caller must never be able to steer any of it with a path of its
own: it names a message, the server names the file.
"""

import os
import time
from pathlib import Path

import anyio
import pytest

from outlook_mcp import downloads
from outlook_mcp.config import ServerConfig

ADA = "ada@example.com"
BOB = "bob@example.com"
MESSAGE = "AAMkAGI2T=="


@pytest.fixture(autouse=True)
def empty_table():
    downloads._pending.clear()
    yield
    downloads._pending.clear()


@pytest.fixture
def clock(monkeypatch):
    """A monotonic clock the test moves by hand."""
    now = {"t": 1000.0}
    monkeypatch.setattr(downloads.time, "monotonic", lambda: now["t"])
    return now


@pytest.fixture
def http(monkeypatch, tmp_path):
    """An HTTP deployment writing under tmp_path, with downloads enabled."""
    config = ServerConfig(
        transport="http",
        public_url="https://outlook-mcp.example.com/",
        download_path=str(tmp_path),
    )
    monkeypatch.setattr(downloads, "get_config", lambda: config)
    return config


@pytest.fixture
def stdio(monkeypatch, tmp_path):
    """A stdio deployment: one user, and no URL to hand anybody."""
    config = ServerConfig(download_path=str(tmp_path))
    monkeypatch.setattr(downloads, "get_config", lambda: config)
    return config


def write(user, message_id, name, body=b"payload"):
    """A file where a download for this user and message would have landed."""
    directory = downloads.message_dir(user, message_id)
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / name
    path.write_bytes(body)
    return path


def age(path, seconds):
    """Backdate a file, the way the passage of time would."""
    when = time.time() - seconds
    os.utime(path, (when, when))
    return path


class _Request:
    """The two things the route reads off a Starlette request.

    No headers: the route reads none. That is the point of the mode, and a
    request object that cannot offer any is the cheapest way to keep it true.
    """

    def __init__(self, token, method="GET"):
        self.path_params = {"token": token}
        self.method = method


def fetch(token, method="GET"):
    """Redeem a link the way a real server does, background task included.

    Starlette runs response.background after the body has gone out, which is
    where the file is deleted. Running it here rather than in the two tests that
    look at the disk keeps every test on the real lifecycle.
    """
    response = anyio.run(downloads.download_attachment, _Request(token, method))
    if response.background is not None:
        anyio.run(response.background)
    return response


class TestWhereFilesLand:
    def test_stdio_writes_straight_into_the_configured_directory(self, stdio, tmp_path):
        assert downloads.download_root(None) == tmp_path

    def test_each_user_gets_a_directory_of_their_own(self, http, tmp_path):
        # The isolation is the directory, exactly as it is for the token caches:
        # without it one listing shows everybody's mail attachments.
        assert downloads.download_root(ADA) != downloads.download_root(BOB)
        assert downloads.download_root(ADA).parent == tmp_path

    def test_the_address_is_not_in_the_path(self, http):
        # A listing should say how many people are served, not who they are.
        assert "ada" not in str(downloads.download_root(ADA)).lower()

    @pytest.mark.parametrize("variant", ["ADA@example.com", " ada@example.com "])
    def test_the_same_person_lands_in_the_same_directory(self, http, variant):
        # The proxy may spell an address differently from one request to the next.
        assert downloads.download_root(variant) == downloads.download_root(ADA)

    def test_each_message_gets_a_directory_under_it(self, http):
        directory = downloads.message_dir(ADA, MESSAGE)
        assert directory.parent == downloads.download_root(ADA)
        assert directory.name.startswith(downloads.MESSAGE_DIR_PREFIX)

    def test_two_messages_do_not_share_one(self, http):
        assert downloads.message_dir(ADA, MESSAGE) != downloads.message_dir(ADA, "other")

    def test_a_message_id_is_never_a_path(self, http):
        # Graph ids carry '/' and '+'; hashing is what keeps them one segment.
        directory = downloads.message_dir(ADA, "AAMk/../../etc+passwd==")
        assert directory.parent == downloads.download_root(ADA)


class TestOfferingALink:
    def test_stdio_is_offered_nothing(self, stdio):
        # The caller shares the filesystem: the path in the answer is the file.
        assert downloads.offer(Path("f.pdf"), "f.pdf", "application/pdf", None) is None

    def test_http_gets_an_absolute_url(self, http):
        url = downloads.offer(Path("f.pdf"), "f.pdf", "application/pdf", ADA)
        assert url.startswith("https://outlook-mcp.example.com/attachments/")

    def test_without_a_public_url_there_is_no_link(self, monkeypatch, tmp_path):
        config = ServerConfig(transport="http", download_path=str(tmp_path))
        monkeypatch.setattr(downloads, "get_config", lambda: config)
        assert downloads.offer(Path("f.pdf"), "f.pdf", "application/pdf", ADA) is None
        assert not downloads._pending


class TestRedeemingALink:
    def test_the_token_alone_fetches_the_file(self, http):
        # No identity header, no bearer, nothing: an agent holding the link can
        # GET it. That is the requirement, since the credential its MCP client
        # uses to reach this server is one the agent never sees.
        path = write(ADA, MESSAGE, "invoice.pdf")
        token = downloads._mint(path, "invoice.pdf", "application/pdf", ADA)

        response = fetch(token)

        assert response.status_code == 200
        assert Path(response.path) == path
        assert "invoice.pdf" in response.headers["content-disposition"]

    def test_it_works_only_once(self, http):
        path = write(ADA, MESSAGE, "invoice.pdf")
        token = downloads._mint(path, "invoice.pdf", "application/pdf", ADA)

        assert fetch(token).status_code == 200
        assert fetch(token).status_code == 404

    def test_serving_it_deletes_the_file(self, http):
        # What makes an unauthenticated link acceptable: by the time the answer
        # holding it has been written, there is nothing left on the other end.
        path = write(ADA, MESSAGE, "invoice.pdf")
        token = downloads._mint(path, "invoice.pdf", "application/pdf", ADA)

        fetch(token)

        assert not path.exists()

    def test_the_message_directory_goes_with_the_last_file(self, http):
        path = write(ADA, MESSAGE, "invoice.pdf")
        token = downloads._mint(path, "invoice.pdf", "application/pdf", ADA)

        fetch(token)

        assert not path.parent.exists()

    def test_a_second_link_to_the_served_file_is_revoked(self, http):
        # Two calls for one attachment mint two tickets over the same path. The
        # second must not survive to promise a file that is gone.
        path = write(ADA, MESSAGE, "invoice.pdf")
        first = downloads._mint(path, "invoice.pdf", "application/pdf", ADA)
        second = downloads._mint(path, "invoice.pdf", "application/pdf", ADA)

        fetch(first)

        assert fetch(second).status_code == 404

    def test_an_expired_link_is_refused(self, http, clock):
        path = write(ADA, MESSAGE, "invoice.pdf")
        token = downloads._mint(path, "invoice.pdf", "application/pdf", ADA)
        clock["t"] += downloads.TICKET_TTL_SECONDS + 1

        assert fetch(token).status_code == 404

    def test_an_unknown_token_is_refused(self, http):
        assert fetch("not-a-token").status_code == 404

    def test_a_deleted_file_says_so(self, http):
        path = write(ADA, MESSAGE, "invoice.pdf")
        token = downloads._mint(path, "invoice.pdf", "application/pdf", ADA)
        path.unlink()

        assert fetch(token).status_code == 410

    def test_head_does_not_burn_the_link(self, http):
        # Starlette answers HEAD on a GET route by itself, and a preflight probe
        # must not spend the one fetch the link is good for.
        path = write(ADA, MESSAGE, "invoice.pdf")
        token = downloads._mint(path, "invoice.pdf", "application/pdf", ADA)

        assert fetch(token, method="HEAD").status_code == 405
        assert fetch(token).status_code == 200

    def test_stdio_serves_nothing_at_all(self, stdio):
        # Nobody is proxying, and there is no public URL to have handed out a
        # link in the first place: the route must not exist.
        token = downloads._mint(Path("f.pdf"), "f.pdf", "application/pdf", ADA)
        assert fetch(token).status_code == 404


class TestRetention:
    """What happens to a file nobody ever fetched."""

    def test_an_old_file_is_swept(self, http):
        old = age(write(ADA, MESSAGE, "invoice.pdf"), http.retention_seconds + 60)

        assert downloads.sweep() == [old]
        assert not old.exists()

    def test_a_fresh_file_is_left_alone(self, http):
        fresh = write(ADA, MESSAGE, "invoice.pdf")

        assert downloads.sweep() == []
        assert fresh.exists()

    def test_the_directories_go_with_the_files(self, http):
        old = age(write(ADA, MESSAGE, "invoice.pdf"), http.retention_seconds + 60)
        message = old.parent
        user = message.parent

        downloads.sweep()

        assert not message.exists()
        assert not user.exists()
        # The configured root is not this module's to remove, and the next
        # download needs it.
        assert http.attachment_dir.is_dir()

    def test_a_link_to_a_swept_file_is_revoked(self, http):
        old = age(write(ADA, MESSAGE, "invoice.pdf"), http.retention_seconds + 60)
        token = downloads._mint(old, "invoice.pdf", "application/pdf", ADA)

        downloads.sweep()

        assert fetch(token).status_code == 404

    def test_a_directory_of_a_still_wanted_file_survives(self, http):
        old = age(write(ADA, MESSAGE, "old.pdf"), http.retention_seconds + 60)
        fresh = write(ADA, MESSAGE, "new.pdf")

        downloads.sweep()

        assert not old.exists()
        assert fresh.exists()

    def test_retention_can_be_turned_off(self, monkeypatch, tmp_path):
        config = ServerConfig(
            transport="http",
            public_url="https://outlook-mcp.example.com/",
            download_path=str(tmp_path),
            retention_minutes=0,
        )
        monkeypatch.setattr(downloads, "get_config", lambda: config)
        old = age(write(ADA, MESSAGE, "invoice.pdf"), 10 * 24 * 3600)

        assert downloads.sweep() == []
        assert old.exists()

    def test_stdio_downloads_are_never_swept(self, stdio):
        # The file the tool wrote is the answer, in the caller's own download
        # directory. Deleting it would be deleting the user's file.
        old = age(write(None, MESSAGE, "invoice.pdf"), 10 * 24 * 3600)

        assert downloads.sweep() == []
        assert old.exists()

    def test_sweeping_an_empty_deployment_is_not_an_error(self, http):
        assert downloads.sweep() == []


class TestTheTableIsBounded:
    def test_expired_tickets_are_pruned_on_the_next_mint(self, http, clock):
        downloads._mint(Path("a"), "a", "text/plain", ADA)
        clock["t"] += downloads.TICKET_TTL_SECONDS + 1
        downloads._mint(Path("b"), "b", "text/plain", ADA)
        assert len(downloads._pending) == 1

    def test_it_never_grows_past_the_limit(self, http):
        for _ in range(downloads.MAX_PENDING_TICKETS + 10):
            downloads._mint(Path("a"), "a", "text/plain", ADA)
        assert len(downloads._pending) <= downloads.MAX_PENDING_TICKETS

    def test_the_oldest_is_dropped_first(self, http, clock):
        first = downloads._mint(Path("a"), "a", "text/plain", ADA)
        for _ in range(downloads.MAX_PENDING_TICKETS):
            clock["t"] += 1
            newest = downloads._mint(Path("b"), "b", "text/plain", ADA)
        assert first not in downloads._pending
        assert newest in downloads._pending


class TestDeletingDownloads:
    def test_every_file_of_one_message_goes(self, http):
        write(ADA, MESSAGE, "one.pdf")
        write(ADA, MESSAGE, "two.pdf")

        removed = downloads.delete_message_downloads(ADA, MESSAGE)

        assert removed == ["one.pdf", "two.pdf"]
        assert not downloads.message_dir(ADA, MESSAGE).exists()

    def test_one_file_can_be_named(self, http):
        write(ADA, MESSAGE, "one.pdf")
        kept = write(ADA, MESSAGE, "two.pdf")

        removed = downloads.delete_message_downloads(ADA, MESSAGE, "one.pdf")

        assert removed == ["one.pdf"]
        assert kept.is_file()

    def test_another_message_is_untouched(self, http):
        write(ADA, MESSAGE, "one.pdf")
        other = write(ADA, "other-message", "one.pdf")

        downloads.delete_message_downloads(ADA, MESSAGE)

        assert other.is_file()

    def test_another_user_is_untouched(self, http):
        theirs = write(BOB, MESSAGE, "one.pdf")

        # Same message id, a different caller: the directory is not the same one.
        assert downloads.delete_message_downloads(ADA, MESSAGE) == []
        assert theirs.is_file()

    def test_a_filename_cannot_climb_out_of_the_directory(self, http, tmp_path):
        # `filename` is echoed back from a tool result, so it is caller input.
        outside = tmp_path / "secret.json"
        outside.write_bytes(b"x")
        write(ADA, MESSAGE, "one.pdf")

        removed = downloads.delete_message_downloads(ADA, MESSAGE, "../../secret.json")

        assert removed == []
        assert outside.is_file()

    def test_deleting_nothing_is_not_an_error(self, http):
        assert downloads.delete_message_downloads(ADA, "never-downloaded") == []

    def test_outstanding_links_to_the_files_are_revoked(self, http):
        path = write(ADA, MESSAGE, "one.pdf")
        token = downloads._mint(path, "one.pdf", "application/pdf", ADA)

        downloads.delete_message_downloads(ADA, MESSAGE)

        assert token not in downloads._pending

    def test_links_are_revoked_even_with_a_relative_download_path(
        self, monkeypatch, tmp_path
    ):
        # The tool mints the absolute path it reported, while deletion walks the
        # configured directory: the two only look alike once resolved.
        monkeypatch.chdir(tmp_path)
        config = ServerConfig(
            transport="http", public_url="https://mcp.example.com", download_path="files"
        )
        monkeypatch.setattr(downloads, "get_config", lambda: config)
        path = write(ADA, MESSAGE, "one.pdf")
        token = downloads._mint(path.absolute(), "one.pdf", "application/pdf", ADA)

        downloads.delete_message_downloads(ADA, MESSAGE)

        assert token not in downloads._pending

    def test_a_link_to_another_message_survives(self, http):
        write(ADA, MESSAGE, "one.pdf")
        other = write(ADA, "other-message", "two.pdf")
        token = downloads._mint(other, "two.pdf", "application/pdf", ADA)

        downloads.delete_message_downloads(ADA, MESSAGE)

        assert token in downloads._pending

    def test_stdio_deletes_from_the_one_download_directory(self, stdio):
        path = write(None, MESSAGE, "one.pdf")

        assert downloads.delete_message_downloads(None, MESSAGE) == ["one.pdf"]
        assert not path.exists()
