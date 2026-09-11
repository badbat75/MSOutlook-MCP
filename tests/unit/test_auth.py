"""Unit tests for AuthManager.get_token, with a stubbed MSAL application.

The property under test is a security one: MSAL keys its cached tokens by
client id and never by client secret, so an AuthManager must prove its secret
to AAD before it is allowed to serve anything the cache already holds.
"""

import asyncio
import logging
import os
import stat
from pathlib import Path

import httpx
import msal
import pytest

from outlook_mcp.auth import (
    CONTACTS_SCOPE_URLS,
    CORE_SCOPE_URLS,
    GRAPH_BASE_URL,
    GRAPH_SCOPE_URLS,
    TOKEN_CACHE_PATH,
    USER_CACHE_DIR,
    AuthManager,
    CredentialsError,
    GraphClient,
    load_token_cache,
    save_token_cache,
    scopes_for,
    shared_cache_path,
    user_cache_path,
)


class StubApp:
    """Stands in for msal.ConfidentialClientApplication, recording every call."""

    def __init__(self, accounts=None, silent_results=None, client_results=None):
        self._accounts = accounts if accounts is not None else [{"home_account_id": "u"}]
        self._silent_results = list(silent_results or [])
        self._client_results = list(client_results or [])
        self.silent_calls = []
        self.silent_scopes = []
        self.client_calls = []

    def get_accounts(self):
        return self._accounts

    def acquire_token_silent_with_error(self, scopes, account=None, force_refresh=False, **kwargs):
        self.silent_calls.append({"force_refresh": force_refresh, "account": account})
        self.silent_scopes.append(list(scopes))
        return self._silent_results.pop(0) if self._silent_results else None

    def acquire_token_for_client(self, scopes=None, **kwargs):
        self.client_calls.append({"scopes": scopes})
        return self._client_results.pop(0) if self._client_results else None


# A cache path nothing ever writes: the manager starts empty and, with the file
# absent throughout, never believes a sign-in has rewritten it.
NO_CACHE = Path(__file__).parent / "no-such-cache.json"


def make_manager(app, probe=None, user=None):
    """An AuthManager wired to a stub app, with disk writes disabled."""
    manager = AuthManager("client-id", "secret", "common", cache_path=NO_CACHE, user=user)
    manager._app = app
    manager._save_cache = lambda: None
    if probe is not None:
        manager._unverified_client_app = lambda: probe
    return manager


TOKEN = {"access_token": "at-1"}
TOKEN2 = {"access_token": "at-2"}


class TestDelegatedPath:
    def test_first_call_forces_a_refresh(self):
        # The forced redemption is the request AAD authenticates the secret on.
        app = StubApp(silent_results=[TOKEN])
        manager = make_manager(app)

        assert asyncio.run(manager.get_token()) == "at-1"
        assert app.silent_calls == [{"force_refresh": True, "account": {"home_account_id": "u"}}]

    def test_later_calls_may_use_the_cache(self):
        app = StubApp(silent_results=[TOKEN, TOKEN2])
        manager = make_manager(app)

        asyncio.run(manager.get_token())
        assert asyncio.run(manager.get_token()) == "at-2"
        assert [c["force_refresh"] for c in app.silent_calls] == [True, False]

    def test_failed_verification_raises_instead_of_serving_a_cached_token(self):
        # This is the bypass being guarded: without the raise, the next call
        # would happily return whatever the shared cache holds for this client id.
        app = StubApp(silent_results=[None])
        manager = make_manager(app)

        with pytest.raises(CredentialsError, match="client secret is wrong"):
            asyncio.run(manager.get_token())

    def test_failed_verification_does_not_fall_through_to_client_credentials(self):
        app = StubApp(silent_results=[None], client_results=[TOKEN])
        manager = make_manager(app)

        with pytest.raises(CredentialsError):
            asyncio.run(manager.get_token())
        assert app.client_calls == []

    def test_a_later_silent_failure_is_not_fatal(self):
        # Once the secret is proven, an expired refresh token falls through to
        # the app-only path exactly as it did before.
        app = StubApp(silent_results=[TOKEN, None], client_results=[TOKEN2])
        manager = make_manager(app)

        asyncio.run(manager.get_token())
        assert asyncio.run(manager.get_token()) == "at-2"
        assert len(app.client_calls) == 1


class TestAppOnlyPath:
    def test_first_call_goes_through_an_empty_cache_app(self):
        # acquire_token_for_client refuses force_refresh, so the only way to
        # guarantee the request reaches AAD is an app with nothing cached.
        shared = StubApp(accounts=[], client_results=[TOKEN2])
        probe = StubApp(accounts=[], client_results=[TOKEN])
        manager = make_manager(shared, probe=probe)

        assert asyncio.run(manager.get_token()) == "at-1"
        assert len(probe.client_calls) == 1
        assert shared.client_calls == []

    def test_later_calls_use_the_shared_app(self):
        shared = StubApp(accounts=[], client_results=[TOKEN2])
        probe = StubApp(accounts=[], client_results=[TOKEN])
        manager = make_manager(shared, probe=probe)

        asyncio.run(manager.get_token())
        assert asyncio.run(manager.get_token()) == "at-2"
        assert len(probe.client_calls) == 1
        assert len(shared.client_calls) == 1

    def test_no_token_anywhere_is_a_runtime_error(self):
        shared = StubApp(accounts=[])
        probe = StubApp(accounts=[])
        manager = make_manager(shared, probe=probe)

        with pytest.raises(RuntimeError, match="outlook_mcp_auth.py"):
            asyncio.run(manager.get_token())


class TestPerUserManager:
    """A manager that belongs to one enrolled person, as in proxy identity mode.

    The dangerous fallback here is app-only: with one app registration shared by
    everyone, letting an unenrolled caller through to client credentials would
    hand them a token that acts as the application rather than as themselves.
    """

    def test_an_unenrolled_user_never_reaches_client_credentials(self):
        app = StubApp(accounts=[], client_results=[TOKEN])
        probe = StubApp(accounts=[], client_results=[TOKEN])
        manager = make_manager(app, probe=probe, user="ada@example.com")

        with pytest.raises(CredentialsError, match="has not authorized"):
            asyncio.run(manager.get_token())
        assert app.client_calls == []
        assert probe.client_calls == []

    def test_the_error_names_the_user_and_the_way_out(self):
        manager = make_manager(StubApp(accounts=[]), user="ada@example.com")
        with pytest.raises(CredentialsError) as excinfo:
            asyncio.run(manager.get_token())
        message = str(excinfo.value)
        assert "ada@example.com" in message
        assert "--user ada@example.com" in message

    def test_an_enrolled_user_is_served_normally(self):
        app = StubApp(silent_results=[TOKEN])
        manager = make_manager(app, user="ada@example.com")
        assert asyncio.run(manager.get_token()) == "at-1"

    def test_a_stale_grant_points_at_enrolling_again(self):
        app = StubApp(silent_results=[None])
        manager = make_manager(app, user="ada@example.com")
        with pytest.raises(CredentialsError, match="/oauth/login"):
            asyncio.run(manager.get_token())

    def test_without_a_user_the_app_only_path_is_unchanged(self):
        shared = StubApp(accounts=[], client_results=[TOKEN2])
        probe = StubApp(accounts=[], client_results=[TOKEN])
        manager = make_manager(shared, probe=probe)
        assert asyncio.run(manager.get_token()) == "at-1"


class TestUserCachePath:
    def test_two_users_get_two_files(self, tmp_path):
        assert user_cache_path("ada@example.com", tmp_path) != user_cache_path(
            "bob@example.com", tmp_path
        )

    def test_the_same_user_is_stable(self, tmp_path):
        assert user_cache_path("ada@example.com", tmp_path) == user_cache_path(
            "ada@example.com", tmp_path
        )

    @pytest.mark.parametrize("variant", ["ADA@example.com", " ada@example.com ", "Ada@Example.com"])
    def test_case_and_padding_do_not_make_a_second_account(self, tmp_path, variant):
        # A proxy that changes the casing between requests must not strand a
        # user with a second, empty cache and a "not authorized" error.
        assert user_cache_path(variant, tmp_path) == user_cache_path("ada@example.com", tmp_path)

    def test_the_filename_does_not_disclose_the_address(self, tmp_path):
        path = user_cache_path("ada@example.com", tmp_path)
        assert "ada" not in path.name
        assert path.suffix == ".json"

    def test_no_directory_means_the_home_default(self):
        # None is what config.cache_directory hands back when [auth].cache_dir
        # is unset, so it has to mean "the default" rather than crash or land in
        # the working directory.
        assert user_cache_path("ada@example.com").parent == USER_CACHE_DIR

    def test_a_directory_puts_the_cache_there_and_nowhere_else(self, tmp_path):
        # The deployment layout: a directory that belongs to this server, with
        # no ".outlook_mcp" level inside it, which would only repeat what the
        # directory already says.
        path = user_cache_path("ada@example.com", tmp_path)
        assert path.parent == tmp_path


class TestSharedCachePath:
    def test_no_directory_means_the_home_default(self):
        assert shared_cache_path() == TOKEN_CACHE_PATH

    def test_a_directory_holds_it_under_a_plain_name(self, tmp_path):
        # No digest: a stdio server has no address to hash, because nothing in
        # the request says who the caller is.
        assert shared_cache_path(tmp_path) == tmp_path / "shared.json"

    def test_it_never_collides_with_a_user_cache(self, tmp_path):
        users = {user_cache_path(u, tmp_path) for u in ("ada@example.com", "shared")}
        assert shared_cache_path(tmp_path) not in users


class TestSaveTokenCache:
    def test_creates_the_directory(self, tmp_path):
        path = tmp_path / "caches" / "u.json"
        cache = msal.SerializableTokenCache()
        save_token_cache(cache, path)
        assert path.is_file()

    def test_round_trips_through_load(self, tmp_path):
        path = tmp_path / "u.json"
        cache = msal.SerializableTokenCache()
        cache.deserialize('{"AccessToken": {}}')
        save_token_cache(cache, path)
        assert load_token_cache(path).serialize() == cache.serialize()

    @pytest.mark.skipif(os.name == "nt", reason="POSIX permission bits")
    def test_the_file_is_private(self, tmp_path):
        # It holds refresh tokens, and it is created 0600 before anything is
        # written into it rather than chmod-ed afterwards.
        path = tmp_path / "u.json"
        save_token_cache(msal.SerializableTokenCache(), path)
        assert stat.S_IMODE(path.stat().st_mode) == 0o600

    def test_an_existing_file_is_overwritten_not_appended(self, tmp_path):
        path = tmp_path / "u.json"
        path.write_text('{"AccessToken": {"stale": {}}}')
        save_token_cache(msal.SerializableTokenCache(), path)
        assert "stale" not in path.read_text()


class TestCacheWriteBack:
    """Where a manager persists its cache, which is the whole isolation."""

    def test_writes_to_the_path_it_was_given(self, tmp_path):
        path = tmp_path / "ada.json"
        manager = AuthManager("client-id", "secret", "common", cache_path=path)
        manager._cache.deserialize('{"AccessToken": {}}')
        manager._cache.has_state_changed = True
        manager._save_cache()
        assert path.is_file()

    def test_leaves_the_shared_cache_alone(self, tmp_path, monkeypatch):
        # The bug this guards: a per-user manager persisting over the shared
        # file would give every user the same account on the next start.
        shared = tmp_path / "shared.json"
        monkeypatch.setattr("outlook_mcp.auth.TOKEN_CACHE_PATH", shared)
        manager = AuthManager("client-id", "secret", "common", cache_path=tmp_path / "ada.json")
        manager._cache.has_state_changed = True
        manager._save_cache()
        assert not shared.exists()


class TestScopesPerResource:
    """A request asks for the scopes of the resource it addresses, never all of them.

    Measured on a personal account: a refresh token asked for one scope its
    user has not consented to fails the whole request (AADSTS70000). Asking for
    everything on every request would have turned mail and calendar off for
    each grant made before contacts existed.
    """

    @pytest.mark.parametrize("endpoint", [
        "/me/contacts",
        "/me/contacts/AAMk==",
        "/me/contactFolders/contacts",
        "/me/contactfolders/F1/childFolders",
        "/me/contactFolders/F1/contacts",
        "https://graph.microsoft.com/v1.0/me/contactFolders/F1/contacts?$skip=500",
    ])
    def test_contacts_ask_for_the_contacts_scope(self, endpoint):
        assert scopes_for(endpoint) == CONTACTS_SCOPE_URLS

    @pytest.mark.parametrize("endpoint", [
        "/me",
        "/me/messages",
        "/me/mailFolders/inbox/messages",
        "/me/events/AAMk==",
        "/me/calendars",
        "https://graph.microsoft.com/v1.0/me/messages?$skip=10",
    ])
    def test_everything_else_asks_for_the_core_scopes(self, endpoint):
        assert scopes_for(endpoint) == CORE_SCOPE_URLS

    def test_a_sign_in_asks_for_everything(self):
        # The one place every scope is requested: the consent screen.
        assert set(GRAPH_SCOPE_URLS) == set(CORE_SCOPE_URLS) | set(CONTACTS_SCOPE_URLS)
        assert "https://graph.microsoft.com/Contacts.ReadWrite" in GRAPH_SCOPE_URLS

    def test_get_token_asks_msal_for_the_scopes_it_is_given(self):
        app = StubApp(silent_results=[TOKEN, TOKEN2])
        manager = make_manager(app)
        asyncio.run(manager.get_token())
        asyncio.run(manager.get_token(CONTACTS_SCOPE_URLS))
        assert app.silent_scopes == [CORE_SCOPE_URLS, CONTACTS_SCOPE_URLS]

    def test_the_client_asks_for_the_scopes_of_its_endpoint(self):
        # Through GraphClient.request, the absolute URL of a nextLink included.
        asked = []

        class RecordingAuth:
            async def get_token(self, scopes=None):
                asked.append(scopes)
                return "at"

        client = GraphClient(RecordingAuth())
        client._client = httpx.AsyncClient(
            base_url=GRAPH_BASE_URL,
            transport=httpx.MockTransport(lambda request: httpx.Response(200, json={"url": str(request.url)})),
        )

        async def run():
            await client.get("/me/messages")
            answer = await client.get(f"{GRAPH_BASE_URL}/me/contactFolders/F1/contacts?$skip=500")
            await client.close()
            return answer

        answer = asyncio.run(run())
        assert asked == [CORE_SCOPE_URLS, CONTACTS_SCOPE_URLS]
        assert answer["url"] == f"{GRAPH_BASE_URL}/me/contactFolders/F1/contacts?$skip=500"


class TestBytes:
    """What Graph serves as bytes rather than JSON, such as a contact's photo."""

    PHOTO = "/me/contacts/C1/photo/$value"

    @staticmethod
    def call(handler, use):
        class Auth:
            async def get_token(self, scopes=None):
                return "at"

        client = GraphClient(Auth())
        client._client = httpx.AsyncClient(base_url=GRAPH_BASE_URL, transport=httpx.MockTransport(handler))

        async def run():
            try:
                return await use(client)
            finally:
                await client.close()

        return asyncio.run(run())

    def test_the_bytes_come_with_their_media_type(self):
        answer = self.call(
            lambda request: httpx.Response(200, content=b"\xff\xd8jpeg", headers={"Content-Type": "image/jpeg"}),
            lambda client: client.get_bytes(self.PHOTO),
        )
        assert answer == (b"\xff\xd8jpeg", "image/jpeg")

    def test_nothing_there_is_none_and_no_error_in_the_log(self, caplog):
        # A contact without a photo answers 404: expected, and a move asks for
        # every contact's, so logging it as an error would bury the real ones.
        with caplog.at_level(logging.ERROR, logger="outlook_mcp"):
            answer = self.call(
                lambda request: httpx.Response(404, json={"error": {"code": "ErrorItemNotFound"}}),
                lambda client: client.get_bytes(self.PHOTO),
            )
        assert answer is None
        assert not caplog.records

    def test_any_other_failure_is_raised_and_logged(self, caplog):
        with caplog.at_level(logging.ERROR, logger="outlook_mcp"):
            with pytest.raises(httpx.HTTPStatusError):
                self.call(
                    lambda request: httpx.Response(500, json={"error": {"code": "boom"}}),
                    lambda client: client.get_bytes(self.PHOTO),
                )
        assert caplog.records

    def test_a_json_request_still_logs_a_404(self, caplog):
        with caplog.at_level(logging.ERROR, logger="outlook_mcp"):
            with pytest.raises(httpx.HTTPStatusError):
                self.call(
                    lambda request: httpx.Response(404, json={"error": {"code": "ErrorItemNotFound"}}),
                    lambda client: client.get("/me/contacts/C1"),
                )
        assert caplog.records


NOT_CONSENTED = {
    "error": "invalid_grant",
    "error_codes": [70000],
    "error_description": "AADSTS70000: The request was denied because one or more scopes requested are unauthorized or expired.",
}


class TestMissingConsent:
    def test_it_names_the_scope_and_the_way_to_grant_it(self):
        app = StubApp(silent_results=[NOT_CONSENTED])
        manager = make_manager(app)
        with pytest.raises(CredentialsError) as excinfo:
            asyncio.run(manager.get_token(CONTACTS_SCOPE_URLS))
        message = str(excinfo.value)
        assert "Contacts.ReadWrite" in message
        assert "AADSTS70000" in message
        assert "outlook_mcp_auth.py" in message
        # Not the message for a wrong secret or an expired grant, which is what
        # acquire_token_silent's None used to turn this into.
        assert "client secret is wrong" not in message

    def test_a_per_user_grant_points_at_enrolling_again(self):
        app = StubApp(silent_results=[NOT_CONSENTED])
        manager = make_manager(app, user="ada@example.com")
        with pytest.raises(CredentialsError, match="/oauth/login") as excinfo:
            asyncio.run(manager.get_token(CONTACTS_SCOPE_URLS))
        assert "ada@example.com" in str(excinfo.value)

    def test_the_work_account_form_is_recognised_too(self):
        consent_required = {"error": "invalid_grant", "error_codes": [65001], "suberror": "consent_required"}
        manager = make_manager(StubApp(silent_results=[consent_required]))
        with pytest.raises(CredentialsError, match="Contacts.ReadWrite"):
            asyncio.run(manager.get_token(CONTACTS_SCOPE_URLS))

    def test_it_never_falls_through_to_client_credentials(self):
        # Even on a stdio manager with a verified secret, where an expired grant
        # would fall through: acting as the application is not a substitute for
        # a scope the user has not granted.
        app = StubApp(silent_results=[TOKEN, NOT_CONSENTED], client_results=[TOKEN2])
        manager = make_manager(app)
        asyncio.run(manager.get_token())
        with pytest.raises(CredentialsError):
            asyncio.run(manager.get_token(CONTACTS_SCOPE_URLS))
        assert app.client_calls == []

    def test_the_other_tools_keep_working(self):
        app = StubApp(silent_results=[NOT_CONSENTED, TOKEN])
        manager = make_manager(app)
        with pytest.raises(CredentialsError):
            asyncio.run(manager.get_token(CONTACTS_SCOPE_URLS))
        assert asyncio.run(manager.get_token()) == "at-1"

    def test_it_does_not_count_as_a_verified_secret(self):
        app = StubApp(silent_results=[NOT_CONSENTED, TOKEN])
        manager = make_manager(app)
        with pytest.raises(CredentialsError):
            asyncio.run(manager.get_token(CONTACTS_SCOPE_URLS))
        asyncio.run(manager.get_token())
        assert app.silent_calls[1]["force_refresh"] is True


def write_cache(path, marker):
    """A token cache file recognisable by a marker, of a size unique to it."""
    cache = msal.SerializableTokenCache()
    cache.deserialize('{"AccessToken": {"%s": {}}}' % marker)
    save_token_cache(cache, path)


class TestASignInWhileRunning:
    """A sign-in rewrites the cache file under a running server.

    To grant a new scope, to replace a grant that stopped working, or to put
    another account behind the file. A manager that kept its first copy would
    never see it, and its next write-back would erase it from disk.
    """

    def manager(self, path, app=None):
        manager = AuthManager("client-id", "secret", "common", cache_path=path)
        manager._app = app or StubApp(silent_results=[TOKEN, TOKEN2])
        return manager

    def test_the_rewritten_file_is_adopted_on_the_next_call(self, tmp_path):
        path = tmp_path / "cache.json"
        write_cache(path, "old-grant")
        manager = self.manager(path)
        app = manager._app

        write_cache(path, "new-grant-with-contacts")
        asyncio.run(manager.get_token())

        assert "new-grant-with-contacts" in manager._cache.serialize()
        # In place: the MSAL application keeps the cache object it was built on.
        assert manager._app is app

    def test_the_old_grant_is_never_written_over_the_new_one(self, tmp_path):
        path = tmp_path / "cache.json"
        write_cache(path, "old-grant")
        manager = self.manager(path)

        write_cache(path, "new-grant-with-contacts")
        # What a refresh of the old grant leaves behind, just after the sign-in.
        manager._cache.has_state_changed = True
        manager._save_cache()

        assert "new-grant-with-contacts" in path.read_text()

    def test_its_own_write_back_is_not_mistaken_for_a_sign_in(self, tmp_path):
        path = tmp_path / "cache.json"
        write_cache(path, "grant")
        manager = self.manager(path)
        manager._cache.deserialize('{"AccessToken": {"refreshed-here-and-longer": {}}}')
        manager._cache.has_state_changed = True
        manager._save_cache()
        assert "refreshed-here-and-longer" in path.read_text()

        # Held in memory only: a reload, which only a rewrite by someone else
        # may trigger, would lose it.
        manager._cache.deserialize('{"AccessToken": {"in-memory-only": {}}}')
        manager._adopt_rewritten_cache()
        assert "in-memory-only" in manager._cache.serialize()

    def test_a_first_enrollment_after_startup_is_picked_up(self, tmp_path):
        # A caller who used a tool before enrolling used to stay "not
        # authorized" until the server restarted.
        path = tmp_path / "ada.json"
        manager = self.manager(path)
        assert manager._cache_stamp is None

        write_cache(path, "enrolled")
        manager._adopt_rewritten_cache()
        assert "enrolled" in manager._cache.serialize()

    def test_a_deleted_file_empties_the_cache(self, tmp_path):
        # How an operator withdraws a grant: delete the file.
        path = tmp_path / "cache.json"
        write_cache(path, "withdrawn")
        manager = self.manager(path)
        path.unlink()

        manager._adopt_rewritten_cache()
        assert "withdrawn" not in manager._cache.serialize()


class TestVerificationIsPerCredentialSet:
    def test_a_second_manager_starts_unverified(self):
        # The pool keys clients on (id, secret, tenant), so a caller sending a
        # different secret gets a different manager, which must verify anew.
        app_a = StubApp(silent_results=[TOKEN])
        manager_a = make_manager(app_a)
        asyncio.run(manager_a.get_token())

        app_b = StubApp(silent_results=[TOKEN2])
        manager_b = make_manager(app_b)
        asyncio.run(manager_b.get_token())

        assert app_b.silent_calls[0]["force_refresh"] is True
