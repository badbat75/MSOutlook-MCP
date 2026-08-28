"""Unit tests for AuthManager.get_token, with a stubbed MSAL application.

The property under test is a security one: MSAL keys its cached tokens by
client id and never by client secret, so an AuthManager must prove its secret
to AAD before it is allowed to serve anything the cache already holds.
"""

import asyncio
import os
import stat

import msal
import pytest

from outlook_mcp.auth import (
    TOKEN_CACHE_PATH,
    USER_CACHE_DIR,
    AuthManager,
    CredentialsError,
    load_token_cache,
    save_token_cache,
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
        self.client_calls = []

    def get_accounts(self):
        return self._accounts

    def acquire_token_silent(self, scopes, account=None, force_refresh=False, **kwargs):
        self.silent_calls.append({"force_refresh": force_refresh, "account": account})
        return self._silent_results.pop(0) if self._silent_results else None

    def acquire_token_for_client(self, scopes=None, **kwargs):
        self.client_calls.append({"scopes": scopes})
        return self._client_results.pop(0) if self._client_results else None


def make_manager(app, cache=None, probe=None, user=None):
    """An AuthManager wired to a stub app, with disk writes disabled."""
    manager = AuthManager("client-id", "secret", "common", token_cache=cache, user=user)
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
