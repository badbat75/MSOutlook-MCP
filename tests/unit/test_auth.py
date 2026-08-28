"""Unit tests for AuthManager.get_token, with a stubbed MSAL application.

The property under test is a security one: MSAL keys its cached tokens by
client id and never by client secret, so an AuthManager must prove its secret
to AAD before it is allowed to serve anything the cache already holds.
"""

import asyncio

import pytest

from outlook_mcp.auth import AuthManager, CredentialsError


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


def make_manager(app, cache=None, probe=None):
    """An AuthManager wired to a stub app, with disk writes disabled."""
    manager = AuthManager("client-id", "secret", "common", token_cache=cache)
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
