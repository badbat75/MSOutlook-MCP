"""Unit tests for outlook_mcp.credentials: whose mailbox a request acts on.

Pure logic, no network. The server runs as one app registration and never takes
one from a caller, so what these tests pin down is the other half: that a
request is attributed to exactly one enrolled person, and that two people never
end up sharing a client, a cache or a token.
"""

import types

import msal
import pytest

from outlook_mcp.auth import CredentialsError
from outlook_mcp.config import ServerConfig
from outlook_mcp import credentials as credentials_module
from outlook_mcp.credentials import (
    Credentials,
    GraphClientPool,
    Principal,
    ProxyAuthPolicy,
    credentials_from_config,
    get_graph,
)

USER_HEADER = "X-Auth-Email"
SERVER_CREDS = Credentials("server-id", "server-secret", "common")


class TestCredentialsFromConfig:
    def test_reads_a_complete_set(self):
        config = ServerConfig(client_id="id", client_secret="secret", tenant_id="tenant")
        assert credentials_from_config(config) == Credentials("id", "secret", "tenant")

    def test_defaults_the_tenant(self):
        config = ServerConfig(client_id="id", client_secret="secret")
        assert credentials_from_config(config).tenant_id == "common"

    @pytest.mark.parametrize("config", [
        ServerConfig(),
        ServerConfig(client_id="id"),
        ServerConfig(client_secret="secret"),
    ])
    def test_returns_none_when_unconfigured(self, config):
        assert credentials_from_config(config) is None


class TestCredentials:
    def test_is_hashable_so_the_pool_can_key_on_it(self):
        a = Credentials("id", "secret", "tenant")
        b = Credentials("id", "secret", "tenant")
        assert {a: 1}[b] == 1

    def test_differing_secrets_are_different_keys(self):
        a = Credentials("id", "secret", "tenant")
        b = Credentials("id", "other", "tenant")
        assert a != b
        assert len({a, b}) == 2


class TestProxyAuthPolicy:
    """What the server reads off a request the reverse proxy has authenticated."""

    def test_reads_the_user(self):
        policy = ProxyAuthPolicy(USER_HEADER)
        assert policy.user_from_headers({USER_HEADER: "ada@example.com"}) == "ada@example.com"

    @pytest.mark.parametrize("transform", [str.lower, str.upper, lambda s: s])
    def test_header_name_is_case_insensitive(self, transform):
        policy = ProxyAuthPolicy(USER_HEADER)
        headers = {transform(USER_HEADER): "ada@example.com"}
        assert policy.user_from_headers(headers) == "ada@example.com"

    def test_trims_whitespace(self):
        policy = ProxyAuthPolicy(USER_HEADER)
        assert policy.user_from_headers({USER_HEADER: "  ada@example.com \n"}) == "ada@example.com"

    def test_a_custom_header_name_is_honoured(self):
        policy = ProxyAuthPolicy("X-Forwarded-User")
        assert policy.user_from_headers({"X-Forwarded-User": "ada@example.com"}) == "ada@example.com"
        with pytest.raises(CredentialsError):
            policy.user_from_headers({USER_HEADER: "ada@example.com"})

    def test_a_missing_user_header_names_it(self):
        policy = ProxyAuthPolicy(USER_HEADER)
        with pytest.raises(CredentialsError, match=USER_HEADER):
            policy.user_from_headers({})

    def test_a_blank_user_header_counts_as_missing(self):
        # A proxy that forgot the proxy_set_header line sends the request
        # anyway; serving it as somebody would be the worst possible guess.
        policy = ProxyAuthPolicy(USER_HEADER)
        with pytest.raises(CredentialsError):
            policy.user_from_headers({USER_HEADER: "   "})


class TestPrincipal:
    def test_two_users_are_different_keys(self):
        a = Principal(SERVER_CREDS, "ada@example.com")
        b = Principal(SERVER_CREDS, "bob@example.com")
        assert a != b
        assert len({a, b}) == 2

    def test_the_same_user_is_the_same_key(self):
        a = Principal(SERVER_CREDS, "ada@example.com")
        b = Principal(SERVER_CREDS, "ada@example.com")
        assert {a: 1}[b] == 1

    def test_a_user_is_distinct_from_no_user(self):
        assert Principal(SERVER_CREDS) != Principal(SERVER_CREDS, "ada@example.com")


@pytest.fixture
def isolated_caches(tmp_path):
    """A cache directory of the test's own, instead of the real one in $HOME.

    Nothing is stubbed: [auth].cache_dir is what a deployment sets for exactly
    this reason, so the tests exercise the path logic they depend on rather than
    a lambda standing in for it.
    """
    return tmp_path


class TestGraphClientPool:
    def test_one_client_per_principal(self, isolated_caches):
        pool = GraphClientPool(isolated_caches)
        ada = pool.get(Principal(SERVER_CREDS, "ada@example.com"))
        assert pool.get(Principal(SERVER_CREDS, "ada@example.com")) is ada
        assert pool.get(Principal(SERVER_CREDS, "bob@example.com")) is not ada

    def test_each_user_writes_back_to_its_own_cache(self, isolated_caches):
        # The property that makes the mode safe: one manager persisting over
        # another's file would hand both users the same account.
        pool = GraphClientPool(isolated_caches)
        ada = pool.get(Principal(SERVER_CREDS, "ada@example.com"))
        bob = pool.get(Principal(SERVER_CREDS, "bob@example.com"))
        assert ada.auth._cache_path != bob.auth._cache_path
        assert ada.auth._cache is not bob.auth._cache

    def test_a_user_manager_knows_who_it_is(self, isolated_caches):
        pool = GraphClientPool(isolated_caches)
        client = pool.get(Principal(SERVER_CREDS, "ada@example.com"))
        assert client.auth.user == "ada@example.com"

    def test_the_stdio_principal_uses_the_shared_cache(self, isolated_caches):
        pool = GraphClientPool(isolated_caches)
        client = pool.get(Principal(SERVER_CREDS))
        assert client.auth._cache is pool._token_cache
        assert client.auth.user is None


def make_ctx(lifespan_context, request=None):
    """The slice of the SDK's Context that get_graph actually reads."""
    return types.SimpleNamespace(
        request_context=types.SimpleNamespace(
            lifespan_context=lifespan_context, request=request
        )
    )


class TestGetGraph:
    def lifespan(self, proxy_auth=None, credentials=SERVER_CREDS, cache_dir=None):
        return {
            "pool": GraphClientPool(cache_dir),
            "credentials": credentials,
            "proxy_auth": proxy_auth,
        }

    def test_stdio_uses_the_configured_registration(self, isolated_caches):
        client = get_graph(make_ctx(self.lifespan(cache_dir=isolated_caches)))
        assert client.auth.client_id == "server-id"
        assert client.auth.user is None

    def test_without_credentials_it_says_where_they_go(self, isolated_caches):
        with pytest.raises(CredentialsError, match=r"\[credentials\]"):
            get_graph(make_ctx(self.lifespan(credentials=None, cache_dir=isolated_caches)))

    def test_http_serves_the_user_the_proxy_named(self, isolated_caches):
        policy = ProxyAuthPolicy(USER_HEADER)
        request = types.SimpleNamespace(headers={USER_HEADER: "ada@example.com"})
        client = get_graph(make_ctx(self.lifespan(policy, cache_dir=isolated_caches), request))
        assert client.auth.client_id == "server-id"
        assert client.auth.user == "ada@example.com"

    def test_a_caller_cannot_bring_its_own_registration(self, isolated_caches):
        # The X-Outlook-* headers are gone; anything resembling them is inert.
        policy = ProxyAuthPolicy(USER_HEADER)
        request = types.SimpleNamespace(headers={
            USER_HEADER: "ada@example.com",
            "X-Outlook-Client-Id": "attacker-id",
            "X-Outlook-Client-Secret": "attacker-secret",
        })
        client = get_graph(make_ctx(self.lifespan(policy, cache_dir=isolated_caches), request))
        assert client.auth.client_id == "server-id"

    def test_http_refuses_a_request_without_the_header(self, isolated_caches):
        policy = ProxyAuthPolicy(USER_HEADER)
        request = types.SimpleNamespace(headers={})
        with pytest.raises(CredentialsError, match=USER_HEADER):
            get_graph(make_ctx(self.lifespan(policy, cache_dir=isolated_caches), request))

    def test_two_users_never_share_a_client(self, isolated_caches):
        policy = ProxyAuthPolicy(USER_HEADER)
        lifespan = self.lifespan(policy, cache_dir=isolated_caches)
        ada = get_graph(make_ctx(lifespan, types.SimpleNamespace(
            headers={USER_HEADER: "ada@example.com"})))
        bob = get_graph(make_ctx(lifespan, types.SimpleNamespace(
            headers={USER_HEADER: "bob@example.com"})))
        assert ada is not bob
        assert ada.auth._cache is not bob.auth._cache
