"""Unit tests for outlook_mcp.credentials: where a request's credentials come from.

Pure logic, no network: the header reader decides what an HTTP caller is
allowed to authenticate as, so its edge cases are worth pinning down.
"""

import pytest

from outlook_mcp.auth import CredentialsError
from outlook_mcp.credentials import (
    DEFAULT_TENANT_ID,
    HEADER_CLIENT_ID,
    HEADER_CLIENT_SECRET,
    HEADER_TENANT_ID,
    Credentials,
    credentials_from_env,
    credentials_from_headers,
)


class TestCredentialsFromEnv:
    def test_reads_a_complete_set(self):
        creds = credentials_from_env({
            "OUTLOOK_CLIENT_ID": "id",
            "OUTLOOK_CLIENT_SECRET": "secret",
            "OUTLOOK_TENANT_ID": "tenant",
        })
        assert creds == Credentials("id", "secret", "tenant")

    def test_defaults_the_tenant(self):
        creds = credentials_from_env({
            "OUTLOOK_CLIENT_ID": "id",
            "OUTLOOK_CLIENT_SECRET": "secret",
        })
        assert creds.tenant_id == DEFAULT_TENANT_ID

    def test_trims_whitespace(self):
        creds = credentials_from_env({
            "OUTLOOK_CLIENT_ID": "  id  ",
            "OUTLOOK_CLIENT_SECRET": "\tsecret\n",
            "OUTLOOK_TENANT_ID": " tenant ",
        })
        assert creds == Credentials("id", "secret", "tenant")

    @pytest.mark.parametrize("environ", [
        {},
        {"OUTLOOK_CLIENT_ID": "id"},
        {"OUTLOOK_CLIENT_SECRET": "secret"},
        {"OUTLOOK_CLIENT_ID": "id", "OUTLOOK_CLIENT_SECRET": "   "},
    ])
    def test_returns_none_when_incomplete(self, environ):
        assert credentials_from_env(environ) is None

    def test_blank_tenant_falls_back_to_the_default(self):
        creds = credentials_from_env({
            "OUTLOOK_CLIENT_ID": "id",
            "OUTLOOK_CLIENT_SECRET": "secret",
            "OUTLOOK_TENANT_ID": "   ",
        })
        assert creds.tenant_id == DEFAULT_TENANT_ID


class TestCredentialsFromHeaders:
    def test_reads_a_complete_set(self):
        creds = credentials_from_headers({
            HEADER_CLIENT_ID: "id",
            HEADER_CLIENT_SECRET: "secret",
            HEADER_TENANT_ID: "tenant",
        })
        assert creds == Credentials("id", "secret", "tenant")

    @pytest.mark.parametrize("transform", [str.lower, str.upper, lambda s: s])
    def test_header_names_are_case_insensitive(self, transform):
        # Starlette normalises, but a plain dict from a test or another
        # framework does not, so the function normalises for itself.
        creds = credentials_from_headers({
            transform(HEADER_CLIENT_ID): "id",
            transform(HEADER_CLIENT_SECRET): "secret",
        })
        assert creds.client_id == "id"

    def test_defaults_the_tenant(self):
        creds = credentials_from_headers({
            HEADER_CLIENT_ID: "id",
            HEADER_CLIENT_SECRET: "secret",
        })
        assert creds.tenant_id == DEFAULT_TENANT_ID

    def test_trims_whitespace(self):
        creds = credentials_from_headers({
            HEADER_CLIENT_ID: "  id ",
            HEADER_CLIENT_SECRET: " secret ",
            HEADER_TENANT_ID: " tenant ",
        })
        assert creds == Credentials("id", "secret", "tenant")

    def test_missing_headers_are_named_in_the_error(self):
        with pytest.raises(CredentialsError) as excinfo:
            credentials_from_headers({})
        message = str(excinfo.value)
        assert HEADER_CLIENT_ID in message
        assert HEADER_CLIENT_SECRET in message

    def test_error_names_only_the_missing_header(self):
        with pytest.raises(CredentialsError) as excinfo:
            credentials_from_headers({HEADER_CLIENT_ID: "id"})
        missing_list = str(excinfo.value).split(".")[0]
        assert HEADER_CLIENT_SECRET in missing_list
        assert HEADER_CLIENT_ID not in missing_list

    def test_whitespace_only_value_counts_as_missing(self):
        with pytest.raises(CredentialsError):
            credentials_from_headers({
                HEADER_CLIENT_ID: "id",
                HEADER_CLIENT_SECRET: "   ",
            })

    def test_error_says_environment_variables_are_not_consulted(self):
        # HTTP mode must never silently fall back to the server's own
        # credentials; the message is what tells a caller so.
        with pytest.raises(CredentialsError, match="environment variables are not consulted"):
            credentials_from_headers({})


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
