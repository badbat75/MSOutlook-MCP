"""Unit tests for outlook_mcp.config: the one file a deployment is read from.

Two expensive failure modes live here. A silent fallback leaves an HTTP
deployment quietly running on stdio, unreachable with nothing in the logs to
say why. And an HTTP server bound anywhere but loopback is reachable around the
reverse proxy, which is the only thing making its identity header believable.
"""

from pathlib import Path

import pytest

from outlook_mcp import config as config_module
from outlook_mcp.config import (
    DEFAULT_DOWNLOAD_PATH,
    ConfigError,
    ServerConfig,
    is_loopback,
    load_config,
    resolve_config_path,
)

HTTP = '[server]\ntransport = "http"\nbind_host = "{host}"\n'


@pytest.fixture(autouse=True)
def isolated_lookup(tmp_path, monkeypatch):
    """Never let a real outlook_mcp.toml or $OUTLOOK_MCP_CONFIG reach the tests."""
    monkeypatch.delenv(config_module.CONFIG_ENV_VAR, raising=False)
    monkeypatch.setattr(
        config_module, "DEFAULT_CONFIG_PATH", tmp_path / "absent" / "outlook_mcp.toml"
    )


def write_config(tmp_path, body: str):
    path = tmp_path / "outlook_mcp.toml"
    path.write_text(body, encoding="utf-8")
    return path


class TestResolveConfigPath:
    def test_returns_none_when_nothing_is_configured(self):
        assert resolve_config_path() is None

    def test_explicit_path_must_exist(self, tmp_path):
        with pytest.raises(ConfigError):
            resolve_config_path(str(tmp_path / "typo.toml"))

    def test_environment_variable_must_exist_too(self, tmp_path, monkeypatch):
        monkeypatch.setenv(config_module.CONFIG_ENV_VAR, str(tmp_path / "typo.toml"))
        with pytest.raises(ConfigError):
            resolve_config_path()

    def test_environment_variable_is_used(self, tmp_path, monkeypatch):
        path = write_config(tmp_path, "[server]\n")
        monkeypatch.setenv(config_module.CONFIG_ENV_VAR, str(path))
        assert resolve_config_path() == path

    def test_project_file_is_used_when_present(self, tmp_path, monkeypatch):
        path = write_config(tmp_path, "[server]\n")
        monkeypatch.setattr(config_module, "DEFAULT_CONFIG_PATH", path)
        assert resolve_config_path() == path


class TestLoadConfig:
    def test_defaults_to_stdio_without_a_file(self):
        config = load_config()
        assert config.transport == "stdio"
        assert config.is_http is False
        assert config.source is None
        assert config.has_credentials is False

    def test_reads_an_http_configuration(self, tmp_path):
        path = write_config(
            tmp_path,
            '[server]\ntransport = "http"\nbind_host = "127.0.0.1"\nbind_port = 9000\n',
        )
        config = load_config(str(path))
        assert (config.transport, config.bind_host, config.bind_port) == (
            "http", "127.0.0.1", 9000,
        )
        assert config.is_http is True
        assert config.source == path

    def test_transport_is_case_insensitive(self, tmp_path):
        path = write_config(tmp_path, '[server]\ntransport = "HTTP"\n')
        assert load_config(str(path)).transport == "http"

    def test_empty_file_keeps_the_defaults(self, tmp_path):
        path = write_config(tmp_path, "")
        config = load_config(str(path))
        assert config.transport == "stdio"
        assert config.bind_port == 8000

    def test_rejects_an_unknown_transport(self, tmp_path):
        path = write_config(tmp_path, '[server]\ntransport = "carrier-pigeon"\n')
        with pytest.raises(ConfigError, match="transport"):
            load_config(str(path))

    def test_rejects_invalid_toml(self, tmp_path):
        path = write_config(tmp_path, "[server\ntransport = ")
        with pytest.raises(ConfigError, match="Invalid TOML"):
            load_config(str(path))

    def test_rejects_a_non_table_server_key(self, tmp_path):
        path = write_config(tmp_path, 'server = "http"\n')
        with pytest.raises(ConfigError, match=r"\[server\] must be a table"):
            load_config(str(path))

    def test_rejects_an_empty_bind_host(self, tmp_path):
        path = write_config(tmp_path, '[server]\nbind_host = "   "\n')
        with pytest.raises(ConfigError, match="bind_host"):
            load_config(str(path))

    def test_strips_whitespace_around_bind_host(self, tmp_path):
        path = write_config(tmp_path, '[server]\nbind_host = "  127.0.0.1  "\n')
        assert load_config(str(path)).bind_host == "127.0.0.1"

    @pytest.mark.parametrize("value", ["true", "0", "65536", '"8000"', "8000.5"])
    def test_rejects_a_bad_bind_port(self, tmp_path, value):
        # `true` matters on its own: bool is a subclass of int, so an unguarded
        # isinstance check would silently turn it into port 1.
        path = write_config(tmp_path, f"[server]\nbind_port = {value}\n")
        with pytest.raises(ConfigError, match="bind_port"):
            load_config(str(path))

    def test_accepts_an_unknown_key_but_warns(self, tmp_path, caplog):
        path = write_config(tmp_path, '[server]\nbnid_port = 9000\n')
        config = load_config(str(path))
        assert config.bind_port == 8000
        assert "bnid_port" in caplog.text


class TestAllowedHosts:
    """The Host header values a proxied server accepts. See _transport_security."""

    def test_absent_means_empty(self, tmp_path):
        # And empty means "pass nothing to the SDK", which is what keeps an
        # unproxied deployment behaving exactly as before.
        path = write_config(tmp_path, '[server]\ntransport = "http"\n')
        assert load_config(str(path)).allowed_hosts == ()

    def test_reads_a_list_as_a_tuple(self, tmp_path):
        # A tuple because ServerConfig is frozen.
        body = '[server]\nallowed_hosts = ["a.example.com", " b.example.com "]\n'
        assert load_config(str(write_config(tmp_path, body))).allowed_hosts == (
            "a.example.com", "b.example.com",
        )

    @pytest.mark.parametrize("value", ['"a.example.com"', '["a", 3]', '["a", ""]', '["  "]'])
    def test_rejects_anything_but_non_empty_strings(self, tmp_path, value):
        path = write_config(tmp_path, f"[server]\nallowed_hosts = {value}\n")
        with pytest.raises(ConfigError, match="allowed_hosts"):
            load_config(str(path))

    @pytest.mark.parametrize(
        "value", ["https://a.example.com", "a.example.com/mcp", "http://a.example.com/"]
    )
    def test_rejects_a_url_and_says_what_to_write(self, tmp_path, value):
        # The likeliest mistake, and one that would otherwise never match: the
        # SDK compares the Host header literally, so a URL here means a 421 that
        # explains nothing.
        path = write_config(tmp_path, f'[server]\nallowed_hosts = ["{value}"]\n')
        with pytest.raises(ConfigError, match="a.example.com"):
            load_config(str(path))


class TestCredentialsSection:
    def test_reads_the_app_registration(self, tmp_path):
        path = write_config(
            tmp_path,
            '[credentials]\nclient_id = "cid"\nclient_secret = "sec"\ntenant_id = "t"\n',
        )
        config = load_config(str(path))
        assert (config.client_id, config.client_secret, config.tenant_id) == ("cid", "sec", "t")
        assert config.has_credentials is True

    def test_defaults_the_tenant(self, tmp_path):
        path = write_config(tmp_path, '[credentials]\nclient_id = "cid"\nclient_secret = "sec"\n')
        assert load_config(str(path)).tenant_id == "common"

    def test_trims_whitespace(self, tmp_path):
        path = write_config(
            tmp_path, '[credentials]\nclient_id = "  cid "\nclient_secret = " sec  "\n'
        )
        config = load_config(str(path))
        assert (config.client_id, config.client_secret) == ("cid", "sec")

    def test_no_section_at_all_is_allowed(self, tmp_path):
        # The server still starts, and says what is missing on the first call.
        path = write_config(tmp_path, '[server]\ntransport = "stdio"\n')
        assert load_config(str(path)).has_credentials is False

    @pytest.mark.parametrize("body", [
        '[credentials]\nclient_id = "cid"\n',
        '[credentials]\nclient_secret = "sec"\n',
    ])
    def test_half_a_credential_set_is_refused(self, tmp_path, body):
        # It would otherwise surface as an opaque failure on the first call.
        path = write_config(tmp_path, body)
        with pytest.raises(ConfigError, match="both client_id and client_secret"):
            load_config(str(path))

    def test_rejects_an_empty_value(self, tmp_path):
        path = write_config(tmp_path, '[credentials]\nclient_id = "  "\nclient_secret = "s"\n')
        with pytest.raises(ConfigError, match="client_id"):
            load_config(str(path))


class TestAuthSection:
    def test_defaults_the_user_header(self, tmp_path):
        path = write_config(tmp_path, '[server]\ntransport = "http"\n')
        assert load_config(str(path)).user_header == "X-Auth-Email"

    def test_reads_a_custom_header_and_public_url(self, tmp_path):
        path = write_config(
            tmp_path,
            HTTP.format(host="127.0.0.1")
            + '[auth]\nuser_header = "X-Forwarded-User"\npublic_url = "https://mcp.example.com/"\n',
        )
        config = load_config(str(path))
        assert config.user_header == "X-Forwarded-User"
        assert config.enroll_callback_url == "https://mcp.example.com/oauth/callback"

    def test_rejects_an_empty_header_name(self, tmp_path):
        path = write_config(tmp_path, '[auth]\nuser_header = "  "\n')
        with pytest.raises(ConfigError, match="user_header"):
            load_config(str(path))

    def test_rejects_a_relative_public_url(self, tmp_path):
        # It is handed to Entra as a redirect URI, so anything Entra cannot send
        # a browser back to is a configuration error, not a runtime surprise.
        path = write_config(tmp_path, '[auth]\npublic_url = "/oauth"\n')
        with pytest.raises(ConfigError, match="public_url"):
            load_config(str(path))

    def test_warns_about_an_unknown_key(self, tmp_path, caplog):
        path = write_config(tmp_path, '[auth]\nuser_headr = "X-Auth-Email"\n')
        config = load_config(str(path))
        assert config.user_header == "X-Auth-Email"
        assert "user_headr" in caplog.text

    def test_no_cache_dir_leaves_the_default_alone(self, tmp_path):
        # None, not a resolved path: auth.py owns the default, so the value can
        # be handed to user_cache_path() as it is.
        path = write_config(tmp_path, '[auth]\nuser_header = "X-Auth-Email"\n')
        assert load_config(str(path)).cache_directory is None

    def test_reads_a_cache_dir(self, tmp_path):
        body = f'[auth]\ncache_dir = "{(tmp_path / "caches").as_posix()}"\n'
        config = load_config(str(write_config(tmp_path, body)))
        assert config.cache_directory == tmp_path / "caches"

    def test_expands_a_home_relative_cache_dir(self, tmp_path):
        config = load_config(str(write_config(tmp_path, '[auth]\ncache_dir = "~/c"\n')))
        assert config.cache_directory == Path.home() / "c"

    def test_rejects_a_relative_cache_dir(self, tmp_path):
        # A stdio server inherits the working directory of whatever spawned it,
        # so a relative path would put the tokens somewhere nobody chose.
        path = write_config(tmp_path, '[auth]\ncache_dir = "caches"\n')
        with pytest.raises(ConfigError, match="cache_dir"):
            load_config(str(path))

    def test_accepts_a_posix_path_from_any_platform(self, tmp_path):
        # A Linux deployment's file is routinely written and checked from
        # Windows, where "/opt/..." has no drive letter and Path calls it
        # relative. It is not: it never follows the CWD on the host that runs it.
        path = write_config(tmp_path, '[auth]\ncache_dir = "/opt/outlook-mcp/data/caches"\n')
        assert load_config(str(path)).cache_directory is not None

    def test_enrollment_needs_both_http_and_a_public_url(self, tmp_path):
        (tmp_path / "a").mkdir()
        (tmp_path / "b").mkdir()
        without = write_config(tmp_path / "a", HTTP.format(host="127.0.0.1"))
        assert load_config(str(without)).enrollment_enabled is False

        stdio = write_config(
            tmp_path / "b", '[auth]\npublic_url = "https://mcp.example.com"\n'
        )
        assert load_config(str(stdio)).enrollment_enabled is False

        both = write_config(
            tmp_path,
            HTTP.format(host="127.0.0.1") + '[auth]\npublic_url = "https://mcp.example.com"\n',
        )
        assert load_config(str(both)).enrollment_enabled is True


class TestAttachmentsSection:
    def test_defaults_to_the_downloads_directory(self, tmp_path):
        path = write_config(tmp_path, "")
        assert load_config(str(path)).attachment_dir == DEFAULT_DOWNLOAD_PATH

    def test_reads_a_configured_path(self, tmp_path):
        path = write_config(tmp_path, '[attachments]\ndownload_path = "/srv/outlook"\n')
        assert str(load_config(str(path)).attachment_dir) in ("/srv/outlook", "\\srv\\outlook")

    def test_handing_a_file_over_needs_both_http_and_a_public_url(self):
        # Over stdio the caller already has the file, and over HTTP the server
        # cannot say what URL it is reachable on without being told.
        assert ServerConfig().downloads_enabled is False
        assert ServerConfig(transport="http").downloads_enabled is False
        assert ServerConfig(public_url="https://mcp.example.com").downloads_enabled is False
        assert ServerConfig(
            transport="http", public_url="https://mcp.example.com"
        ).downloads_enabled is True

    def test_the_download_url_hangs_off_the_public_url(self):
        config = ServerConfig(transport="http", public_url="https://mcp.example.com/")
        assert config.download_url("tok") == "https://mcp.example.com/attachments/tok"

    def test_there_is_no_download_url_without_a_public_one(self):
        assert ServerConfig(transport="http").download_url("tok") is None

    def test_expands_a_home_relative_path(self, tmp_path):
        path = write_config(tmp_path, '[attachments]\ndownload_path = "~/mail"\n')
        assert "~" not in str(load_config(str(path)).attachment_dir)


class TestRetention:
    """How long a download nobody fetched stays on the server."""

    def test_an_hour_by_default(self):
        assert ServerConfig(transport="http").retention_seconds == 3600

    def test_a_configured_number_of_minutes(self, tmp_path):
        path = write_config(tmp_path, "[attachments]\nretention_minutes = 5\n")
        config = load_config(str(path))
        assert config.retention_minutes == 5

    def test_zero_means_never(self):
        assert ServerConfig(transport="http", retention_minutes=0).retention_seconds is None

    def test_stdio_downloads_never_expire(self):
        # There the file the tool wrote is the answer, in the caller's own
        # download directory: an expiry would delete the user's file.
        assert ServerConfig().retention_seconds is None
        assert ServerConfig(retention_minutes=5).retention_seconds is None

    @pytest.mark.parametrize("value", ["-1", '"60"', "true", "1.5"])
    def test_a_bad_value_is_refused(self, tmp_path, value):
        path = write_config(tmp_path, f"[attachments]\nretention_minutes = {value}\n")
        with pytest.raises(ConfigError, match="retention_minutes"):
            load_config(str(path))


class TestHttpIsRefusedOffLoopback:
    """The identity header is only believable if the proxy is the only way in.

    Over HTTP a request carries no credential at all, so a server reachable
    around the proxy would let anyone name any mailbox and be served that
    user's tokens. Listening on loopback is what establishes it, and the server
    refuses to start without it.
    """

    @pytest.mark.parametrize("host", ["127.0.0.1", "::1", "localhost", "127.0.1.1"])
    def test_a_loopback_bind_is_accepted(self, tmp_path, host):
        path = write_config(tmp_path, HTTP.format(host=host))
        assert load_config(str(path)).is_http is True

    @pytest.mark.parametrize("host", ["0.0.0.0", "::", "10.0.0.5", "192.168.1.20"])
    def test_a_routable_bind_is_refused(self, tmp_path, host):
        path = write_config(tmp_path, HTTP.format(host=host))
        with pytest.raises(ConfigError, match="loopback"):
            load_config(str(path))

    def test_the_refusal_explains_the_exposure(self, tmp_path):
        path = write_config(tmp_path, HTTP.format(host="0.0.0.0"))
        with pytest.raises(ConfigError, match="any mailbox"):
            load_config(str(path))

    def test_a_hostname_bind_is_treated_as_reachable(self, tmp_path):
        # Not resolvable here, and assuming it means loopback is exactly the
        # mistake that exposes every mailbox on the server.
        path = write_config(tmp_path, HTTP.format(host="mcp.internal"))
        with pytest.raises(ConfigError):
            load_config(str(path))

    def test_the_default_bind_host_is_already_loopback(self, tmp_path):
        path = write_config(tmp_path, '[server]\ntransport = "http"\n')
        assert load_config(str(path)).is_http is True

    def test_stdio_is_unaffected(self, tmp_path):
        # No socket, no exposure: bind_host is meaningless there.
        path = write_config(tmp_path, '[server]\nbind_host = "0.0.0.0"\n')
        assert load_config(str(path)).transport == "stdio"


class TestHttpUrl:
    def test_names_the_mcp_endpoint(self):
        config = ServerConfig(transport="http", bind_host="127.0.0.1", bind_port=8000)
        assert config.http_url == "http://127.0.0.1:8000/mcp"

    def test_ipv6_host_is_bracketed(self):
        config = ServerConfig(transport="http", bind_host="::1", bind_port=8000)
        assert config.http_url == "http://[::1]:8000/mcp"


class TestIsLoopback:
    @pytest.mark.parametrize("host", ["127.0.0.1", "127.5.5.5", "::1", "[::1]", "localhost"])
    def test_recognises_loopback(self, host):
        assert is_loopback(host) is True

    @pytest.mark.parametrize("host", ["0.0.0.0", "::", "10.0.0.1", "mcp.internal", ""])
    def test_everything_else_is_reachable(self, host):
        assert is_loopback(host) is False
