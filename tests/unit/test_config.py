"""Unit tests for outlook_mcp.config: the TOML file that picks the transport.

A silent fallback here is the expensive failure mode: an HTTP deployment that
quietly starts on stdio is unreachable with nothing in the logs to say why.
"""

import pytest

from outlook_mcp import config as config_module
from outlook_mcp.config import (
    ConfigError,
    ServerConfig,
    load_config,
    resolve_config_path,
)


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

    def test_reads_an_http_configuration(self, tmp_path):
        path = write_config(
            tmp_path,
            '[server]\ntransport = "http"\nbind_host = "0.0.0.0"\nbind_port = 9000\n',
        )
        config = load_config(str(path))
        assert (config.transport, config.bind_host, config.bind_port) == (
            "http", "0.0.0.0", 9000,
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


class TestHttpUrl:
    def test_names_the_mcp_endpoint(self):
        config = ServerConfig(transport="http", bind_host="127.0.0.1", bind_port=8000)
        assert config.http_url == "http://127.0.0.1:8000/mcp"

    @pytest.mark.parametrize("wildcard", ["0.0.0.0", "::"])
    def test_wildcard_bind_is_shown_as_loopback(self, wildcard):
        # A client cannot dial 0.0.0.0, so printing it as the URL is useless.
        config = ServerConfig(transport="http", bind_host=wildcard, bind_port=8000)
        assert config.http_url == "http://127.0.0.1:8000/mcp"

    def test_ipv6_host_is_bracketed(self):
        config = ServerConfig(transport="http", bind_host="::1", bind_port=8000)
        assert config.http_url == "http://[::1]:8000/mcp"
