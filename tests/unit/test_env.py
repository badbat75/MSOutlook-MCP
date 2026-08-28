"""Unit tests for outlook_mcp.env: the one .env parser every entry point uses."""

import pytest

from outlook_mcp import env as env_module
from outlook_mcp.env import (
    EnvFileError,
    load_project_env,
    parse_env_file,
    resolve_env_path,
)


class TestParseEnvFile:
    def test_reads_plain_assignments(self):
        assert parse_env_file("A=1\nB=two\n") == {"A": "1", "B": "two"}

    def test_skips_blanks_and_comments(self):
        text = "\n# a comment\n\nA=1\n   # indented comment\n"
        assert parse_env_file(text) == {"A": "1"}

    def test_tolerates_export_prefix(self):
        assert parse_env_file("export A=1\n") == {"A": "1"}

    def test_strips_one_layer_of_matching_quotes(self):
        parsed = parse_env_file("D=\"a b\"\nS='c d'\n")
        assert parsed == {"D": "a b", "S": "c d"}

    def test_keeps_hash_inside_a_quoted_value(self):
        # A client secret is allowed to contain '#'; it is not a comment there.
        assert parse_env_file('K="se#cret"\n') == {"K": "se#cret"}

    def test_leaves_unbalanced_quotes_alone(self):
        assert parse_env_file("K=\"oops'\n") == {"K": "\"oops'"}

    def test_trims_whitespace_around_key_and_value(self):
        assert parse_env_file("  K  =  v  \n") == {"K": "v"}

    def test_keeps_empty_value(self):
        assert parse_env_file("K=\n") == {"K": ""}

    def test_ignores_lines_without_an_equals_sign(self):
        assert parse_env_file("not an assignment\nK=v\n") == {"K": "v"}

    def test_ignores_an_empty_key(self):
        assert parse_env_file("=orphan\nK=v\n") == {"K": "v"}

    def test_last_assignment_wins(self):
        assert parse_env_file("K=1\nK=2\n") == {"K": "2"}


class TestResolveEnvPath:
    def test_explicit_path_must_exist(self, tmp_path):
        with pytest.raises(EnvFileError):
            resolve_env_path(str(tmp_path / "absent.env"))

    def test_explicit_path_is_returned(self, tmp_path):
        path = tmp_path / "present.env"
        path.write_text("A=1\n", encoding="utf-8")
        assert resolve_env_path(str(path)) == path

    def test_environment_variable_is_honoured(self, tmp_path):
        path = tmp_path / "from_var.env"
        path.write_text("A=1\n", encoding="utf-8")
        environ = {env_module.ENV_PATH_VAR: str(path)}
        assert resolve_env_path(environ=environ) == path

    def test_argument_wins_over_environment_variable(self, tmp_path):
        chosen = tmp_path / "chosen.env"
        chosen.write_text("A=1\n", encoding="utf-8")
        environ = {env_module.ENV_PATH_VAR: str(tmp_path / "ignored.env")}
        assert resolve_env_path(str(chosen), environ=environ) == chosen

    def test_returns_none_when_project_file_is_absent(self, tmp_path, monkeypatch):
        monkeypatch.setattr(env_module, "DEFAULT_ENV_PATH", tmp_path / "nothing.env")
        assert resolve_env_path(environ={}) is None


class TestLoadProjectEnv:
    def test_merges_values_and_reports_the_file(self, tmp_path):
        path = tmp_path / ".env"
        path.write_text("OUTLOOK_CLIENT_ID=abc\n", encoding="utf-8")
        environ = {}
        assert load_project_env(str(path), environ=environ) == path
        assert environ == {"OUTLOOK_CLIENT_ID": "abc"}

    def test_never_overwrites_an_existing_variable(self, tmp_path):
        # The "env" block of a Claude Desktop config, or a systemd unit, wins.
        path = tmp_path / ".env"
        path.write_text("OUTLOOK_CLIENT_ID=from-file\n", encoding="utf-8")
        environ = {"OUTLOOK_CLIENT_ID": "from-host"}
        load_project_env(str(path), environ=environ)
        assert environ["OUTLOOK_CLIENT_ID"] == "from-host"

    def test_is_a_no_op_without_a_file(self, tmp_path, monkeypatch):
        monkeypatch.setattr(env_module, "DEFAULT_ENV_PATH", tmp_path / "nothing.env")
        environ = {}
        assert load_project_env(environ=environ) is None
        assert environ == {}

    def test_project_root_holds_the_entry_point(self):
        # Everything the server resolves hangs off this, so a wrong value here
        # would silently make it read files from the wrong place.
        assert (env_module.PROJECT_ROOT / "outlook_mcp_server.py").is_file()
