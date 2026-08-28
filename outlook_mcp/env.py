"""Loading of the project ``.env`` file, shared by every entry point.

The server, the authorization script and the contrib client all need the same
``OUTLOOK_*`` variables from the same file, so they all come through here.
Before this module each of them carried its own parser and they disagreed:
one stripped surrounding quotes from a value and one did not, which made a
quoted client secret work in one process and fail in the next.

Two rules the parser keeps:

* the real environment always wins. A variable already set (the ``env`` block
  of a Claude Desktop config, an ``export`` in a service unit) is never
  overwritten by the file.
* the file is found relative to the project, never to the current directory.
  A stdio server inherits the CWD of whatever host spawned it, which may not
  even be traversable by the server's user.
"""

import os
from pathlib import Path
from typing import Dict, Mapping, MutableMapping, Optional

PROJECT_ROOT = Path(__file__).resolve().parent.parent
"""Directory holding outlook_mcp_server.py, resolved from this file's location."""

ENV_FILENAME = ".env"
ENV_PATH_VAR = "OUTLOOK_ENV_FILE"
DEFAULT_ENV_PATH = PROJECT_ROOT / ENV_FILENAME


class EnvFileError(ValueError):
    """Raised when an explicitly requested .env file cannot be used."""


def parse_env_file(text: str) -> Dict[str, str]:
    """Parse ``KEY=VALUE`` lines into a dict.

    Blank lines and ``#`` comments are skipped, an optional leading ``export``
    is tolerated, and one layer of matching single or double quotes is removed
    from the value so a secret containing ``#`` or spaces survives.
    """
    values: Dict[str, str] = {}
    for line in text.splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        if line.startswith("export ") or line.startswith("export\t"):
            line = line[len("export"):].lstrip()
        key, _, value = line.partition("=")
        key = key.strip()
        if not key:
            continue
        value = value.strip()
        if len(value) >= 2 and value[0] == value[-1] and value[0] in ("'", '"'):
            value = value[1:-1]
        values[key] = value
    return values


def resolve_env_path(
    explicit: Optional[str] = None, environ: Mapping[str, str] = os.environ
) -> Optional[Path]:
    """Pick the .env file to read, or None when there is none.

    A path given explicitly (argument or ``OUTLOOK_ENV_FILE``) must exist: a
    typo there should fail loudly instead of silently starting a server with
    no credentials. The implicit project-root file is optional.
    """
    requested = explicit or environ.get(ENV_PATH_VAR)
    if requested:
        path = Path(requested).expanduser()
        if not path.is_file():
            raise EnvFileError(f"Environment file not found: {path}")
        return path
    if DEFAULT_ENV_PATH.is_file():
        return DEFAULT_ENV_PATH
    return None


def load_project_env(
    explicit: Optional[str] = None,
    environ: MutableMapping[str, str] = os.environ,
) -> Optional[Path]:
    """Merge the project .env into ``environ``. Returns the file used, if any.

    Call this before anything reads ``OUTLOOK_*``: it is the only reason the
    variables are present when the server is started by a host that passes no
    environment of its own.
    """
    path = resolve_env_path(explicit, environ)
    if path is None:
        return None
    for key, value in parse_env_file(path.read_text(encoding="utf-8")).items():
        environ.setdefault(key, value)
    return path
