#!/usr/bin/env bash
# Outlook MCP - deploy to a Linux host as a systemd service
# =========================================================
# Copies this checkout to a remote host, builds a virtual environment there,
# installs the configuration and a systemd unit, and starts the service.
#
#   ./scripts/deploy.sh --target root@server.example.com
#   ./scripts/deploy.sh --target me@host --dir /srv/outlook-mcp \
#                       --service-user outlook --config deploy/prod.toml
#
# On Windows run it from Git Bash; it needs ssh, scp, git and tar.
#
# What ends up on the host:
#
#   <dir>/app/               this checkout, replaced whole on every deploy
#   <dir>/venv/              the virtual environment, built once and reused
#   <dir>/outlook_mcp.toml   the configuration, 0600, owned by the service user
#   <service user home>/     the MSAL token caches, one file per enrolled user
#
# The unit runs the HTTP transport, which is the mode with something for a
# service manager to supervise: a stdio server is spawned per client by
# whatever MCP host talks to it, and needs neither a unit nor a service user.
#
# Only files git knows about are uploaded, so a local outlook_mcp.toml or .env
# never travels by accident. The configuration the service runs on is the one
# named by --config, or one already present on the host, or a starter file this
# script writes for you to fill in.

set -euo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"

TARGET=""
DIR="/opt/outlook-mcp"
DATA_DIR=""
SERVICE_USER="outlook-mcp"
SERVICE_NAME="outlook-mcp"
CONFIG_FILE=""
PORT="8000"
PYTHON="python3"
RESTART="yes"
DRY_RUN="no"

CYAN=$'\033[0;36m'; RED=$'\033[0;31m'; YELLOW=$'\033[0;33m'
GREEN=$'\033[0;32m'; GRAY=$'\033[0;90m'; NC=$'\033[0m'

# Progress goes to stderr, so that --dry-run can put the rendered script on
# stdout and nothing else.
log()  { printf '%s==>%s %s\n' "$CYAN" "$NC" "$*" >&2; }
warn() { printf '%sWARNING:%s %s\n' "$YELLOW" "$NC" "$*" >&2; }
die()  { printf '%sERROR:%s %s\n' "$RED" "$NC" "$*" >&2; exit 1; }

usage() {
    cat <<'USAGE'
Usage: scripts/deploy.sh --target USER@HOST [options]

Required:
  --target USER@HOST      SSH destination. Must be root or able to sudo.

Options:
  --dir PATH              Install directory on the host  (default /opt/outlook-mcp)
  --data-dir PATH         Token caches and attachments   (default <dir>/data).
                          It is the service's HOME, and the one writable path
                          the unit grants.
  --service-user NAME     System user the service runs as (default outlook-mcp)
  --service-name NAME     systemd unit name, without .service (default outlook-mcp)
  --config PATH           Local TOML to install as the server configuration
  --port N                Port for the starter configuration (default 8000)
  --python PATH           Python interpreter on the host   (default python3)
  --no-restart            Install everything, leave the running service alone
  --dry-run               Print the script that would run on the host, and stop
  -h, --help              This text
USAGE
}

while [ $# -gt 0 ]; do
    case "$1" in
        --target)       TARGET="${2:-}"; shift 2 ;;
        --dir)          DIR="${2:-}"; shift 2 ;;
        --data-dir)     DATA_DIR="${2:-}"; shift 2 ;;
        --service-user) SERVICE_USER="${2:-}"; shift 2 ;;
        --service-name) SERVICE_NAME="${2:-}"; shift 2 ;;
        --config)       CONFIG_FILE="${2:-}"; shift 2 ;;
        --port)         PORT="${2:-}"; shift 2 ;;
        --python)       PYTHON="${2:-}"; shift 2 ;;
        --no-restart)   RESTART="no"; shift ;;
        --dry-run)      DRY_RUN="yes"; shift ;;
        -h|--help)      usage; exit 0 ;;
        *)              usage >&2; die "Unknown option: $1" ;;
    esac
done

# --- Check the arguments before touching anything ---------------------------

[ -n "$TARGET" ] || { usage >&2; die "--target is required"; }
case "$DIR" in /*) ;; *) die "--dir must be an absolute path, not '$DIR'" ;; esac

# Everything the service writes lives here: the MSAL token caches under
# .outlook_mcp/caches/, and the attachment downloads. Keeping it inside the
# install directory makes the deployment one tree to back up or remove, and the
# deploy never touches it: only app/ is replaced.
[ -n "$DATA_DIR" ] || DATA_DIR="$DIR/data"
case "$DATA_DIR" in /*) ;; *) die "--data-dir must be an absolute path, not '$DATA_DIR'" ;; esac
case "$DATA_DIR" in
    "$DIR"/app|"$DIR"/app/*|"$DIR"/venv|"$DIR"/venv/*)
        die "--data-dir must not be inside app/ or venv/: a deploy replaces those" ;;
esac

for name in "$SERVICE_USER" "$SERVICE_NAME"; do
    case "$name" in
        *[!A-Za-z0-9_.-]*|"")
            die "'$name' is not a usable user or unit name: letters, digits, '_', '.' and '-' only" ;;
    esac
done

case "$PORT" in
    ''|*[!0-9]*) die "--port must be a number, not '$PORT'" ;;
esac
[ "$PORT" -ge 1 ] && [ "$PORT" -le 65535 ] || die "--port must be between 1 and 65535"

if [ -n "$CONFIG_FILE" ]; then
    [ -f "$CONFIG_FILE" ] || die "--config file not found: $CONFIG_FILE"
    HAVE_CONFIG="yes"
else
    HAVE_CONFIG="no"
fi

for tool in git tar ssh scp; do
    command -v "$tool" >/dev/null 2>&1 || die "$tool is not on PATH"
done

cd "$ROOT"
git rev-parse --git-dir >/dev/null 2>&1 || die "$ROOT is not a git checkout"

# --- Pack the checkout ------------------------------------------------------
# Tracked files plus untracked ones git would keep, which is everything in the
# checkout except what .gitignore covers: no venv, no .env, no outlook_mcp.toml.

TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

log "Packing the checkout"
git ls-files -z --cached --others --exclude-standard > "$TMP/files.z"
# Written through stdout rather than -f PATH: GNU tar reads a colon in a
# filename as host:path and tries to open a remote archive, which is what a
# Windows temporary directory looks like. bsdtar has no --force-local, so
# redirection is the portable way to say "this is a local file".
tar --null --files-from "$TMP/files.z" -czf - > "$TMP/app.tar.gz"
printf '%s    %s files, %s\n' "$GRAY" \
    "$(tr -cd '\0' < "$TMP/files.z" | wc -c | tr -d ' ')" \
    "$(du -h "$TMP/app.tar.gz" | cut -f1)$NC" >&2

if ! git diff --quiet || ! git diff --cached --quiet; then
    warn "the working tree has uncommitted changes, and they are being deployed"
fi

# --- Render the script that runs on the host --------------------------------

render_bootstrap() {
    local stage="$1"
    printf '#!/usr/bin/env bash\n'
    printf '# Rendered by scripts/deploy.sh on %s. Runs on the target host, as root.\n' "$(date -u '+%Y-%m-%dT%H:%M:%SZ')"
    printf 'DIR=%q\n' "$DIR"
    printf 'DATA_DIR=%q\n' "$DATA_DIR"
    printf 'SERVICE_USER=%q\n' "$SERVICE_USER"
    printf 'SERVICE_NAME=%q\n' "$SERVICE_NAME"
    printf 'PYTHON=%q\n' "$PYTHON"
    printf 'PORT=%q\n' "$PORT"
    printf 'STAGE=%q\n' "$stage"
    printf 'HAVE_CONFIG=%q\n' "$HAVE_CONFIG"
    printf 'RESTART=%q\n' "$RESTART"
    cat <<'REMOTE'

set -euo pipefail

# Re-run under sudo when the SSH user is not root. ssh -t gives it a terminal,
# so a password prompt works.
if [ "$(id -u)" -ne 0 ]; then
    exec sudo -p 'sudo password for %u on %H: ' bash "$0" "$@"
fi

log()  { printf '\033[0;36m==>\033[0m %s\n' "$*"; }
warn() { printf '\033[0;33mWARNING:\033[0m %s\n' "$*" >&2; }

CONFIG="$DIR/outlook_mcp.toml"
TARBALL="$STAGE/app.tar.gz"
UNIT="/etc/systemd/system/$SERVICE_NAME.service"

command -v systemctl >/dev/null 2>&1 || { echo "This host has no systemd." >&2; exit 1; }

# --- The service user -------------------------------------------------------

if id "$SERVICE_USER" >/dev/null 2>&1; then
    log "Service user $SERVICE_USER already exists"
else
    login_shell=/bin/false
    for candidate in /usr/sbin/nologin /sbin/nologin; do
        if [ -x "$candidate" ]; then login_shell="$candidate"; break; fi
    done
    log "Creating system user $SERVICE_USER (home $DATA_DIR)"
    useradd --system --home-dir "$DATA_DIR" --shell "$login_shell" "$SERVICE_USER"
fi

SERVICE_GROUP="$(id -gn "$SERVICE_USER")"

# The data directory is the service's HOME, whatever passwd says for this user:
# auth.py builds the token cache paths from Path.home(), and the unit sets HOME
# to exactly this. One file per enrolled user lands in .outlook_mcp/caches/, so
# nobody else has any business reading the tree.
install -d -o "$SERVICE_USER" -g "$SERVICE_GROUP" -m 700 "$DATA_DIR"

# --- The interpreter --------------------------------------------------------
# Checked before the venv, because "pip install ." on an old Python fails with a
# requires-python error that reads like a packaging problem.

if ! "$PYTHON" -c 'import sys; sys.exit(0 if sys.version_info >= (3, 11) else 1)'; then
    echo "$PYTHON is $("$PYTHON" -V 2>&1), and this server needs 3.11 or newer." >&2
    echo "Install a newer interpreter and pass it with --python." >&2
    exit 1
fi

# --- The code ---------------------------------------------------------------
# Unpacked beside the venv rather than around it, so a deploy can replace the
# tree wholesale: no file survives a rename, and the venv is not rebuilt.

log "Unpacking the checkout into $DIR/app"
install -d -m 755 "$DIR"
rm -rf "$DIR/app.new" "$DIR/app.old"
install -d -m 755 "$DIR/app.new"
tar -xzf - -C "$DIR/app.new" < "$TARBALL"
if [ -d "$DIR/app" ]; then mv "$DIR/app" "$DIR/app.old"; fi
mv "$DIR/app.new" "$DIR/app"
rm -rf "$DIR/app.old"

if [ ! -x "$DIR/venv/bin/python" ]; then
    log "Creating the virtual environment ($PYTHON)"
    "$PYTHON" -m venv "$DIR/venv"
fi

log "Installing the package and its dependencies"
"$DIR/venv/bin/pip" install --quiet --upgrade pip
"$DIR/venv/bin/pip" install --quiet --editable "$DIR/app"

# --- The configuration ------------------------------------------------------

if [ "$HAVE_CONFIG" = yes ]; then
    log "Installing the configuration given with --config"
    install -o "$SERVICE_USER" -g "$SERVICE_GROUP" -m 600 \
            "$STAGE/outlook_mcp.toml" "$CONFIG"
    rm -f "$STAGE/outlook_mcp.toml"
elif [ -f "$CONFIG" ]; then
    log "Keeping the configuration already on this host"
else
    warn "no configuration here and none given, writing a starter file"
    umask 077
    cat > "$CONFIG" <<TOML
# Outlook MCP - written by scripts/deploy.sh. Fill in [credentials].
#
# Loopback only: over HTTP a request carries no credential, just the identity
# header the reverse proxy appends, and that is believable only because the
# proxy is the sole route in.

[server]
transport = "http"
bind_host = "127.0.0.1"
bind_port = $PORT

[credentials]
client_id = ""
client_secret = ""
tenant_id = "common"

[auth]
user_header = "X-Auth-Email"
# public_url = "https://outlook-mcp.example.com"   # turns on /oauth/login

# One MSAL token cache per enrolled mailbox, here rather than under a home
# directory: this directory already belongs to this server, so a dotted
# .outlook_mcp level inside it would repeat what the path already says.
cache_dir = "$DATA_DIR/caches"

[attachments]
# Inside the data directory, which is the only path the unit lets the service
# write to. A remote caller never gets to choose where files land.
download_path = "$DATA_DIR/attachments"
TOML
    chown "$SERVICE_USER:$SERVICE_GROUP" "$CONFIG"
    chmod 600 "$CONFIG"
fi

# The server's own loader, so this fails here rather than in a restart loop:
# it is what rejects an unknown transport and a non-loopback bind.
log "Validating $CONFIG"
if ! "$DIR/venv/bin/python" - "$CONFIG" "$DATA_DIR" <<'PY'
import os
import sys
from pathlib import Path

from outlook_mcp.config import ConfigError, load_config

path, data_dir = sys.argv[1], Path(sys.argv[2])

# Resolve "~" the way the service will, not the way root would: the unit sets
# HOME to the data directory, so the default ~/Downloads/outlook_attachments
# lands inside it. Judging it as /root/... would fail a correct configuration.
os.environ["HOME"] = str(data_dir)

try:
    config = load_config(path)
except ConfigError as e:
    sys.exit(f"    {e}")

if config.transport != "http":
    sys.exit(
        f"    {path} sets transport = {config.transport!r}. A systemd service "
        "runs the HTTP transport; a stdio server is spawned once per client and "
        "would exit immediately here."
    )

# The unit grants exactly one writable path. Anything outside it would only fail
# when someone downloads their first attachment or enrols, which is a long way
# from here.
#
# Imported here rather than at the top, and it has to stay here: auth.py builds
# USER_CACHE_DIR from Path.home() at import time, so importing it before the
# assignment above would capture root's home and judge a correct default wrong.
from outlook_mcp.auth import USER_CACHE_DIR  # noqa: E402

caches = config.cache_directory or USER_CACHE_DIR
for label, directory in (("[attachments].download_path", config.attachment_dir),
                         ("[auth].cache_dir", caches)):
    if directory != data_dir and data_dir not in directory.parents:
        sys.exit(
            f"    {label} is {directory}, outside the data directory "
            f"{data_dir}. The service cannot write there: ProtectSystem=strict "
            "makes everything else read-only."
        )

if not config.has_credentials:
    print("    no [credentials] yet: the server will start and refuse every call")
print(f"    http://{config.bind_host}:{config.bind_port}/mcp, user header "
      f"{config.user_header}")
print(f"    token caches -> {caches}")
print(f"    attachments  -> {config.attachment_dir}")
PY
then
    echo "Configuration unusable; the service was not started." >&2
    exit 1
fi

# The service reads the code and the config, and writes neither. Only the
# config carries a secret, so only it is restricted to the service user.
chown -R root:root "$DIR/app" "$DIR/venv"
chown "$SERVICE_USER:$SERVICE_GROUP" "$CONFIG"
chmod 600 "$CONFIG"

# --- The unit ---------------------------------------------------------------

log "Installing $UNIT"
sed -e "s|@SERVICE_NAME@|$SERVICE_NAME|g" \
    -e "s|@SERVICE_USER@|$SERVICE_USER|g" \
    -e "s|@SERVICE_GROUP@|$SERVICE_GROUP|g" \
    -e "s|@DATA_DIR@|$DATA_DIR|g" \
    -e "s|@DIR@|$DIR|g" \
    -e "s|@CONFIG@|$CONFIG|g" \
    "$DIR/app/contrib/systemd/outlook-mcp.service" > "$UNIT"
chmod 644 "$UNIT"
if grep -q '@[A-Z_]\+@' "$UNIT"; then
    warn "$UNIT still has an unsubstituted placeholder"
fi

systemctl daemon-reload
systemctl enable "$SERVICE_NAME" >/dev/null

rm -f "$TARBALL"

if [ "$RESTART" != yes ]; then
    log "Not restarting (--no-restart). Apply it with: systemctl restart $SERVICE_NAME"
    exit 0
fi

log "Restarting $SERVICE_NAME"
systemctl restart "$SERVICE_NAME"
sleep 1
if systemctl is-active --quiet "$SERVICE_NAME"; then
    log "$SERVICE_NAME is running"
else
    warn "$SERVICE_NAME did not come up. Its last lines:"
    journalctl -u "$SERVICE_NAME" -n 20 --no-pager || true
    exit 1
fi
REMOTE
}

if [ "$DRY_RUN" = yes ]; then
    render_bootstrap '<remote home>/.cache/outlook-mcp-deploy'
    exit 0
fi

# --- Ship it ----------------------------------------------------------------

log "Preparing the staging directory on $TARGET"
STAGE="$(ssh "$TARGET" 'umask 077; d="$HOME/.cache/outlook-mcp-deploy"; mkdir -p "$d"; printf %s "$d"')"
[ -n "$STAGE" ] || die "could not create a staging directory on $TARGET"

log "Uploading to $TARGET:$STAGE"
scp -q "$TMP/app.tar.gz" "$TARGET:$STAGE/app.tar.gz"
if [ "$HAVE_CONFIG" = yes ]; then
    # Into a 0700 directory, moved into place and removed there. It holds the
    # client secret, so it never sits in a world-readable /tmp.
    scp -q "$CONFIG_FILE" "$TARGET:$STAGE/outlook_mcp.toml"
fi

render_bootstrap "$STAGE" > "$TMP/bootstrap.sh"
scp -q "$TMP/bootstrap.sh" "$TARGET:$STAGE/bootstrap.sh"

log "Running the install on $TARGET"
ssh -t "$TARGET" "bash $(printf '%q' "$STAGE/bootstrap.sh")"

cat <<EOF

${GREEN}Deployed.${NC}  ${GRAY}$TARGET:$DIR${NC}

  Endpoint      http://127.0.0.1:$PORT/mcp   (loopback: reach it through your proxy)
  Unit          systemctl status $SERVICE_NAME
  Logs          journalctl -u $SERVICE_NAME -f
  Config        $DIR/outlook_mcp.toml
  Data          $DATA_DIR   (token caches and attachments; survives a redeploy)

Next, if you have not already:
  1. Put a reverse proxy in front of it, setting the identity header on every
     location it forwards from. Snippet: docs/SETUP.md#multi-user-behind-a-reverse-proxy
  2. Enrol each user, either at <public_url>/oauth/login in their browser, or on
     the host. The same --config the service runs on is what says where the
     cache goes, so it lands where the service will look for it:

       sudo -u $SERVICE_USER ${DIR}/venv/bin/outlook-mcp-auth \\
            --config ${DIR}/outlook_mcp.toml --user them@example.com --no-browser

     Give each installation its own sign-in. Copying a cache from another host
     looks like a shortcut: Entra rotates the refresh token as it redeems it, so
     two copies keep invalidating each other.
EOF
