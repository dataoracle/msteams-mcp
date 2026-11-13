#!/usr/bin/env bash
# Sign in to Microsoft once (delegated OAuth2: authorization code + PKCE) and store
# the refresh token for msteams-mcp, encrypted with the TPM-sealed vault key.
# Run it again whenever the server reports that the sign-in is no longer valid.
#
#   sudo msteams-mcp-authorize          (installed by install.sh)
#   sudo ./deploy/authorize.sh          (from a clone of the repo)
set -euo pipefail
[[ $EUID -eq 0 ]] || { echo "run with sudo" >&2; exit 1; }

APP=/opt/msteams-mcp
ENV_FILE=/etc/msteams-mcp/msteams-mcp.env
KEY=/etc/credstore.encrypted/msteams-mcp.vault_key.cred
STATE=/var/lib/msteams-mcp
HERE="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"

[[ -x "$APP/venv/bin/python" ]] || { echo "Not installed yet: sudo ./deploy/install.sh" >&2; exit 1; }
[[ -f "$ENV_FILE" ]] || { echo "Missing $ENV_FILE: run install.sh" >&2; exit 1; }
id msteams-mcp >/dev/null 2>&1 || { echo "Missing user msteams-mcp: run install.sh" >&2; exit 1; }
if [[ ! -f "$KEY" ]]; then
  echo "==> No vault key yet, generating one"
  "$HERE/seal-credential.sh" vault_key --generate
fi
install -d -o msteams-mcp -g msteams-mcp -m 0700 "$STATE"

cd "$APP/app"
"$APP/venv/bin/python" "$APP/app/authorize.py" \
  --env-file "$ENV_FILE" --key-cred "$KEY" \
  --vault "$STATE/vault.json.enc" --owner msteams-mcp

if ! systemctl is-active --quiet msteams-mcp; then
  echo "msteams-mcp is not running. Start it with: sudo systemctl enable --now msteams-mcp"
fi
