#!/usr/bin/env bash
#
# Creates all Azure resources needed for teams-admin-agent:
#   1. Entra ID app registration (multi-tenant)
#   2. Client secret
#   3. Microsoft Graph API permissions + admin consent
#   4. Resource group
#   5. Azure Bot resource
#   6. Teams channel on the bot
#
# Prerequisites: az cli, logged in (az login), jq
#
# Usage:
#   ./infra/setup.sh
#   ./infra/setup.sh --worker-url https://teams-admin-agent.workers.dev
#
set -euo pipefail

# ── Defaults ──────────────────────────────────────────────────────

APP_NAME="teams-admin-agent"
RG_NAME="teams-admin-agent-rg"
LOCATION="eastus"
BOT_SKU="F0"  # Free tier

# Microsoft Graph well-known app ID
GRAPH_API="00000003-0000-0000-c000-000000000000"

# Application permission GUIDs (Role type)
PERM_USER_READ_ALL="df021288-bdef-4463-88db-98f22de89214"
PERM_TEAM_READBASIC_ALL="2280dda6-0bfd-44ee-a2f4-cb867571a9d4"
PERM_TEAMMEMBER_READWRITE_ALL="0121dc95-1b9f-4aed-8bac-58c5ac466f35"

# ── Parse args ────────────────────────────────────────────────────

WORKER_URL=""
while [[ $# -gt 0 ]]; do
  case "$1" in
    --worker-url) WORKER_URL="$2"; shift 2 ;;
    --name)       APP_NAME="$2"; shift 2 ;;
    --rg)         RG_NAME="$2"; shift 2 ;;
    --location)   LOCATION="$2"; shift 2 ;;
    *) echo "Unknown option: $1"; exit 1 ;;
  esac
done

# ── Preflight checks ─────────────────────────────────────────────

for cmd in az jq; do
  if ! command -v "$cmd" &>/dev/null; then
    echo "Error: $cmd is required but not installed." >&2
    exit 1
  fi
done

if ! az account show &>/dev/null; then
  echo "Error: not logged in. Run 'az login' first." >&2
  exit 1
fi

TENANT_ID=$(az account show --query tenantId -o tsv)
echo "Tenant:   $TENANT_ID"
echo "App name: $APP_NAME"
echo ""

if [[ -z "$WORKER_URL" ]]; then
  read -rp "Enter your Cloudflare Worker URL (e.g. https://teams-admin-agent.workers.dev): " WORKER_URL
fi

WORKER_URL="${WORKER_URL%/}"  # strip trailing slash
MESSAGING_ENDPOINT="${WORKER_URL}/api/messages"
echo "Messaging endpoint: $MESSAGING_ENDPOINT"
echo ""

# ── 1. App registration ──────────────────────────────────────────

echo "==> Creating app registration..."

APP_ID=$(az ad app create \
  --display-name "$APP_NAME" \
  --sign-in-audience AzureADMultipleOrgs \
  --query appId -o tsv)

echo "    App (client) ID: $APP_ID"

# Create service principal (required for admin consent)
echo "==> Creating service principal..."
az ad sp create --id "$APP_ID" -o none 2>/dev/null || true

# ── 2. Client secret ─────────────────────────────────────────────

echo "==> Creating client secret..."

CLIENT_SECRET=$(az ad app credential reset \
  --id "$APP_ID" \
  --display-name "teams-admin-agent-secret" \
  --years 2 \
  --query password -o tsv)

echo "    Client secret created (valid for 2 years)"

# ── 3. API permissions ───────────────────────────────────────────

echo "==> Adding Microsoft Graph permissions..."

az ad app permission add \
  --id "$APP_ID" \
  --api "$GRAPH_API" \
  --api-permissions \
    "${PERM_USER_READ_ALL}=Role" \
    "${PERM_TEAM_READBASIC_ALL}=Role" \
    "${PERM_TEAMMEMBER_READWRITE_ALL}=Role" \
  -o none

echo "    Added: User.Read.All, Team.ReadBasic.All, TeamMember.ReadWrite.All"

# Brief pause for propagation
sleep 5

echo "==> Granting admin consent..."
az ad app permission admin-consent --id "$APP_ID" -o none
echo "    Admin consent granted"

# ── 4. Resource group ─────────────────────────────────────────────

echo "==> Creating resource group '$RG_NAME' in '$LOCATION'..."
az group create --name "$RG_NAME" --location "$LOCATION" -o none

# ── 5. Azure Bot ──────────────────────────────────────────────────

echo "==> Creating Azure Bot..."

az bot create \
  --resource-group "$RG_NAME" \
  --name "$APP_NAME" \
  --app-type MultiTenant \
  --appid "$APP_ID" \
  --endpoint "$MESSAGING_ENDPOINT" \
  --sku "$BOT_SKU" \
  -o none

echo "    Bot created: $APP_NAME"

# ── 6. Teams channel ─────────────────────────────────────────────

echo "==> Enabling Microsoft Teams channel..."
az bot msteams create \
  --resource-group "$RG_NAME" \
  --name "$APP_NAME" \
  -o none 2>/dev/null || true

echo "    Teams channel enabled"

# ── Save state (no secrets in this file — just resource IDs) ──────

SCRIPT_DIR="$(dirname "$0")"

cat > "$SCRIPT_DIR/.azure-env" <<EOF
APP_ID=$APP_ID
APP_NAME=$APP_NAME
RG_NAME=$RG_NAME
TENANT_ID=$TENANT_ID
EOF

# ── Push secrets directly into Cloudflare Worker ──────────────────

echo ""
echo "============================================================"
echo "  Azure setup complete"
echo "============================================================"
echo ""
echo "App (client) ID : $APP_ID"
echo "Tenant ID       : $TENANT_ID"
echo "Resource group  : $RG_NAME"
echo "Messaging URL   : $MESSAGING_ENDPOINT"
echo ""
echo "────────────────────────────────────────────────────────────"
echo "  Pushing Azure secrets to Cloudflare Worker..."
echo "────────────────────────────────────────────────────────────"
echo ""

if ! command -v wrangler &>/dev/null; then
  echo "  wrangler CLI not found — saving secrets to a local file instead."
  echo "  Install wrangler and run:  source infra/.secrets && ./infra/push-secrets.sh"
  echo ""

  # Write secrets to a local file (gitignored via infra/.azure-env pattern)
  SECRETS_FILE="$SCRIPT_DIR/.secrets"
  cat > "$SECRETS_FILE" <<EOF
export _BOT_ID='$APP_ID'
export _BOT_PASSWORD='$CLIENT_SECRET'
export _MS_TENANT_ID='$TENANT_ID'
export _MS_CLIENT_ID='$APP_ID'
export _MS_CLIENT_SECRET='$CLIENT_SECRET'
EOF
  chmod 600 "$SECRETS_FILE"
  echo "  Secrets written to infra/.secrets (mode 600, gitignored)"
else
  echo "$APP_ID"        | wrangler secret put BOT_ID 2>/dev/null        && echo "  ✓ BOT_ID"
  echo "$CLIENT_SECRET" | wrangler secret put BOT_PASSWORD 2>/dev/null  && echo "  ✓ BOT_PASSWORD"
  echo "$TENANT_ID"     | wrangler secret put MS_TENANT_ID 2>/dev/null  && echo "  ✓ MS_TENANT_ID"
  echo "$APP_ID"        | wrangler secret put MS_CLIENT_ID 2>/dev/null  && echo "  ✓ MS_CLIENT_ID"
  echo "$CLIENT_SECRET" | wrangler secret put MS_CLIENT_SECRET 2>/dev/null && echo "  ✓ MS_CLIENT_SECRET"
  echo ""
  echo "  All Azure secrets pushed to Cloudflare Worker."
fi

# Clear secret from shell memory
unset CLIENT_SECRET

echo ""
echo "────────────────────────────────────────────────────────────"
echo "  Next steps:"
echo "────────────────────────────────────────────────────────────"
echo ""
echo "  1. Set your Slack secrets:"
echo "     wrangler secret put SLACK_BOT_TOKEN"
echo "     wrangler secret put SLACK_SIGNING_SECRET"
echo "     wrangler secret put SLACK_ADMIN_CHANNEL_ID"
echo ""
echo "  2. Deploy the worker:"
echo "     npm run deploy"
echo ""
echo "  3. Package the Teams manifest:"
echo "     ./infra/package-manifest.sh"
echo ""
echo "  4. Upload the manifest ZIP in Teams Admin Center"
echo ""
echo "(Resource IDs saved to infra/.azure-env for teardown)"
