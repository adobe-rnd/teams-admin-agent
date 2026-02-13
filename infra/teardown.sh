#!/usr/bin/env bash
#
# Removes all Azure resources created by setup.sh:
#   - Resource group (which deletes the Bot resource)
#   - App registration + service principal
#
# Usage:
#   ./infra/teardown.sh
#   ./infra/teardown.sh --app-id <APP_ID> --rg <RG_NAME>
#
set -euo pipefail

SCRIPT_DIR="$(dirname "$0")"
ENV_FILE="$SCRIPT_DIR/.azure-env"

# Load saved values if available
APP_ID=""
RG_NAME=""
APP_NAME=""

if [[ -f "$ENV_FILE" ]]; then
  source "$ENV_FILE"
  echo "Loaded resource IDs from $ENV_FILE"
fi

# Override with args
while [[ $# -gt 0 ]]; do
  case "$1" in
    --app-id) APP_ID="$2"; shift 2 ;;
    --rg)     RG_NAME="$2"; shift 2 ;;
    *) echo "Unknown option: $1"; exit 1 ;;
  esac
done

if [[ -z "$APP_ID" || -z "$RG_NAME" ]]; then
  echo "Error: APP_ID and RG_NAME are required."
  echo "Either run setup.sh first or pass --app-id and --rg."
  exit 1
fi

echo "This will permanently delete:"
echo "  - Resource group: $RG_NAME (and all resources in it)"
echo "  - App registration: $APP_ID"
echo ""
read -rp "Are you sure? (y/N) " confirm
if [[ "$confirm" != "y" && "$confirm" != "Y" ]]; then
  echo "Aborted."
  exit 0
fi

echo ""
echo "==> Deleting resource group '$RG_NAME'..."
az group delete --name "$RG_NAME" --yes --no-wait
echo "    Resource group deletion initiated (runs in background)"

echo "==> Deleting app registration..."
az ad app delete --id "$APP_ID"
echo "    App registration deleted"

# Clean up local state
rm -f "$ENV_FILE"

echo ""
echo "Done. All Azure resources have been removed."
