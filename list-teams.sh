#!/usr/bin/env bash
# List all M365 teams via Graph and print a Slack-canvas-ready markdown table.
# Usage: TENANT_ID=your-tenant-id GRAPH_TOKEN=your-bearer-token ./list-teams.sh

set -e
BASE_URL="${GRAPH_BASE_URL:-https://graph.microsoft.com/v1.0}"
TENANT_ID="${TENANT_ID:?Set TENANT_ID}"
GRAPH_TOKEN="${GRAPH_TOKEN:?Set GRAPH_TOKEN}"

escape_md() {
  # Escape ] and \ in link text so markdown doesn't break
  sed 's/\\/\\\\/g; s/\]/\\]/g'
}

next_url="$BASE_URL/groups?\$filter=resourceProvisioningOptions/Any(x:x%20eq%20'Team')&\$select=id,displayName"
echo "| Team | Agent added | Agent announced |"
echo "| --- | --- | --- |"

while [[ -n "$next_url" ]]; do
  resp=$(curl -sS -H "Authorization: Bearer $GRAPH_TOKEN" "$next_url")
  next_url=""
  echo "$resp" | jq -r '.value[]? | "\(.displayName)|\(.id)"' | while IFS='|' read -r name id; do
    [[ -z "$id" ]] && continue
    escaped_name=$(printf '%s' "$name" | escape_md)
    url="https://teams.microsoft.com/l/team/${id}?tenantId=${TENANT_ID}"
    echo "| [$escaped_name]($url) | ☐ | ☐ |"
  done
  next_url=$(echo "$resp" | jq -r '.["@odata.nextLink"] // empty')
  if [[ -n "$next_url" ]]; then
    next_url=$(echo "$next_url" | sed 's/ /%20/g')
  fi
done
