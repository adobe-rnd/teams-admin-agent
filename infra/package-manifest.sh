#!/usr/bin/env bash
#
# Generates the Teams app manifest ZIP ready for upload.
#
#   1. Substitutes the BOT_ID into manifest.json
#   2. Generates placeholder icon PNGs (color 192x192, outline 32x32)
#   3. Zips everything into teams-admin-agent.zip
#
# Usage:
#   ./infra/package-manifest.sh <BOT_ID>
#   ./infra/package-manifest.sh           # reads from infra/.azure-env
#
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
OUT_DIR="$SCRIPT_DIR/../dist"
ENV_FILE="$SCRIPT_DIR/.azure-env"

# Resolve BOT_ID
BOT_ID="${1:-}"
if [[ -z "$BOT_ID" && -f "$ENV_FILE" ]]; then
  source "$ENV_FILE"
  BOT_ID="${APP_ID:-}"
fi
if [[ -z "$BOT_ID" ]]; then
  echo "Usage: $0 <BOT_ID>"
  echo "   or: run setup.sh first (saves BOT_ID to .azure-env)"
  exit 1
fi

echo "BOT_ID: $BOT_ID"

# Prepare output directory
rm -rf "$OUT_DIR"
mkdir -p "$OUT_DIR/manifest"

# ── 1. Substitute BOT_ID into manifest.json ──────────────────────

sed "s/{{BOT_ID}}/$BOT_ID/g" "$SCRIPT_DIR/manifest.json" \
  > "$OUT_DIR/manifest/manifest.json"

echo "==> manifest.json written"

# ── 2. Icons: use custom admin bot icons or generate placeholders ───

ICONS_DIR="$SCRIPT_DIR/icons"
COLOR_SRC="$ICONS_DIR/color.png"
OUTLINE_SRC="$ICONS_DIR/outline.png"

if [[ -f "$COLOR_SRC" && -f "$OUTLINE_SRC" ]]; then
  cp "$COLOR_SRC" "$OUT_DIR/manifest/color.png"
  cp "$OUTLINE_SRC" "$OUT_DIR/manifest/outline.png"
  echo "==> icons copied from $ICONS_DIR (color 192×192, outline 32×32)"
else
  python3 -c "
import struct, zlib, sys

def png(w, h, r, g, b, path):
    def chunk(t, d):
        c = t + d
        return struct.pack('>I', len(d)) + c + struct.pack('>I', zlib.crc32(c) & 0xFFFFFFFF)
    raw = b''
    for _ in range(h):
        raw += b'\x00' + bytes([r, g, b]) * w
    data = (
        b'\x89PNG\r\n\x1a\n'
        + chunk(b'IHDR', struct.pack('>IIBBBBB', w, h, 8, 2, 0, 0, 0))
        + chunk(b'IDAT', zlib.compress(raw))
        + chunk(b'IEND', b'')
    )
    with open(path, 'wb') as f:
        f.write(data)

png(192, 192, 74, 21, 75, sys.argv[1])   # color.png  — purple (#4A154B)
png(32,  32,  74, 21, 75, sys.argv[2])   # outline.png — same color
" "$OUT_DIR/manifest/color.png" "$OUT_DIR/manifest/outline.png"
  echo "==> placeholder icons generated (add infra/icons/color.png and outline.png for admin bot icon)"
fi

# ── 3. Zip ────────────────────────────────────────────────────────

(cd "$OUT_DIR/manifest" && zip -qr ../teams-admin-agent.zip .)

echo "==> dist/teams-admin-agent.zip created"
echo ""
echo "Upload this ZIP in Teams Admin Center → Manage apps → Upload."
echo "Or sideload it in Teams → Apps → Manage your apps → Upload a custom app."
