#!/bin/sh
# Launch the axlook MCP server from the plugin (or a checkout): pick a JS runtime, make sure the
# Accessibility helper runs on this Mac, then exec the bundled server.
ROOT="${CLAUDE_PLUGIN_ROOT:-$(cd "$(dirname "$0")/.." && pwd)}"
# GUI hosts start MCP servers with a minimal PATH.
export PATH="$HOME/.bun/bin:/opt/homebrew/bin:/usr/local/bin:$PATH"

if [ "$(uname -s)" != "Darwin" ]; then
  echo "axlook: macOS only (it drives New Outlook for Mac)" >&2
  exit 1
fi

HELPER="$ROOT/dist/axlook-helper"
chmod +x "$HELPER" 2>/dev/null
if ! "$HELPER" --version >/dev/null 2>&1; then
  # Prebuilt helper unusable (stripped exec bit, quarantine, ...): build one from source.
  DATA="${AXLOOK_DATA:-$HOME/.cache/axlook}"
  mkdir -p "$DATA"
  HELPER="$DATA/axlook-helper"
  if ! "$HELPER" --version >/dev/null 2>&1; then
    if ! swiftc -O -swift-version 5 "$ROOT/helper/OutlookAX.swift" -o "$HELPER" >/dev/null 2>&1; then
      echo "axlook: the bundled helper cannot run and building it needs the Xcode Command Line Tools (xcode-select --install)" >&2
      exit 1
    fi
  fi
fi
export AXLOOK_HELPER="$HELPER"

if command -v node >/dev/null 2>&1; then
  exec node "$ROOT/dist/axlook-mcp.js"
elif command -v bun >/dev/null 2>&1; then
  exec bun "$ROOT/dist/axlook-mcp.js"
fi
echo "axlook: needs Node.js 18+ or Bun on PATH" >&2
exit 1
