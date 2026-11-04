#!/bin/sh
# Build everything that ships: a universal (arm64 + x86_64), ad-hoc signed helper and a single-file
# server bundle that runs under Node >= 18 or Bun. Output: dist/axlook-helper, dist/axlook-mcp.js
set -e
cd "$(dirname "$0")/.."
mkdir -p dist build
for arch in arm64 x86_64; do
  swiftc -O -swift-version 5 -target "$arch-apple-macos13" helper/OutlookAX.swift -o "build/axlook-helper-$arch"
done
lipo -create build/axlook-helper-arm64 build/axlook-helper-x86_64 -output dist/axlook-helper
codesign --force --sign - --identifier dev.axlook.helper dist/axlook-helper
bun build src/server.ts --target=node --format=esm --minify-syntax --outfile dist/axlook-mcp.js
# Give the bundle a node shebang so `npx axlook-mcp` / `node dist/axlook-mcp.js` work.
{ printf '#!/usr/bin/env node\n'; grep -v '^#!' dist/axlook-mcp.js; } > build/axlook-mcp.js && mv build/axlook-mcp.js dist/axlook-mcp.js
chmod +x dist/axlook-helper dist/axlook-mcp.js
echo "built dist/axlook-helper ($(lipo -archs dist/axlook-helper)) and dist/axlook-mcp.js ($(wc -c < dist/axlook-mcp.js | tr -d ' ') bytes)"
