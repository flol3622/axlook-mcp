#!/bin/sh
# Build the Swift Accessibility helper into bin/axlook-helper.
set -e
cd "$(dirname "$0")/.."
mkdir -p bin
swiftc -O -swift-version 5 helper/OutlookAX.swift -o bin/axlook-helper
echo "built bin/axlook-helper"
