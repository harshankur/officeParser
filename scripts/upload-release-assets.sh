#!/usr/bin/env bash
# =============================================================================
# upload-release-assets.sh
# -----------------------------------------------------------------------------
# Fallback script to use when the GitHub Actions "Upload Release Assets" step
# fails.  Builds the browser bundles, generates the SBOM, and uploads all
# assets directly to an already-published GitHub Release using the gh CLI.
#
# Usage:
#   ./scripts/upload-release-assets.sh [TAG]
#
#   TAG  – Optional. Git tag to upload to, e.g. v7.2.3.
#          If omitted the version from package.json is used (prefixed with v).
#
# Prerequisites:
#   • Node.js + npm  (already required for the project)
#   • gh CLI         (brew install gh)
#   • gh auth login  (authenticate once; stored in .gh-config/ inside the repo)
# =============================================================================

set -euo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
GH_CFG_DIR="$REPO_ROOT/.gh-config"
ASSETS_DIR="$REPO_ROOT/release_assets"

# ── Resolve tag ───────────────────────────────────────────────────────────────
if [[ $# -ge 1 && -n "$1" ]]; then
    TAG="$1"
else
    VERSION=$(node -p "require('$REPO_ROOT/package.json').version")
    TAG="v$VERSION"
fi

echo "▶ Target release: $TAG"

# Strip leading 'v' for versioned filenames (e.g. officeparser@7.2.3.browser.mjs)
VERSION_CLEAN="${TAG#v}"

# ── Ensure gh CLI is available ────────────────────────────────────────────────
if ! command -v gh &>/dev/null; then
    echo "✖ gh CLI not found.  Install it with:  brew install gh"
    exit 1
fi

# ── Authenticate gh (use local config dir to avoid /root-owned ~/.config) ─────
# Unset any stale GITHUB_TOKEN from the environment so gh uses the stored creds.
unset GITHUB_TOKEN
export GH_CONFIG_DIR="$GH_CFG_DIR"

if ! gh auth status &>/dev/null; then
    echo "▶ gh not authenticated – starting browser login…"
    gh auth login --git-protocol https --hostname github.com --web
fi

# ── Build ─────────────────────────────────────────────────────────────────────
echo "▶ Building browser bundles…"
cd "$REPO_ROOT"
npm run build

# ── Stage assets ──────────────────────────────────────────────────────────────
echo "▶ Staging assets in release_assets/…"
mkdir -p "$ASSETS_DIR"

cp dist/officeparser.browser.iife.js      "$ASSETS_DIR/officeparser@$VERSION_CLEAN.browser.iife.js"
cp dist/officeparser.browser.mjs          "$ASSETS_DIR/officeparser@$VERSION_CLEAN.browser.mjs"
cp dist/officeparser.browser.d.ts         "$ASSETS_DIR/officeparser@$VERSION_CLEAN.browser.d.ts"
cp dist/officeparser.browser.slim.iife.js "$ASSETS_DIR/officeparser@$VERSION_CLEAN.browser.slim.iife.js"
cp dist/officeparser.browser.slim.mjs     "$ASSETS_DIR/officeparser@$VERSION_CLEAN.browser.slim.mjs"
cp dist/officeparser.browser.slim.d.ts    "$ASSETS_DIR/officeparser@$VERSION_CLEAN.browser.slim.d.ts"

# ── Generate SBOM ─────────────────────────────────────────────────────────────
echo "▶ Generating SBOM…"
npx --yes @cyclonedx/cyclonedx-npm \
    --output-format json \
    --output-file "$ASSETS_DIR/sbom.cdx.json"

# ── Upload ────────────────────────────────────────────────────────────────────
echo "▶ Uploading assets to GitHub Release $TAG…"
gh release upload "$TAG" "$ASSETS_DIR"/* \
    --clobber \
    --repo harshankur/officeParser

echo ""
echo "✔ All assets uploaded to https://github.com/harshankur/officeParser/releases/tag/$TAG"
echo ""
echo "Assets uploaded:"
for f in "$ASSETS_DIR"/*; do
    echo "  • $(basename "$f")"
done
