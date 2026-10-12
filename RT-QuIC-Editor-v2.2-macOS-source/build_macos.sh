#!/bin/zsh
set -euo pipefail

RTQUIC_ROOT="${0:A:h}"
RTQUIC_PYTHON="${PYTHON:-python3}"
RTQUIC_VENV="$RTQUIC_ROOT/.venv-build"
RTQUIC_RELEASE="$RTQUIC_ROOT/release"

cd "$RTQUIC_ROOT"

echo "[1/5] Creating the isolated build environment..."
if [[ ! -x "$RTQUIC_VENV/bin/python" ]]; then
  "$RTQUIC_PYTHON" -m venv "$RTQUIC_VENV"
fi

echo "[2/5] Installing build dependencies..."
"$RTQUIC_VENV/bin/python" -m pip install --upgrade pip
"$RTQUIC_VENV/bin/python" -m pip install -r requirements-macos.txt

echo "[3/5] Building RT-QuIC Editor.app..."
"$RTQUIC_VENV/bin/python" -m PyInstaller --noconfirm --clean RT-QuIC-Editor-macOS.spec

echo "[4/5] Verifying the application bundle..."
codesign --verify --deep --strict "$RTQUIC_ROOT/dist/RT-QuIC-Editor.app"

echo "[5/5] Creating the compressed DMG..."
mkdir -p "$RTQUIC_RELEASE"
RTQUIC_ARCH="$(uname -m)"
if [[ "$RTQUIC_ARCH" == "arm64" ]]; then
  RTQUIC_ARCH_LABEL="Apple-Silicon"
else
  RTQUIC_ARCH_LABEL="Intel"
fi
RTQUIC_OUTPUT="$RTQUIC_RELEASE/RT-QuIC-Editor-v2.2-macOS-$RTQUIC_ARCH_LABEL.dmg"
RTQUIC_STAGE="$(mktemp -d /tmp/rtquic-dmg.XXXXXX)"
trap 'rm -rf "$RTQUIC_STAGE"' EXIT

ditto "$RTQUIC_ROOT/dist/RT-QuIC-Editor.app" "$RTQUIC_STAGE/RT-QuIC-Editor.app"
ln -s /Applications "$RTQUIC_STAGE/Applications"
rm -f "$RTQUIC_OUTPUT"
hdiutil create \
  -volname "RT-QuIC Editor" \
  -srcfolder "$RTQUIC_STAGE" \
  -ov \
  -format UDZO \
  "$RTQUIC_OUTPUT"

echo "Build complete:"
echo "$RTQUIC_OUTPUT"
shasum -a 256 "$RTQUIC_OUTPUT"
