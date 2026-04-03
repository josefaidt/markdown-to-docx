#!/usr/bin/env sh
set -e

REPO="josefaidt/markdown-to-docx"
BIN_NAME="markdown-to-docx"

# XDG Base Directory spec: https://specifications.freedesktop.org/basedir-spec/latest/
# Binary lives in $XDG_DATA_HOME/<bin-name>/bin/ and is symlinked into $XDG_BIN_HOME.
# Set INSTALL_DIR to override the entire install location (skips symlink).
XDG_DATA_HOME="${XDG_DATA_HOME:-$HOME/.local/share}"
XDG_BIN_HOME="${XDG_BIN_HOME:-$HOME/.local/bin}"

if [ -n "$INSTALL_DIR" ]; then
  # Manual override — install directly, no symlink
  DATA_BIN_DIR="$INSTALL_DIR"
  LINK_DIR=""
else
  DATA_BIN_DIR="${XDG_DATA_HOME}/${BIN_NAME}/bin"
  LINK_DIR="$XDG_BIN_HOME"
fi

INSTALL_PATH="${DATA_BIN_DIR}/${BIN_NAME}"
LINK_PATH="${LINK_DIR:+${LINK_DIR}/${BIN_NAME}}"

OS=$(uname -s)
ARCH=$(uname -m)

case "$OS-$ARCH" in
  Linux-x86_64)  ARTIFACT="${BIN_NAME}-linux-x64" ;;
  Linux-aarch64) ARTIFACT="${BIN_NAME}-linux-arm64" ;;
  Darwin-x86_64) ARTIFACT="${BIN_NAME}-darwin-x64" ;;
  Darwin-arm64)  ARTIFACT="${BIN_NAME}-darwin-arm64" ;;
  *)
    echo "Unsupported platform: $OS-$ARCH"
    exit 1
    ;;
esac

LATEST=$(curl -fsSL "https://api.github.com/repos/${REPO}/releases/latest" | grep '"tag_name"' | sed 's/.*"tag_name": *"\([^"]*\)".*/\1/')

if [ -z "$LATEST" ]; then
  echo "Failed to fetch latest release version"
  exit 1
fi

IS_UPDATE=0
if [ -x "$INSTALL_PATH" ]; then
  CURRENT=$("$INSTALL_PATH" --version 2>/dev/null || echo "")
  if [ "$CURRENT" = "$LATEST" ]; then
    echo "${BIN_NAME} is already up to date (${LATEST})"
    exit 0
  fi
  IS_UPDATE=1
fi

URL="https://github.com/${REPO}/releases/download/${LATEST}/${ARTIFACT}"

mkdir -p "$DATA_BIN_DIR"
TMP=$(mktemp)
curl -fsSL "$URL" -o "$TMP"
chmod +x "$TMP"
mv "$TMP" "$INSTALL_PATH"

if [ -n "$LINK_PATH" ]; then
  mkdir -p "$LINK_DIR"
  ln -sf "$INSTALL_PATH" "$LINK_PATH"
fi

DISPLAY_PATH="${LINK_PATH:-$INSTALL_PATH}"
# Shorten $HOME to ~
DISPLAY_PATH_SHORT=$(echo "$DISPLAY_PATH" | sed "s|^$HOME|~|")

echo ""
if [ "$IS_UPDATE" = "1" ]; then
  echo "✔ ${BIN_NAME} successfully updated!"
else
  echo "✔ ${BIN_NAME} successfully installed!"
fi
echo ""
echo "  Version:  ${LATEST}"
echo "  Location: ${DISPLAY_PATH_SHORT}"
echo ""
echo "  Next: Run ${BIN_NAME} --help to get started"
echo ""

case ":$PATH:" in
  *":${XDG_BIN_HOME}:"*) ;;
  *)
    echo "Note: ${XDG_BIN_HOME} is not in your PATH."
    echo "Add the following to your shell profile:"
    echo "  export PATH=\"\$PATH:${XDG_BIN_HOME}\""
    echo ""
    ;;
esac
