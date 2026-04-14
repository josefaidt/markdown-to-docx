#!/usr/bin/env sh
set -e

REPO="josefaidt/markdown-to-docx"
BIN_NAME="markdown-to-docx"

# XDG Base Directory spec: https://specifications.freedesktop.org/basedir-spec/latest/
# Binary lives in $XDG_DATA_HOME/<bin-name>/bin/ and is symlinked into $XDG_BIN_HOME.
# Set INSTALL_DIR to override the entire install location (skips symlink).
XDG_DATA_HOME="${XDG_DATA_HOME:-$HOME/.local/share}"
XDG_BIN_HOME="${XDG_BIN_HOME:-$HOME/.local/bin}"

# ANSI colors — only when stdout is a terminal
if [ -t 1 ]; then
  BOLD=$(printf '\033[1m')
  GREEN=$(printf '\033[1;32m')
  CYAN=$(printf '\033[0;36m')
  YELLOW=$(printf '\033[0;33m')
  DIM=$(printf '\033[2m')
  RESET=$(printf '\033[0m')
else
  BOLD=""
  GREEN=""
  CYAN=""
  YELLOW=""
  DIM=""
  RESET=""
fi

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
    printf "Unsupported platform: %s\n" "$OS-$ARCH"
    exit 1
    ;;
esac

LATEST=$(curl -fsSL "https://api.github.com/repos/${REPO}/releases/latest" | grep '"tag_name"' | sed 's/.*"tag_name": *"\([^"]*\)".*/\1/')

if [ -z "$LATEST" ]; then
  printf "Failed to fetch latest release version\n"
  exit 1
fi

IS_UPDATE=0
if [ -x "$INSTALL_PATH" ]; then
  CURRENT=$("$INSTALL_PATH" --version 2>/dev/null || echo "")
  if [ "$CURRENT" = "$LATEST" ]; then
    printf "%s✔%s %s is already up to date (%s)\n" "$GREEN" "$RESET" "$BIN_NAME" "$LATEST"
    exit 0
  fi
  IS_UPDATE=1
fi

URL="https://github.com/${REPO}/releases/download/${LATEST}/${ARTIFACT}"

printf "Downloading %s %s...\n" "$BIN_NAME" "$LATEST"
mkdir -p "$DATA_BIN_DIR"
TMP=$(mktemp)
# Show a progress bar when stderr is a terminal, silent otherwise
if [ -t 2 ]; then
  curl -fL --progress-bar "$URL" -o "$TMP"
else
  curl -fsSL "$URL" -o "$TMP"
fi
chmod +x "$TMP"
mv "$TMP" "$INSTALL_PATH"

if [ -n "$LINK_PATH" ]; then
  mkdir -p "$LINK_DIR"
  ln -sf "$INSTALL_PATH" "$LINK_PATH"
fi

DISPLAY_PATH="${LINK_PATH:-$INSTALL_PATH}"
# Shorten $HOME to ~
DISPLAY_PATH_SHORT=$(echo "$DISPLAY_PATH" | sed "s|^$HOME|~|")

printf "\n"
if [ "$IS_UPDATE" = "1" ]; then
  printf "%s✔%s %s%s successfully updated!%s\n" "$GREEN" "$RESET" "$BOLD" "$BIN_NAME" "$RESET"
else
  printf "%s✔%s %s%s successfully installed!%s\n" "$GREEN" "$RESET" "$BOLD" "$BIN_NAME" "$RESET"
fi
printf "\n"
printf "  Version:  %s%s%s\n" "$DIM" "$LATEST" "$RESET"
printf "  Location: %s%s%s\n" "$DIM" "$DISPLAY_PATH_SHORT" "$RESET"
printf "\n"
printf "  Next: Run %s%s --help%s to get started\n" "$CYAN" "$BIN_NAME" "$RESET"
printf "\n"

case ":$PATH:" in
  *":${XDG_BIN_HOME}:"*) ;;
  *)
    printf "%sNote:%s %s is not in your PATH.\n" "$YELLOW" "$RESET" "$XDG_BIN_HOME"
    printf "Add the following to your shell profile:\n"
    printf "  %sexport PATH=\"\$PATH:%s\"%s\n" "$CYAN" "$XDG_BIN_HOME" "$RESET"
    printf "\n"
    ;;
esac
