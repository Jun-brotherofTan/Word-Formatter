#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

install_python() {
  echo "Python3 not found. Attempting to install Python3..."

  if command -v apt >/dev/null 2>&1; then
    sudo apt update
    sudo apt install -y python3 python3-venv
    return
  fi

  if command -v dnf >/dev/null 2>&1; then
    sudo dnf install -y python3
    return
  fi

  if command -v yum >/dev/null 2>&1; then
    sudo yum install -y python3
    return
  fi

  if command -v pacman >/dev/null 2>&1; then
    sudo pacman -Sy --noconfirm python
    return
  fi

  if command -v zypper >/dev/null 2>&1; then
    sudo zypper install -y python3
    return
  fi

  if command -v brew >/dev/null 2>&1; then
    brew install python
    return
  fi

  echo "No supported package manager found."
  echo "Install Python 3 manually, then re-run this script."
  exit 1
}

if ! command -v python3 >/dev/null 2>&1; then
  install_python
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo "Python3 installation failed."
  exit 1
fi

exec python3 "$SCRIPT_DIR/run_app.py" "$@"
