#!/bin/sh
set -eu

if ! command -v node >/dev/null 2>&1; then
  echo "Node.js not found."

  if command -v brew >/dev/null 2>&1; then
    answer="Y"
    if [ -r /dev/tty ]; then
      printf "Install Node.js via Homebrew now? [Y/n] " > /dev/tty
      read -r answer < /dev/tty || answer="Y"
    fi

    case "${answer:-Y}" in
      y|Y|yes|YES)
        brew install node
        ;;
      *)
        echo "Please install Node.js from https://nodejs.org and re-run this script."
        exit 1
        ;;
    esac
  else
    echo "Please install Node.js from https://nodejs.org and re-run this script."
    exit 1
  fi
fi

exec npx pi-for-excel-proxy "$@"
