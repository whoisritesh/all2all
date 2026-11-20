#!/usr/bin/env bash
set -e
cd "$(dirname "$0")/backend"

if [ ! -d ".venv" ]; then
  python3 -m venv .venv
fi

source .venv/bin/activate

if [ ! -f ".venv/.installed" ] || [ "$1" = "--update" ]; then
  pip install -r requirements.txt
  touch .venv/.installed
fi

echo "Starting all2all at http://127.0.0.1:8000 ..."
uvicorn main:app --reload --host 0.0.0.0 --port 8000
