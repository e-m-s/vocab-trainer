#!/bin/bash
set -e
cd "$(dirname "$0")"

if [ ! -d .venv ]; then
    python3 -m venv .venv
    .venv/bin/pip install -q --upgrade pip
    .venv/bin/pip install -q -r requirements.txt
fi

exec .venv/bin/python3 vocab_trainer.py
