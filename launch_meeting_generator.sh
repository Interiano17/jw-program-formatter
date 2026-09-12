#!/usr/bin/env bash

set -euo pipefail

project_dir="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
cd "$project_dir"

if [[ -x "$project_dir/.venv-app/bin/python" ]]; then
    python_executable="$project_dir/.venv-app/bin/python"
elif [[ -x "$project_dir/.venv/bin/python" ]]; then
    python_executable="$project_dir/.venv/bin/python"
elif [[ -x "$project_dir/.venv/Scripts/python.exe" ]]; then
    python_executable="$project_dir/.venv/Scripts/python.exe"
elif command -v python3 >/dev/null 2>&1; then
    python_executable="$(command -v python3)"
else
    printf 'Error: no se encontró Python 3 ni un entorno virtual .venv.\n' >&2
    exit 1
fi

exec "$python_executable" -m meeting_generator
