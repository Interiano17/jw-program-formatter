#!/usr/bin/env bash

set -euo pipefail

project_dir="/home/interiano/Projects/program_automation"
cd "$project_dir"

export PYTHONPATH="$project_dir${PYTHONPATH:+:$PYTHONPATH}"
exec "$project_dir/.venv/bin/python" -m meeting_generator.main
