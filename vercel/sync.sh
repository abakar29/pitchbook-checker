#!/bin/sh
# Copies the checker and sample files next to the Vercel function. Run after changing pitchbook_checker.py.
cd "$(dirname "$0")" && cp ../pitchbook_checker.py . && cp ../samples/project_atlas_* samples/
