#!/usr/bin/env bash
# Build editable PPTX and preview PDF slide decks from the markdown in this folder.
#
#   slides/marp/build.sh                         build every deck
#   slides/marp/build.sh 01-*.md                 build selected decks
#   slides/marp/build.sh templates/slides.md     build any slide markdown file
#
# PPTX: slides/build_pptx.py (needs uv), written to slides/ (git-ignored,
# attached to releases). It also regenerates the Marp theme from course.toml.
# PDF: Marp CLI (npm install -g @marp-team/marp-cli), written to slides/output/.
set -euo pipefail
shopt -s nullglob
here="$(cd "$(dirname "$0")" && pwd)"

# Arguments may be paths, or unexpanded globs naming decks in this folder.
decks=()
if [ "$#" -gt 0 ]; then
  for arg in "$@"; do
    if [ -f "$arg" ]; then
      decks+=("$(realpath "$arg")")
    else
      matches=("$here"/$(basename "$arg"))
      [ "${#matches[@]}" -gt 0 ] || { echo "No slide markdown matches $arg" >&2; exit 1; }
      decks+=("${matches[@]}")
    fi
  done
else
  decks=("$here"/[0-9]*.md)
fi
if [ "${#decks[@]}" -eq 0 ]; then
  echo "No slide markdown in $here yet: run scripts/new_lesson.py to add a lesson"
  exit 0
fi
mkdir -p "$here/../output"

uv run "$here/../build_pptx.py" "${decks[@]}"

if command -v marp >/dev/null; then
  theme="$(sed -n 's/^name *= *"\(.*\)".*/\1/p' "$here/../../course.toml")"
  for md in "${decks[@]}"; do
    name="$(basename "${md%.md}")"
    marp --theme "$here/$theme.css" --allow-local-files -o "$here/../output/$name.pdf" "$md" </dev/null
  done
else
  echo "Marp CLI not found; skipping PDF previews"
fi
