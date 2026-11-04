#!/usr/bin/env bash
# List possible spelling mistakes in markdown files, using the dictionary set
# by `language` in course.toml.
#
#   scripts/check_spelling.sh              markdown files changed since the last commit
#   scripts/check_spelling.sh FILE.md ...  the files given
#
# Names and technical terms are flagged too, so review the list rather than
# changing every word. Needs aspell and its dictionary (e.g. aspell-en).
set -euo pipefail
root="$(cd "$(dirname "$0")/.." && pwd)"
lang="$(sed -n 's/^language *= *"\(.*\)".*/\1/p' "$root/course.toml")"

if [ "$#" -gt 0 ]; then
  files=("$@")
else
  cd "$root"
  mapfile -t files < <({ git diff --name-only HEAD; git ls-files --others --exclude-standard; } | grep '\.md$' | sort -u)
fi

for f in "${files[@]}"; do
  [ -f "$f" ] || continue
  words="$(aspell list --lang="${lang:-en_GB}" --mode=markdown < "$f" | sort -u | tr '\n' ' ')"
  [ -n "$words" ] && echo "$f: $words"
done
true
