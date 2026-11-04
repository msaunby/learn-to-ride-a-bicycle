# /// script
# requires-python = ">=3.11"
# ///
"""Start a new lesson and its slide summary from the templates.

    uv run scripts/new_lesson.py 3 "Writing in Your Own Voice"

creates lessons/03-writing-in-your-own-voice.md and
slides/marp/03-writing-in-your-own-voice.md from templates/, filled in with
the lesson number and title and the course details in course.toml. Existing
files are never overwritten. Add the lesson to the lesson map in
course-overview.md yourself.
"""

import re
import sys
import tomllib
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent


def main():
    if len(sys.argv) != 3 or not sys.argv[1].isdigit():
        sys.exit(__doc__)
    number, title = int(sys.argv[1]), sys.argv[2].strip()
    config = tomllib.loads((ROOT / "course.toml").read_text(encoding="utf-8"))
    slug = f"{number:02d}-" + re.sub(r"[^a-z0-9]+", "-", title.lower()).strip("-")
    fields = {
        "number": str(number),
        "next_number": str(number + 1),
        "title": title,
        "programme": config["course"]["programme"],
        "institution": config["course"]["institution"],
        "theme": config["theme"]["name"],
    }
    for template, target in [("lesson.md", ROOT / "lessons"), ("slides.md", ROOT / "slides" / "marp")]:
        path = target / f"{slug}.md"
        if path.exists():
            print(f"{path.relative_to(ROOT)} already exists; left unchanged")
            continue
        text = (ROOT / "templates" / template).read_text(encoding="utf-8")
        text = re.sub(r"\{\{(\w+)\}\}", lambda m: fields[m.group(1)], text)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
        print(f"created {path.relative_to(ROOT)}")


if __name__ == "__main__":
    main()
