# Course kit

A GitHub template for short, tutor-led courses made of one-hour lessons, each with a mini-lecture, a paired workshop, a self-assessment quiz and a discussion, plus self-study notes for learners working alone. Lessons are written in Markdown, and each has a slide summary that builds into an editable PowerPoint deck and a PDF preview.

The structure and tooling come from the [Generative AI for Researchers](https://github.com/coding-for-reproducible-research/genai-lesson-plan) course. Its repository is a complete worked example of a course built this way.

## What's included

| Path | What it is |
|---|---|
| [AGENTS.md](AGENTS.md) | The course brief (to fill in) and the conventions for writing lessons, slides and releases. AI coding agents read it automatically |
| [course.toml](course.toml) | Course title, programme label, institution, spelling language, slide colours and fonts |
| [templates/](templates/) | Lesson and slide summary templates |
| [course-overview.md](course-overview.md), [ways-of-working.md](ways-of-working.md), [TODO.md](TODO.md) | Templates for the tutor's overview, the students' guide and the work log |
| [slides/](slides/) | Slide build tools, theme and [style guide](slides/style-guide.md) |
| [scripts/](scripts/) | `new_lesson.py` to start a lesson, `check_spelling.sh` to check spelling |

## Requirements

- [uv](https://docs.astral.sh/uv/) to build PowerPoint decks and run the scripts (it installs Python packages automatically).
- Optional: [Marp CLI](https://github.com/marp-team/marp-cli) (`npm install -g @marp-team/marp-cli`) for PDF previews, `aspell` for the spelling check, and the [GitHub CLI](https://cli.github.com/) for releases.
- The slide fonts (Calibri and Georgia by default) installed locally, so that text wraps as it will in PowerPoint. They come with Windows and macOS. Under WSL, the Windows fonts are found automatically.

## Starting a new course

1. On GitHub, click **Use this template** → **Create a new repository**, then clone it.
2. Edit [course.toml](course.toml): the course title, programme label, institution and, if you like, the colours and fonts.
3. Fill in the **Course brief** at the top of [AGENTS.md](AGENTS.md), and set the release repository in its Releases section.
4. Draft the course. By hand, or with an AI agent such as Claude Code, for example:

   ```
   Read AGENTS.md. Draft course-overview.md with a lesson map that covers every requirement
   in the course brief. Then, for each lesson, run scripts/new_lesson.py and write the lesson
   and its slide summary, following the templates. Build the slides and fix any overflow warnings.
   ```

   Work through the draft lesson by lesson, rather than accepting it in one go.

5. Build the slides and check them:

   ```bash
   slides/marp/build.sh
   ```

6. Before each commit, run `scripts/check_spelling.sh`. When the course is ready to pilot, create a release, following [AGENTS.md](AGENTS.md#releases).

To start a single lesson by hand:

```bash
uv run scripts/new_lesson.py 1 "What Is This Course About?"
```

To see every slide layout, build the slide template: `slides/marp/build.sh templates/slides.md` (writes `slides/slides.pptx`).

## Updating an existing course

The tooling (`slides/build_pptx.py`, `slides/theme-base.css`, `slides/marp/build.sh` and `scripts/`) is the same in every course built from the kit. To pick up improvements, copy those files from the latest kit and rebuild. If the decks change unexpectedly, compare them with the previous release.
