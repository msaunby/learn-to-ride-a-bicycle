# AGENTS.md

Instructions for anyone, human or AI agent, writing or maintaining this course. The **Course brief** describes what this course is and is the part to rewrite for each new course. The rest describes how the repository works and applies to every course built from the course kit.

## Course brief

> Replace the bracketed text in this section before drafting any lessons. Keep it short: it is the specification that the lessons are written from.

### Overview

[What the course is for, who asked for it, and who teaches it. For example: "A short course introducing X is required. It will be taught by a tutor, with a mix of lectures, workshops, exercises and self-assessment quizzes."]

### Audience and prerequisites

[Who the students are and what they already know. For example: "Assume all participants are recent graduates in a STEM discipline."]

### Lesson format

Students work in pairs or small groups, but allow for students who wish to use the material for self-study. Each lesson is 1 hour, with a short introduction and a review of the previous lesson. [Change this if the course differs, and update the segment timings in `templates/lesson.md` and `course-overview.md` to match.]

### Requirements

[The outcomes the course must cover, as a list. Each usually becomes one or two lessons.]

- [Requirement]
- [Requirement]

### Tools and policies

[Tools students will use, any approved-tool or data protection rules, and institutional policies that lessons must follow. Delete this section if there are none.]

### Course-wide threads

[Optional: anything that should run through every lesson, such as a short "spotlight" on a recurring theme in each wrap-up, or a portfolio item carried from lesson to lesson into a final capstone.]

## Repository layout

| Path | Contents |
|---|---|
| `course.toml` | Course title, programme label, institution, language, and the slide colours and fonts |
| `course-overview.md` | Aims, audience, format, lesson map, assessment, materials, maintenance and changelog |
| `ways-of-working.md` | The students' guide: how lessons run, what to bring, the portfolio |
| `lessons/NN-name.md` | One file per lesson: the full lesson plan, including tutor notes and quiz answers |
| `slides/marp/NN-name.md` | One slide summary per lesson, safe to show students |
| `slides/` | Slide tooling, style guide and README |
| `templates/` | Lesson and slide templates used by `scripts/new_lesson.py` |
| `scripts/` | `new_lesson.py` (start a lesson) and `check_spelling.sh` |
| `TODO.md` | Outstanding work, each item tagged with the version and date it was raised |

## Writing lessons

Start each lesson with `uv run scripts/new_lesson.py N "Title"`, which creates the lesson and its slide summary from the templates. Then add it to the lesson map in `course-overview.md`.

- **Keep the lesson structure.** Every lesson has the same headings as `templates/lesson.md`: learning outcomes, preparation, a timing table, the five timed segments, self-study notes and tutor notes. Tutors rely on finding things in the same place.
- **Students and tutors read the same file.** Quiz answers go in a collapsible `<details>` block, and tutor-only guidance goes under **Tutor notes**.
- **Write for self-study too.** The *Self-study notes* section adapts every paired exercise for one person.
- **Connect the lessons.** If an exercise uses an output from an earlier one, say so in both places, record it in the "How the lessons connect" table in `course-overview.md`, and suggest a fallback in the tutor notes for students who missed the earlier lesson.
- **Don't assume the gap between lessons.** Lessons may run back to back in one session, so avoid "next week", "tomorrow" or "today".
- **Leave room for tutors.** Every lesson ends with a **Your examples** section for tutors to fill in with examples from their cohort's discipline, and a **Pilot findings** section.
- **Write plainly.** Use short sentences and the spelling set by `language` in `course.toml` (British English by default: "summarise", "organise").

## Slides

Each lesson has a slide summary in `slides/marp/`, which `slides/marp/build.sh` builds into editable PPTX decks (via `slides/build_pptx.py`) and PDF previews. See `slides/README.md` for the slide classes and for how to update a summary when its lesson changes. Fix any overflow warnings the build prints.

- Slide summaries leave out tutor notes and quiz answers.
- Any `<!-- comment -->` in a slide file becomes a speaker note, so don't use comments for drafting notes.
- Colours and fonts come from `course.toml`. Don't edit the generated `slides/marp/<theme name>.css`. Edit `slides/theme-base.css` for layout changes, and make the matching change in `slides/build_pptx.py`.

## Tracking work

- `course-overview.md` starts with a status line (`DRAFT v0.x`) and ends with a changelog. Update both when the course changes.
- Record outstanding work in `TODO.md` as `- [ ] **[v0.x, YYYY-MM-DD]** description`. When an item is resolved, move it to **Done** with the date and the version in which it was resolved.
- Use commit messages starting `DRAFT v0.x:` for changes that move the course to a new draft version.

## Before committing

Check every changed file for spelling and grammar mistakes before committing, and fix them or point them out to the user:

```bash
scripts/check_spelling.sh            # markdown files changed since the last commit
scripts/check_spelling.sh FILE.md    # particular files
```

This also flags names and technical terms, so review the list rather than changing every word. Leave quoted text, such as prompts recorded exactly as they were used, unchanged.

## Releases

PowerPoint slide decks (`slides/*.pptx`) are generated from the markdown and are distributed as GitHub Release assets rather than committed to the repository. Keep the markdown as the source of truth.

Releases are published to `[OWNER/REPO]` (the `origin` remote). Always pass `--repo` explicitly, in case a clone has other remotes. Creating a release is public and outward-facing, so agents must confirm with the user before running any `gh release` command.

### Before releasing

1. Make sure `main` is clean and pushed: `git status` and `git push origin main`.
2. Build every deck with `slides/marp/build.sh` and fix any overflow warnings.
3. Check every deck for content students should not see, in particular **speaker notes containing tutor notes or quiz answers**, and for document properties (author, company) that should not be published. The author is set by `author` in `course.toml`.
4. Choose a version number. Use `v0.x` while the course is in draft, matching the `DRAFT v0.x` commit messages, and `v1.0` for the first version taught.

### Creating a release

```bash
gh release create v0.1 slides/*.pptx \
  --repo [OWNER/REPO] \
  --target main \
  --title "v0.1 — Draft course and slides" \
  --notes "Summary of what changed in the lessons and slides."
```

This creates the tag on the current `main` commit and attaches the decks as downloadable assets. Add `--draft` to create it without publishing, so it can be reviewed on GitHub first, or `--prerelease` to mark it as a draft course version.

### Updating a release

To replace or add decks on an existing release without creating a new version:

```bash
gh release upload v0.1 slides/*.pptx --repo [OWNER/REPO] --clobber
```

Prefer a new version when lesson content has changed, so each release's slides match the tagged markdown.
