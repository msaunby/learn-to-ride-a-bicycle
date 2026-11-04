# Slides

Each lesson in `lessons/` has a slide summary in `slides/marp/`, written in Markdown. The summary leaves out tutor notes and quiz answers, so it is safe to build into decks for students. Start a new one with `scripts/new_lesson.py` (see the main [README](../README.md)), which copies [templates/slides.md](../templates/slides.md).

## How to build the slides

```bash
slides/marp/build.sh                        # every deck
slides/marp/build.sh 01-*.md                # one deck
slides/marp/build.sh templates/slides.md    # any slide file, e.g. to preview the template
```

This produces:

- `slides/NN-name.pptx`: an editable PowerPoint deck built by `slides/build_pptx.py` with native text boxes, shapes and tables. It needs [uv](https://docs.astral.sh/uv/), which installs python-pptx automatically. Decks are git-ignored and published as release assets (see `AGENTS.md`).
- `slides/output/NN-name.pdf`: a quick preview built with [Marp](https://marp.app/). PDFs are git-ignored, and this step is skipped if Marp is not installed.

The layout follows [style-guide.md](style-guide.md). Text is measured with the course fonts when they are installed (Windows, macOS, or WSL with Windows fonts), so text wraps as it will in PowerPoint. The script prints a warning if a slide may overflow.

## Changing the look

Colours, fonts and the deck author are set in [course.toml](../course.toml). Every build regenerates the Marp theme, `slides/marp/<theme name>.css`, from them and [theme-base.css](theme-base.css). Commit the generated theme so that Marp previews in an editor (such as the Marp for VS Code extension) match the PDFs, but don't edit it by hand.

To change a layout, edit `theme-base.css` (for the PDF preview) and the matching function in `build_pptx.py` (for the PPTX decks).

## Writing slide markdown

Slides are separated by `---`. Choose a layout by putting a class comment at the top of a slide:

| Class | Layout | Markdown |
|---|---|---|
| `title` | Dark title slide | `# Title`, `## Programme label`, then paragraphs for the lesson number and institution |
| (none) | Light content slide | Headings, paragraphs, bullets, tables, `>` callout panels |
| `cards` | Grid of white cards with numbered badges | A `-` list; each item is `**Title**` with the description on an indented next line. Add `cols-3` for three columns and `plain` for coloured top bars instead of numbers |
| `steps` | Numbered rows with accent badges | A `1.` list in the same title/description form. A `>` callout becomes a side panel. Add `teal` for boxed rows with secondary-colour badges |
| `compare` | Two columns: ✓ and ✗ | `### Strong at` + list, then `### Weak at` + list (any headings work) |
| `exercises` | Exercise cards with time pills | `### Exercise 1.1: Title *7 min*`, then a one-paragraph description |
| `dark` | Dark emphasis slide | A `>` callout and a `1.` list of steps |
| `takeaway` | Dark closing slide | A `1.` list in title/description form |

Every deck ends with the same four slides: Self-Assessment Quiz, Discussion (`steps teal`), Key Takeaways (`takeaway`) and Next Lesson (`dark`). The last lesson replaces Next Lesson with a "Keep Going" takeaway slide. [templates/slides.md](../templates/slides.md) shows every class.

For example:

```markdown
<!-- _class: cards -->

# Learning Outcomes

- **How it works**
  Explain in plain terms how the method produces its results
```

Put an image on a line of its own as `![alt text](images/name.png)`, with the path relative to the slide markdown. The image is scaled to fit the space left on the slide, and the alt text becomes its description in PowerPoint. Marp size keywords such as `h:540` in the alt text are used by the PDF preview and ignored in the PPTX.

Diagrams written in [Mermaid](https://mermaid.js.org/) are rendered to PNG once and committed, so the build doesn't need Mermaid. Keep the source (`.mmd`) next to the image, and render it with the course colours (set them in `images/mermaid-config.json` to match `course.toml`):

```bash
cd slides/marp/images
npx -y @mermaid-js/mermaid-cli@11.4.2 -i diagram.mmd -o diagram.png \
  -c mermaid-config.json -w 1600 -b white -s 2
```

Wrap a paragraph in `<mark>…</mark>` to show it as a lead panel. Any other `<!-- comment -->` becomes a speaker note. Remove speaker notes before release if they contain tutor notes.

## Updating slides with AI assistance

When a lesson changes, ask an AI assistant to update the slide summary, for example:

```
Update slides/marp/03-name.md to match lessons/03-name.md.
Keep the slide classes described in slides/README.md, leave out tutor notes and quiz answers,
keep to 6 or fewer items per slide, then run slides/marp/build.sh and fix any overflow warnings.
```
