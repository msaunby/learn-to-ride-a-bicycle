# Slide Deck Style Guide

Design reference for the lesson slide decks. `slides/build_pptx.py` and the Marp theme implement it. Colours and fonts are set in [course.toml](../course.toml) and referred to here by role. The defaults are shown in brackets.

## Slide dimensions

- Widescreen 16:9: 13.33" × 7.5" (laid out as 1280 × 720 px).

## Colour roles

| Role | Default | Usage |
|---|---|---|
| `primary` | `#003C3C` deep teal-green | Title, dark and takeaway slide backgrounds; heading text on content slides |
| `secondary` | `#00695C` teal | Table headers, card badges, callout headings |
| `accent` | `#C75B12` copper | Step badges, exercise time pills, accent bars, programme label |
| `background` | `#FAF8F5` cream | Content slide backgrounds |
| `white` | `#FFFFFF` | Cards; text on dark slides |
| `text` | `#2D2D2D` | Body text |
| `muted` | `#666666` | Description text beneath bold titles |
| `callout` | `#E0F2EF` pale mint | Callout panels ("Key reminder", "Example prompt") |
| `accent_light` | `#FBEDE3` | Time-pill backgrounds |
| `positive` / `negative` | `#2E7D32` / `#C62828` | ✓ and ✗ in comparison columns |

When choosing new colours, keep `primary` dark enough for white text, `background` and `callout` light enough for `text`, and `accent` distinct from `secondary`.

## Typography

| Element | Font | Size (approx.) | Colour |
|---|---|---|---|
| Slide title | heading, bold | 28–32 pt | `primary`; white on dark slides |
| Card or section heading | heading, bold | 18–20 pt | `primary` |
| Body text | body | 14–16 pt | `text` |
| Description text | body | 12–14 pt | `muted` |
| Table header | body, bold | 13–14 pt | white on `secondary` |
| Badge numbers | body, bold | 14–16 pt | white |
| Slide number | body | 10 pt | grey, bottom right |

Defaults: Georgia for headings and Calibri for body text. Pair one serif heading font with one sans-serif body font, and never mix in others.

## Slide types

1. **Title** (`title`): `primary` background with subtle circles; programme label in `accent` capitals, the lesson title, then the lesson number and institution.
2. **Learning outcomes** (`cards`): 2×2 grid of white cards with numbered badges.
3. **Numbered list** (`steps`): rows with `accent` numbered badges, bold title and `muted` description; optional `callout` side panel.
4. **Table** (default): `secondary` header row with white text, alternating rows.
5. **Comparison** (`compare`): two columns, ✓ in `positive` and ✗ in `negative`.
6. **Live demo** (`dark`): `primary` background, a callout with what is shown, then numbered steps.
7. **Workshop exercises** (`exercises`): cards with an `accent` bar and a time pill.
8. **Discussion** (`steps teal`): three boxed rows, each a short topic and a question.
9. **Key takeaways** (`takeaway`): `primary` background, three numbered cards. Takeaways only: the next lesson has its own slide.
10. **Next lesson** (`dark`): a callout with the next lesson's number, title and one-line summary, then what to bring or prepare. Don't say when the next lesson is: lessons may run back to back. The final lesson ends with a "Keep Going" takeaway slide instead.

## General rules

1. **Six or seven items per slide at most.** Keep text concise.
2. **One idea per slide.**
3. **Title, dark and takeaway slides use `primary` backgrounds**; all other slides use `background`.
4. **Fragments rather than full sentences in bullets** where possible.
5. **Tables always have a header row.**
6. **Decoration is subtle.** It supports the content and doesn't distract.
