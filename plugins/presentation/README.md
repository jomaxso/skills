# presentation

Build presentation decks the way you build software: from a single scripted source of truth that
can be diffed, reviewed and regenerated.

## Install

```bash
copilot plugin install presentation@jomaxso
```

## Skills

### `deck-as-code`

Generates `.pptx` and `.pdf` from one PowerShell script driving PowerPoint COM, on a design-token
system with reusable shape primitives.

The script is the source of truth; the deck files are disposable build output. That makes a
stakeholder deck behave like a codebase — a status change is a one-line diff, spacing stays correct
by construction across every slide, and the whole deck can be regenerated from scratch at any time.

Covers:

- A design-token system (canvas grid, vertical rhythm, palette) and a primitive library
- Slide archetypes that work in front of an audience: status tables, zoned flow diagrams,
  architecture zones, chapter dividers, demo interstitials
- Layout maths — row pitch, vertical offsets, and a character budget for text that will not fit
- SVG-generated assets for maps and charts, with native PowerPoint shapes layered on top
- A mandatory build, render, and visually inspect loop, because PowerPoint COM fails silently
- The traps that produce corrupted or silently broken output: character-encoding mojibake,
  case-insensitive variable shadowing, and no-match string replacements

Ships a runnable starter deck and a render helper.

**Requires** Windows and Microsoft PowerPoint. COM automation needs the real application.

## Licence

MIT
