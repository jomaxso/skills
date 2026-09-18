---
name: deck-as-code
description: >
  Use when building or iterating on a PowerPoint presentation from a script instead of by hand.
  Generates .pptx and .pdf from a single PowerShell build script driving PowerPoint COM, on a
  design-token system with reusable shape primitives. Covers slide layout maths, status tables,
  flow diagrams, chapter dividers, SVG-generated map and chart assets, and a build-render-inspect
  verification loop. Use for stakeholder decks, status reports, architecture and roadmap slides
  that will be revised many times.
license: MIT
---

# Presentation decks as code

## Overview

Build the deck from **one PowerShell script** that drives PowerPoint through COM. The script is
the source of truth. The `.pptx` and `.pdf` are disposable build output, deleted and regenerated
on every run.

This pays off the moment a deck gets revised more than twice. A stakeholder deck typically absorbs
dozens of small content changes — a row reordered, a status changed, a box moved. Done by hand,
each one risks breaking alignment somewhere else. Done from a script with shared design tokens,
spacing and colour stay correct by construction, and every change is diffable.

**Use this skill when:**

- Building a deck that will be revised repeatedly (status reports, steering decks, roadmaps)
- Consistency across 20+ slides matters more than per-slide hand-tuning
- The content is structured — tables, flows, timelines, architecture diagrams
- The deck must be reproducible or reviewable as text

**Do not use it when:**

- It is a one-off single slide, or the user wants to keep editing in PowerPoint afterwards
- The deck is mostly photos and prose with little structure
- PowerPoint is not installed (COM requires the real application, Windows only)

**Prerequisites:** Windows, Microsoft PowerPoint installed, Windows PowerShell 5.1 or PowerShell 7.

## The non-negotiable rule: look at every slide you change

**Never trust an edit you have not seen rendered.** COM fails silently in ways that never raise an
error: text overflows its box, a white card lands on a near-white background and disappears, a
shape sits behind another, spacing collapses when a row is added.

Every single change follows this loop:

```
edit the script  ->  build  ->  export the affected slides to PNG  ->  VIEW the images  ->  export PDF
```

Use `scripts/render-slides.ps1` to do the middle three steps in one call:

```powershell
.\render-slides.ps1 -Script .\build-deck.ps1 -Slides 16,24 -Pdf
```

Then open the PNGs and actually look at them. In practice this catches a real problem
roughly one time in three — far too often to skip. Delete the `shots\` folder afterwards.

## Getting started

Copy `scripts/new-deck.ps1` into the working folder, rename it `build-deck.ps1`, and start
replacing the example slide blocks. It is a complete, runnable deck: design tokens, the full
primitive library, chapter dividers, and one worked example of each slide archetype.

```powershell
Copy-Item .\new-deck.ps1 .\build-deck.ps1
.\build-deck.ps1
```

## How the script is organised

Keep this order. It is what makes the file navigable once it passes 1,000 lines.

1. **Design tokens** — canvas size, the horizontal band, vertical rhythm, colour palette, font
2. **Primitives** — `Add-Txt`, `Add-Rect`, `Add-Line`, `Add-Oval`, `Add-Pill`, `Add-Head`, ...
3. **Chapter data + `Add-Divider`** — dividers generated from a `$CHAPTERS` array
4. **Build section** — COM startup, then one block per slide, then save, inside `try/finally`

Every slide is a numbered, commented block:

```powershell
# ===== 16. New QSPLM workflow =====
$s = New-Slide $pres
Add-Chrome $s 'New QSPLM'
Add-Head $s 'New QSPLM' 'A familiar mechanical workflow' 'From finding the product to the ERP.'
# ...slide body...
```

Those `# ===== N. Title =====` markers are load-bearing. They are how you locate a slide to edit,
and how you scope a string replacement so it cannot hit the wrong slide. Renumber with:

```powershell
$n = 0
$t = [regex]::Replace($t, '(?m)^# ===== \d+\. ', { param($m) $script:n++; "# ===== $script:n. " })
```

See `references/PRIMITIVES.md` for the full token set and primitive signatures.

## Editing an existing deck script

Content changes arrive as prose — "move that row up", "mark it in progress". Translating them into
safe edits is most of the work.

**Identify the slide by its title, never by the page number the user quotes.** Footer numbers do
not match slide indices whenever some slides skip the footer (title, dividers, interstitials).

**Guard every replacement with an occurrence count.** `.Replace()` returns the string unchanged
when nothing matches, so a typo in the search string silently does nothing:

```powershell
$n = ([regex]::Matches($t, [regex]::Escape($p))).Count
if ($n -ne 1) { throw "occurrences = $n for: $p" }
```

**Scope the replacement to one slide block** when the string could appear on several slides:

```powershell
$m = [regex]::Match($t, '(?ms)^# ===== 25\..*?(?=^# ===== 26\.)')
$blk = $m.Value
# ...guarded swaps inside $blk...
$t = $t.Remove($m.Index, $m.Length).Insert($m.Index, $blk)
```

**Replace whole arrays rather than editing rows one by one.** Reordering or restatusing table rows
is far safer as a wholesale swap of the `@( ... )` literal:

```powershell
$m = [regex]::Match($t, '(?ms)\$rmRows = @\(\r?\n.*?\r?\n\)')
$t = $t.Remove($m.Index, $m.Length).Insert($m.Index, $newArrayText)
```

`references/PITFALLS.md` documents the failure modes behind each of these rules, plus the
character-encoding and variable-shadowing traps that produce corrupted or silently broken output.

## Layout maths

`Add-Txt` sets `AutoSize = 0`, so boxes never grow. **Oversized text overflows silently** — it will
not error and may not even be obvious in the PNG if it overlaps something pale. Estimate before
writing:

```
characters per line  ~=  width_inches / (0.5 * font_pt / 72)
```

A 3.45" column at 11.5pt holds roughly 43 characters. Check any string you add against its column.

For an evenly-pitched table of rows with pitch `h`:

- a 0.22"-high text box sits at `row_y + ((h / 2) - 0.11)`
- a 0.20"-high box at `row_y + ((h / 2) - 0.10)`
- `Add-Oval` takes a **centre**, so its centre is `row_y + (h / 2)`

When rows are added to a table, everything below must be retuned together: pitch, the per-column
vertical offsets, and the y of any closing strip. Work backwards from the footer rule — the table
plus its strip must finish above it. Reduce pitch first, then the offsets, then the font size.

## Slide archetypes that work

These are the patterns that survived contact with a real stakeholder audience.

**Status table.** Four columns: item, planned phase, status pill, plain-language note. Colour the
pill by status and let colour carry the meaning. Group rows by status (done first) rather than by
topic — the audience wants progress, not taxonomy. Close with a tinted strip explaining any marker
used in the table.

**Flow diagram.** Numbered circles on a horizontal spine, a card under each step. Put **zone bands**
above the spine to group steps by where they happen ("in the browser", "in CATIA", "in the ERP").
Colour the arrows that cross a zone boundary and leave the within-zone arrows grey — the handovers
are the story.

**Architecture diagram.** Two or three labelled zone columns, boxes stacked inside them, connectors
between. Keep every box in the zone it truly belongs to; a box in the wrong column invites exactly
the question you do not want. Label the important connectors ("sign-in", "push").

**Chapter divider.** Full-bleed dark slide, oversized chapter number, and a progress bar of all
chapters with the current one highlighted. Generate these from a `$CHAPTERS` array so adding a
chapter updates every divider at once.

**Dark interstitial** before a live demo, stating what the audience is about to see and roughly how
long it takes.

**Prefer a drawn diagram over a screenshot placeholder.** A placeholder that never gets filled is
worse than nothing, and a diagram you draw explains the concept better than a screenshot of a UI
the audience has not learned yet. Reserve screenshots for things you cannot draw.

## Generated SVG assets

For maps, charts or anything coordinate-driven, generate an SVG with a companion script and import
it as a picture. Keep the generator separate (`build-map.ps1`) and rerun it only when the underlying
data changes.

**Force `InvariantCulture` at the top of any SVG generator.** On a German or French locale, .NET
formats `12.5` as `12,5`, which silently produces a broken SVG:

```powershell
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture
$INV = [System.Globalization.CultureInfo]::InvariantCulture
function N([double]$v) { $v.ToString('0.##', $INV) }
```

Draw only the base geography in the SVG. Add markers, labels and callouts as **native PowerPoint
shapes** on top, so they use the deck's fonts and colours and stay editable. Map SVG coordinates to
inches with a helper pinned to the picture's placement:

```powershell
$pic = $s.Shapes.AddPicture($mapSvg, $false, $true, (U $mapX), (U $mapY), (U $mapW), (U $mapH))
function MapX([double]$svgx) { $mapX + ($svgx / 1600.0) * $mapW }   # 1600 = SVG viewBox width
function MapY([double]$svgy) { $mapY + ($svgy / 760.0) * $mapH }
```

## Content integrity for status decks

When a deck reports progress, verify its claims against the source — the repository, the issue
tracker, the actual running system. Read the code rather than accepting a summary.

Two distinct risks are worth separating when reporting back:

- **The deck says something untrue.** Fix it.
- **The deck is accurate but the product cannot demonstrate it.** The deck is fine; the live demo
  is the hazard. Say so explicitly, and name the screens to avoid.

Flag any silent omission too. A feature that exists today and appears nowhere on the roadmap will
be noticed by whoever relies on it, and it is much better raised in preparation than from the floor.

## Reference

- `references/PRIMITIVES.md` — design tokens, primitive signatures, COM constants
- `references/PITFALLS.md` — encoding, variable shadowing, replacement recipes, COM quirks
- `scripts/new-deck.ps1` — runnable starter deck with every archetype
- `scripts/render-slides.ps1` — build, export slide PNGs, export PDF
