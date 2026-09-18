# Design tokens, primitives and COM constants

## Canvas and grid

16:9 at 13.3333" x 7.5". One horizontal band governs every slide — `$X0` left edge, `$XW` usable
width, `$X1` right edge. Nothing sits outside it except full-bleed chrome.

```powershell
$SlideW = 13.3333; $SlideH = 7.5
$X0 = 0.95; $XW = 11.45; $X1 = $X0 + $XW

$Y_EYEBROW = 0.44        # small spaced caps above the title
$Y_TITLE   = 0.70        # 27pt slide title
$Y_SUB     = 1.32        # 12.5pt one-line subtitle
$Y_BODY    = 1.98        # content start when a subtitle is present
$Y_BODY_NOSUB = 1.72     # content start without one
$Y_RULE    = 6.80        # footer hairline; all content must finish above this
$Y_FOOT    = 6.90
```

Working within a fixed vertical rhythm is what keeps 25 slides looking like one deck. Resist
nudging `$Y_TITLE` on a single slide.

## Colour

PowerPoint COM stores colour as **BGR**, not RGB. Always go through the helper:

```powershell
function RGB([int]$r, [int]$g, [int]$b) { $r + ($g -shl 8) + ($b -shl 16) }
```

A palette of this size is enough for any business deck. More colours means less meaning per colour.

```powershell
$INK    = RGB 16 42 67       # near-black body and headings; also the dark slide background
$INK2   = RGB 26 58 88       # secondary ink; cards on dark slides
$BLUE   = RGB 23 105 170     # primary accent
$TEAL   = RGB 14 158 142     # secondary accent, success, "done"
$GREEN  = RGB 74 157 82
$AMBER  = RGB 226 148 38     # attention, "in progress", pulled forward
$RED    = RGB 205 80 76
$SLATE  = RGB 82 110 138     # secondary text
$MUTE   = RGB 140 162 182    # footer text
$LINE   = RGB 213 227 238    # hairlines, inactive connectors
$BG     = RGB 246 250 253    # light slide background
$WHITE  = RGB 255 255 255
$C_BLUE = RGB 232 242 250    # tints for filled callout strips
$C_TEAL = RGB 228 245 240
$C_AMBR = RGB 253 241 224
$C_RED  = RGB 252 236 235
$DK_SUB = RGB 150 178 200    # secondary text on dark slides
```

Font: `$F = 'Aptos'`. Any single clean sans works; set it once and let the primitives apply it.

## Units

COM works in points. Author everything in inches and convert at the boundary:

```powershell
function U([double]$v) { $v * 72 }
```

## Primitives

Signatures as used throughout the template. `$align`: 1 = left, 2 = centre, 3 = right.

```powershell
Add-Txt   $slide $text $x $y $w $h $size $color [$bold] [$align] [$lineSpacing] [$spaceAfter]
Add-Rect  $slide $x $y $w $h $fill [$rounded] [$radius]
Add-Panel $slide $x $y $w $h $fill $accent          # card with a coloured left bar
Add-Line  $slide $x1 $y1 $x2 $y2 $color [$weight] [$arrow]
Add-Oval  $slide $cx $cy $diameter $fill [$ring]    # CENTRE x/y, not top-left
Add-Pill  $slide $text $x $y $w $color [$size]      # fixed 0.26" high rounded label
Add-Eyebrow $slide $text [$color]                   # letter-spaced caps at $Y_EYEBROW
Add-Head  $slide $eyebrow $title [$subtitle]        # eyebrow + title + subtitle in one call
Add-Chrome $slide $section                          # side bar, top rule, footer, page number
New-Slide $pres [$dark]
Add-Bullets $slide $items $x $y $w $dotColor [$size] [$step] [$textColor]
Add-Placeholder $slide $label $x $y $w $h $accent   # dashed screenshot frame
Add-Divider $pres $idx $title $sub                  # full chapter divider slide
```

Two behaviours to remember:

- **`Add-Txt` sets `AutoSize = 0`.** Boxes never resize. Oversized text overflows in silence.
- **`Add-Oval` takes a centre point and a diameter.** Every other primitive takes a top-left corner
  and a size. This is the single most common source of misplaced shapes.

## Letter-spaced small caps

Used for eyebrows and zone labels. There is no letter-spacing in the COM object model, so insert
the spaces:

```powershell
$spaced = ($text.ToUpper().ToCharArray() -join ' ')
```

## Page numbering

`Add-Chrome` increments `$script:PageNo`. Dark slides — title, dividers, demo interstitials,
closing — deliberately do not call it, so **the printed footer number lags the slide index**. This
is intentional: the audience counts content slides. It also means you cannot locate a slide from
the number a reviewer quotes. Find it by title.

## Chapter dividers

Drive them from one array so adding a chapter renumbers every divider and its progress bar:

```powershell
$CHAPTERS = @('Legacy system', 'Infrastructure', 'New platform', 'What comes next')
Add-Divider $pres 1 'Legacy system' 'Where we started and what we inherited.'
```

## COM constants

| Purpose | Value |
|---|---|
| Rectangle | `AddShape(1, ...)` |
| Rounded rectangle | `AddShape(5, ...)` — set corner radius via `.Adjustments.Item(1)` |
| Oval | `AddShape(9, ...)` |
| Textbox | `AddTextbox(1, ...)` (1 = horizontal) |
| Blank slide layout | `Slides.Add($index, 12)` |
| Slide size 16:9 | `PageSetup.SlideSize = 15` |
| Save as `.pptx` | `SaveAs($path, 24)` |
| Save as `.pdf` | `SaveAs($path, 32)` |
| Dashed line style | `Line.DashStyle = 4` |
| Arrowhead on | `Line.EndArrowheadStyle = 2` (1 = none) |

Turn `Shadow.Visible = 0` off on every shape. The default drop shadow instantly dates a deck.

## Startup and teardown

Always wrap the build in `try/finally` so a mid-script error cannot leave an orphaned
`POWERPNT.EXE` holding the file open.

```powershell
$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1                      # required; PowerPoint refuses to stay hidden
$pres = $app.Presentations.Add()
$pres.PageSetup.SlideSize = 15
$pres.PageSetup.SlideWidth  = U $SlideW
$pres.PageSetup.SlideHeight = U $SlideH
try {
    # ...slide blocks...
    $pres.SaveAs($target, 24)
} finally {
    if ($pres) { $pres.Close() }
    $app.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
```

## Exporting

```powershell
$slide.Export($pngPath, 'PNG', 1400, 788)     # 1400x788 renders 16:9 crisply for review
$pres.SaveAs($pdfPath, 32)
```

Reopening the saved `.pptx` read-only and calling `SaveAs(..., 32)` in the same COM session is the
most reliable way to produce the PDF.
