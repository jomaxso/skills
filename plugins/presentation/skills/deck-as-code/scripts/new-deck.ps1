# Starter deck. Copy to build-deck.ps1 and replace the example slide blocks.
# Keep this file pure ASCII - see references/PITFALLS.md #1.
$ErrorActionPreference = 'Stop'

# ---------- design tokens ----------
$SlideW = 13.3333; $SlideH = 7.5
$X0 = 0.95; $XW = 11.45; $X1 = $X0 + $XW

$Y_EYEBROW = 0.44; $Y_TITLE = 0.70; $Y_SUB = 1.32
$Y_BODY = 1.98; $Y_BODY_NOSUB = 1.72
$Y_RULE = 6.80; $Y_FOOT = 6.90

$F = 'Aptos'

function RGB([int]$r, [int]$g, [int]$b) { $r + ($g -shl 8) + ($b -shl 16) }
$INK    = RGB 16 42 67
$INK2   = RGB 26 58 88
$BLUE   = RGB 23 105 170
$TEAL   = RGB 14 158 142
$GREEN  = RGB 74 157 82
$AMBER  = RGB 226 148 38
$RED    = RGB 205 80 76
$SLATE  = RGB 82 110 138
$MUTE   = RGB 140 162 182
$LINE   = RGB 213 227 238
$BG     = RGB 246 250 253
$WHITE  = RGB 255 255 255
$C_BLUE = RGB 232 242 250
$C_TEAL = RGB 228 245 240
$C_AMBR = RGB 253 241 224
$C_RED  = RGB 252 236 235
$DK_SUB = RGB 150 178 200

function U([double]$v) { $v * 72 }

# ---------- deck identity ----------
$DECK_NAME = 'Example deck'
$DECK_DATE = '01.01.2026'
$OUT_FILE  = 'Example-Deck.pptx'

# ---------- primitives ----------
function Add-Txt($slide, [string]$text, [double]$x, [double]$y, [double]$w, [double]$h,
                 [double]$size, [int]$color, [bool]$bold = $false, [int]$align = 1,
                 [double]$lineSpacing = 1.0, [double]$spaceAfter = 0) {
    $t = $slide.Shapes.AddTextbox(1, (U $x), (U $y), (U $w), (U $h))
    $tf = $t.TextFrame
    $tf.MarginLeft = 0; $tf.MarginRight = 0; $tf.MarginTop = 0; $tf.MarginBottom = 0
    $tf.WordWrap = -1
    try { $tf.AutoSize = 0 } catch {}
    $tr = $tf.TextRange
    $tr.Text = $text
    $tr.Font.Name = $F
    $tr.Font.Size = $size
    $tr.Font.Bold = $bold
    $tr.Font.Color.RGB = $color
    $tr.ParagraphFormat.Alignment = $align
    if ($lineSpacing -ne 1.0) {
        $tr.ParagraphFormat.LineRuleWithin = -1
        $tr.ParagraphFormat.SpaceWithin = $lineSpacing
    }
    if ($spaceAfter -gt 0) {
        $tr.ParagraphFormat.LineRuleAfter = 0
        $tr.ParagraphFormat.SpaceAfter = $spaceAfter
    }
    $t
}

function Add-Rect($slide, [double]$x, [double]$y, [double]$w, [double]$h, [int]$fill,
                  [bool]$rounded = $true, [double]$radius = 0.04) {
    $shapeType = if ($rounded) { 5 } else { 1 }
    $s = $slide.Shapes.AddShape($shapeType, (U $x), (U $y), (U $w), (U $h))
    if ($rounded) { try { $s.Adjustments.Item(1) = $radius } catch {} }
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB = $fill
    $s.Line.Visible = 0
    $s.Shadow.Visible = 0
    $s
}

function Add-Panel($slide, [double]$x, [double]$y, [double]$w, [double]$h, [int]$fill, [int]$accent) {
    $card = Add-Rect $slide $x $y $w $h $fill $true 0.03
    $bar = Add-Rect $slide $x $y 0.055 $h $accent $false
    @{ Card = $card; Bar = $bar }
}

function Add-Line($slide, [double]$x1, [double]$y1, [double]$x2, [double]$y2,
                  [int]$color, [double]$weight = 1.25, [int]$arrow = 0) {
    $l = $slide.Shapes.AddLine((U $x1), (U $y1), (U $x2), (U $y2))
    $l.Line.ForeColor.RGB = $color
    $l.Line.Weight = $weight
    $l.Line.EndArrowheadStyle = if ($arrow -eq 1) { 2 } else { 1 }
    $l.Shadow.Visible = 0
    $l
}

# NOTE: takes the CENTRE point and a diameter, unlike every other primitive.
function Add-Oval($slide, [double]$cx, [double]$cy, [double]$d, [int]$fill, [int]$ring = -1) {
    $s = $slide.Shapes.AddShape(9, (U ($cx - $d / 2)), (U ($cy - $d / 2)), (U $d), (U $d))
    $s.Fill.Solid(); $s.Fill.ForeColor.RGB = $fill
    if ($ring -ge 0) { $s.Line.Visible = -1; $s.Line.ForeColor.RGB = $ring; $s.Line.Weight = 2.5 }
    else { $s.Line.Visible = 0 }
    $s.Shadow.Visible = 0
    $s
}

function Add-Pill($slide, [string]$text, [double]$x, [double]$y, [double]$w, [int]$color, [double]$size = 7.5) {
    $p = Add-Rect $slide $x $y $w 0.26 $color $true 0.5
    Add-Txt $slide $text $x ($y + 0.055) $w 0.16 $size $WHITE $true 2 | Out-Null
    $p
}

function Add-Eyebrow($slide, [string]$text, [int]$color = $BLUE) {
    $spaced = ($text.ToUpper().ToCharArray() -join ' ')
    Add-Txt $slide $spaced $X0 $Y_EYEBROW $XW 0.18 8.5 $color $true 1 | Out-Null
}

function Add-Head($slide, [string]$eyebrow, [string]$title, [string]$subtitle = '') {
    Add-Eyebrow $slide $eyebrow
    Add-Txt $slide $title $X0 $Y_TITLE $XW 0.52 27 $INK $true 1 | Out-Null
    if ($subtitle) { Add-Txt $slide $subtitle $X0 $Y_SUB $XW 0.30 12.5 $SLATE $false 1 | Out-Null }
}

$script:PageNo = 0
function Add-Chrome($slide, [string]$section) {
    $script:PageNo++
    Add-Rect $slide 0 0 0.30 $SlideH $INK $false | Out-Null
    Add-Rect $slide 0.30 0 ($SlideW - 0.30) 0.075 $TEAL $false | Out-Null
    Add-Line $slide $X0 $Y_RULE $X1 $Y_RULE $LINE 1 | Out-Null
    Add-Txt $slide $section.ToUpper() $X0 $Y_FOOT 6 0.16 8 $MUTE $true 1 | Out-Null
    $stamp = '{0}  |  {1}  |  {2:D2}' -f $DECK_NAME, $DECK_DATE, $script:PageNo
    Add-Txt $slide $stamp ($X1 - 4) $Y_FOOT 4 0.16 8 $MUTE $false 3 | Out-Null
}

function New-Slide($pres, [bool]$dark = $false) {
    $slide = $pres.Slides.Add($pres.Slides.Count + 1, 12)
    $slide.FollowMasterBackground = 0
    $slide.Background.Fill.Solid()
    $slide.Background.Fill.ForeColor.RGB = if ($dark) { $INK } else { $BG }
    $slide
}

function Add-Bullets($slide, [string[]]$items, [double]$x, [double]$y, [double]$w,
                     [int]$dotColor, [double]$size = 11, [double]$step = 0.30, [int]$textColor = $INK2) {
    for ($i = 0; $i -lt $items.Count; $i++) {
        $ry = $y + ($i * $step)
        Add-Oval $slide ($x + 0.055) ($ry + 0.095) 0.075 $dotColor | Out-Null
        Add-Txt $slide $items[$i] ($x + 0.20) $ry ($w - 0.20) 0.26 $size $textColor $false 1 | Out-Null
    }
}

function Add-Placeholder($slide, [string]$label, [double]$x, [double]$y, [double]$w, [double]$h, [int]$accent) {
    Add-Rect $slide $x $y $w $h $WHITE $true 0.03 | Out-Null
    $frame = $slide.Shapes.AddShape(5, (U $x), (U $y), (U $w), (U $h))
    try { $frame.Adjustments.Item(1) = 0.03 } catch {}
    $frame.Fill.Visible = 0
    $frame.Line.ForeColor.RGB = $LINE; $frame.Line.Weight = 1; $frame.Line.DashStyle = 4
    $frame.Shadow.Visible = 0
    Add-Rect $slide $x $y $w 0.10 $accent $false | Out-Null
    $cy = $y + ($h / 2)
    Add-Txt $slide $label $x ($cy + 0.16) $w 0.20 10.5 $INK $true 2 | Out-Null
    Add-Txt $slide 'screenshot to be inserted' $x ($cy + 0.40) $w 0.18 8.5 $MUTE $false 2 | Out-Null
}

$CHAPTERS = @('Where we started', 'Where we are', 'What comes next')

function Add-Divider($pres, [int]$idx, [string]$title, [string]$sub) {
    $slide = New-Slide $pres $true
    Add-Rect $slide 0.30 0 ($SlideW - 0.30) 0.075 $TEAL $false | Out-Null
    Add-Rect $slide 0 0 0.30 $SlideH $TEAL $false | Out-Null
    Add-Txt $slide ('{0:D2}' -f $idx) 7.90 1.80 4.50 3.20 170 $INK2 $true 3 | Out-Null
    $eyebrow = ('CHAPTER {0:D2} OF {1:D2}' -f $idx, $CHAPTERS.Count).ToCharArray() -join ' '
    Add-Txt $slide $eyebrow $X0 2.30 6.90 0.22 10 (RGB 63 182 222) $true 1 | Out-Null
    Add-Txt $slide $title $X0 2.64 6.90 1.10 34 $WHITE $true 1 1.1 | Out-Null
    Add-Rect $slide $X0 3.88 1.60 0.055 $TEAL $false | Out-Null
    Add-Txt $slide $sub $X0 4.14 6.90 0.90 14 $DK_SUB $false 1 1.28 | Out-Null
    $segGap = 0.20
    $segW = ($XW - ($CHAPTERS.Count - 1) * $segGap) / $CHAPTERS.Count
    for ($i = 0; $i -lt $CHAPTERS.Count; $i++) {
        $sx = $X0 + $i * ($segW + $segGap)
        $on = (($i + 1) -eq $idx)
        $barC = if ($on) { $TEAL } else { $INK2 }
        $txtC = if ($on) { $WHITE } else { RGB 88 120 150 }
        Add-Rect $slide $sx 5.94 $segW 0.06 $barC $false | Out-Null
        Add-Txt $slide $CHAPTERS[$i] $sx 6.10 $segW 0.22 9.5 $txtC $on 1 | Out-Null
    }
    $slide
}

# ---------- shared column widths ----------
# Never reassign these inside a slide block - see references/PITFALLS.md #2.
$c2w = ($XW - 0.40) / 2
$c3w = ($XW - 0.80) / 3
$rx  = $X0 + $c2w + 0.40

# ---------- build ----------
$root = $PSScriptRoot
$target = Join-Path $root $OUT_FILE
if (Test-Path $target) { Remove-Item -LiteralPath $target -Force }

$app = New-Object -ComObject PowerPoint.Application
$app.Visible = -1
$pres = $app.Presentations.Add()
$pres.PageSetup.SlideSize = 15
$pres.PageSetup.SlideWidth = U $SlideW
$pres.PageSetup.SlideHeight = U $SlideH

try {

# ===== 1. Title =====
$s = New-Slide $pres $true
Add-Rect $s 9.55 0 ($SlideW - 9.55) $SlideH $INK2 $false | Out-Null
Add-Oval $s 11.85 6.05 2.30 $TEAL | Out-Null
Add-Txt $s 'E X A M P L E   D E C K' $X0 1.28 7.5 0.20 9 (RGB 63 182 222) $true 1 | Out-Null
Add-Txt $s "Where we started.`rWhere we are now.`rWhere we are going." $X0 1.75 8.2 2.30 34 $WHITE $true 1 1.12 | Out-Null
Add-Txt $s "Legacy system  $([char]0x00B7)  infrastructure  $([char]0x00B7)  new platform" $X0 4.42 8.2 0.30 13 $DK_SUB $false 1 | Out-Null
Add-Rect $s $X0 5.85 1.55 0.035 $TEAL $false | Out-Null
Add-Txt $s "Stakeholder presentation  |  $DECK_DATE" $X0 6.12 6.5 0.24 11.5 $WHITE $false 1 | Out-Null

# ===== 2. Chapter divider =====
Add-Divider $pres 1 'Where we started' 'The system we inherited and the constraints it put on us.' | Out-Null

# ===== 3. Two-column content =====
$s = New-Slide $pres
Add-Chrome $s 'Where we started'
Add-Head $s 'Where we started' 'The system we inherited' 'One sentence framing what this slide proves.'
$leftItems  = @('First point in plain language', 'Second point', 'Third point', 'Fourth point')
$rightItems = @('Consequence one', 'Consequence two', 'Consequence three')
Add-Panel $s $X0 $Y_BODY $c2w 2.60 $WHITE $BLUE | Out-Null
Add-Txt $s 'What existed' ($X0 + 0.30) ($Y_BODY + 0.24) ($c2w - 0.60) 0.30 14 $INK $true 1 | Out-Null
Add-Bullets $s $leftItems ($X0 + 0.30) ($Y_BODY + 0.70) ($c2w - 0.60) $BLUE
Add-Panel $s $rx $Y_BODY $c2w 2.60 $WHITE $AMBER | Out-Null
Add-Txt $s 'What it cost us' ($rx + 0.30) ($Y_BODY + 0.24) ($c2w - 0.60) 0.30 14 $INK $true 1 | Out-Null
Add-Bullets $s $rightItems ($rx + 0.30) ($Y_BODY + 0.70) ($c2w - 0.60) $AMBER
Add-Rect $s $X0 5.20 $XW 0.82 $C_BLUE $true 0.06 | Out-Null
Add-Txt $s 'The one sentence you want the room to remember from this slide.' $X0 5.46 $XW 0.30 12 $BLUE $true 2 | Out-Null

# ===== 4. Status table =====
$s = New-Slide $pres
Add-Chrome $s 'Where we are'
Add-Head $s 'Roadmap' 'Where the roadmap stands today' 'What we planned, what is built, and what is still ahead.'
$up = [char]0x2191
# Group rows by status: done first, then in progress. Colour carries the meaning.
$rmRows = @(
    @('Delivered item one',   'Phase 1',       'DONE',        'Live today', $TEAL),
    @('Delivered item two',   'Phase 1',       'DONE',        'Live today', $TEAL),
    @('Item under way',       'Phase 1',       'IN PROGRESS', 'Build started', $BLUE),
    @('Pulled forward item',  "Phase 3$up",    'IN PROGRESS', 'Started early because users need it', $BLUE),
    @('Item being reworked',  'Phase 1',       'REDESIGN',    'Works today, being simplified', $AMBER)
)
$rmCol = @(3.45, 1.35, 1.28)
$rmHeads = @('SCOPE ITEM', 'PLANNED FOR', 'STATUS', 'WHERE WE ARE')
$rmX = @($X0, ($X0 + $rmCol[0]), ($X0 + $rmCol[0] + $rmCol[1]), ($X0 + $rmCol[0] + $rmCol[1] + $rmCol[2]))
for ($i = 0; $i -lt 4; $i++) {
    Add-Txt $s (($rmHeads[$i].ToCharArray() -join ' ')) $rmX[$i] 1.94 3.0 0.18 8 $SLATE $true 1 | Out-Null
}
Add-Line $s $X0 2.22 $X1 2.22 $LINE 1.25 | Out-Null
$rmPitch = 0.44
for ($i = 0; $i -lt $rmRows.Count; $i++) {
    $ry = 2.32 + $i * $rmPitch
    Add-Txt $s $rmRows[$i][0] $rmX[0] ($ry + 0.13) $rmCol[0] 0.26 11.5 $INK $true 1 | Out-Null
    $phC = if ($rmRows[$i][1] -match [regex]::Escape($up)) { $AMBER } else { $SLATE }
    Add-Txt $s $rmRows[$i][1] $rmX[1] ($ry + 0.14) $rmCol[1] 0.24 10.5 $phC $false 1 | Out-Null
    Add-Pill $s $rmRows[$i][2] $rmX[2] ($ry + 0.09) 1.15 $rmRows[$i][4] | Out-Null
    Add-Txt $s $rmRows[$i][3] $rmX[3] ($ry + 0.14) ($X1 - $rmX[3]) 0.24 10 $INK2 $false 1 | Out-Null
}
Add-Rect $s $X0 5.96 $XW 0.68 $C_AMBR $true 0.06 | Out-Null
Add-Txt $s "$up  Pulled forward from a later phase, because it helps users now." $X0 6.17 $XW 0.28 12 $AMBER $true 2 | Out-Null

# ===== 5. Flow with zone bands =====
$s = New-Slide $pres
Add-Chrome $s 'Where we are'
Add-Head $s 'Where we are' 'A familiar workflow' 'From finding the record to pushing the result downstream.'
$wfW = 1.75; $wfGap = 0.19; $wfPitch = $wfW + $wfGap
$wfSteps = @(
    @('Find',   'Search by name or number',        $BLUE),
    @('See',    'All the data in one view',        $BLUE),
    @('Open',   'One click starts the tool',       $TEAL),
    @('Change', 'Work exactly as you do today',    $TEAL),
    @('Save',   'Back as a new revision',          $TEAL),
    @('Push',   'Data goes to the ERP',            $AMBER)
)
# start index, span, tint, ink
$wfZones = @(
    @('IN THE BROWSER', 0, 2, $C_BLUE, $BLUE),
    @('IN THE TOOL',    2, 3, $C_TEAL, $TEAL),
    @('DOWNSTREAM',     5, 1, $C_AMBR, $AMBER)
)
foreach ($z in $wfZones) {
    $zx = $X0 + $z[1] * $wfPitch
    $zw = $z[2] * $wfW + ($z[2] - 1) * $wfGap
    Add-Rect $s $zx 2.45 $zw 0.34 $z[3] $true 0.10 | Out-Null
    Add-Txt $s (($z[0].ToCharArray() -join ' ')) $zx 2.53 $zw 0.18 8.5 $z[4] $true 2 | Out-Null
}
for ($i = 0; $i -lt $wfSteps.Count; $i++) {
    $wfX = $X0 + $i * $wfPitch
    $wfCx = $wfX + $wfW / 2
    $wfCol = $wfSteps[$i][2]
    Add-Rect $s $wfX 3.57 $wfW 1.45 $WHITE $true 0.06 | Out-Null
    Add-Rect $s $wfX 3.57 $wfW 0.05 $wfCol $false | Out-Null
    Add-Txt $s $wfSteps[$i][0] $wfX 3.81 $wfW 0.32 16 $INK $true 2 | Out-Null
    Add-Txt $s $wfSteps[$i][1] ($wfX + 0.12) 4.27 ($wfW - 0.24) 0.62 10.5 $SLATE $false 2 1.2 | Out-Null
    Add-Oval $s $wfCx 3.25 0.46 $wfCol | Out-Null
    Add-Txt $s ([string]($i + 1)) ($wfCx - 0.23) 3.14 0.46 0.22 11.5 $WHITE $true 2 | Out-Null
    if ($i -lt $wfSteps.Count - 1) {
        # Colour only the arrows that cross a zone boundary - those are the story.
        $wfNext = $wfSteps[$i + 1][2]
        $wfArrow = if ($wfNext -ne $wfCol) { $wfNext } else { $LINE }
        Add-Line $s ($wfCx + 0.31) 3.25 ($wfCx + $wfPitch - 0.31) 3.25 $wfArrow 1.75 1 | Out-Null
    }
}
Add-Rect $s $X0 5.60 $XW 0.82 $C_TEAL $true 0.06 | Out-Null
Add-Txt $s 'One record is the single source - nothing is re-typed downstream.' $X0 5.86 $XW 0.30 12 $TEAL $true 2 | Out-Null

# ===== 6. Architecture zones =====
$s = New-Slide $pres
Add-Chrome $s 'Where we are'
Add-Head $s 'Where we are' 'How the new system is built' 'Three worlds, and the connections between them.'
# label, x, width
$arZones = @(
    @('ON THE USER PC', 0.95, 3.70),
    @('CLOUD',          5.15, 2.95),
    @('INTEGRATIONS',   8.55, 3.85)
)
foreach ($z in $arZones) {
    # Height must clear the last box row (4.70 + 0.74) with padding, or boxes spill out.
    Add-Rect $s $z[1] 2.05 $z[2] 3.55 $WHITE $true 0.03 | Out-Null
    Add-Rect $s $z[1] 2.05 $z[2] 0.05 $BLUE $false | Out-Null
    Add-Txt $s (($z[0].ToCharArray() -join ' ')) $z[1] 2.22 $z[2] 0.18 8.5 $SLATE $true 2 | Out-Null
}
# label, x, y, width
$arBoxes = @(
    @('Browser client',  1.15, 2.66, 3.30, $BLUE),
    @('Local agent',     1.15, 3.68, 3.30, $BLUE),
    @('Desktop tool',    1.15, 4.70, 3.30, $SLATE),
    @('Web application', 5.37, 2.66, 2.51, $TEAL),
    @('Web API',         5.37, 3.68, 2.51, $TEAL),
    @('Database',        5.37, 4.70, 2.51, $TEAL),
    @('Identity',        8.77, 2.66, 3.41, $BLUE),
    @('ERP',             8.77, 3.68, 3.41, $AMBER)
)
foreach ($b in $arBoxes) {
    Add-Rect $s $b[1] $b[2] $b[3] 0.74 $BG $true 0.05 | Out-Null
    Add-Txt $s $b[0] $b[1] ($b[2] + 0.25) $b[3] 0.26 11.5 $b[4] $true 2 | Out-Null
}
Add-Line $s 2.80 3.42 2.80 3.66 $LINE 1.5 1 | Out-Null
Add-Line $s 2.80 4.44 2.80 4.68 $LINE 1.5 1 | Out-Null
Add-Line $s 6.625 3.42 6.625 3.66 $LINE 1.5 1 | Out-Null
Add-Line $s 6.625 4.44 6.625 4.68 $LINE 1.5 1 | Out-Null
Add-Line $s 4.43 3.03 5.35 3.03 $TEAL 1.75 1 | Out-Null
# Connector labels must be one short word - the gap between zones will not hold more.
Add-Txt $s 'HTTPS' 4.40 2.80 0.95 0.18 8 $TEAL $false 2 | Out-Null
Add-Line $s 7.90 4.05 8.77 4.05 $AMBER 1.75 1 | Out-Null
Add-Txt $s 'push' 7.89 3.82 0.90 0.18 8 $AMBER $false 2 | Out-Null

# ===== 7. Dark demo interstitial =====
$s = New-Slide $pres $true
Add-Rect $s 0.30 0 ($SlideW - 0.30) 0.075 $BLUE $false | Out-Null
Add-Rect $s 0 0 0.30 $SlideH $BLUE $false | Out-Null
Add-Txt $s 'L I V E   D E M O N S T R A T I O N' $X0 0.95 $XW 0.22 10 (RGB 63 182 222) $true 1 | Out-Null
Add-Txt $s 'The new system, end to end' $X0 1.32 $XW 0.60 32 $WHITE $true 1 | Out-Null
Add-Txt $s 'The question this demo answers, in one line.' $X0 2.02 $XW 0.30 14 $DK_SUB $false 1 | Out-Null
$demoW = ($XW - 3 * 0.42) / 4
$demoSteps = @(
    @('STEP ONE',   'What we show first'),
    @('STEP TWO',   'What we show next'),
    @('STEP THREE', 'And then this'),
    @('STEP FOUR',  'Finishing here')
)
for ($i = 0; $i -lt 4; $i++) {
    $dx = $X0 + $i * ($demoW + 0.42)
    Add-Rect $s $dx 3.05 $demoW 1.30 $INK2 $true 0.07 | Out-Null
    Add-Rect $s $dx 3.05 $demoW 0.05 $BLUE $false | Out-Null
    Add-Txt $s $demoSteps[$i][0] $dx 3.42 $demoW 0.24 11.5 (RGB 92 176 232) $true 2 | Out-Null
    Add-Txt $s $demoSteps[$i][1] $dx 3.76 $demoW 0.40 10.5 $DK_SUB $false 2 1.2 | Out-Null
    if ($i -lt 3) { Add-Line $s ($dx + $demoW + 0.09) 3.70 ($dx + $demoW + 0.33) 3.70 (RGB 63 182 222) 2 1 | Out-Null }
}
Add-Rect $s ($X0 + $XW / 2 - 1.85) 4.95 3.70 1.00 $BLUE $true 0.14 | Out-Null
Add-Txt $s 'LIVE DEMO' ($X0 + $XW / 2 - 1.85) 5.18 3.70 0.55 30 $WHITE $true 2 | Out-Null
Add-Txt $s 'approx. 10-15 minutes' $X0 6.14 $XW 0.24 11 $DK_SUB $false 2 | Out-Null

# ===== 8. Closing =====
$s = New-Slide $pres $true
Add-Rect $s 9.55 0 ($SlideW - 9.55) $SlideH $INK2 $false | Out-Null
Add-Oval $s 11.85 6.05 2.30 $TEAL | Out-Null
Add-Txt $s 'T H A N K   Y O U' $X0 1.15 7.5 0.22 10 (RGB 63 182 222) $true 1 | Out-Null
Add-Txt $s "We keep today's work safe.`rWe build tomorrow's system together." $X0 1.58 8.2 1.55 30 $WHITE $true 1 1.15 | Out-Null
Add-Txt $s 'Thank you to everyone who contributed.' $X0 3.32 8.0 0.50 12 $DK_SUB $false 1 1.2 | Out-Null
Add-Txt $s 'T A L K   T O   U S' $X0 4.30 8.0 0.20 9 $TEAL $true 1 | Out-Null
$clWays = @('Ask now', 'Find us in the office', 'Write to us any time')
$clW = (8.0 - 2 * 0.22) / 3
for ($i = 0; $i -lt $clWays.Count; $i++) {
    $cx = $X0 + $i * ($clW + 0.22)
    Add-Rect $s $cx 4.66 $clW 0.62 $INK2 $true 0.10 | Out-Null
    Add-Txt $s $clWays[$i] $cx 4.86 $clW 0.24 10.5 $WHITE $true 2 | Out-Null
}
Add-Rect $s $X0 5.90 1.55 0.035 $TEAL $false | Out-Null
Add-Txt $s 'We capture every request and follow up.' $X0 6.16 8.0 0.26 11 $DK_SUB $false 1 | Out-Null

$pres.SaveAs($target, 24)
"Saved: $target ($($pres.Slides.Count) slides)"

} finally {
    if ($pres) { $pres.Close() }
    $app.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
}
