# Pitfalls and recipes

Every item here comes from a bug that actually shipped into a rendered deck. Most produce no error
message at all — they are visible only in the PNG, which is why the render-and-look loop is not
optional.

## 1. Character encoding corruption

**Symptom.** A middot renders as `A·`, an times sign as `A—`. The script looks perfectly correct
when you read it back.

**Cause.** PowerShell 7 writes UTF-8 **without a BOM**. If the build then runs under Windows
PowerShell 5.1 — which is what `powershell.exe` is — 5.1 reads a BOM-less file as CP1252 and mangles
every multi-byte character. Every non-ASCII literal in the script is affected at once.

**Fix.** Keep the script **pure ASCII** and build non-ASCII characters from code points:

```powershell
Add-Txt $s "Legacy system  $([char]0x00B7)  infrastructure  $([char]0x00B7)  new platform" ...
Add-Txt $s "Quantum Systems  $([char]0x00D7)  Partner" ...
$up = [char]0x2191    # arrow
$em = [char]0x2003    # em space
```

Note the switch to double quotes — `$([char]0x00B7)` is subexpression interpolation and does not
expand inside single quotes.

Write the file with a BOM as belt and braces, and assert the invariant after every edit:

```powershell
if (([regex]::Matches($t, '[^\x00-\x7F]')).Count -ne 0) { throw 'non-ASCII introduced' }
[IO.File]::WriteAllText($path, $t, (New-Object Text.UTF8Encoding $true))
```

To find offending lines in an existing script:

```powershell
[regex]::Matches($t, '(?m)^.*[^\x00-\x7F].*$') | ForEach-Object { $_.Value }
```

## 2. Variable shadowing (PowerShell is case-insensitive)

**Symptom.** A slide far below the one you edited breaks — shapes stack, or a loop runs the wrong
number of times. Nothing errors.

**Cause.** `$rX` and `$rx` are the *same variable*. Slide blocks share one script scope, so a
plausibly-named local in slide 12 silently overwrites a shared value and corrupts every *later*
slide that relies on it.

**Fix.** Prefix every slide-local variable with a short slide tag: `$rmRows`, `$osPhases`,
`$wfSteps`, `$p2Rows`, `$ukItems`. Reserve bare short names for genuinely shared layout values
(`$c3w`, `$centers`, `$demoW`) and never reassign them inside a slide block.

When a slide breaks that you did not touch, suspect this first.

## 3. `.Replace()` returns silently on no match

**Symptom.** The build succeeds, the PNG is unchanged, and it looks as though the edit did not run.
It did — it just matched nothing.

**Fix.** Guard every replacement:

```powershell
$n = ([regex]::Matches($t, [regex]::Escape($p))).Count
if ($n -ne 1) { throw "occurrences = $n for: $p" }
$t = $t.Replace($p, $replacement)
```

The count check catches both halves of the problem: `0` means the search string is wrong, `2+`
means it is ambiguous and would corrupt another slide.

## 4. Backtick escapes inside search strings

**Symptom.** A search string copied verbatim out of the file never matches.

**Cause.** Deck scripts contain PowerShell double-quoted literals with an embedded `` `r `` for a
line break, e.g. `"Products, BOMs and documents,`rread-only in most areas"`. Searching for that
using a **double-quoted** PowerShell string turns `` `r `` into a real carriage return, which is
not what is on disk.

**Fix.** Use a single-quoted search string, double the backtick, or — simplest and most robust —
match a shorter substring that avoids the escape entirely.

## 5. CRLF in here-strings

**Symptom.** A multi-line search block built with `@' ... '@` fails to match a file that plainly
contains it.

**Cause.** Line-ending mismatch between the here-string and the file.

**Fix.** Normalise before comparing or inserting:

```powershell
$new = ($new -split "`r?`n") -join "`r`n"
```

Or build the search block from an array joined with `"`r`n"`.

## 6. Ambiguous matches across slides

**Symptom.** An edit lands on the wrong slide, or an occurrence guard throws `occurrences = 2`.

**Cause.** Structural lines repeat. `Add-Line $s $X0 2.22 $X1 2.22 $LINE 1.25` is a table rule and
appears on every table slide.

**Fix.** Extract one slide block by regex, mutate the substring, splice it back:

```powershell
$m = [regex]::Match($t, '(?ms)^# ===== 25\..*?(?=^# ===== 26\.)')
if (-not $m.Success) { throw 'block 25 not found' }
$blk = $m.Value
# ...guarded swaps inside $blk...
$t = $t.Remove($m.Index, $m.Length).Insert($m.Index, $blk)
```

The same pattern replaces an entire slide block when a redesign changes more than a line or two.
For the final block in the file, anchor the lookahead on the closing section instead:

```powershell
[regex]::Match($t, '(?ms)^# ===== 27\..*?(?=^\$pres\.SaveAs)')
```

## 7. Reordering table rows

Editing rows individually invites both ambiguity and drift. Swap the whole array literal:

```powershell
$m = [regex]::Match($t, '(?ms)\$rmRows = @\(\r?\n.*?\r?\n\)')
if (-not $m.Success) { throw 'rmRows not found' }
$new = @'
$rmRows = @(
    @('Product management', 'Phase 1', 'DONE',        'Create, find and edit products'),
    @('Mechanical BOMs',    'Phase 1', 'DONE',        'Multi-level structure and revisions')
)
'@
$new = ($new -split "`r?`n") -join "`r`n"
$t = $t.Remove($m.Index, $m.Length).Insert($m.Index, $new)
```

## 8. Silent text overflow

**Symptom.** Text runs past its box, over a neighbouring column, or off the slide.

**Cause.** `Add-Txt` sets `AutoSize = 0`, so nothing ever resizes or warns.

**Fix.** Budget before writing. Roughly:

```
characters per line  ~=  width_inches / (0.5 * font_pt / 72)
```

3.45" at 11.5pt is about 43 characters; 4.77" at 10pt about 68. Check every string you add, and
look at the render regardless.

## 9. Cards invisible against the background

**Symptom.** The layout renders but the cards cannot be seen.

**Cause.** The light slide background is `RGB 246 250 253`. A card filled with the same near-white
tint vanishes.

**Fix.** Fill cards `$WHITE` on the light background and give each one a coloured top or left bar
for definition. On dark slides use `$INK2` cards on the `$INK` background.

## 10. Dead space inside cards

**Symptom.** Cards look empty and bottom-heavy; the slide feels unbalanced.

**Cause.** Card height was sized for the longest possible text rather than the actual text.

**Fix.** Size the card to its real content, then recentre the whole group vertically in the space
between the subtitle and the footer rule. Expect two or three render-and-look iterations to settle
a new layout — that is normal, not a sign something is wrong.

## 11. Retuning a table after adding a row

Adding one row pushes the table into the footer. Everything must move together: pitch, the
per-column vertical offsets, the zebra band height, and the y of any closing strip.

Reference points that fit above a 6.80 footer rule, for a table starting at y 2.32:

| Rows | Header y | Rule | Start y | Pitch | Strip |
|---|---|---|---|---|---|
| 8  | 1.94 | 2.22 | 2.32 | 0.44 | y 5.96 h 0.68 |
| 9  | 1.94 | 2.22 | 2.32 | 0.41 | y 6.08 h 0.62 |
| 11 | 1.94 | 2.22 | 2.32 | 0.34 | y 6.12 h 0.58 |
| 12 | 1.86 | 2.14 | 2.22 | 0.32 | y 6.12 h 0.58 |
| 13 | 1.80 | 2.08 | 2.16 | 0.30 | y 6.12 h 0.58 |

Past 13 rows, split the slide rather than shrinking further.

## 12. Shapes overflowing their container

**Symptom.** The last box in a group hangs out of the bottom of the zone panel behind it, or a
connector label runs underneath a neighbouring card.

**Cause.** Container size was chosen before the contents were finalised, then a row was added or a
label got longer. Nothing clips in PowerPoint, so it renders as an overhang rather than an error.

**Fix.** Derive the container from its contents, not the other way round. A zone holding rows that
end at `4.70 + 0.74` needs a height of at least `5.44 - 2.05 = 3.39`, plus padding.

Keep **connector labels to one short word** (`push`, `HTTPS`, `sync`). The gap between two zone
columns is typically half an inch, which at 8pt holds about nine characters. Anything longer will
sit on top of the next zone.

## 13. Decimal separators in generated SVG

**Symptom.** An SVG generated on a German or French machine renders as an empty or broken image.

**Cause.** `12.5` formats as `12,5`, which makes SVG path data invalid.

**Fix.** Force `InvariantCulture` at the top of every generator and format through a helper:

```powershell
[System.Threading.Thread]::CurrentThread.CurrentCulture = [System.Globalization.CultureInfo]::InvariantCulture
$INV = [System.Globalization.CultureInfo]::InvariantCulture
function N([double]$v) { $v.ToString('0.##', $INV) }
```

## 14. PowerPoint already open

The build deletes and rewrites the `.pptx`. If the user has it open, the build fails or writes a
locked temporary copy.

Check first, and **never kill the process** — it may hold unsaved work in other files. Ask the user
to close it:

```powershell
Get-Process POWERPNT -ErrorAction SilentlyContinue
```

## 15. Orphaned COM processes

An error between `Presentations.Add()` and `Close()` leaves `POWERPNT.EXE` running and holding the
file. Always use `try/finally` with `Close()`, `Quit()` and `ReleaseComObject`.

## 16. `$app.Visible` must be true

PowerPoint automation refuses to run reliably hidden. Set `$app.Visible = -1`. The window will
flicker during the build; this is expected.
