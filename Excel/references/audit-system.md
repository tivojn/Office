# Post-Generation Audit System for Excel Workbooks

## Table of Contents

1. [Cascading Fix Problem](#cascading-fix-problem)
2. [Checks 1-10](#check-1-column-width)
3. [Iterative Fix Loop](#iterative-fix-loop)
4. [Fix Strategies](#fix-strategies)
5. [False Positive Avoidance](#false-positive-avoidance)
6. [Output Format](#output-format)
7. [Key Lessons Learned](#key-lessons-learned)

---

## Cascading Fix Problem

Fixing one issue often creates another. This is the #1 reason audits fail:
- Widening a column to fit text → may push total width beyond print area (CHECK 6)
- Increasing font size to fix minimum → changes row heights and print pagination (CHECK 6)
- Fixing a number format → may change column width needed (CHECK 1)
- Adding conditional formatting → may introduce off-palette colors (CHECK 5)

**The iterative loop is NON-NEGOTIABLE. A single-pass audit is useless.**

---

## CHECK 1: COLUMN WIDTH
```
For every column with data:
  - Column width > 0 (not hidden unintentionally)
  - Text columns: width >= max(len(cell_value) * 1.1, 8) characters
    (allow 10% buffer over longest content)
  - Number columns: no cells displaying "######" (width too narrow for format)
  - Merged cell ranges: sum of merged column widths >= content width
  - No column left at default width (8.43) if it contains styled data
  - Spacer columns: width 2-3 (intentionally narrow)

STYLE-AWARE: Not affected by style — width is content-driven.
```

## CHECK 2: FONT COMPLIANCE
```
For every cell with content or formatting:
  - Font size >= 9pt for captions/labels
  - Font size >= 10pt for body text
  - Font size >= 11pt for headers
  - Font name is explicitly set (not inherited default)

STYLE-AWARE: If a style is active:
  - Title cells use the style's title font name and size
  - Header cells use the style's header font name and size
  - Body cells use the style's body font name and size
  - KPI values use the style's kpi_value font name and size
  - FLAG WARNING for any font not in the active style's fonts dict
```

## CHECK 3: NUMBER FORMAT
```
For every column with numeric data:
  - Currency values: format is '$#,##0.00' or '$#,##0' (not "General")
  - Percentage values: format is '0.0%' or '0%' (not "General")
  - Date values: format is 'YYYY-MM-DD' or consistent date format (not number serial)
  - Integer values: format is '#,##0' with thousands separator
  - All cells in a column use the same number format (consistency)
  - No "General" format on clearly typed data (currencies, dates, percentages)

EXCEPTION: Input cells may intentionally use different formats.
```

## CHECK 4: FORMULA INTEGRITY
```
For every cell with a formula:
  - Value is not #REF! (broken reference)
  - Value is not #NAME? (unknown function/name)
  - Value is not #VALUE! (type mismatch)
  - Value is not #DIV/0! (division by zero without protection)
  - Value is not #N/A (lookup failure without IFERROR wrapper)
  - No circular references detected

For named ranges:
  - All defined names reference valid ranges (not #REF!)
  - No duplicate named ranges with conflicting scopes

FLAG CRITICAL for any error value.
FLAG WARNING for #N/A if IFERROR wrapper exists (may be intentional).
```

## CHECK 5: COLOR / FILL COMPLIANCE
```
Collect all unique colors from:
  - Cell background fills
  - Font colors
  - Border colors
  - Conditional formatting colors
  - Tab colors

For each color:
  - Not accidentally default blue from theme (Excel's default blue #4472C4)
  - Not accidentally default black when style specifies a different dark color

STYLE-AWARE: If a style is active:
  - Every color must match a value in the style's palette dict
    Tolerance: ±15 per RGB channel
  - Exception: pure white (255,255,255) and pure black (0,0,0) always allowed
  - Header row background matches style header_bg
  - Alternating row fill matches style alt_row
  - KPI panel background matches style kpi_bg
  - Border colors match style border
  - FLAG WARNING for off-palette colors
```

## CHECK 6: PRINT LAYOUT
```
For every visible sheet with data:
  - Print area is set (not empty)
  - Page orientation is set explicitly (landscape for dashboards, portrait for tables)
  - Margins are set (not default 0.75" — should match style or be explicitly chosen)
  - FitToPagesWide = 1 (content fits page width without horizontal scrolling)
  - Print title rows set for data tables (headers repeat on every printed page)
  - Grid lines: PrintGridlines = False for dashboards, True for data sheets (optional)
  - Header/footer set if sheet has > 1 printed page

STYLE-AWARE: If a style is active:
  - Orientation matches style print config
  - Paper size matches style print config
```

## CHECK 7: CHART THEMING
```
For every chart in the workbook:
  - Chart has a title (HasTitle = True)
  - Chart title uses heading font (not default)
  - Chart border is removed (ChartArea border = no line)
  - Gridlines removed or set to very light color
  - Legend exists and is positioned at bottom (not overlapping data)
  - Axis text uses body or muted color (not default black)
  - Series colors are explicitly set (not Excel defaults)

STYLE-AWARE: If a style is active:
  - Series 1 color matches style chart_1
  - Series 2 color matches style chart_2
  - Series 3 color matches style chart_3
  - Series 4 color matches style chart_4
  - Tolerance: ±15 per RGB channel
  - Plot area background is white or style card_bg
  - FLAG WARNING for default Excel series colors (indicates theme not applied)
```

## CHECK 8: CONDITIONAL FORMATTING
```
For every sheet with conditional formatting rules:
  - Rules don't conflict (same range with contradictory conditions)
  - Positive color matches palette positive (or cond_format.positive_text)
  - Negative color matches palette negative (or cond_format.negative_text)
  - Data bar color matches palette accent or cond_format.data_bar
  - No rules reference deleted/moved ranges

STYLE-AWARE: If a style is active:
  - All conditional formatting colors must appear in style cond_format dict
  - Tolerance: ±15 per RGB channel
  - FLAG WARNING for off-palette conditional formatting colors
```

## CHECK 9: STYLE COMPLIANCE (only when a design style is active)

```
Skip this check entirely if no design style was specified.

Load the active style dict from references/style-xlsx-mapping.md.

9a — TAB COLORS:
  All visible sheet tabs match style tab_colors dict
  (tolerance: ±15 per RGB channel)

9b — FONT FAMILIES:
  Collect all unique font names across the workbook:
    Every font must appear in the active style's fonts dict values
    FLAG WARNING for each font not in the style

9c — COLOR PALETTE:
  Collect all unique colors from cells, fills, borders:
    Each color must match a style palette value (tolerance: ±15 per RGB channel)
    FLAG WARNING for off-palette colors
    Exception: white and black always allowed

9d — TABLE STYLING:
  All data tables follow the style's table_style pattern
  Header row colors match style header_bg/header_text
  Banded row colors match style alt_row

9e — KPI PANELS:
  KPI panel background matches style kpi_bg
  KPI value font matches style kpi_value font
  KPI label color matches style kpi_label

9f — CHART THEMING:
  (Covered by CHECK 7 — cross-referenced here for completeness)

9g — STYLE-SPECIFIC RULES:
  XSTYLE-01 (Consulting): Max 2-3 colors per sheet, no decorative images
  XSTYLE-04 (Data Science): Monospace font (Consolas) on all data cells
  XSTYLE-06 (Finance): Blue font on input cells, black on formulas, green on links
  XSTYLE-07 (Marketing): No pure black text — use dark brown #2D1B0E
```

## CHECK 10: DATA STRUCTURE
```
For every data region in the workbook:
  - Header row exists (first row of each data block has bold/styled text)
  - No mixed data types within a column (text + numbers in same column)
  - Stacked tables have >= 3 empty rows between them
  - No data in designated spacer rows or spacer columns
  - Named ranges reference valid, non-empty ranges
  - Sheet order follows recommended pattern:
    Dashboard → Data → Charts → Reference → Config → Audit
  - Frozen panes set on data sheets (freeze below header row)
  - AutoFilter enabled on data tables (optional but recommended)
```

---

## Iterative Fix Loop

```python
MAX_PASSES = 5

for pass_num in range(1, MAX_PASSES + 1):
    issues = run_all_checks(wb)  # Checks 1-10 (9 only if style active)
    critical = [i for i in issues if i.severity == 'CRITICAL']

    if not critical:
        print(f"Clean after {pass_num - 1} fix passes")
        break

    for issue in issues:
        apply_fix(issue)

    wb.save(path)
    # If using openpyxl, reload: wb = load_workbook(path)
    # Reload after save: wb = load_workbook(path)

    print(f"Pass {pass_num}: fixed {len(issues)} issues, re-auditing...")
else:
    print(f"{len(critical)} critical issues remain after {MAX_PASSES} passes")
```

---

## Fix Strategies

**COLUMN WIDTH (CHECK 1):**
1. Calculate max content length per column (including formatted numbers)
2. Set width = max(content_length * 1.1, min_width_for_type)
3. For merged cells: distribute width proportionally across merged columns
4. **After width fix → re-run CHECK 6 (print layout may change)**

**FONT COMPLIANCE (CHECK 2):**
1. Set missing font.name to style body font (or "Calibri" if no style)
2. Set missing font.size to style body size (or 11 if no style)
3. If font size < minimum, increase to minimum
4. **After font fix → re-run CHECK 1 (column widths may need adjustment)**

**NUMBER FORMAT (CHECK 3):**
1. Detect column data type by sampling first 10 non-empty cells
2. Apply appropriate format: currency, percentage, date, integer
3. Ensure consistency within each column
4. **After format fix → re-run CHECK 1 (formatted numbers may need wider columns)**

**FORMULA INTEGRITY (CHECK 4):**
1. #REF! → attempt to identify broken reference and fix, or flag for manual review
2. #DIV/0! → wrap in IFERROR: `=IFERROR(original_formula, 0)`
3. #N/A → wrap in IFERROR: `=IFERROR(original_formula, "")`
4. Circular reference → flag for manual review (cannot auto-fix)
5. **After formula fix → re-run CHECK 3 (values may change format needs)**

**COLOR COMPLIANCE (CHECK 5 / CHECK 9c):**
1. Map off-palette colors to nearest palette color by Euclidean RGB distance
2. Replace via openpyxl PatternFill with nearest palette color
3. **After color fix → no cascading checks needed**

**PRINT LAYOUT (CHECK 6):**
1. Set print area to encompass all data on the sheet
2. Set FitToPagesWide = 1
3. Set repeat rows for data tables (header row)
4. Set orientation based on content width (>10 columns = landscape)
5. **After print fix → no cascading checks needed**

**CHART THEMING (CHECK 7):**
1. Apply style series colors via VBA macro or openpyxl chart styling
2. Set chart title font to style heading font
3. Remove chart border
4. Position legend at bottom
5. **After chart fix → no cascading checks needed**

**CONDITIONAL FORMATTING (CHECK 8):**
1. Update rule colors to match palette cond_format values
2. Remove conflicting rules on same range
3. **After cond format fix → re-run CHECK 5 (new colors introduced)**

**SPACING (CHECK 10):**
1. Insert empty rows between stacked tables if gap < 3 rows
2. Ensure spacer rows have no data and no formatting (except background)
3. **After spacing fix → re-run CHECK 6 (print area may change)**

---

## False Positive Avoidance

1. **Hidden sheets:** Skip hidden sheets from all visual checks (font, color, print). Still check formula integrity.
2. **Spacer rows/columns:** Rows/columns with intentional empty space and specific height/width are NOT errors. Only flag if height=0 (hidden) unintentionally.
3. **Default font inheritance:** A cell with no explicit font may inherit from the column/row/workbook default. Only flag if the inherited font doesn't match the style.
4. **Conditional formatting colors:** These are dynamic — a cell may show green now but red later. Audit the RULE colors, not the current cell display color.
5. **Input cells in financial models:** XSTYLE-06 intentionally uses blue font for inputs. Don't flag blue as "off-palette" when finance style is active.
6. **Chart default colors:** If a chart was just created and styling hasn't been applied yet, all series will have Excel defaults. Flag this as WARNING, not CRITICAL — it's expected mid-build and gets fixed during theming.
7. **Named range scope:** Workbook-scoped and sheet-scoped named ranges can have the same name. This is not an error — it's a valid Excel feature.

---

## Output Format

Per-sheet report:
```
[Sheet] [SEVERITY] [CHECK#] — Description → Fix applied / Remaining
```

Final summary:
```
CRITICAL: N (must be 0 before delivery)
WARNING: N (should be 0, acceptable if documented)
INFO: N (advisory)

STYLE: XSTYLE-XX (Name) or "Default (no style)"
STYLE COMPLIANCE: All checks passed / N issues
PASSES: X until clean
TOTAL FIXES: N applied
```

---

## Key Lessons Learned

1. **Column width is the #1 visual bug in Excel** — too narrow causes ######, too wide wastes space. Always calculate from content.

2. **Fixing one issue often creates another** — the iterative loop is essential. Width changes affect print layout. Font changes affect column width. Format changes affect width.

3. **Number format "General" is the silent killer** — dates show as serials, currencies show without symbols, percentages show as decimals. Always set explicit formats.

4. **Formula errors are CRITICAL** — a single #REF! or #DIV/0! destroys credibility. Check every formula cell.

5. **Excel's default blue (#4472C4) creeps in everywhere** — theme-inherited colors are the most common off-palette violation. Always audit all colors.

6. **Print layout is often forgotten** — a beautiful on-screen dashboard that prints on 3 pages with cut-off columns is not professional. Always set print area, fit-to-page, and repeat rows.

7. **Chart styling uses openpyxl or VBA** — openpyxl handles chart theming natively. Use VBA macros for advanced chart operations that require a live Excel instance.

8. **Conditional formatting rules can conflict** — two rules on the same range with overlapping conditions cause unpredictable display. Audit rule order and conditions.

9. **The audit must work generically** — discover structure dynamically, don't hardcode sheet names or cell ranges. Scan all sheets, all data regions.

10. **Tab colors are cheap but impactful** — they take 1 line of code but immediately signal a polished, intentional workbook. Always set them.
