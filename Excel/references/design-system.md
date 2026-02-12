# Design System Reference — Excel

## Table of Contents

1. [Typography](#typography)
2. [Color Palettes](#color-palettes)
3. [Layout Rules](#layout-rules)
4. [Visual Hierarchy](#visual-hierarchy)
5. [Dashboard Layout Patterns](#dashboard-layout-patterns)
6. [KPI Panel Design](#kpi-panel-design)
7. [Table Design](#table-design)
8. [Chart Theming](#chart-theming)
9. [Conditional Formatting Patterns](#conditional-formatting-patterns)
10. [Print Layout](#print-layout)
11. [Financial Model Conventions](#financial-model-conventions)
12. [Palette Template](#palette-template)

---

## Typography

- **Minimum font size: 9pt** on any element. Body text minimum 10pt. Headers minimum 11pt.
- Preferred body text: 10-11pt. Headers: 11-14pt. Dashboard titles: 16-20pt. KPI values: 24-32pt.
- Always use a deliberate font pairing:
  - **Montserrat + Georgia** (modern headers + classic body)
  - **Calibri + Cambria** (safe cross-platform default)
  - **Segoe UI + Segoe UI** (clean, modern, Windows-native)
  - **Helvetica Neue + Georgia** (macOS elegant)
- Dashboard titles: bold, uppercase or mixed case, with accent color.
- Header rows: bold, white text on dark background, centered.
- Body text: left-aligned for text, right-aligned for numbers, centered for short labels.
- **Never use default Calibri 11pt everywhere.** Always differentiate header vs body styling.

## Color Palettes

Define a complete palette before building any workbook. All colors as RGB tuples for openpyxl.

### 8 Curated Styles

The skill includes 8 curated design styles (XSTYLE-01 through XSTYLE-08), each with a complete palette, font config, table styling, chart theming, KPI design, and conditional formatting colors.

**For full style descriptions:** See [Design Styles Catalog](design-styles-catalog.md)
**For implementation-ready values:** See [Style → Excel Mapping](style-xlsx-mapping.md)

| Style | Vibe | Best For |
|---|---|---|
| XSTYLE-01 — Consulting & Strategy | Navy + gold, McKinsey | Financial models, consulting |
| XSTYLE-02 — Executive Dashboard | Electric blue, Apple | SaaS KPIs, exec summaries |
| XSTYLE-03 — Corporate Report | Royal blue + emerald | Board packs, quarterly reviews |
| XSTYLE-04 — Data Science | Dark + cyan/magenta | Research data, technical |
| XSTYLE-05 — Sales & Pipeline | Coral + teal, bold | Pipeline trackers, quotas |
| XSTYLE-06 — Finance & Accounting | Forest green + gold | Budgets, P&L, audit papers |
| XSTYLE-07 — Marketing & Creative | Terracotta + sage | Campaign reports, attribution |
| XSTYLE-08 — Operations & Logistics | Cobalt + amber | Inventory, project tracking |

### Custom Palette

If none of the 8 styles fit, define a custom palette using the template below.

## Layout Rules

### Grid Principles

- **Plan the grid before writing data.** Decide column layout, widths, and zones before any content.
- Standard dashboard width: 10-14 columns (A through J-N).
- Leave column A and the last column as narrow spacers (width 2-3) for visual breathing room.
- Group related data: KPI row, then chart row, then data table row.
- Maximum visible columns without scrolling: ~14 on a standard 1920px monitor at 100% zoom.

### Column Width Guidelines

| Content Type | Width (chars) | Example |
|-------------|--------------|---------|
| Narrow spacer | 2-3 | Column dividers |
| Short label | 8-10 | "Status", "ID" |
| Name/Title | 18-25 | Full names, descriptions |
| Number | 12-15 | Currency, percentages |
| Date | 12-14 | "2026-02-10" |
| Wide text | 30-40 | Comments, notes |
| KPI value (merged) | 20-30 | Large display numbers |

### Row Height Guidelines

| Content Type | Height (pt) | Example |
|-------------|-------------|---------|
| Title bar | 45-60 | Dashboard title |
| Subtitle | 28-35 | Section subtitle |
| KPI panel | 50-80 | Large metric display |
| KPI label | 25-30 | Metric name/label |
| Header row | 30-40 | Column headers |
| Body row | 18-22 | Data rows |
| Spacer row | 8-12 | Visual separation |
| Chart row | 200-300 | Chart placeholder |

### Spacing & Separation

- Use empty spacer rows (height 8-12pt, no borders) between sections.
- Use empty spacer columns (width 2-3) between KPI panels.
- Never butt two styled sections directly against each other — always have a breathing gap.
- Frozen header rows: freeze at row 2 or 3 (below title bar).

## Visual Hierarchy

Layer order in a dashboard (top to bottom):

1. **Title bar** — merged across full width, dark background, large bold text
2. **Subtitle/filter bar** — date range, filter labels, muted text
3. **Spacer row** — thin gap
4. **KPI panels** — 3-5 merged-cell blocks with large numbers, accent labels
5. **Spacer row** — thin gap
6. **Charts** — 1-2 charts side by side
7. **Spacer row** — thin gap
8. **Data tables** — header row + banded body rows + totals row
9. **Footer** — source notes, last updated timestamp

## Dashboard Layout Patterns

### Executive Dashboard (1-sheet)

```
Row 1:     [============ TITLE BAR (merged A1:N1) ============]
Row 2:     [======== Subtitle / Last Updated (merged) ========]
Row 3:     [spacer]
Row 4-6:   [KPI 1] [sp] [KPI 2] [sp] [KPI 3] [sp] [KPI 4]
Row 7:     [spacer]
Row 8-22:  [   Chart 1 (A8:G22)   ] [sp] [   Chart 2 (I8:N22)   ]
Row 23:    [spacer]
Row 24:    [============ TABLE HEADER (A24:N24) ==============]
Row 25-40: [banded data rows...]
Row 41:    [============ TOTALS ROW (bold, accent top border) =]
Row 42:    [footer: source note, timestamp]
```

### Multi-Sheet Workbook

| Sheet | Purpose | Design |
|-------|---------|--------|
| Dashboard | Visual summary | KPIs, charts, hide gridlines |
| Data | Raw data | Clean table, filters, no styling overkill |
| Charts | Additional charts | Full-page chart views |
| Reference | Lookups, constants | Simple tables, named ranges |
| Config | Parameters, assumptions | Input cells highlighted (blue font) |

### Financial Model Layout

```
Row 1:     [====== Model Title ======]
Row 2:     [spacer]
Row 3:     [--- ASSUMPTIONS ---]
Row 4-10:  [assumption_label | input_value (blue font)]
Row 11:    [spacer]
Row 12:    [--- PROJECTIONS ---]
Row 13:    [header: Year | 2024 | 2025 | 2026 | 2027 | 2028]
Row 14-30: [line items with formulas]
Row 31:    [bold totals row, accent top border]
```

## KPI Panel Design

### Single KPI Panel (Merged Cells)

```
+---------------------------+
|                           |  ← Rows 1-2: Dark BG
|        $2.4M              |     Value: 28pt bold white, centered
|                           |
+---------------------------+
|     TOTAL REVENUE         |  ← Row 3: Same dark BG
|     +12.5% vs LY          |     Label: 10pt bold accent, centered
+---------------------------+     Subtitle: 9pt muted, centered
```

Typically spans 2-3 columns wide, 3-4 rows tall. Use merged cells for the value area and label area.

### KPI Row Layout (4 Panels)

```
Columns:  A    B-C    D    E-F    G    H-I    J    K-L    M
          sp   KPI1   sp   KPI2   sp   KPI3   sp   KPI4   sp
```

- Spacer columns (A, D, G, J, M): width 2
- KPI columns (B-C, E-F, H-I, K-L): width 12-15 each

### Accent Borders on KPI Panels

- Top border: thick accent color (weight 3-4, gold/blue)
- No side/bottom borders (clean floating look)
- Or: full box border with accent color, no inner borders

## Table Design

### Professional Banded Table

| Element | Style |
|---------|-------|
| Header row | Dark bg (navy/blue), white bold text, 11pt, centered, height 35pt |
| Header bottom border | Thick accent color (gold/blue), weight 3 |
| Body rows | 10pt Georgia/Calibri, alternating white/#F1F5F9, height 20-22pt |
| Body text alignment | Left for text, right for numbers, center for short labels |
| Borders | Thin light gray (#E2E8F0) on all inner + outer edges |
| Totals row | Bold, accent top border (thick), gray background |
| Number formats | Explicit: currency, %, integer, date — never "General" |

### Table Column Alignment Rules

| Data Type | Horizontal | Vertical | Number Format |
|-----------|-----------|----------|---------------|
| Text/Labels | Left | Middle | @ or General |
| Currency | Right | Middle | $#,##0.00 |
| Percentage | Right | Middle | 0.0% |
| Integer | Right | Middle | #,##0 |
| Date | Center | Middle | YYYY-MM-DD |
| Status/Category | Center | Middle | @ |
| Boolean/Flag | Center | Middle | General |

### Totals Row Pattern

```python
# Totals row: bold, accent border, slightly different background
totals_rng = ws.range((total_row, 1), (total_row, n_cols))
totals_rng.font.bold = True
totals_rng.font.size = 11
totals_rng.color = (248, 250, 252)   # Slightly off-white
set_border(totals_rng, edges=[8], weight=3, color=pal['accent'])  # Top border accent
```

## Chart Theming

### Chart Style Rules

1. **Remove chart border** — no outline on the chart area.
2. **Remove gridlines** or make them very light gray.
3. **Use palette colors** for series — never use Excel defaults.
4. **Title**: Montserrat/Calibri bold, 14pt, dark text color.
5. **Legend**: bottom position, 10pt, remove border.
6. **Axis text**: 10pt, muted color, no axis titles unless essential.
7. **Data labels**: only if chart has few data points, 10pt.
8. **Plot area**: white or very light background.

### Chart Type Selection Guide

| Data | Best Chart | Avoid |
|------|-----------|-------|
| Trend over time | Line / Area | Pie |
| Category comparison | Column (clustered) | Line |
| Part of whole | Pie / Donut (max 5-6 slices) | Stacked bar |
| Ranking | Horizontal bar | Pie |
| Correlation | Scatter | Column |
| Distribution | Histogram | Pie |
| Budget vs Actual | Stacked / Clustered column | Area |
| Multiple metrics | Combo (column + line) | Pie |

### Themed Chart Setup

```python
# After creating chart, apply theme
chart.api[1].HasTitle = True
chart.api[1].ChartTitle.Text = 'Monthly Revenue'
chart.api[1].ChartTitle.Format.TextFrame2.TextRange.Font.Size = 14
chart.api[1].ChartTitle.Format.TextFrame2.TextRange.Font.Name = 'Montserrat'
chart.api[1].ChartTitle.Format.TextFrame2.TextRange.Font.Bold = True

# Remove border
chart.api[1].ChartArea.Format.Line.Visible = False

# White plot area
chart.api[1].PlotArea.Format.Fill.ForeColor.RGB = rgb_to_excel(255, 255, 255)

# Remove gridlines
chart.api[1].Axes(2).HasMajorGridlines = False

# Apply palette colors to series
palette_colors = [pal['chart_1'], pal['chart_2'], pal['chart_3'], pal['chart_4']]
for i, color in enumerate(palette_colors, 1):
    try:
        s = chart.api[1].SeriesCollection(i)
        s.Format.Fill.ForeColor.RGB = rgb_to_excel(*color)
    except:
        break

# Legend at bottom
chart.api[1].HasLegend = True
chart.api[1].Legend.Position = -4107  # xlBottom
chart.api[1].Legend.Format.Line.Visible = False
```

## Conditional Formatting Patterns

### Traffic Light (Green/Yellow/Red)

```python
# Green for values > 80%, Yellow 50-80%, Red < 50%
rng = ws['D2:D20']

# Red: < 50%
rng.api.FormatConditions.Add(Type=1, Operator=6, Formula1="0.5")  # xlLess
fc1 = rng.api.FormatConditions(rng.api.FormatConditions.Count)
fc1.Interior.Color = rgb_to_excel(254, 226, 226)  # Light red bg
fc1.Font.Color = rgb_to_excel(220, 38, 38)          # Red text

# Yellow: 50-80%
rng.api.FormatConditions.Add(Type=1, Operator=1, Formula1="0.5", Formula2="0.8")  # xlBetween
fc2 = rng.api.FormatConditions(rng.api.FormatConditions.Count)
fc2.Interior.Color = rgb_to_excel(254, 249, 195)  # Light yellow bg
fc2.Font.Color = rgb_to_excel(161, 98, 7)           # Dark yellow text

# Green: > 80%
rng.api.FormatConditions.Add(Type=1, Operator=5, Formula1="0.8")  # xlGreater
fc3 = rng.api.FormatConditions(rng.api.FormatConditions.Count)
fc3.Interior.Color = rgb_to_excel(220, 252, 231)  # Light green bg
fc3.Font.Color = rgb_to_excel(22, 163, 74)          # Green text
```

### Data Bars (accent color)

```python
rng = ws['C2:C20']
rng.api.FormatConditions.AddDatabar()
db = rng.api.FormatConditions(rng.api.FormatConditions.Count)
db.BarColor.Color = rgb_to_excel(*pal['accent'])
db.BarBorder.Type = 1  # xlDataBarBorderSolid
```

### Color Scale (Red to Green)

```python
rng = ws['E2:E20']
rng.api.FormatConditions.AddColorScale(3)  # 3-color scale
cs = rng.api.FormatConditions(rng.api.FormatConditions.Count)
cs.ColorScaleCriteria(1).FormatColor.Color = rgb_to_excel(248, 113, 113)  # Red
cs.ColorScaleCriteria(2).FormatColor.Color = rgb_to_excel(253, 224, 71)   # Yellow
cs.ColorScaleCriteria(3).FormatColor.Color = rgb_to_excel(74, 222, 128)   # Green
```

### Highlight Positive/Negative

```python
rng = ws['F2:F20']

# Positive (green text)
rng.api.FormatConditions.Add(Type=1, Operator=5, Formula1="0")
fc = rng.api.FormatConditions(rng.api.FormatConditions.Count)
fc.Font.Color = rgb_to_excel(*pal['positive'])

# Negative (red text)
rng.api.FormatConditions.Add(Type=1, Operator=6, Formula1="0")
fc = rng.api.FormatConditions(rng.api.FormatConditions.Count)
fc.Font.Color = rgb_to_excel(*pal['negative'])
```

## Print Layout

### Landscape Dashboard (Letter)

```python
ps = ws.api.PageSetup
ps.Orientation = 2                    # Landscape
ps.PaperSize = 1                      # Letter
ps.TopMargin = 36                     # 0.5"
ps.BottomMargin = 36
ps.LeftMargin = 36
ps.RightMargin = 36
ps.CenterHorizontally = True
ps.FitToPagesWide = 1                 # Fit to 1 page wide
ps.FitToPagesTall = False             # Natural page breaks
ps.PrintArea = '$A$1:$N$50'
ps.PrintTitleRows = '$1:$2'           # Repeat title on each page
ps.PrintGridlines = False
ps.CenterHeader = '&"Montserrat,Bold"&14Dashboard Title'
ps.CenterFooter = 'Page &P of &N'
ps.RightFooter = '&D'                # Date
```

### Portrait Data Table (A4)

```python
ps = ws.api.PageSetup
ps.Orientation = 1                    # Portrait
ps.PaperSize = 9                      # A4
ps.TopMargin = 54                     # 0.75"
ps.BottomMargin = 54
ps.LeftMargin = 54
ps.RightMargin = 54
ps.CenterHorizontally = True
ps.FitToPagesWide = 1
ps.FitToPagesTall = False
ps.PrintTitleRows = '$1:$1'           # Repeat header
ps.PrintGridlines = False
```

## Financial Model Conventions

### Cell Color Coding (Industry Standard)

| Color | Meaning | Example |
|-------|---------|---------|
| Blue font | Input / Assumption | Editable inputs |
| Black font | Formula / Calculation | Computed values |
| Green font | Link to other sheet | Cross-sheet references |
| Red font | Link to external source | External data feeds |
| Yellow highlight | Key assumption | Important inputs |
| Light blue highlight | Input cell | Editable zones |
| Gray highlight | Protected / Locked | Non-editable |

### Number Format Standards

| Type | Format | Display |
|------|--------|---------|
| Years | `@` (text) | 2024, 2025, 2026 |
| Currency (M) | `$#,##0.0,,"M"` | $2.4M |
| Currency (K) | `$#,##0,"K"` | $450K |
| Currency (exact) | `$#,##0.00` | $1,234.56 |
| Percentage | `0.0%` | 12.5% |
| Ratio | `0.00x` | 1.25x |
| Zero as dash | `#,##0;-#,##0;"-"` | - |
| Growth rate | `+0.0%;-0.0%` | +12.5% |

## Palette Template

Use this template to define a custom palette:

```python
pal = {
    # Core colors
    'header_bg':    (R, G, B),      # Dark background for headers
    'header_text':  (R, G, B),      # Light text on headers (usually white)
    'accent':       (R, G, B),      # Primary accent (borders, highlights)
    'accent2':      (R, G, B),      # Secondary accent
    'text':         (R, G, B),      # Body text color
    'muted':        (R, G, B),      # Secondary text, labels

    # Table colors
    'alt_row':      (R, G, B),      # Alternating row background
    'border':       (R, G, B),      # Table/cell borders
    'card_bg':      (R, G, B),      # Card/panel background

    # Status colors
    'positive':     (R, G, B),      # Green (good, up, profit)
    'negative':     (R, G, B),      # Red (bad, down, loss)

    # Chart series colors (4 minimum)
    'chart_1':      (R, G, B),      # Primary series
    'chart_2':      (R, G, B),      # Secondary series
    'chart_3':      (R, G, B),      # Tertiary series
    'chart_4':      (R, G, B),      # Quaternary series

    # KPI panel colors
    'kpi_bg':       (R, G, B),      # Panel background
    'kpi_text':     (R, G, B),      # Large value text
    'kpi_label':    (R, G, B),      # Small label text
}
```
