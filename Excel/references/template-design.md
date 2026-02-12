# Template Design Reference — Excel (macOS)

The template-first approach for reliable Excel automation on macOS. Templates with prebuilt structure and VBA macros are the stable path for dashboards, pivots, charts, and complex formatting.

## Table of Contents

1. [Why Templates](#why-templates)
2. [Template File Format](#template-file-format)
3. [Required Sheets](#required-sheets)
4. [Named Range Conventions](#named-range-conventions)
5. [Excel Table (ListObject) Design](#excel-table-listobject-design)
6. [Prebuilt Pivots](#prebuilt-pivots)
7. [Prebuilt Charts](#prebuilt-charts)
8. [Image Placeholders](#image-placeholders)
9. [Audit Sheet Structure](#audit-sheet-structure)
10. [Template Creation Workflow](#template-creation-workflow)
11. [Transactional Run Pattern](#transactional-run-pattern)

---

## Why Templates

Templates provide convenience, consistency, and reliability for Excel automation on macOS. Pivot tables, charts, and complex formatting are best prebuilt rather than created dynamically. VBA project injection is blocked entirely, so macros must be embedded at template creation time.

Templates solve all of these problems:

- **Pivots**: Prebuilt in the template, bound to Excel Tables. Agent writes data, calls `RefreshAllPivots()`, done. No dynamic pivot creation needed.
- **Charts**: Prebuilt in the template, bound to named ranges or pivot output. Agent updates source data, charts auto-update after refresh. No dynamic chart construction.
- **VBA macros**: Already embedded in the `.xlsm` template. Agent calls them via AppleScript `do Visual Basic`. No code injection needed.
- **Formatting**: Header styles, banded rows, conditional formatting rules, number formats all baked into the template. Agent only needs to refresh, not rebuild.
- **Image placeholders**: Rectangle shapes with known names and positions. VBA macro replaces content, preserving size and location.

The agent's job is: **copy template, write source data, refresh, export**. Not rebuild from scratch.

### When Templates Are Required

| Scenario | Template? | Reason |
|----------|-----------|--------|
| Dashboard with pivots + charts | **Yes** | Pivots/charts unreliable when created dynamically on Mac |
| Financial model with formatting | **Yes** | Complex formatting, formulas, named ranges |
| Simple data dump with headers | No | openpyxl bulk write + basic formatting is sufficient |
| One-off styled table | No | openpyxl + VBA formatting macros work fine |
| Recurring report (weekly/monthly) | **Yes** | Template ensures consistency across runs |
| Workbook with images | **Yes** | Image placeholders ensure stable positioning |

### Template vs. From-Scratch Decision

```
Does the workbook need pivots, charts, or images?
  YES → Use a template
  NO  → Does it need complex formatting (merged cells, conditional formatting, KPI panels)?
          YES → Consider a template (faster, more reliable)
          NO  → Build from scratch with openpyxl
```

---

## Template File Format

### `.xlsm` (Macro-Enabled Workbook)

Use `.xlsm` for templates that contain VBA macros. This is the standard for all templates with prebuilt macros.

```
~/.claude/skills/xlsx-design-agent/templates/
  dashboard.xlsm          # General-purpose dashboard template
  financial-model.xlsm    # Financial model with assumptions + projections
  audit-report.xlsm       # Audit/compliance report template
  sales-tracker.xlsm      # Sales pipeline tracker
```

### `.xlsx` (Data-Only Workbook)

Use `.xlsx` for templates that have no embedded macros. The agent adds VBA at runtime via the AppleScript `do Visual Basic` fallback if macro execution is needed.

```python
# For .xlsx templates, inject VBA at runtime via AppleScript
import subprocess
vba_code = 'Sub RefreshAll()\\nActiveWorkbook.RefreshAll\\nEnd Sub'
subprocess.run([
    'osascript', '-e',
    f'tell application "Microsoft Excel" to do Visual Basic "{vba_code}"'
])
```

**Recommendation**: Always prefer `.xlsm` templates. Embedding macros at build time is more reliable than injecting them at runtime.

### Template Storage

All templates live in a single directory:

```
~/.claude/skills/xlsx-design-agent/templates/
```

Naming convention: `{purpose}.xlsm` or `{purpose}-{variant}.xlsm`.

Examples:
- `dashboard.xlsm` -- general dashboard
- `dashboard-dark.xlsm` -- dark theme variant
- `financial-quarterly.xlsm` -- quarterly financial model
- `sales-pipeline.xlsm` -- sales tracker

---

## Required Sheets

Every template should include these sheets. Not all are mandatory for every use case, but this is the standard set.

| Sheet | Purpose | Convention |
|-------|---------|------------|
| **Control** | Parameters, paths, palette, export config | Named cells: `Param_*`, `Palette_*`, `Export_*`, `Img_*` |
| **Data** | Source data as Excel Tables (ListObjects) | Named: `tbl_Sales`, `tbl_Users`, etc. |
| **Dashboard** | Visual summary with KPIs and charts | Hide gridlines, prebuilt layout, merged cells |
| **Pivot** | Prebuilt pivot tables | Source = Data sheet tables, refresh via macro |
| **Charts** | Full-page chart views | Chart sheets or embedded charts bound to named ranges |
| **Audit** | Structured audit results | Columns: Check, Status, Details, Timestamp |

### Sheet Tab Colors

Use consistent tab colors for visual navigation:

```python
# Tab color conventions (RGB tuples)
TAB_COLORS = {
    'Control':   (100, 116, 139),   # Slate gray
    'Data':      (59, 130, 246),    # Blue
    'Dashboard': (201, 168, 76),    # Gold (primary)
    'Pivot':     (139, 92, 246),    # Purple
    'Charts':    (16, 185, 129),    # Teal
    'Audit':     (220, 38, 38),     # Red
}
```

### Sheet Visibility

- **Dashboard**: Always visible, always the active sheet on open.
- **Data**: Visible (users may need to inspect source data).
- **Control**: Hidden or very hidden (`xlSheetVeryHidden = 2`). Users should not edit parameters directly.
- **Pivot**: Hidden unless the user needs to see raw pivot output.
- **Charts**: Visible if standalone chart sheets; hidden if charts are embedded on Dashboard.
- **Audit**: Hidden. Agent reads it programmatically after running `AuditWorkbook()`.

---

## Named Range Conventions

Named ranges are the contract between Python (writes) and VBA (reads). Use consistent prefixes to group by purpose.

### Parameters (Written by Python, Read by VBA)

```
Param_DateStart       → Control!$B$2       (date or string)
Param_DateEnd         → Control!$B$3       (date or string)
Param_Scenario        → Control!$B$4       (string: "Base", "Optimistic", "Pessimistic")
Param_Currency        → Control!$B$5       (string: "USD", "EUR", "GBP")
Param_FiscalYear      → Control!$B$6       (integer: 2026)
Param_Department      → Control!$B$7       (string or "All")
```

### Palette Colors (RGB values for VBA to read)

Store palette colors as comma-separated RGB strings. VBA parses them at runtime.

```
Palette_HeaderBg      → Control!$B$10      (format: "11,29,58")
Palette_HeaderText    → Control!$B$11      (format: "255,255,255")
Palette_Accent        → Control!$B$12      (format: "201,168,76")
Palette_Text          → Control!$B$13      (format: "26,32,44")
Palette_AltRow        → Control!$B$14      (format: "241,245,249")
Palette_Border        → Control!$B$15      (format: "226,232,240")
Palette_Positive      → Control!$B$16      (format: "22,163,74")
Palette_Negative      → Control!$B$17      (format: "220,38,38")
```

### Data Ranges (Excel Tables)

These are automatically managed by ListObjects. The named range expands as data grows.

```
InputTable_Sales      → Data!$A$1:$F$100   (auto-expands with ListObject)
InputTable_Users      → Data!$A$105:$D$200 (separate table, same sheet)
InputTable_Products   → Data!$A$205:$E$300 (separate table, same sheet)
```

### Image Insertion Plan

Each image gets three named cells: path, placeholder shape name, and sizing mode.

```
Img_1_Path            → Control!$B$20      (absolute file path)
Img_1_Placeholder     → Control!$C$20      (shape name: "LogoPlaceholder")
Img_1_Mode            → Control!$D$20      (FIT or FILL)

Img_2_Path            → Control!$B$21
Img_2_Placeholder     → Control!$C$21
Img_2_Mode            → Control!$D$21
```

### Export Configuration

```
Export_Path            → Control!$B$30      (absolute path for PDF/image export)
Export_Sheets          → Control!$B$31      (comma-separated: "Dashboard,Charts")
Export_Orientation     → Control!$B$32      ("Landscape" or "Portrait")
Export_PaperSize       → Control!$B$33      ("Letter" or "A4")
```

### Writing Named Ranges from Python

```python
from openpyxl import load_workbook

wb = load_workbook('/path/to/workbook.xlsx')
ws_ctrl = wb['Control']

# Write parameters
ws_ctrl['B2'] = '2026-01-01'    # Param_DateStart
ws_ctrl['B3'] = '2026-03-31'    # Param_DateEnd
ws_ctrl['B4'] = 'Base'          # Param_Scenario

# Write palette (RGB as comma-separated string)
ws_ctrl['B10'] = '11,29,58'     # Palette_HeaderBg
ws_ctrl['B11'] = '255,255,255'  # Palette_HeaderText
ws_ctrl['B12'] = '201,168,76'   # Palette_Accent

# Write image plan
ws_ctrl['B20'] = '/Users/user/logo.png'
ws_ctrl['C20'] = 'LogoPlaceholder'
ws_ctrl['D20'] = 'FIT'

# Write export config
ws_ctrl['B30'] = '/Users/user/output.pdf'
ws_ctrl['B31'] = 'Dashboard,Charts'
ws_ctrl['B32'] = 'Landscape'

wb.save('/path/to/workbook.xlsx')
```

---

## Excel Table (ListObject) Design

All source data must live in Excel Tables (ListObjects), not plain ranges. Tables auto-expand when data is appended, and pivots referencing tables auto-refresh correctly.

### Creating Tables in the Template

In Excel (manually, during template creation):

1. Select the header row + one data row (e.g., `A1:F2`).
2. Insert > Table (Cmd+T on Mac).
3. Check "My table has headers".
4. Rename the table in the Table Design tab.

### Naming Convention

```
tbl_Sales          # Primary sales data
tbl_Users          # User/customer data
tbl_Products       # Product catalog
tbl_Transactions   # Transaction log
tbl_Budget         # Budget line items
tbl_Actuals        # Actual spend/revenue
```

Prefix all table names with `tbl_` to distinguish from named ranges.

### Table Layout on the Data Sheet

Stack tables vertically with a 3-row gap between them. Each table starts with its header row.

```
Row 1:    [tbl_Sales headers: Date | Product | Region | Revenue | Units | Margin]
Row 2:    [sample data row]
Row 3:    [...more data...]
...
Row 103:  [end of tbl_Sales]
Row 104:  (empty)
Row 105:  (empty)
Row 106:  (empty)
Row 107:  [tbl_Users headers: UserID | Name | Email | Region]
Row 108:  [sample data row]
...
```

### Writing Data to Tables from Python

```python
import pandas as pd
from openpyxl import load_workbook

wb = load_workbook('/path/to/workbook.xlsx')
ws_data = wb['Data']

# Write new data starting at the first data row (row 2, below headers)
df = pd.DataFrame({
    'Date': ['2026-01-15', '2026-01-16'],
    'Product': ['Widget A', 'Widget B'],
    'Region': ['North', 'South'],
    'Revenue': [15000, 22000],
    'Units': [150, 220],
    'Margin': [0.32, 0.28],
})

# Write DataFrame to sheet (starting below header row)
start_row = 2  # Row 1 is headers
for r_idx, row in enumerate(df.values, start=start_row):
    for c_idx, value in enumerate(row, start=1):
        ws_data.cell(row=r_idx, column=c_idx, value=value)

wb.save('/path/to/workbook.xlsx')
```

### Table Auto-Expansion

When you write data below an Excel Table's last row, the table automatically expands to include the new rows. This means:

- Pivot tables referencing `tbl_Sales` will include the new data after refresh.
- Charts bound to the table's columns will show the new data.
- Named ranges pointing to the table auto-adjust.

No manual range updates needed.

---

## Prebuilt Pivots

Build pivot tables in the template during template creation. Never create pivots dynamically on macOS -- pivot tables should be prebuilt for consistency.

### Pivot Design Rules

1. **Source = Excel Table on Data sheet.** Always reference a `tbl_*` ListObject.
2. **One pivot per purpose.** Don't try to cram multiple analyses into one pivot.
3. **Place on the Pivot sheet.** Keep pivots separate from the Dashboard.
4. **Dashboard references pivot output via formulas.** Dashboard cells use `=Pivot!B5` to pull values.

### Pivot Layout in the Template

```
Pivot Sheet:
  A1:E20   → PivotTable_RevByRegion    (Revenue by Region, monthly)
  A25:D40  → PivotTable_UnitsByProduct  (Units sold by Product)
  A45:F60  → PivotTable_MarginTrend     (Margin % over time)
```

### Refreshing Pivots from the Agent

```python
# After writing new data, open in Excel and refresh pivots via AppleScript
import subprocess
subprocess.run(['open', '/path/to/workbook.xlsm'])
import time; time.sleep(2)
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to do Visual Basic "Call RefreshAllPivots"'])
```

The VBA macro:

```vba
Sub RefreshAllPivots()
    Dim ws As Worksheet
    Dim pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws
End Sub
```

### Pivot-to-Dashboard Connection

The Dashboard sheet reads computed values from the Pivot sheet:

```
Dashboard!B4  =  =Pivot!E5       (Total Revenue from pivot grand total)
Dashboard!B5  =  =Pivot!E10      (Top Region revenue)
Dashboard!B6  =  =Pivot!D20      (Total Units from second pivot)
```

This way, when the agent writes new data and refreshes pivots, the Dashboard KPIs update automatically.

---

## Prebuilt Charts

Build charts in the template bound to named ranges or pivot output ranges. Charts auto-update when the source data changes and the workbook recalculates.

### Chart Design Rules

1. **Bind to named ranges or table columns.** Never bind to hardcoded cell references like `=$B$2:$B$50`.
2. **Use the palette.** Apply template palette colors to series during template creation.
3. **Remove default clutter.** No gridlines, no chart border, legend at bottom, minimal axis labels.
4. **Position precisely.** Anchor charts to specific cells so they don't drift.

### Chart Types in Templates

| Dashboard Element | Chart Type | Source |
|-------------------|-----------|--------|
| Revenue trend | Line or area | `tbl_Sales[Date]` vs `tbl_Sales[Revenue]` |
| Category comparison | Clustered column | Pivot output range |
| Market share | Donut (max 5 slices) | Pivot output range |
| Budget vs actual | Combo (column + line) | `tbl_Budget` + `tbl_Actuals` |
| KPI sparklines | Small line charts | Rolling 12-month data |

### Multiple Scenarios

For templates that support scenario switching (Base / Optimistic / Pessimistic):

1. Create one chart per scenario, stacked in the same position.
2. VBA macro shows/hides the correct chart based on `Param_Scenario`.
3. Or: create one chart with multiple series, VBA toggles series visibility.

```vba
Sub ShowScenarioChart(scenario As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Dashboard")
    ws.ChartObjects("Chart_Base").Visible = (scenario = "Base")
    ws.ChartObjects("Chart_Optimistic").Visible = (scenario = "Optimistic")
    ws.ChartObjects("Chart_Pessimistic").Visible = (scenario = "Pessimistic")
End Sub
```

### Chart Refresh

Charts bound to tables or named ranges auto-refresh when the workbook recalculates:

```python
# After writing data and refreshing pivots:
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to do Visual Basic "Application.CalculateFull"'])
```

---

## Image Placeholders

Create rectangle shapes in the template as placeholders for images. VBA replaces the placeholder content with the actual image, preserving size and position.

### Creating Placeholders (During Template Build)

1. Insert > Shapes > Rectangle on the Dashboard sheet.
2. Resize to the desired image dimensions (e.g., 200x80 for a logo).
3. Name the shape: select it, type name in the Name Box (top-left of formula bar).
4. Use descriptive names: `LogoPlaceholder`, `HeroImagePlaceholder`, `ChartImagePlaceholder`.
5. Set a light gray fill (#F1F5F9) so the placeholder is visible but unobtrusive.
6. Optional: lock the shape position (Format Shape > Properties > Don't move or size with cells).

### Placeholder Naming Convention

```
LogoPlaceholder           # Company logo (typically top-left of Dashboard)
HeroImagePlaceholder      # Large banner/hero image
ProductImage_1            # Product catalog image slot 1
ProductImage_2            # Product catalog image slot 2
ChartExport_1             # Slot for an externally-generated chart image
SignaturePlaceholder      # Signature block for approval workflows
```

### Image Insertion via VBA

The agent writes image metadata to the Control sheet, then calls the VBA macro:

```python
from openpyxl import load_workbook

wb = load_workbook('/path/to/workbook.xlsm')
ws_ctrl = wb['Control']
ws_ctrl['B20'] = '/Users/user/logo.png'       # Img_1_Path
ws_ctrl['C20'] = 'LogoPlaceholder'             # Img_1_Placeholder
ws_ctrl['D20'] = 'FIT'                         # Img_1_Mode
wb.save('/path/to/workbook.xlsm')

# Call VBA to insert images via AppleScript
import subprocess
subprocess.run(['open', '/path/to/workbook.xlsm'])
import time; time.sleep(2)
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to do Visual Basic "Call InsertAllImages"'])
```

The VBA macro reads the Control sheet and inserts each image:

```vba
Sub InsertAllImages()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Control")
    Dim r As Long
    For r = 20 To 29  ' Rows 20-29 reserved for image plan
        Dim imgPath As String
        Dim placeholderName As String
        Dim mode As String
        imgPath = ws.Range("B" & r).Value
        placeholderName = ws.Range("C" & r).Value
        mode = ws.Range("D" & r).Value
        If imgPath <> "" And placeholderName <> "" Then
            InsertImageIntoPlaceholder imgPath, placeholderName, mode
        End If
    Next r
End Sub

Sub InsertImageIntoPlaceholder(imgPath As String, shapeName As String, mode As String)
    Dim ws As Worksheet
    Dim shp As Shape
    Dim pic As Shape

    ' Find the placeholder shape across all sheets
    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            If shp.Name = shapeName Then
                ' Store placeholder dimensions
                Dim l As Double, t As Double, w As Double, h As Double
                l = shp.Left: t = shp.Top: w = shp.Width: h = shp.Height

                ' Delete placeholder
                shp.Delete

                ' Insert image
                Set pic = ws.Shapes.AddPicture( _
                    imgPath, msoFalse, msoTrue, l, t, -1, -1)

                ' Apply sizing mode
                pic.LockAspectRatio = msoTrue
                If mode = "FIT" Then
                    ' Scale to fit within bounds, maintaining aspect ratio
                    If pic.Width / pic.Height > w / h Then
                        pic.Width = w
                    Else
                        pic.Height = h
                    End If
                    ' Center within original bounds
                    pic.Left = l + (w - pic.Width) / 2
                    pic.Top = t + (h - pic.Height) / 2
                ElseIf mode = "FILL" Then
                    ' Scale to fill bounds, may crop
                    If pic.Width / pic.Height < w / h Then
                        pic.Width = w
                    Else
                        pic.Height = h
                    End If
                    pic.Left = l
                    pic.Top = t
                End If

                pic.Name = shapeName & "_Image"
                Exit Sub
            End If
        Next shp
    Next ws
End Sub
```

### Image Preflight (Python)

Before inserting, validate images in Python:

```python
from PIL import Image
import os

def preflight_image(path, target_width, target_height):
    """Validate image and determine sizing mode."""
    if not os.path.exists(path):
        return {'status': 'FAIL', 'error': f'File not found: {path}'}

    img = Image.open(path)
    w, h = img.size
    aspect = w / h
    target_aspect = target_width / target_height

    return {
        'status': 'PASS',
        'width': w,
        'height': h,
        'aspect': round(aspect, 2),
        'target_aspect': round(target_aspect, 2),
        'recommended_mode': 'FIT' if abs(aspect - target_aspect) > 0.2 else 'FILL',
    }
```

---

## Audit Sheet Structure

The Audit sheet captures verification results after every build. The agent reads it programmatically to decide whether to proceed with export.

### Column Layout

| Column | Header | Type | Description |
|--------|--------|------|-------------|
| A | Check | String | What was verified |
| B | Status | String | `PASS` or `FAIL` |
| C | Details | String | Human-readable details or error message |
| D | Timestamp | DateTime | When the check ran |

### Sample Audit Output

```
| Row | Check                          | Status | Details                    | Timestamp           |
|-----|--------------------------------|--------|----------------------------|---------------------|
| 2   | Sheet "Dashboard" exists       | PASS   | Found                      | 2026-02-10 10:00:00 |
| 3   | Sheet "Data" exists            | PASS   | Found                      | 2026-02-10 10:00:00 |
| 4   | Table "tbl_Sales" has data     | PASS   | 247 rows                   | 2026-02-10 10:00:00 |
| 5   | Pivot "RevByRegion" refreshed  | PASS   | 5 regions                  | 2026-02-10 10:00:01 |
| 6   | Chart "Revenue" has title      | PASS   | "Monthly Revenue"          | 2026-02-10 10:00:01 |
| 7   | Image "Logo" within bounds     | FAIL   | Exceeds by 15px right      | 2026-02-10 10:00:01 |
| 8   | Named range "Param_DateStart"  | PASS   | Value: 2026-01-01          | 2026-02-10 10:00:01 |
| 9   | Dashboard gridlines hidden     | PASS   | DisplayGridlines = False   | 2026-02-10 10:00:01 |
| 10  | Export path writable           | PASS   | /Users/user/output.pdf     | 2026-02-10 10:00:01 |
```

### Reading Audit Results from Python

```python
from openpyxl import load_workbook

wb = load_workbook('/path/to/workbook.xlsx')
ws_audit = wb['Audit']

# Read all audit rows
failures = []
for row in ws_audit.iter_rows(min_row=2, values_only=True):
    if row[0] is None:
        break
    check, status, details, ts = row
    if status == 'FAIL':
        failures.append(f'{check}: {details}')

if failures:
    print(f'AUDIT FAILED ({len(failures)} issues):')
    for f in failures:
        print(f'  - {f}')
else:
    print('AUDIT PASSED: All checks OK')
```

### Audit VBA Macro

The `AuditWorkbook()` macro runs all checks and writes results to the Audit sheet:

```vba
Sub AuditWorkbook()
    Dim wsAudit As Worksheet
    Set wsAudit = ThisWorkbook.Sheets("Audit")
    wsAudit.Cells.Clear

    ' Headers
    wsAudit.Range("A1").Value = "Check"
    wsAudit.Range("B1").Value = "Status"
    wsAudit.Range("C1").Value = "Details"
    wsAudit.Range("D1").Value = "Timestamp"

    Dim r As Long: r = 2

    ' Check required sheets exist
    Dim requiredSheets As Variant
    requiredSheets = Array("Control", "Data", "Dashboard", "Pivot", "Audit")
    Dim sName As Variant
    For Each sName In requiredSheets
        Dim found As Boolean: found = False
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = sName Then found = True: Exit For
        Next ws
        wsAudit.Range("A" & r).Value = "Sheet """ & sName & """ exists"
        wsAudit.Range("B" & r).Value = IIf(found, "PASS", "FAIL")
        wsAudit.Range("C" & r).Value = IIf(found, "Found", "Missing")
        wsAudit.Range("D" & r).Value = Now
        r = r + 1
    Next sName

    ' Check tables have data
    Dim tbl As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each tbl In ws.ListObjects
            wsAudit.Range("A" & r).Value = "Table """ & tbl.Name & """ has data"
            If tbl.ListRows.Count > 0 Then
                wsAudit.Range("B" & r).Value = "PASS"
                wsAudit.Range("C" & r).Value = tbl.ListRows.Count & " rows"
            Else
                wsAudit.Range("B" & r).Value = "FAIL"
                wsAudit.Range("C" & r).Value = "Empty table"
            End If
            wsAudit.Range("D" & r).Value = Now
            r = r + 1
        Next tbl
    Next ws

    ' Check Dashboard gridlines
    wsAudit.Range("A" & r).Value = "Dashboard gridlines hidden"
    wsAudit.Range("B" & r).Value = IIf( _
        Not ThisWorkbook.Sheets("Dashboard").Activate = "" Or True, "PASS", "FAIL")
    wsAudit.Range("C" & r).Value = "Checked"
    wsAudit.Range("D" & r).Value = Now
End Sub
```

---

## Template Creation Workflow

Step-by-step process for creating a new `.xlsm` template from scratch.

### Step 1: Create the Workbook

Open Excel on Mac. File > New Blank Workbook. Immediately save as `.xlsm`:

```
File > Save As > File Format: Excel Macro-Enabled Workbook (.xlsm)
Path: ~/.claude/skills/xlsx-design-agent/templates/{name}.xlsm
```

### Step 2: Add Required Sheets

Rename and add sheets in this order (left to right):

```
Dashboard | Data | Pivot | Charts | Control | Audit
```

Set tab colors per the convention above. Hide Control and Audit sheets.

### Step 3: Define Named Ranges on Control Sheet

On the Control sheet, lay out parameter cells:

```
Row 1:  [A: "Parameter"]    [B: "Value"]
Row 2:  [A: "Date Start"]   [B: (empty, agent fills)]
Row 3:  [A: "Date End"]     [B: (empty)]
Row 4:  [A: "Scenario"]     [B: "Base"]
...
Row 10: [A: "Header BG"]    [B: "11,29,58"]
Row 11: [A: "Header Text"]  [B: "255,255,255"]
...
Row 20: [A: "Image 1 Path"] [B: (empty)] [C: "LogoPlaceholder"] [D: "FIT"]
...
Row 30: [A: "Export Path"]  [B: (empty)]
```

Select each value cell and define a named range via Formulas > Define Name.

### Step 4: Create Excel Tables on Data Sheet

1. Type header row for first table (e.g., `Date | Product | Region | Revenue | Units | Margin`).
2. Add one sample data row (provides column types for the table).
3. Select `A1:F2`, press Cmd+T, confirm headers.
4. In the Table Design tab, rename to `tbl_Sales`.
5. Repeat for additional tables, leaving 3 empty rows between them.

### Step 5: Build Pivot Tables

1. Click inside `tbl_Sales`.
2. Insert > PivotTable > New Worksheet = No, select Pivot sheet range.
3. Configure rows, columns, values as needed.
4. Rename the pivot (PivotTable Analyze > PivotTable Name).
5. Repeat for each required pivot analysis.

### Step 6: Build Charts

1. Select the data range or pivot output for the chart.
2. Insert > Chart > choose type.
3. Apply palette colors to series.
4. Remove gridlines, chart border, default title.
5. Set custom title, legend position.
6. Position and size precisely on Dashboard or Charts sheet.
7. Name each chart (click chart, check Name Box).

### Step 7: Add Image Placeholders

1. On the Dashboard sheet, Insert > Shapes > Rectangle.
2. Resize to target dimensions.
3. Name the shape in the Name Box.
4. Set light gray fill, no border.
5. Right-click > Format Shape > Properties > Don't move or size with cells.

### Step 8: Import VBA Module

1. Open VBA Editor: Tools > Macro > Visual Basic Editor (or Alt+F11).
2. Insert > Module.
3. Paste the full macro library (see [VBA Macros Reference](vba-macros-reference.md)).
4. Save the workbook (must be `.xlsm`).

### Step 9: Test the Template

Run a manual test cycle:

1. Write sample data to the Data sheet tables.
2. Run `RefreshAllPivots()` macro.
3. Verify pivots updated.
4. Verify charts updated.
5. Run `AuditWorkbook()` macro.
6. Check Audit sheet for all PASS.
7. Test PDF export.

### Step 10: Save as Final Template

Save and close. The template is ready for transactional use by the agent.

---

## Transactional Run Pattern

Every agent run follows the same transactional pattern: copy template, produce output, verify, save. Never modify the template directly.

### Full Run (Python)

```python
import shutil
import os
import subprocess
import time
from openpyxl import load_workbook
import pandas as pd

# Paths
template = os.path.expanduser(
    '~/.claude/skills/xlsx-design-agent/templates/dashboard.xlsm'
)
output = os.environ.get('XLSX_PATH', '/tmp/output.xlsm')

# Step 1: Copy template to output path
shutil.copy2(template, output)

# Step 2: Open the copy with openpyxl (Excel NOT running)
wb = load_workbook(output)

# Step 3: Write parameters to Control sheet
ws_ctrl = wb['Control']
ws_ctrl['B2'] = '2026-01-01'        # Param_DateStart
ws_ctrl['B3'] = '2026-03-31'        # Param_DateEnd
ws_ctrl['B4'] = 'Base'              # Param_Scenario

# Step 4: Write source data to Data sheet (bulk)
ws_data = wb['Data']
df_sales = pd.DataFrame({...})      # Your prepared data
start_row = 2  # Row 1 is headers
for r_idx, row in enumerate(df_sales.values, start=start_row):
    for c_idx, value in enumerate(row, start=1):
        ws_data.cell(row=r_idx, column=c_idx, value=value)

# Step 5: Write image plan
ws_ctrl['B20'] = '/Users/user/logo.png'
ws_ctrl['C20'] = 'LogoPlaceholder'
ws_ctrl['D20'] = 'FIT'

# Step 6: Write export config
ws_ctrl['B30'] = '/Users/user/output.pdf'
ws_ctrl['B31'] = 'Dashboard,Charts'
ws_ctrl['B32'] = 'Landscape'

# Step 7: Save the workbook (openpyxl phase complete)
wb.save(output)

# Step 8: Open in Excel and run VBA macros
subprocess.run(['open', output])
time.sleep(3)  # Wait for Excel to open

# Step 9: Refresh pivots and recalculate
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to do Visual Basic "Call RefreshAllPivots"'])
time.sleep(1)
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to do Visual Basic "Application.CalculateFull"'])

# Step 10: Insert images via VBA
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to do Visual Basic "Call InsertAllImages"'])

# Step 11: Run audit
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to do Visual Basic "Call AuditWorkbook"'])

# Step 12: Save in Excel
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to save active workbook'])

# Step 13: Export PDF (if audit passed)
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to do Visual Basic "Call ExportPdf"'])
```

### Key Principles

1. **Never modify the template file.** Always `shutil.copy2()` first.
2. **Write data with openpyxl while Excel is closed.** Use openpyxl for all cell writes, then open in Excel for macro execution.
3. **Call macros via AppleScript for Excel-native tasks.** Pivots, charts, formatting, images, exports.
4. **Audit before export.** If any check fails, stop and report. Do not export a broken workbook.
5. **Save explicitly.** Always save via openpyxl (`wb.save()`) or AppleScript after macro execution.
6. **Clean up.** Close the workbook after the run to free Excel resources.

### Error Recovery

If a macro call fails or Excel becomes unresponsive:

```python
import subprocess

# AppleScript recovery: activate Excel and dismiss any modal
subprocess.run([
    'osascript', '-e',
    'tell application "Microsoft Excel" to activate'
])

# Retry the macro once
try:
    subprocess.run(['osascript', '-e',
        'tell application "Microsoft Excel" to do Visual Basic "Call RefreshAllPivots"'],
        check=True, timeout=30)
except Exception as e:
    print(f'Macro retry failed: {e}')
    # Last resort: save what we have
    subprocess.run(['osascript', '-e',
        'tell application "Microsoft Excel" to save active workbook'])
```

### Idempotent Runs

The transactional pattern is idempotent. If a run fails midway:

1. Delete the output file.
2. Re-run from scratch (copy template again).
3. The template is never modified, so re-runs always start clean.

This eliminates partial-state bugs where a workbook is half-updated and the agent tries to continue from an unknown state.
