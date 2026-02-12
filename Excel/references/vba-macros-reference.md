# VBA Macros Reference — Excel-Native API for macOS

The **essential companion** to openpyxl on macOS. VBA macros handle Excel-native operations that openpyxl cannot do after the file is open in Excel — pivot refresh, complex exports, and any operation requiring a live Excel instance.

## Table of Contents

1. [Why VBA Is Essential on macOS](#why-vba-is-essential-on-macos)
2. [How to Call VBA from Python](#how-to-call-vba-from-python)
3. [Macro Library — Workbook Management](#macro-library--workbook-management)
4. [Macro Library — Formatting](#macro-library--formatting)
5. [Macro Library — Charts](#macro-library--charts)
6. [Macro Library — Images](#macro-library--images)
7. [Macro Library — Export](#macro-library--export)
8. [Macro Library — Audit](#macro-library--audit)
9. [macOS-Specific VBA Notes](#macos-specific-vba-notes)
10. [Error Handling Patterns](#error-handling-patterns)
11. [Passing Palette Colors from Python to VBA](#passing-palette-colors-from-python-to-vba)

---

## Why VBA Is Essential on macOS

openpyxl handles file-based operations (create workbook, write data, format, charts, images, borders, conditional formatting) but cannot interact with a running Excel instance. VBA macros fill this gap — pivot refresh, recalculation, complex exports, and operations that require Excel's calculation engine.

| Operation | openpyxl | VBA Macro |
|-----------|----------|-----------|
| Borders on a range | Works (file-level) | Works (live Excel) |
| Chart series fill color | Works (file-level) | Full control (live) |
| Chart title font/size | Works (file-level) | Full control (live) |
| Tab (sheet) color | Works (file-level) | Works (live) |
| Pivot table refresh + style | Cannot refresh (no engine) | Full control |
| Image insert + lock aspect | Works (file-level) | Reliable (live) |
| PDF export with options | Not supported | Full page setup |
| Conditional formatting icons | Works (file-level) | Full icon sets |
| Recalculate formulas | Not supported (no engine) | Full control |

**Rule of thumb:** If the operation needs Excel's calculation engine running (pivot refresh, recalculate, PDF export with live data), use VBA. If it's file-level creation (write data, format cells, add charts, borders, images, conditional formatting), use openpyxl.

---

## How to Call VBA from Python

> **Note:** All examples use the `run_vba_macro()` helper from [AppleScript Patterns](applescript-patterns.md).

### Option 1: AppleScript `do Visual Basic` — Primary Method

The primary way to call VBA macros from Python on macOS. Works with macros embedded in `.xlsm` workbooks or for one-off VBA execution.

**Simple call (no arguments):**

```python
import subprocess

def run_vba_macro(macro_name, *args):
    """Call a VBA macro via AppleScript. See applescript-patterns.md for full helper."""
    if args:
        arg_str = ', '.join(
            f'"{a}"' if isinstance(a, str) else str(a) for a in args
        )
        vba_call = f"Call {macro_name}({arg_str})"
    else:
        vba_call = f"Call {macro_name}"

    applescript = f'''
    tell application "Microsoft Excel"
        do Visual Basic "{vba_call}"
    end tell
    '''
    subprocess.run(['osascript', '-e', applescript], check=True)

# Call macro with no arguments
run_vba_macro('ResetWorkbook')

# Call macro with arguments
run_vba_macro('ApplyTableFormatting', 'Sheet1', 2, 50, 1, 8, 11, 29, 58, 241, 245, 249)

# Call macro with string arguments
run_vba_macro('SetColumnWidths', 'Dashboard', 'A:4,B:20,C:15,D:12,E:12')
```

**Direct subprocess call (without helper):**

```python
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to do Visual Basic "Call ResetWorkbook"'],
    check=True)
```

**One-off inline VBA execution:**

```python
import subprocess

vba_code = '''
Sub TempMacro()
    Dim ws As Worksheet
    Set ws = ActiveWorkbook.Sheets("Dashboard")
    ws.Range("A1:H1").Interior.Color = RGB(11, 29, 58)
    ws.Range("A1:H1").Font.Color = RGB(255, 255, 255)
    ws.Range("A1:H1").Font.Bold = True
End Sub
'''

# Execute via AppleScript
applescript = f'''
tell application "Microsoft Excel"
    do Visual Basic "{vba_code.replace('"', '\\"').replace(chr(10), '\\n')}"
end tell
'''
subprocess.run(['osascript', '-e', applescript], check=True)
```

**Note:** For inline VBA with `do Visual Basic`, string escaping can be tricky. For complex macros, prefer the template approach (Option 2) with macros already embedded in the `.xlsm` file.

### Option 2: Template `.xlsm` Approach — Most Reliable

1. Create a `.xlsm` template with all macros in a standard VBA module.
2. Python copies the template to the output path.
3. Open with openpyxl, write data, save, close openpyxl.
4. Open in Excel via AppleScript.
5. Run VBA macros via AppleScript `do Visual Basic`.
6. Save and close via AppleScript.

```python
import shutil
from openpyxl import load_workbook

template = '/path/to/template.xlsm'
output = '/tmp/output.xlsm'
shutil.copy2(template, output)

# Step 1: Write data with openpyxl (file-level operations)
wb = load_workbook(output, keep_vba=True)
ws = wb['Data']
for row_idx, row_data in enumerate(data, start=2):
    for col_idx, value in enumerate(row_data, start=1):
        ws.cell(row=row_idx, column=col_idx, value=value)
wb.save(output)
wb.close()

# Step 2: Open in Excel and run VBA macros via AppleScript
import subprocess

# Open workbook in Excel
subprocess.run(['osascript', '-e', f'''
tell application "Microsoft Excel"
    activate
    open "{output}"
end tell
'''], check=True)

import time
time.sleep(2)  # Wait for Excel to open the file

# Call macros
run_vba_macro('WriteModePrep')
run_vba_macro('ApplyTableFormatting', 'Data', 2, 100, 1, 8, 11, 29, 58, 241, 245, 249)
run_vba_macro('CreateBarChart', 'Dashboard', 'Data!A1:B10', 50, 50, 400, 250, 'Revenue by Quarter', 59, 130, 246)
run_vba_macro('WriteModeEnd')
run_vba_macro('AuditWorkbook')

# Save and close via AppleScript
subprocess.run(['osascript', '-e', '''
tell application "Microsoft Excel"
    save active workbook
    close active workbook
end tell
'''], check=True)
```

### Option 3: Inject VBA Module via AppleScript

For dynamic macro injection when you do not have a template:

```python
import subprocess
import tempfile
import os

def inject_vba_module(workbook_path, module_name, vba_code):
    """Inject a VBA module into an open workbook via AppleScript."""
    # Write VBA to temp file
    tmp = tempfile.NamedTemporaryFile(mode='w', suffix='.bas', delete=False)
    tmp.write(vba_code)
    tmp.close()

    applescript = f'''
    tell application "Microsoft Excel"
        activate
        tell active workbook
            set vbProj to VBProject
            tell vbProj
                import file "{tmp.name}"
            end tell
        end tell
    end tell
    '''
    try:
        subprocess.run(['osascript', '-e', applescript], check=True, capture_output=True)
    finally:
        os.unlink(tmp.name)
```

**Note:** This requires "Trust access to the VBA project object model" to be enabled in Excel preferences (Trust Center > Macro Settings). Not always available. Prefer the template approach.

---

## Macro Library -- Workbook Management

### `ResetWorkbook()`

Clears data ranges, removes filters, resets scroll position. Use before a fresh data load.

```vba
Sub ResetWorkbook()
    On Error Resume Next
    Dim ws As Worksheet
    Dim lo As ListObject

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each ws In ThisWorkbook.Worksheets
        ' Remove AutoFilter
        If ws.AutoFilterMode Then ws.AutoFilterMode = False

        ' Clear ListObject data (preserve headers)
        For Each lo In ws.ListObjects
            If lo.DataBodyRange Is Nothing Then
                ' Table has no data rows, skip
            Else
                lo.DataBodyRange.Delete
            End If
        Next lo

        ' Reset scroll position
        ws.Activate
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
        ws.Range("A1").Select
    Next ws

    ' Re-enable
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    On Error GoTo 0
End Sub
```

**Python call:**
```python
run_vba_macro('ResetWorkbook')
# Or directly:
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to do Visual Basic "Call ResetWorkbook"'])
```

### `SanityCheck()`

Verifies that expected sheets, tables, and named ranges exist. Writes results to an Audit sheet.

```vba
Function SanityCheck() As Boolean
    On Error Resume Next
    Dim auditWs As Worksheet
    Dim row As Long
    Dim allPass As Boolean
    allPass = True
    row = 1

    ' Create or clear Audit sheet
    Set auditWs = Nothing
    Set auditWs = ThisWorkbook.Sheets("Audit")
    If auditWs Is Nothing Then
        Set auditWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        auditWs.Name = "Audit"
    Else
        auditWs.Cells.Clear
    End If

    ' Header
    auditWs.Range("A1").Value = "Check"
    auditWs.Range("B1").Value = "Item"
    auditWs.Range("C1").Value = "Status"
    auditWs.Range("D1").Value = "Details"
    auditWs.Range("A1:D1").Font.Bold = True
    row = 2

    ' Check expected sheets (read from Control sheet if exists)
    Dim controlWs As Worksheet
    Set controlWs = Nothing
    Set controlWs = ThisWorkbook.Sheets("Control")

    If Not controlWs Is Nothing Then
        Dim expectedSheets As String
        expectedSheets = "" & controlWs.Range("ExpectedSheets").Value
        If Len(expectedSheets) > 0 Then
            Dim sheetNames() As String
            sheetNames = Split(expectedSheets, ",")
            Dim i As Long
            For i = LBound(sheetNames) To UBound(sheetNames)
                Dim trimName As String
                trimName = Trim(sheetNames(i))
                Dim testWs As Worksheet
                Set testWs = Nothing
                Set testWs = ThisWorkbook.Sheets(trimName)
                auditWs.Cells(row, 1).Value = "Sheet"
                auditWs.Cells(row, 2).Value = trimName
                If testWs Is Nothing Then
                    auditWs.Cells(row, 3).Value = "FAIL"
                    auditWs.Cells(row, 4).Value = "Sheet not found"
                    allPass = False
                Else
                    auditWs.Cells(row, 3).Value = "PASS"
                End If
                row = row + 1
            Next i
        End If
    End If

    ' Check named ranges
    Dim nm As Name
    For Each nm In ThisWorkbook.Names
        auditWs.Cells(row, 1).Value = "NamedRange"
        auditWs.Cells(row, 2).Value = nm.Name
        Dim testVal As Variant
        testVal = Empty
        testVal = nm.RefersToRange.Value
        If Err.Number <> 0 Then
            auditWs.Cells(row, 3).Value = "FAIL"
            auditWs.Cells(row, 4).Value = "Ref error: " & Err.Description
            allPass = False
            Err.Clear
        Else
            auditWs.Cells(row, 3).Value = "PASS"
            auditWs.Cells(row, 4).Value = nm.RefersTo
        End If
        row = row + 1
    Next nm

    ' Summary
    auditWs.Cells(row + 1, 1).Value = "OVERALL"
    auditWs.Cells(row + 1, 3).Value = IIf(allPass, "ALL PASS", "FAILURES FOUND")
    auditWs.Cells(row + 1, 3).Font.Bold = True

    SanityCheck = allPass
    On Error GoTo 0
End Function
```

**Python call:**
```python
run_vba_macro('SanityCheck')
```

### `WriteModePrep()`

Disables screen updating, events, and sets manual calculation for bulk write performance.

```vba
Sub WriteModePrep()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
End Sub
```

**Python call:**
```python
run_vba_macro('WriteModePrep')
```

### `WriteModeEnd()`

Re-enables everything and forces a full recalculation.

```vba
Sub WriteModeEnd()
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.Calculate
End Sub
```

**Python call:**
```python
run_vba_macro('WriteModeEnd')
```

---

## Macro Library -- Formatting

### `ApplyTableFormatting()`

Full table formatting: header row with background color and white bold text, banded rows, thin borders, proper fonts. This is the workhorse formatting macro.

```vba
Sub ApplyTableFormatting(sheetName As String, startRow As Long, endRow As Long, _
                         startCol As Long, endCol As Long, _
                         headerBgR As Long, headerBgG As Long, headerBgB As Long, _
                         accentR As Long, accentG As Long, accentB As Long)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim headerRng As Range
    Dim dataRng As Range
    Dim fullRng As Range

    Set headerRng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, endCol))
    Set fullRng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(endRow, endCol))

    ' --- Header row ---
    With headerRng
        .Interior.Color = RGB(headerBgR, headerBgG, headerBgB)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 11
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 32
    End With

    ' --- Data rows ---
    Dim r As Long
    For r = startRow + 1 To endRow
        Dim rowRng As Range
        Set rowRng = ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol))

        ' Banded rows (even rows get accent color)
        If (r - startRow) Mod 2 = 0 Then
            rowRng.Interior.Color = RGB(accentR, accentG, accentB)
        Else
            rowRng.Interior.Color = RGB(255, 255, 255)
        End If

        ' Data font
        rowRng.Font.Size = 10
        rowRng.Font.Name = "Calibri"
        rowRng.Font.Color = RGB(30, 41, 59)
        rowRng.RowHeight = 20
    Next r

    ' --- Borders on full range ---
    With fullRng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(226, 232, 240)
    End With

    ' Thicker bottom border on header
    With headerRng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(headerBgR, headerBgG, headerBgB)
    End With

    ' Outer border on full range
    With fullRng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(headerBgR, headerBgG, headerBgB)
    End With
    With fullRng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(226, 232, 240)
    End With
    With fullRng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(226, 232, 240)
    End With
    With fullRng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(226, 232, 240)
    End With

    Exit Sub
ErrHandler:
    Debug.Print "ApplyTableFormatting error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('ApplyTableFormatting', 'Sheet1', 2, 50, 1, 8, 11, 29, 58, 241, 245, 249)
#                                      sheet   sRow eRow sC  eC  hdrR hdrG hdrB altR altG altB
```

### `ApplyKPIPanel()`

Styles a merged cell region as a KPI display panel with background color and large bold text.

```vba
Sub ApplyKPIPanel(sheetName As String, row As Long, col As Long, endCol As Long, _
                  bgR As Long, bgG As Long, bgB As Long, _
                  textR As Long, textG As Long, textB As Long)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(row, col), ws.Cells(row, endCol))

    With rng
        .Merge
        .Interior.Color = RGB(bgR, bgG, bgB)
        .Font.Color = RGB(textR, textG, textB)
        .Font.Bold = True
        .Font.Size = 28
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 60
    End With

    ' Subtle bottom accent border
    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = RGB(textR, textG, textB)
    End With

    Exit Sub
ErrHandler:
    Debug.Print "ApplyKPIPanel error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('ApplyKPIPanel', 'Dashboard', 3, 2, 4, 11, 29, 58, 201, 168, 76)
#                               sheet      row col endC bgR bgG bgB txtR txtG txtB
```

### `ApplyBorders()`

Applies borders to any rectangular range. Supports custom border color.

```vba
Sub ApplyBorders(sheetName As String, startRow As Long, endRow As Long, _
                 startCol As Long, endCol As Long, _
                 borderR As Long, borderG As Long, borderB As Long)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(endRow, endCol))

    Dim borderColor As Long
    borderColor = RGB(borderR, borderG, borderB)

    ' All inner and outer borders
    Dim edge As Variant
    For Each edge In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom, xlInsideVertical, xlInsideHorizontal)
        With rng.Borders(edge)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = borderColor
        End With
    Next edge

    Exit Sub
ErrHandler:
    Debug.Print "ApplyBorders error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('ApplyBorders', 'Sheet1', 1, 50, 1, 8, 226, 232, 240)
```

### `HideGridlines()`

Hides gridlines on a specific sheet for a clean dashboard look.

```vba
Sub HideGridlines(sheetName As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ws.Activate
    ActiveWindow.DisplayGridlines = False

    Exit Sub
ErrHandler:
    Debug.Print "HideGridlines error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('HideGridlines', 'Dashboard')
```

### `SetColumnWidths()`

Sets column widths from a comma-separated string. Format: `"A:12,B:20,C:15"`.

```vba
Sub SetColumnWidths(sheetName As String, widthsString As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim pairs() As String
    pairs = Split(widthsString, ",")

    Dim i As Long
    For i = LBound(pairs) To UBound(pairs)
        Dim parts() As String
        parts = Split(Trim(pairs(i)), ":")
        If UBound(parts) >= 1 Then
            Dim colLetter As String
            Dim colWidth As Double
            colLetter = Trim(parts(0))
            colWidth = CDbl(Trim(parts(1)))
            ws.Columns(colLetter).ColumnWidth = colWidth
        End If
    Next i

    Exit Sub
ErrHandler:
    Debug.Print "SetColumnWidths error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('SetColumnWidths', 'Dashboard', 'A:4,B:20,C:15,D:12,E:12,F:12,G:15,H:4')
```

### `SetRowHeight()`

Sets the height of a specific row.

```vba
Sub SetRowHeight(sheetName As String, row As Long, height As Double)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Rows(row).RowHeight = height
    Exit Sub
ErrHandler:
    Debug.Print "SetRowHeight error: " & Err.Description
End Sub
```

### `SetSheetTabColor()`

Sets the sheet tab color.

```vba
Sub SetSheetTabColor(sheetName As String, r As Long, g As Long, b As Long)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Tab.Color = RGB(r, g, b)
    Exit Sub
ErrHandler:
    Debug.Print "SetSheetTabColor error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('SetSheetTabColor', 'Dashboard', 59, 130, 246)
```

### `ApplyHeaderBar()`

Applies a full-width colored header bar to a single row (useful for section dividers on dashboards).

```vba
Sub ApplyHeaderBar(sheetName As String, row As Long, startCol As Long, endCol As Long, _
                   text As String, bgR As Long, bgG As Long, bgB As Long, _
                   textR As Long, textG As Long, textB As Long, fontSize As Long)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim rng As Range
    Set rng = ws.Range(ws.Cells(row, startCol), ws.Cells(row, endCol))

    ' Merge and style
    rng.Merge
    ws.Cells(row, startCol).Value = text

    With rng
        .Interior.Color = RGB(bgR, bgG, bgB)
        .Font.Color = RGB(textR, textG, textB)
        .Font.Bold = True
        .Font.Size = fontSize
        .Font.Name = "Calibri"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .RowHeight = 36
        .IndentLevel = 1
    End With

    Exit Sub
ErrHandler:
    Debug.Print "ApplyHeaderBar error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('ApplyHeaderBar', 'Dashboard', 10, 1, 8, 'Revenue Breakdown', 11, 29, 58, 255, 255, 255, 13)
```

---

## Macro Library -- Charts

### `CreateBarChart()`

Creates a bar (column) chart from a data range with custom color, title, and positioning.

```vba
Sub CreateBarChart(sheetName As String, dataRange As String, _
                   chartLeft As Double, chartTop As Double, _
                   chartWidth As Double, chartHeight As Double, _
                   chartTitle As String, _
                   color1R As Long, color1G As Long, color1B As Long)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Determine data source sheet and range
    Dim srcRange As Range
    If InStr(dataRange, "!") > 0 Then
        ' Cross-sheet reference like "Data!A1:B10"
        Dim sheetPart As String
        Dim rangePart As String
        sheetPart = Left(dataRange, InStr(dataRange, "!") - 1)
        rangePart = Mid(dataRange, InStr(dataRange, "!") + 1)
        Set srcRange = ThisWorkbook.Sheets(sheetPart).Range(rangePart)
    Else
        Set srcRange = ws.Range(dataRange)
    End If

    ' Create chart object on target sheet
    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)

    With co.Chart
        .ChartType = xlColumnClustered
        .SetSourceData srcRange
        .HasTitle = True
        .ChartTitle.text = chartTitle

        ' Style the title
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(30, 41, 59)
        .ChartTitle.Font.Name = "Calibri"

        ' Color the first series
        If .SeriesCollection.Count >= 1 Then
            .SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(color1R, color1G, color1B)
            .SeriesCollection(1).Format.Line.ForeColor.RGB = RGB(color1R, color1G, color1B)
        End If

        ' Remove legend if single series
        If .SeriesCollection.Count = 1 Then
            .HasLegend = False
        End If

        ' Clean up axes
        .Axes(xlCategory).TickLabels.Font.Size = 9
        .Axes(xlCategory).TickLabels.Font.Color = RGB(100, 116, 139)
        .Axes(xlValue).TickLabels.Font.Size = 9
        .Axes(xlValue).TickLabels.Font.Color = RGB(100, 116, 139)
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(226, 232, 240)

        ' Remove chart border
        .ChartArea.Format.Line.Visible = msoFalse

        ' White background
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

    Exit Sub
ErrHandler:
    Debug.Print "CreateBarChart error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('CreateBarChart', 'Dashboard', 'Data!A1:B10', 50, 200, 400, 250,
              'Revenue by Quarter', 59, 130, 246)
```

### `CreateLineChart()`

Creates a line chart with support for multiple series, each with its own color.

```vba
Sub CreateLineChart(sheetName As String, dataRange As String, _
                    chartLeft As Double, chartTop As Double, _
                    chartWidth As Double, chartHeight As Double, _
                    chartTitle As String, _
                    colorsString As String)
    ' colorsString format: "R,G,B;R,G,B;R,G,B" — one RGB triplet per series separated by semicolons
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim srcRange As Range
    If InStr(dataRange, "!") > 0 Then
        Dim sheetPart As String
        Dim rangePart As String
        sheetPart = Left(dataRange, InStr(dataRange, "!") - 1)
        rangePart = Mid(dataRange, InStr(dataRange, "!") + 1)
        Set srcRange = ThisWorkbook.Sheets(sheetPart).Range(rangePart)
    Else
        Set srcRange = ws.Range(dataRange)
    End If

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)

    With co.Chart
        .ChartType = xlLine
        .SetSourceData srcRange
        .HasTitle = True
        .ChartTitle.text = chartTitle
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(30, 41, 59)
        .ChartTitle.Font.Name = "Calibri"

        ' Apply per-series colors
        If Len(colorsString) > 0 Then
            Dim colorSets() As String
            colorSets = Split(colorsString, ";")
            Dim s As Long
            For s = 0 To UBound(colorSets)
                If s + 1 <= .SeriesCollection.Count Then
                    Dim rgb_parts() As String
                    rgb_parts = Split(Trim(colorSets(s)), ",")
                    If UBound(rgb_parts) >= 2 Then
                        Dim cr As Long, cg As Long, cb As Long
                        cr = CLng(Trim(rgb_parts(0)))
                        cg = CLng(Trim(rgb_parts(1)))
                        cb = CLng(Trim(rgb_parts(2)))
                        .SeriesCollection(s + 1).Format.Line.ForeColor.RGB = RGB(cr, cg, cb)
                        .SeriesCollection(s + 1).Format.Line.Weight = 2.5
                        .SeriesCollection(s + 1).MarkerStyle = xlMarkerStyleCircle
                        .SeriesCollection(s + 1).MarkerSize = 6
                        .SeriesCollection(s + 1).MarkerForegroundColor = RGB(cr, cg, cb)
                        .SeriesCollection(s + 1).MarkerBackgroundColor = RGB(255, 255, 255)
                    End If
                End If
            Next s
        End If

        ' Gridlines
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(226, 232, 240)
        .Axes(xlCategory).TickLabels.Font.Size = 9
        .Axes(xlCategory).TickLabels.Font.Color = RGB(100, 116, 139)
        .Axes(xlValue).TickLabels.Font.Size = 9
        .Axes(xlValue).TickLabels.Font.Color = RGB(100, 116, 139)

        ' Clean border and background
        .ChartArea.Format.Line.Visible = msoFalse
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)

        ' Legend at bottom
        If .SeriesCollection.Count > 1 Then
            .HasLegend = True
            .Legend.Position = xlLegendPositionBottom
            .Legend.Font.Size = 9
            .Legend.Font.Color = RGB(100, 116, 139)
        Else
            .HasLegend = False
        End If
    End With

    Exit Sub
ErrHandler:
    Debug.Print "CreateLineChart error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('CreateLineChart', 'Dashboard', 'Data!A1:D13', 50, 300, 500, 280,
              'Monthly Trends',
              '59,130,246;139,92,246;16,185,129')
#              series1=blue  series2=purple  series3=teal
```

### `CreateDoughnutChart()`

Creates a doughnut chart with per-slice colors.

```vba
Sub CreateDoughnutChart(sheetName As String, dataRange As String, _
                        chartLeft As Double, chartTop As Double, _
                        chartWidth As Double, chartHeight As Double, _
                        chartTitle As String, _
                        colorsString As String)
    ' colorsString format: "R,G,B;R,G,B;R,G,B" — one RGB triplet per data point (slice)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim srcRange As Range
    If InStr(dataRange, "!") > 0 Then
        Dim sheetPart As String
        Dim rangePart As String
        sheetPart = Left(dataRange, InStr(dataRange, "!") - 1)
        rangePart = Mid(dataRange, InStr(dataRange, "!") + 1)
        Set srcRange = ThisWorkbook.Sheets(sheetPart).Range(rangePart)
    Else
        Set srcRange = ws.Range(dataRange)
    End If

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)

    With co.Chart
        .ChartType = xlDoughnut
        .SetSourceData srcRange
        .HasTitle = True
        .ChartTitle.text = chartTitle
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(30, 41, 59)
        .ChartTitle.Font.Name = "Calibri"

        ' Color each slice
        If Len(colorsString) > 0 And .SeriesCollection.Count >= 1 Then
            Dim colorSets() As String
            colorSets = Split(colorsString, ";")
            Dim pt As Long
            For pt = 0 To UBound(colorSets)
                If pt + 1 <= .SeriesCollection(1).Points.Count Then
                    Dim rgb_parts() As String
                    rgb_parts = Split(Trim(colorSets(pt)), ",")
                    If UBound(rgb_parts) >= 2 Then
                        Dim cr As Long, cg As Long, cb As Long
                        cr = CLng(Trim(rgb_parts(0)))
                        cg = CLng(Trim(rgb_parts(1)))
                        cb = CLng(Trim(rgb_parts(2)))
                        .SeriesCollection(1).Points(pt + 1).Format.Fill.ForeColor.RGB = RGB(cr, cg, cb)
                    End If
                End If
            Next pt
        End If

        ' Doughnut hole size
        .SeriesCollection(1).DoughnutHoleSize = 55

        ' Legend
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 9
        .Legend.Font.Color = RGB(100, 116, 139)

        ' Clean frame
        .ChartArea.Format.Line.Visible = msoFalse
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

    Exit Sub
ErrHandler:
    Debug.Print "CreateDoughnutChart error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('CreateDoughnutChart', 'Dashboard', 'Data!A1:B5', 470, 200, 280, 250,
              'Revenue Split',
              '59,130,246;139,92,246;16,185,129;201,168,76;100,116,139')
```

### `CreatePieChart()`

Creates a pie chart with per-slice colors and optional data labels.

```vba
Sub CreatePieChart(sheetName As String, dataRange As String, _
                   chartLeft As Double, chartTop As Double, _
                   chartWidth As Double, chartHeight As Double, _
                   chartTitle As String, _
                   colorsString As String, _
                   showLabels As Boolean)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim srcRange As Range
    If InStr(dataRange, "!") > 0 Then
        Dim sheetPart As String
        Dim rangePart As String
        sheetPart = Left(dataRange, InStr(dataRange, "!") - 1)
        rangePart = Mid(dataRange, InStr(dataRange, "!") + 1)
        Set srcRange = ThisWorkbook.Sheets(sheetPart).Range(rangePart)
    Else
        Set srcRange = ws.Range(dataRange)
    End If

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)

    With co.Chart
        .ChartType = xlPie
        .SetSourceData srcRange
        .HasTitle = True
        .ChartTitle.text = chartTitle
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(30, 41, 59)
        .ChartTitle.Font.Name = "Calibri"

        ' Color each slice
        If Len(colorsString) > 0 And .SeriesCollection.Count >= 1 Then
            Dim colorSets() As String
            colorSets = Split(colorsString, ";")
            Dim pt As Long
            For pt = 0 To UBound(colorSets)
                If pt + 1 <= .SeriesCollection(1).Points.Count Then
                    Dim rgb_parts() As String
                    rgb_parts = Split(Trim(colorSets(pt)), ",")
                    If UBound(rgb_parts) >= 2 Then
                        .SeriesCollection(1).Points(pt + 1).Format.Fill.ForeColor.RGB = _
                            RGB(CLng(Trim(rgb_parts(0))), CLng(Trim(rgb_parts(1))), CLng(Trim(rgb_parts(2))))
                    End If
                End If
            Next pt
        End If

        ' Data labels
        If showLabels Then
            .SeriesCollection(1).HasDataLabels = True
            With .SeriesCollection(1).DataLabels
                .ShowPercentage = True
                .ShowValue = False
                .ShowCategoryName = False
                .Font.Size = 9
                .Font.Color = RGB(30, 41, 59)
            End With
        End If

        ' Legend at bottom
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 9
        .Legend.Font.Color = RGB(100, 116, 139)

        ' Clean frame
        .ChartArea.Format.Line.Visible = msoFalse
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

    Exit Sub
ErrHandler:
    Debug.Print "CreatePieChart error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('CreatePieChart', 'Dashboard', 'Data!A1:B5', 50, 500, 280, 250,
              'Market Share',
              '59,130,246;139,92,246;16,185,129;201,168,76', True)
```

### `StyleAllCharts()`

Applies a consistent visual theme to every chart on a sheet: white background, clean borders, consistent fonts.

```vba
Sub StyleAllCharts(sheetName As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim co As ChartObject
    For Each co In ws.ChartObjects
        With co.Chart
            ' Background
            .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ChartArea.Format.Line.Visible = msoFalse
            .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)

            ' Title
            If .HasTitle Then
                .ChartTitle.Font.Name = "Calibri"
                .ChartTitle.Font.Size = 12
                .ChartTitle.Font.Bold = True
                .ChartTitle.Font.Color = RGB(30, 41, 59)
            End If

            ' Legend
            If .HasLegend Then
                .Legend.Font.Name = "Calibri"
                .Legend.Font.Size = 9
                .Legend.Font.Color = RGB(100, 116, 139)
            End If

            ' Axes (only for charts that have them)
            On Error Resume Next
            If Not .Axes(xlCategory) Is Nothing Then
                .Axes(xlCategory).TickLabels.Font.Size = 9
                .Axes(xlCategory).TickLabels.Font.Color = RGB(100, 116, 139)
                .Axes(xlCategory).TickLabels.Font.Name = "Calibri"
            End If
            If Not .Axes(xlValue) Is Nothing Then
                .Axes(xlValue).TickLabels.Font.Size = 9
                .Axes(xlValue).TickLabels.Font.Color = RGB(100, 116, 139)
                .Axes(xlValue).TickLabels.Font.Name = "Calibri"
                .Axes(xlValue).HasMajorGridlines = True
                .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(226, 232, 240)
                .Axes(xlValue).MajorGridlines.Format.Line.Weight = 0.5
            End If
            On Error GoTo ErrHandler
        End With
    Next co

    Exit Sub
ErrHandler:
    Debug.Print "StyleAllCharts error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('StyleAllCharts', 'Dashboard')
```

### `RefreshAllPivots()`

Refreshes all pivot table caches in the workbook. Essential after writing new source data.

```vba
Sub RefreshAllPivots()
    On Error GoTo ErrHandler
    Dim pc As PivotCache
    For Each pc In ThisWorkbook.PivotCaches
        pc.Refresh
    Next pc

    ' Also refresh each pivot table individually (belt and suspenders)
    Dim ws As Worksheet
    Dim pt As PivotTable
    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
        Next pt
    Next ws

    Exit Sub
ErrHandler:
    Debug.Print "RefreshAllPivots error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('RefreshAllPivots')
```

### `SetChartTitle()`

Sets or changes a chart title by chart index (1-based) on a sheet.

```vba
Sub SetChartTitle(sheetName As String, chartIndex As Long, titleText As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    If chartIndex < 1 Or chartIndex > ws.ChartObjects.Count Then
        Debug.Print "SetChartTitle: chartIndex " & chartIndex & " out of range (1-" & ws.ChartObjects.Count & ")"
        Exit Sub
    End If

    With ws.ChartObjects(chartIndex).Chart
        .HasTitle = True
        .ChartTitle.text = titleText
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(30, 41, 59)
        .ChartTitle.Font.Name = "Calibri"
    End With

    Exit Sub
ErrHandler:
    Debug.Print "SetChartTitle error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('SetChartTitle', 'Dashboard', 1, 'Updated Revenue Chart')
```

### `SetChartSeriesColor()`

Changes the color of a specific series in a specific chart. Useful for post-creation color adjustments.

```vba
Sub SetChartSeriesColor(sheetName As String, chartIndex As Long, seriesIndex As Long, _
                        r As Long, g As Long, b As Long)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    With ws.ChartObjects(chartIndex).Chart.SeriesCollection(seriesIndex)
        .Format.Fill.ForeColor.RGB = RGB(r, g, b)
        .Format.Line.ForeColor.RGB = RGB(r, g, b)
    End With

    Exit Sub
ErrHandler:
    Debug.Print "SetChartSeriesColor error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('SetChartSeriesColor', 'Dashboard', 1, 1, 59, 130, 246)
```

### `CreateStackedBarChart()`

Creates a stacked bar chart with multiple series colors.

```vba
Sub CreateStackedBarChart(sheetName As String, dataRange As String, _
                          chartLeft As Double, chartTop As Double, _
                          chartWidth As Double, chartHeight As Double, _
                          chartTitle As String, colorsString As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim srcRange As Range
    If InStr(dataRange, "!") > 0 Then
        Dim sheetPart As String, rangePart As String
        sheetPart = Left(dataRange, InStr(dataRange, "!") - 1)
        rangePart = Mid(dataRange, InStr(dataRange, "!") + 1)
        Set srcRange = ThisWorkbook.Sheets(sheetPart).Range(rangePart)
    Else
        Set srcRange = ws.Range(dataRange)
    End If

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(chartLeft, chartTop, chartWidth, chartHeight)

    With co.Chart
        .ChartType = xlBarStacked
        .SetSourceData srcRange
        .HasTitle = True
        .ChartTitle.text = chartTitle
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        .ChartTitle.Font.Color = RGB(30, 41, 59)
        .ChartTitle.Font.Name = "Calibri"

        ' Apply per-series colors
        If Len(colorsString) > 0 Then
            Dim colorSets() As String
            colorSets = Split(colorsString, ";")
            Dim s As Long
            For s = 0 To UBound(colorSets)
                If s + 1 <= .SeriesCollection.Count Then
                    Dim rgb_parts() As String
                    rgb_parts = Split(Trim(colorSets(s)), ",")
                    If UBound(rgb_parts) >= 2 Then
                        .SeriesCollection(s + 1).Format.Fill.ForeColor.RGB = _
                            RGB(CLng(Trim(rgb_parts(0))), CLng(Trim(rgb_parts(1))), CLng(Trim(rgb_parts(2))))
                    End If
                End If
            Next s
        End If

        ' Gridlines and axes
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(226, 232, 240)
        .Axes(xlCategory).TickLabels.Font.Size = 9
        .Axes(xlCategory).TickLabels.Font.Color = RGB(100, 116, 139)
        .Axes(xlValue).TickLabels.Font.Size = 9
        .Axes(xlValue).TickLabels.Font.Color = RGB(100, 116, 139)

        ' Legend at bottom
        .HasLegend = True
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 9
        .Legend.Font.Color = RGB(100, 116, 139)

        ' Clean frame
        .ChartArea.Format.Line.Visible = msoFalse
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With

    Exit Sub
ErrHandler:
    Debug.Print "CreateStackedBarChart error: " & Err.Description
End Sub
```

---

## Macro Library -- Images

### `InsertImageIntoPlaceholder()`

Inserts an image at a specific position with either FIT (maintain aspect ratio, letterbox) or FILL (crop to fill) mode.

```vba
Sub InsertImageIntoPlaceholder(sheetName As String, imagePath As String, _
                               leftPt As Double, topPt As Double, _
                               widthPt As Double, heightPt As Double, _
                               mode As String)
    ' mode = "FIT" (maintain aspect, fit within bounds) or "FILL" (crop to fill bounds)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Insert the image
    Dim pic As Shape
    Set pic = ws.Shapes.AddPicture( _
        Filename:=imagePath, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=0, Top:=0, Width:=-1, Height:=-1)

    ' Get original dimensions
    pic.LockAspectRatio = msoFalse
    Dim origWidth As Double, origHeight As Double
    origWidth = pic.Width
    origHeight = pic.Height

    If UCase(mode) = "FIT" Then
        ' Maintain aspect ratio, fit within bounds
        pic.LockAspectRatio = msoTrue
        Dim scaleW As Double, scaleH As Double
        scaleW = widthPt / origWidth
        scaleH = heightPt / origHeight

        If scaleW < scaleH Then
            ' Width is the constraining dimension
            pic.Width = widthPt
            ' Center vertically
            pic.Left = leftPt
            pic.Top = topPt + (heightPt - pic.Height) / 2
        Else
            ' Height is the constraining dimension
            pic.Height = heightPt
            ' Center horizontally
            pic.Left = leftPt + (widthPt - pic.Width) / 2
            pic.Top = topPt
        End If

    ElseIf UCase(mode) = "FILL" Then
        ' Scale to fill, then crop overflow
        pic.LockAspectRatio = msoTrue
        scaleW = widthPt / origWidth
        scaleH = heightPt / origHeight

        If scaleW > scaleH Then
            ' Scale by width (will overflow height)
            pic.Width = widthPt
            pic.Left = leftPt
            pic.Top = topPt - (pic.Height - heightPt) / 2
        Else
            ' Scale by height (will overflow width)
            pic.Height = heightPt
            pic.Left = leftPt - (pic.Width - widthPt) / 2
            pic.Top = topPt
        End If

        ' Crop to exact bounds
        With pic.PictureFormat
            If pic.Width > widthPt Then
                Dim excessW As Double
                excessW = pic.Width - widthPt
                .CropLeft = excessW / 2
                .CropRight = excessW / 2
            End If
            If pic.Height > heightPt Then
                Dim excessH As Double
                excessH = pic.Height - heightPt
                .CropTop = excessH / 2
                .CropBottom = excessH / 2
            End If
        End With
        pic.Left = leftPt
        pic.Top = topPt

    Else
        ' Default: stretch to exact dimensions (no aspect lock)
        pic.LockAspectRatio = msoFalse
        pic.Width = widthPt
        pic.Height = heightPt
        pic.Left = leftPt
        pic.Top = topPt
    End If

    ' Name the shape for later reference
    pic.Name = "IMG_" & Replace(Replace(sheetName, " ", "_"), "'", "") & "_" & _
               CLng(leftPt) & "_" & CLng(topPt)

    Exit Sub
ErrHandler:
    Debug.Print "InsertImageIntoPlaceholder error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('InsertImageIntoPlaceholder',
    'Dashboard', '/tmp/logo.png',
    50, 20, 200, 80, 'FIT')
```

### `AuditImages()`

Checks all images on a sheet: verifies they are within the printable area, reports aspect ratios, and writes results to the Audit sheet.

```vba
Sub AuditImages(sheetName As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Get or create Audit sheet
    Dim auditWs As Worksheet
    On Error Resume Next
    Set auditWs = ThisWorkbook.Sheets("Audit")
    On Error GoTo ErrHandler
    If auditWs Is Nothing Then
        Set auditWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        auditWs.Name = "Audit"
    End If

    ' Find next empty row on Audit sheet
    Dim row As Long
    row = auditWs.Cells(auditWs.Rows.Count, 1).End(xlUp).row + 1
    If row < 2 Then row = 2

    ' Write section header
    auditWs.Cells(row, 1).Value = "--- Image Audit: " & sheetName & " ---"
    auditWs.Cells(row, 1).Font.Bold = True
    row = row + 1

    ' Column headers
    auditWs.Cells(row, 1).Value = "Shape Name"
    auditWs.Cells(row, 2).Value = "Left (pt)"
    auditWs.Cells(row, 3).Value = "Top (pt)"
    auditWs.Cells(row, 4).Value = "Width (pt)"
    auditWs.Cells(row, 5).Value = "Height (pt)"
    auditWs.Cells(row, 6).Value = "Aspect Ratio"
    auditWs.Cells(row, 7).Value = "In Bounds?"
    auditWs.Range(auditWs.Cells(row, 1), auditWs.Cells(row, 7)).Font.Bold = True
    row = row + 1

    ' Page width/height for bounds checking (approximate for Letter size)
    Dim pageWidth As Double, pageHeight As Double
    pageWidth = 720  ' ~10 inches at 72 dpi
    pageHeight = 936 ' ~13 inches at 72 dpi

    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            auditWs.Cells(row, 1).Value = shp.Name
            auditWs.Cells(row, 2).Value = Round(shp.Left, 1)
            auditWs.Cells(row, 3).Value = Round(shp.Top, 1)
            auditWs.Cells(row, 4).Value = Round(shp.Width, 1)
            auditWs.Cells(row, 5).Value = Round(shp.Height, 1)

            If shp.Height > 0 Then
                auditWs.Cells(row, 6).Value = Round(shp.Width / shp.Height, 3)
            Else
                auditWs.Cells(row, 6).Value = "N/A"
            End If

            ' Check bounds
            If shp.Left >= 0 And shp.Top >= 0 And _
               (shp.Left + shp.Width) <= pageWidth And _
               (shp.Top + shp.Height) <= pageHeight Then
                auditWs.Cells(row, 7).Value = "PASS"
            Else
                auditWs.Cells(row, 7).Value = "FAIL - out of bounds"
            End If
            row = row + 1
        End If
    Next shp

    If row = row Then
        ' No images found (row didn't advance)
        auditWs.Cells(row, 1).Value = "(no images found on this sheet)"
    End If

    Exit Sub
ErrHandler:
    Debug.Print "AuditImages error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('AuditImages', 'Dashboard')
```

### `RemoveAllImages()`

Removes all picture shapes from a sheet. Useful before re-inserting updated images.

```vba
Sub RemoveAllImages(sheetName As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If ws.Shapes(i).Type = msoPicture Or ws.Shapes(i).Type = msoLinkedPicture Then
            ws.Shapes(i).Delete
        End If
    Next i

    Exit Sub
ErrHandler:
    Debug.Print "RemoveAllImages error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('RemoveAllImages', 'Dashboard')
```

---

## Macro Library -- Export

### `ExportPdf()`

Exports the active sheet to PDF at the specified path.

```vba
Sub ExportPdf(outputPath As String)
    On Error GoTo ErrHandler

    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    Exit Sub
ErrHandler:
    Debug.Print "ExportPdf error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('ExportPdf', '/tmp/report.pdf')
```

### `ExportSheetAsPdf()`

Exports a specific named sheet to PDF.

```vba
Sub ExportSheetAsPdf(sheetName As String, outputPath As String)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    Exit Sub
ErrHandler:
    Debug.Print "ExportSheetAsPdf error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('ExportSheetAsPdf', 'Dashboard', '/tmp/dashboard.pdf')
```

### `ExportAllSheetsPdf()`

Exports the entire workbook (all sheets) to a single PDF.

```vba
Sub ExportAllSheetsPdf(outputPath As String)
    On Error GoTo ErrHandler

    ThisWorkbook.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=outputPath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=False

    Exit Sub
ErrHandler:
    Debug.Print "ExportAllSheetsPdf error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('ExportAllSheetsPdf', '/tmp/full_report.pdf')
```

### `SetPageSetup()`

Configures page setup for printing/PDF export: orientation, paper size, and fit-to-page scaling.

```vba
Sub SetPageSetup(sheetName As String, orientation As String, _
                 paperSize As String, fitWide As Long, fitTall As Long)
    ' orientation: "portrait" or "landscape"
    ' paperSize: "letter", "a4", "legal", "tabloid"
    ' fitWide: number of pages wide (0 = no fit)
    ' fitTall: number of pages tall (0 = no fit)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    With ws.PageSetup
        ' Orientation
        If LCase(orientation) = "landscape" Then
            .orientation = xlLandscape
        Else
            .orientation = xlPortrait
        End If

        ' Paper size
        Select Case LCase(paperSize)
            Case "letter"
                .paperSize = xlPaperLetter
            Case "a4"
                .paperSize = xlPaperA4
            Case "legal"
                .paperSize = xlPaperLegal
            Case "tabloid"
                .paperSize = xlPaperTabloid
            Case Else
                .paperSize = xlPaperLetter
        End Select

        ' Fit to page
        If fitWide > 0 Or fitTall > 0 Then
            .Zoom = False
            If fitWide > 0 Then .FitToPagesWide = fitWide
            If fitTall > 0 Then .FitToPagesTall = fitTall
        End If

        ' Standard margins (inches converted to points internally by Excel)
        .LeftMargin = Application.InchesToPoints(0.5)
        .RightMargin = Application.InchesToPoints(0.5)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .HeaderMargin = Application.InchesToPoints(0.3)
        .FooterMargin = Application.InchesToPoints(0.3)

        ' Center on page
        .CenterHorizontally = True
        .CenterVertically = False
    End With

    Exit Sub
ErrHandler:
    Debug.Print "SetPageSetup error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('SetPageSetup', 'Dashboard', 'landscape', 'letter', 1, 1)
#                              sheet       orient      paper    fitW fitT
```

### `SetPrintArea()`

Sets the print area for a sheet.

```vba
Sub SetPrintArea(sheetName As String, printRange As String)
    ' printRange: e.g. "A1:H50"
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    ws.PageSetup.PrintArea = printRange
    Exit Sub
ErrHandler:
    Debug.Print "SetPrintArea error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('SetPrintArea', 'Dashboard', 'A1:H50')
```

### `SetPrintTitles()`

Sets repeating rows and columns for multi-page print output.

```vba
Sub SetPrintTitles(sheetName As String, titleRows As String, titleCols As String)
    ' titleRows: e.g. "$1:$2" or "" for none
    ' titleCols: e.g. "$A:$A" or "" for none
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    With ws.PageSetup
        If Len(titleRows) > 0 Then .PrintTitleRows = titleRows
        If Len(titleCols) > 0 Then .PrintTitleColumns = titleCols
    End With

    Exit Sub
ErrHandler:
    Debug.Print "SetPrintTitles error: " & Err.Description
End Sub
```

**Python call:**
```python
run_vba_macro('SetPrintTitles', 'Data', '$1:$1', '')
```

---

## Macro Library -- Audit

### `AuditWorkbook()`

Comprehensive workbook audit: checks all sheets exist, all charts have titles, all images are within bounds, all named ranges resolve. Writes structured pass/fail results to the Audit sheet.

```vba
Sub AuditWorkbook()
    On Error Resume Next
    Dim auditWs As Worksheet
    Dim row As Long
    Dim passCount As Long, failCount As Long
    passCount = 0
    failCount = 0

    ' Create or clear Audit sheet
    Set auditWs = Nothing
    Set auditWs = ThisWorkbook.Sheets("Audit")
    If auditWs Is Nothing Then
        Set auditWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        auditWs.Name = "Audit"
    Else
        auditWs.Cells.Clear
    End If
    On Error GoTo 0

    ' Header
    auditWs.Range("A1").Value = "Category"
    auditWs.Range("B1").Value = "Item"
    auditWs.Range("C1").Value = "Status"
    auditWs.Range("D1").Value = "Details"
    auditWs.Range("A1:D1").Font.Bold = True
    auditWs.Range("A1:D1").Interior.Color = RGB(11, 29, 58)
    auditWs.Range("A1:D1").Font.Color = RGB(255, 255, 255)
    row = 2

    ' === 1. Sheet Existence ===
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        auditWs.Cells(row, 1).Value = "Sheet"
        auditWs.Cells(row, 2).Value = ws.Name
        auditWs.Cells(row, 3).Value = "PASS"
        auditWs.Cells(row, 4).Value = "Exists, " & ws.UsedRange.Rows.Count & " rows used"
        passCount = passCount + 1
        row = row + 1
    Next ws

    ' === 2. Named Ranges ===
    Dim nm As Name
    For Each nm In ThisWorkbook.Names
        auditWs.Cells(row, 1).Value = "NamedRange"
        auditWs.Cells(row, 2).Value = nm.Name

        On Error Resume Next
        Dim testVal As Variant
        testVal = nm.RefersToRange.Value
        If Err.Number <> 0 Then
            auditWs.Cells(row, 3).Value = "FAIL"
            auditWs.Cells(row, 4).Value = "Cannot resolve: " & nm.RefersTo
            failCount = failCount + 1
            Err.Clear
        Else
            auditWs.Cells(row, 3).Value = "PASS"
            auditWs.Cells(row, 4).Value = nm.RefersTo
            passCount = passCount + 1
        End If
        On Error GoTo 0
        row = row + 1
    Next nm

    ' === 3. Charts — verify each has a title ===
    For Each ws In ThisWorkbook.Worksheets
        Dim co As ChartObject
        For Each co In ws.ChartObjects
            auditWs.Cells(row, 1).Value = "Chart"
            auditWs.Cells(row, 2).Value = ws.Name & " / " & co.Name

            On Error Resume Next
            Dim hasTitle As Boolean
            hasTitle = co.Chart.HasTitle
            Dim titleText As String
            titleText = ""
            If hasTitle Then titleText = co.Chart.ChartTitle.text
            On Error GoTo 0

            If hasTitle And Len(titleText) > 0 Then
                auditWs.Cells(row, 3).Value = "PASS"
                auditWs.Cells(row, 4).Value = "Title: " & titleText
                passCount = passCount + 1
            Else
                auditWs.Cells(row, 3).Value = "FAIL"
                auditWs.Cells(row, 4).Value = "Missing chart title"
                failCount = failCount + 1
            End If
            row = row + 1

            ' Check chart has at least one series
            auditWs.Cells(row, 1).Value = "ChartData"
            auditWs.Cells(row, 2).Value = ws.Name & " / " & co.Name

            On Error Resume Next
            Dim seriesCount As Long
            seriesCount = co.Chart.SeriesCollection.Count
            On Error GoTo 0

            If seriesCount > 0 Then
                auditWs.Cells(row, 3).Value = "PASS"
                auditWs.Cells(row, 4).Value = seriesCount & " series"
                passCount = passCount + 1
            Else
                auditWs.Cells(row, 3).Value = "FAIL"
                auditWs.Cells(row, 4).Value = "No data series"
                failCount = failCount + 1
            End If
            row = row + 1
        Next co
    Next ws

    ' === 4. Images — verify within printable bounds ===
    Dim pageWidth As Double, pageHeight As Double
    pageWidth = 720
    pageHeight = 936

    For Each ws In ThisWorkbook.Worksheets
        Dim shp As Shape
        For Each shp In ws.Shapes
            If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
                auditWs.Cells(row, 1).Value = "Image"
                auditWs.Cells(row, 2).Value = ws.Name & " / " & shp.Name

                If shp.Left >= 0 And shp.Top >= 0 And _
                   (shp.Left + shp.Width) <= pageWidth And _
                   (shp.Top + shp.Height) <= pageHeight Then
                    auditWs.Cells(row, 3).Value = "PASS"
                    auditWs.Cells(row, 4).Value = "In bounds (" & _
                        Round(shp.Width, 0) & "x" & Round(shp.Height, 0) & " pt)"
                    passCount = passCount + 1
                Else
                    auditWs.Cells(row, 3).Value = "FAIL"
                    auditWs.Cells(row, 4).Value = "Out of bounds at L=" & _
                        Round(shp.Left, 0) & " T=" & Round(shp.Top, 0) & _
                        " W=" & Round(shp.Width, 0) & " H=" & Round(shp.Height, 0)
                    failCount = failCount + 1
                End If
                row = row + 1
            End If
        Next shp
    Next ws

    ' === 5. Summary ===
    row = row + 1
    auditWs.Cells(row, 1).Value = "SUMMARY"
    auditWs.Cells(row, 1).Font.Bold = True
    auditWs.Cells(row, 2).Value = passCount & " passed, " & failCount & " failed"
    auditWs.Cells(row, 3).Value = IIf(failCount = 0, "ALL PASS", "FAILURES FOUND")
    auditWs.Cells(row, 3).Font.Bold = True

    If failCount = 0 Then
        auditWs.Cells(row, 3).Font.Color = RGB(22, 163, 74) ' green
    Else
        auditWs.Cells(row, 3).Font.Color = RGB(220, 38, 38) ' red
    End If

    ' Color the status column
    Dim r As Long
    For r = 2 To row - 1
        If auditWs.Cells(r, 3).Value = "PASS" Then
            auditWs.Cells(r, 3).Font.Color = RGB(22, 163, 74)
        ElseIf auditWs.Cells(r, 3).Value = "FAIL" Then
            auditWs.Cells(r, 3).Font.Color = RGB(220, 38, 38)
            auditWs.Cells(r, 3).Font.Bold = True
        End If
    Next r

    ' Auto-fit columns
    auditWs.Columns("A:D").AutoFit
End Sub
```

**Python call and result reading:**
```python
run_vba_macro('AuditWorkbook')

# Read audit results back into Python via AppleScript
import subprocess, json
result = subprocess.run(['osascript', '-e', '''
tell application "Microsoft Excel"
    tell active workbook
        tell sheet "Audit"
            set lastRow to (count of rows of used range)
            set summaryVal to value of cell ("C" & lastRow)
            return summaryVal as string
        end tell
    end tell
end tell
'''], capture_output=True, text=True)
if 'FAILURES FOUND' in result.stdout:
    print("AUDIT FAILED -- check Audit sheet")
else:
    print("AUDIT PASSED")
```

### `AuditSheet()`

Audits a single sheet: checks data ranges are populated, formatting is applied, charts exist.

```vba
Sub AuditSheet(sheetName As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    If ws Is Nothing Then
        Debug.Print "AuditSheet: Sheet '" & sheetName & "' not found"
        Exit Sub
    End If

    Dim auditWs As Worksheet
    Set auditWs = ThisWorkbook.Sheets("Audit")
    If auditWs Is Nothing Then
        Set auditWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        auditWs.Name = "Audit"
    End If
    On Error GoTo 0

    Dim row As Long
    row = auditWs.Cells(auditWs.Rows.Count, 1).End(xlUp).row + 1
    If row < 2 Then row = 2

    ' Section header
    auditWs.Cells(row, 1).Value = "=== Sheet Audit: " & sheetName & " ==="
    auditWs.Cells(row, 1).Font.Bold = True
    row = row + 1

    ' Check used range
    auditWs.Cells(row, 1).Value = "UsedRange"
    auditWs.Cells(row, 2).Value = sheetName
    Dim usedRows As Long, usedCols As Long
    usedRows = ws.UsedRange.Rows.Count
    usedCols = ws.UsedRange.Columns.Count
    If usedRows > 1 Or usedCols > 1 Then
        auditWs.Cells(row, 3).Value = "PASS"
        auditWs.Cells(row, 4).Value = usedRows & " rows x " & usedCols & " cols"
    Else
        auditWs.Cells(row, 3).Value = "WARN"
        auditWs.Cells(row, 4).Value = "Sheet appears empty"
    End If
    row = row + 1

    ' Check charts
    auditWs.Cells(row, 1).Value = "Charts"
    auditWs.Cells(row, 2).Value = sheetName
    auditWs.Cells(row, 3).Value = "INFO"
    auditWs.Cells(row, 4).Value = ws.ChartObjects.Count & " chart(s)"
    row = row + 1

    ' Check images
    Dim imgCount As Long
    imgCount = 0
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Type = msoPicture Or shp.Type = msoLinkedPicture Then
            imgCount = imgCount + 1
        End If
    Next shp
    auditWs.Cells(row, 1).Value = "Images"
    auditWs.Cells(row, 2).Value = sheetName
    auditWs.Cells(row, 3).Value = "INFO"
    auditWs.Cells(row, 4).Value = imgCount & " image(s)"
    row = row + 1
End Sub
```

---

## macOS-Specific VBA Notes

### What Works Differently on Mac vs Windows

| Feature | Windows VBA | Mac VBA | Notes |
|---------|------------|---------|-------|
| File paths | Backslash `C:\Users\...` | Forward slash or colon `/Users/...` or `Macintosh HD:Users:...` | Mac VBA accepts POSIX paths in most contexts; use `MacScript()` for HFS paths if needed |
| `Application.FileDialog` | Full support | Limited/broken | Use `Application.GetOpenFilename` or hardcode paths passed from Python |
| `Shell()` function | Runs CMD commands | Runs Mac shell | Syntax differs; prefer passing commands from Python instead |
| Clipboard operations | `MSForms.DataObject` | Not available | Use Python for clipboard; avoid VBA clipboard on Mac |
| ActiveX controls | Full support | Not supported | Use form controls only, or avoid controls entirely |
| `SendKeys` | Works | Does not work | Never use on Mac; use AppleScript for UI automation if needed |
| UserForms | Full support | Partial support | Basic forms work; some controls are missing or buggy |
| `Environ()` | Returns Windows env vars | Returns Mac env vars | Works, but paths differ |
| Ribbon/Add-ins | COM add-ins | Different model | Not relevant for automation scripts |
| `CreateObject()` | Creates COM objects | Very limited | Most COM objects unavailable on Mac |
| `WScript.Shell` | Available | Not available | Use `MacScript()` or pass commands from Python |
| File I/O | `Open "C:\..."` | `Open "/Users/..."` | POSIX paths work in `Open` statement |
| `ThisWorkbook.Path` | Returns `C:\...` | Returns `/Users/...` | POSIX format on Mac |

### Mac-Specific VBA Tips

1. **Always use POSIX paths** when passing file paths from Python to VBA macros. Mac Excel VBA handles `/Users/...` paths correctly in most file operations.

2. **Avoid late binding** (`CreateObject`) on Mac. It fails for most objects. If you need external libraries, handle them in Python instead.

3. **`RGB()` function works identically** on Mac and Windows. No changes needed for color handling.

4. **`Application.Wait`** works on Mac but `Sleep` (from `kernel32.dll`) does not. For delays:
   ```vba
   ' Works on Mac:
   Application.Wait Now + TimeValue("00:00:02")
   ' Does NOT work on Mac:
   ' Declare Sub Sleep Lib "kernel32" ...  ' Windows-only DLL
   ```

5. **`Dir()` function** works with POSIX paths on Mac:
   ```vba
   If Dir("/Users/someone/file.xlsx") <> "" Then
       ' File exists
   End If
   ```

6. **Macro security** on Mac: By default, macros are disabled. The workbook must be `.xlsm` format, and the user must enable macros. When calling macros via AppleScript `do Visual Basic`, Excel must have macros enabled for that workbook.

7. **`Application.ScreenUpdating`** works on Mac but is less impactful than Windows (Mac Excel rendering is already smoother). Still use it for large batch operations.

8. **Chart formatting** is identical between Mac and Windows VBA. The `Chart` object model is fully supported. This is the key advantage of using VBA for live Excel operations.

9. **Pivot tables** work the same in VBA on Mac. The `PivotTable` and `PivotCache` objects behave identically. Refresh patterns are unchanged.

10. **PDF export** via `ExportAsFixedFormat` works on Mac but requires a PDF engine. Modern Mac Excel has this built in. No external PDF printer needed.

---

## Error Handling Patterns

### Standard Error Handler Template

Every macro should use this pattern:

```vba
Sub MyMacro(param1 As String)
    On Error GoTo ErrHandler

    ' ... macro body ...

    Exit Sub
ErrHandler:
    Dim errMsg As String
    errMsg = "MyMacro error: " & Err.Description & " (code " & Err.Number & ")"
    Debug.Print errMsg

    ' Optionally write to Audit sheet
    On Error Resume Next
    Dim auditWs As Worksheet
    Set auditWs = ThisWorkbook.Sheets("Audit")
    If Not auditWs Is Nothing Then
        Dim row As Long
        row = auditWs.Cells(auditWs.Rows.Count, 1).End(xlUp).row + 1
        auditWs.Cells(row, 1).Value = "ERROR"
        auditWs.Cells(row, 2).Value = "MyMacro"
        auditWs.Cells(row, 3).Value = "FAIL"
        auditWs.Cells(row, 4).Value = errMsg
        auditWs.Cells(row, 3).Font.Color = RGB(220, 38, 38)
    End If
    On Error GoTo 0
End Sub
```

### Safe Sheet Access

Always validate sheet existence before operating:

```vba
Function GetSheet(sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    If GetSheet Is Nothing Then
        Debug.Print "Sheet '" & sheetName & "' not found"
    End If
End Function
```

Usage:
```vba
Dim ws As Worksheet
Set ws = GetSheet("Dashboard")
If ws Is Nothing Then Exit Sub
```

### Safe Range Access

Validate range before formatting:

```vba
Function SafeRange(ws As Worksheet, addr As String) As Range
    On Error Resume Next
    Set SafeRange = ws.Range(addr)
    On Error GoTo 0
End Function
```

### WriteModePrep / WriteModeEnd Guard

Always wrap bulk operations with the performance guards, and ensure `WriteModeEnd` runs even if errors occur:

```vba
Sub BulkOperation()
    On Error GoTo ErrHandler
    Call WriteModePrep

    ' ... bulk work ...

    Call WriteModeEnd
    Exit Sub

ErrHandler:
    ' CRITICAL: always re-enable even on error
    Call WriteModeEnd
    Debug.Print "BulkOperation error: " & Err.Description
End Sub
```

### Retry Pattern

For operations that occasionally fail due to Excel timing (common on Mac):

```vba
Function RetryMacro(macroName As String, maxRetries As Long) As Boolean
    Dim attempt As Long
    For attempt = 1 To maxRetries
        On Error Resume Next
        Application.Run macroName
        If Err.Number = 0 Then
            RetryMacro = True
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
        Application.Wait Now + TimeValue("00:00:01")
    Next attempt
    RetryMacro = False
End Function
```

### Function Return for Python

When you need VBA to return a status to Python, use a Function instead of Sub:

```vba
Function DoSomethingWithStatus() As String
    On Error GoTo ErrHandler

    ' ... work ...

    DoSomethingWithStatus = "OK"
    Exit Function

ErrHandler:
    DoSomethingWithStatus = "ERROR: " & Err.Description
End Function
```

Python side:
```python
# Call the function via AppleScript and capture return value
result = subprocess.run(['osascript', '-e', '''
tell application "Microsoft Excel"
    set res to do Visual Basic "DoSomethingWithStatus()"
    return res
end tell
'''], capture_output=True, text=True)
if result.stdout.strip().startswith('ERROR'):
    print(f"Macro failed: {result.stdout.strip()}")
```

---

## Passing Palette Colors from Python to VBA

The recommended approach is to write color values to a Control sheet as named ranges, then read them in VBA. This avoids passing many color arguments to each macro.

### Python Side: Write Palette to Control Sheet

```python
from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName

wb = load_workbook('/tmp/workbook.xlsm', keep_vba=True)

# Create or get Control sheet
if 'Control' not in wb.sheetnames:
    wb.create_sheet('Control')
ctrl = wb['Control']

# Define palette
pal = {
    'header_bg':   (11, 29, 58),
    'header_text': (255, 255, 255),
    'accent':      (201, 168, 76),
    'text':        (30, 41, 59),
    'muted':       (100, 116, 139),
    'alt_row':     (241, 245, 249),
    'border':      (226, 232, 240),
    'chart_1':     (59, 130, 246),
    'chart_2':     (139, 92, 246),
    'chart_3':     (16, 185, 129),
    'chart_4':     (201, 168, 76),
    'positive':    (22, 163, 74),
    'negative':    (220, 38, 38),
}

# Write palette to Control sheet
ctrl['A1'].value = 'Color Name'
ctrl['B1'].value = 'R'
ctrl['C1'].value = 'G'
ctrl['D1'].value = 'B'
row = 2

for name, (r, g, b) in pal.items():
    ctrl[f'A{row}'].value = name
    ctrl[f'B{row}'].value = r
    ctrl[f'C{row}'].value = g
    ctrl[f'D{row}'].value = b
    # Create named range for easy VBA access
    dn = DefinedName(f'PAL_{name.upper()}', attr_text=f"Control!$B${row}:$D${row}")
    wb.defined_names.add(dn)
    row += 1

# Also write expected sheets list
ctrl[f'A{row + 1}'].value = 'ExpectedSheets'
ctrl[f'B{row + 1}'].value = 'Dashboard,Data,Charts,Audit'
dn = DefinedName('ExpectedSheets', attr_text=f"Control!$B${row + 1}")
wb.defined_names.add(dn)

wb.save('/tmp/workbook.xlsm')
wb.close()
```

### VBA Side: Read Palette from Control Sheet

```vba
Function GetPaletteColor(colorName As String) As Long
    ' Reads an RGB color from the Control sheet named range PAL_<NAME>
    ' Returns the RGB Long value for use in .Interior.Color, .Font.Color, etc.
    On Error GoTo ErrHandler

    Dim rng As Range
    Set rng = ThisWorkbook.Names("PAL_" & UCase(colorName)).RefersToRange

    Dim r As Long, g As Long, b As Long
    r = CLng(rng.Cells(1, 1).Value)
    g = CLng(rng.Cells(1, 2).Value)
    b = CLng(rng.Cells(1, 3).Value)

    GetPaletteColor = RGB(r, g, b)
    Exit Function

ErrHandler:
    ' Default to black if color not found
    GetPaletteColor = RGB(0, 0, 0)
    Debug.Print "GetPaletteColor: '" & colorName & "' not found, defaulting to black"
End Function
```

### Usage in Macros

```vba
Sub ApplyPaletteFormatting(sheetName As String, startRow As Long, endRow As Long, _
                           startCol As Long, endCol As Long)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    ' Read colors from palette
    Dim headerBg As Long, headerText As Long
    Dim altRow As Long, borderColor As Long, textColor As Long
    headerBg = GetPaletteColor("HEADER_BG")
    headerText = GetPaletteColor("HEADER_TEXT")
    altRow = GetPaletteColor("ALT_ROW")
    borderColor = GetPaletteColor("BORDER")
    textColor = GetPaletteColor("TEXT")

    ' Header
    Dim headerRng As Range
    Set headerRng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, endCol))
    headerRng.Interior.Color = headerBg
    headerRng.Font.Color = headerText
    headerRng.Font.Bold = True
    headerRng.Font.Size = 11

    ' Data rows with banding
    Dim r As Long
    For r = startRow + 1 To endRow
        Dim rowRng As Range
        Set rowRng = ws.Range(ws.Cells(r, startCol), ws.Cells(r, endCol))
        If (r - startRow) Mod 2 = 0 Then
            rowRng.Interior.Color = altRow
        Else
            rowRng.Interior.Color = RGB(255, 255, 255)
        End If
        rowRng.Font.Color = textColor
        rowRng.Font.Size = 10
    Next r

    ' Borders
    Dim fullRng As Range
    Set fullRng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(endRow, endCol))
    With fullRng.Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = borderColor
    End With

    Exit Sub
ErrHandler:
    Debug.Print "ApplyPaletteFormatting error: " & Err.Description
End Sub
```

### Chart Colors from Palette

```vba
Sub ApplyPaletteChartColors(sheetName As String, chartIndex As Long)
    ' Applies chart_1, chart_2, chart_3, chart_4 colors from the palette
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)

    Dim cht As Chart
    Set cht = ws.ChartObjects(chartIndex).Chart

    Dim colors(1 To 4) As Long
    colors(1) = GetPaletteColor("CHART_1")
    colors(2) = GetPaletteColor("CHART_2")
    colors(3) = GetPaletteColor("CHART_3")
    colors(4) = GetPaletteColor("CHART_4")

    Dim s As Long
    For s = 1 To cht.SeriesCollection.Count
        If s <= 4 Then
            cht.SeriesCollection(s).Format.Fill.ForeColor.RGB = colors(s)
            On Error Resume Next
            cht.SeriesCollection(s).Format.Line.ForeColor.RGB = colors(s)
            On Error GoTo ErrHandler
        End If
    Next s

    Exit Sub
ErrHandler:
    Debug.Print "ApplyPaletteChartColors error: " & Err.Description
End Sub
```

### Colors String Builder (Python Helper)

Convenience function to build the semicolon-separated colors string for chart macros:

```python
def build_colors_string(pal, keys):
    """Build 'R,G,B;R,G,B;...' string from palette dict and key list.

    Args:
        pal: dict mapping color names to (R, G, B) tuples
        keys: list of palette key names, e.g. ['chart_1', 'chart_2', 'chart_3']

    Returns:
        str like '59,130,246;139,92,246;16,185,129'
    """
    parts = []
    for key in keys:
        r, g, b = pal[key]
        parts.append(f'{r},{g},{b}')
    return ';'.join(parts)

# Example usage:
colors = build_colors_string(pal, ['chart_1', 'chart_2', 'chart_3'])
run_vba_macro('CreateLineChart', 'Dashboard', 'Data!A1:D13', 50, 300, 500, 280,
              'Monthly Trends', colors)
```

---

## Quick Reference: Macro Signatures

All macros are called via `run_vba_macro()` helper or AppleScript `do Visual Basic`. See [AppleScript Patterns](applescript-patterns.md) for the full helper implementation.

| Macro | Arguments | Returns |
|-------|-----------|---------|
| `ResetWorkbook` | (none) | - |
| `SanityCheck` | (none) | Boolean |
| `WriteModePrep` | (none) | - |
| `WriteModeEnd` | (none) | - |
| `ApplyTableFormatting` | sheetName, startRow, endRow, startCol, endCol, headerBgR/G/B, accentR/G/B | - |
| `ApplyKPIPanel` | sheetName, row, col, endCol, bgR/G/B, textR/G/B | - |
| `ApplyBorders` | sheetName, startRow, endRow, startCol, endCol, borderR/G/B | - |
| `HideGridlines` | sheetName | - |
| `SetColumnWidths` | sheetName, widthsString (`"A:12,B:20"`) | - |
| `SetRowHeight` | sheetName, row, height | - |
| `SetSheetTabColor` | sheetName, R, G, B | - |
| `ApplyHeaderBar` | sheetName, row, startCol, endCol, text, bgR/G/B, textR/G/B, fontSize | - |
| `CreateBarChart` | sheetName, dataRange, left, top, width, height, title, colorR/G/B | - |
| `CreateLineChart` | sheetName, dataRange, left, top, width, height, title, colorsString | - |
| `CreateDoughnutChart` | sheetName, dataRange, left, top, width, height, title, colorsString | - |
| `CreatePieChart` | sheetName, dataRange, left, top, width, height, title, colorsString, showLabels | - |
| `CreateStackedBarChart` | sheetName, dataRange, left, top, width, height, title, colorsString | - |
| `StyleAllCharts` | sheetName | - |
| `RefreshAllPivots` | (none) | - |
| `SetChartTitle` | sheetName, chartIndex, titleText | - |
| `SetChartSeriesColor` | sheetName, chartIndex, seriesIndex, R, G, B | - |
| `InsertImageIntoPlaceholder` | sheetName, imagePath, leftPt, topPt, widthPt, heightPt, mode (`"FIT"`/`"FILL"`) | - |
| `AuditImages` | sheetName | - |
| `RemoveAllImages` | sheetName | - |
| `ExportPdf` | outputPath | - |
| `ExportSheetAsPdf` | sheetName, outputPath | - |
| `ExportAllSheetsPdf` | outputPath | - |
| `SetPageSetup` | sheetName, orientation, paperSize, fitWide, fitTall | - |
| `SetPrintArea` | sheetName, printRange | - |
| `SetPrintTitles` | sheetName, titleRows, titleCols | - |
| `AuditWorkbook` | (none) | - |
| `AuditSheet` | sheetName | - |
| `GetPaletteColor` | colorName | Long (RGB) |
| `ApplyPaletteFormatting` | sheetName, startRow, endRow, startCol, endCol | - |
| `ApplyPaletteChartColors` | sheetName, chartIndex | - |
