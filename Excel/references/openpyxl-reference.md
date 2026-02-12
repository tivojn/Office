# openpyxl Reference — File-Based Excel Engine

Complete reference for Excel automation with openpyxl. openpyxl works directly on `.xlsx` files without needing Excel to be running. It handles ALL formatting, charts, images, borders, conditional formatting, and data I/O. For operations requiring a live Excel instance (pivot refresh, recalculation, VBA macros), use AppleScript after openpyxl is done.

## Table of Contents

1. [Standard Imports](#1-standard-imports)
2. [Workbook Management](#2-workbook-management)
3. [Sheet Operations](#3-sheet-operations)
4. [Data I/O (Core Section)](#4-data-io-core-section)
5. [Formatting](#5-formatting)
6. [Borders](#6-borders)
7. [Charts](#7-charts)
8. [Images](#8-images)
9. [Conditional Formatting](#9-conditional-formatting)
10. [Named Ranges and Styles](#10-named-ranges-and-styles)
11. [Page Setup and Print](#11-page-setup-and-print)
12. [View Settings](#12-view-settings)
13. [Helper Functions](#13-helper-functions)
14. [Opening in Excel When Done](#14-opening-in-excel-when-done)

---

## 1. Standard Imports

```python
from openpyxl import Workbook, load_workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side,
    NamedStyle, numbers
)
from openpyxl.chart import (
    BarChart, LineChart, PieChart, DoughnutChart, AreaChart,
    ScatterChart, Reference
)
from openpyxl.chart.series import SeriesLabel
from openpyxl.chart.label import DataLabelList
from openpyxl.drawing.image import Image as XlImage
from openpyxl.formatting.rule import (
    ColorScaleRule, DataBarRule, IconSetRule, CellIsRule,
    FormulaRule
)
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.workbook.defined_name import DefinedName
import pandas as pd
import numpy as np
from PIL import Image
import os
import subprocess
import time

xlsx_path = os.environ.get('XLSX_PATH', '/tmp/workbook.xlsx')
```

---

## 2. Workbook Management

### Create New Workbook

```python
wb = Workbook()
ws = wb.active  # First sheet created by default
ws.title = 'Dashboard'
wb.save(xlsx_path)
```

### Open Existing Workbook

```python
wb = load_workbook(xlsx_path)
# With data_only=True to read cached formula values (not formulas)
wb = load_workbook(xlsx_path, data_only=True)
```

### Copy Template for Transactional Runs

```python
import shutil

template_path = '/path/to/template.xlsx'
output_path = '/tmp/output_report.xlsx'

shutil.copy2(template_path, output_path)
wb = load_workbook(output_path)
# ... do work ...
wb.save(output_path)
```

### Save

```python
# Save to current path
wb.save(xlsx_path)

# Save As (new path)
wb.save('/path/to/new_file.xlsx')
```

Note: openpyxl has no concept of "close" — the workbook is just a Python object. Save when done, and let it go out of scope.

---

## 3. Sheet Operations

### Create, Access, Delete

```python
# Access by name
ws = wb['Dashboard']

# Access active sheet
ws = wb.active

# List all sheet names
names = wb.sheetnames  # ['Sheet1', 'Sheet2', ...]

# Add new sheet at end
ws = wb.create_sheet('NewSheet')

# Add at specific position (0-indexed)
ws = wb.create_sheet('Cover', 0)  # First position
ws = wb.create_sheet('Data', 1)   # Second position

# Delete a sheet
del wb['Temp']
# Or: wb.remove(wb['Temp'])

# Rename
ws.title = 'Sales Dashboard'

# Copy sheet within workbook
source = wb['Template']
target = wb.copy_worksheet(source)
target.title = 'Q1 Report'
```

### Sheet Tab Color

```python
# Set tab color (hex RGB string, no # prefix)
ws.sheet_properties.tabColor = '0B1D3A'   # Navy
ws.sheet_properties.tabColor = 'C9A84C'   # Gold
ws.sheet_properties.tabColor = '3B82F6'   # Blue

# Tab color conventions
TAB_COLORS = {
    'Control':   '64748B',  # Slate gray
    'Data':      '3B82F6',  # Blue
    'Dashboard': 'C9A84C',  # Gold
    'Pivot':     '8B5CF6',  # Purple
    'Charts':    '10B981',  # Teal
    'Audit':     'DC2626',  # Red
}
for sheet_name, color in TAB_COLORS.items():
    if sheet_name in wb.sheetnames:
        wb[sheet_name].sheet_properties.tabColor = color
```

### Reorder Sheets

```python
# Move sheet to a specific position
wb.move_sheet('Dashboard', offset=-2)  # Move 2 positions left
# Or set order directly:
wb._sheets.sort(key=lambda s: ['Dashboard', 'Data', 'Charts', 'Audit'].index(s.title)
    if s.title in ['Dashboard', 'Data', 'Charts', 'Audit'] else 99)
```

### Sheet Count

```python
n = len(wb.sheetnames)
```

---

## 4. Data I/O (Core Section)

### Write DataFrame to Sheet

```python
import pandas as pd
from openpyxl.utils.dataframe import dataframe_to_rows

df = pd.DataFrame({
    'Product': ['Widget A', 'Widget B', 'Widget C', 'Widget D'],
    'Revenue': [125000, 89000, 210000, 67000],
    'Units': [1250, 890, 2100, 670],
    'Growth': [0.12, -0.05, 0.23, 0.08],
})

# Write with headers, no index
for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
    for c_idx, value in enumerate(row, 1):
        ws.cell(row=r_idx, column=c_idx, value=value)
```

### Write DataFrame Starting at Specific Cell

```python
start_row, start_col = 5, 3  # C5

for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start_row):
    for c_idx, value in enumerate(row, start_col):
        ws.cell(row=r_idx, column=c_idx, value=value)
```

### Write 2D List

```python
data = [
    ['Name', 'Age', 'City'],
    ['Alice', 30, 'New York'],
    ['Bob', 25, 'Los Angeles'],
    ['Charlie', 35, 'Chicago'],
]
for r_idx, row in enumerate(data, 1):
    for c_idx, value in enumerate(row, 1):
        ws.cell(row=r_idx, column=c_idx, value=value)
```

### Write Single Row

```python
headers = ['Jan', 'Feb', 'Mar', 'Apr', 'May']
for c_idx, val in enumerate(headers, 1):
    ws.cell(row=1, column=c_idx, value=val)
```

### Write Single Column

```python
months = ['Jan', 'Feb', 'Mar', 'Apr', 'May']
for r_idx, val in enumerate(months, 1):
    ws.cell(row=r_idx, column=1, value=val)
```

### Read Range as DataFrame

```python
data = []
for row in ws.iter_rows(min_row=1, max_row=100, max_col=6, values_only=True):
    data.append(row)
df = pd.DataFrame(data[1:], columns=data[0])  # First row = headers
```

### Read Range as 2D List

```python
data = []
for row in ws.iter_rows(min_row=1, max_row=10, min_col=1, max_col=4, values_only=True):
    data.append(list(row))
```

### Write Formulas

```python
# Single formula
ws['F2'] = '=SUM(B2:B101)'

# Column of formulas
for i in range(2, 102):
    ws[f'E{i}'] = f'=B{i}*C{i}'
```

### Cell Value Types

```python
ws['A1'] = 'Text string'       # String
ws['B1'] = 42                   # Integer
ws['C1'] = 3.14                 # Float
ws['D1'] = True                 # Boolean
ws['E1'] = datetime.now()       # DateTime
ws['F1'] = '=SUM(A1:A10)'      # Formula
ws['G1'] = None                 # Clear cell
```

---

## 5. Formatting

### Font

```python
from openpyxl.styles import Font

# Header font
header_font = Font(
    name='Montserrat',
    size=12,
    bold=True,
    italic=False,
    color='FFFFFF',  # White (hex, no #)
)

# Body font
body_font = Font(name='Calibri', size=10, color='1A202C')

# Apply to range
for cell in ws[1]:  # Row 1
    cell.font = header_font

# Apply to specific cell
ws['A1'].font = Font(name='Calibri', size=14, bold=True, color='0B1D3A')

# Underline
ws['A1'].font = Font(underline='single')  # 'single', 'double'

# Strikethrough
ws['A1'].font = Font(strikethrough=True)
```

### Cell Background Color (Fill)

```python
from openpyxl.styles import PatternFill

# Solid fill
navy_fill = PatternFill(start_color='0B1D3A', end_color='0B1D3A', fill_type='solid')
light_gray_fill = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
gold_fill = PatternFill(start_color='C9A84C', end_color='C9A84C', fill_type='solid')

# Apply to row
for cell in ws[1]:
    cell.fill = navy_fill

# Apply to range
for row in ws['A1:F20']:
    for cell in row:
        cell.fill = light_gray_fill

# Remove fill
ws['A1'].fill = PatternFill(fill_type=None)
```

### Number Format

```python
# Apply to cells
ws['B2'].number_format = '#,##0'           # 1,234
ws['B2'].number_format = '#,##0.00'        # 1,234.56
ws['B2'].number_format = '0.0%'            # 12.3%
ws['B2'].number_format = '$#,##0.00'       # $1,234.56
ws['B2'].number_format = '_($* #,##0.00_)' # Accounting
ws['B2'].number_format = 'YYYY-MM-DD'      # 2026-02-10
ws['B2'].number_format = 'MMM DD, YYYY'    # Feb 10, 2026
ws['B2'].number_format = '0.00E+00'        # Scientific
ws['B2'].number_format = '@'               # Force text
ws['B2'].number_format = '#,##0.00;-#,##0.00;"-"'  # Dash for zero
ws['B2'].number_format = '#,##0.00;[Red]-#,##0.00'  # Red negative
ws['B2'].number_format = '#,##0.0,,"M"'    # Millions with M
ws['B2'].number_format = '#,##0,"K"'       # Thousands with K

# Apply to entire column range
for row in ws.iter_rows(min_row=2, max_row=101, min_col=2, max_col=2):
    for cell in row:
        cell.number_format = '$#,##0.00'
```

### Row Height and Column Width

```python
# Column width (in approximate character units)
ws.column_dimensions['A'].width = 25
ws.column_dimensions['B'].width = 15

# Set multiple columns
for col, width in [('A', 25), ('B', 15), ('C', 15), ('D', 12), ('E', 12)]:
    ws.column_dimensions[col].width = width

# Row height (in points)
ws.row_dimensions[1].height = 40    # Header row
for r in range(2, 101):
    ws.row_dimensions[r].height = 22  # Body rows
```

### Alignment

```python
from openpyxl.styles import Alignment

# Center alignment
center = Alignment(horizontal='center', vertical='center')
ws['A1'].alignment = center

# Left + middle
ws['A1'].alignment = Alignment(horizontal='left', vertical='center')

# Right + middle (for numbers)
ws['B2'].alignment = Alignment(horizontal='right', vertical='center')

# Wrap text
ws['A1'].alignment = Alignment(wrap_text=True, vertical='top')

# Indent
ws['A1'].alignment = Alignment(indent=1)

# Text rotation
ws['A1'].alignment = Alignment(text_rotation=45)
```

### Merged Cells

```python
# Merge
ws.merge_cells('A1:F1')

# Write to merged cell (always top-left)
ws['A1'] = 'Dashboard Title'

# Unmerge
ws.unmerge_cells('A1:F1')

# Merge with range notation
ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
```

### Hide/Unhide Rows and Columns

```python
# Hide column
ws.column_dimensions['E'].hidden = True

# Unhide column
ws.column_dimensions['E'].hidden = False

# Hide row
ws.row_dimensions[50].hidden = True

# Unhide row
ws.row_dimensions[50].hidden = False
```

---

## 6. Borders

openpyxl has full border support — no workarounds needed.

```python
from openpyxl.styles import Border, Side

# Define border styles
thin_border = Border(
    left=Side(style='thin', color='E2E8F0'),
    right=Side(style='thin', color='E2E8F0'),
    top=Side(style='thin', color='E2E8F0'),
    bottom=Side(style='thin', color='E2E8F0'),
)

# Header bottom border (thick accent)
header_bottom = Border(
    bottom=Side(style='medium', color='0B1D3A')
)

# Apply to range
for row in ws['A1:F20']:
    for cell in row:
        cell.border = thin_border

# Apply thick bottom border to header
for cell in ws[1]:
    cell.border = Border(
        left=Side(style='thin', color='E2E8F0'),
        right=Side(style='thin', color='E2E8F0'),
        top=Side(style='thin', color='E2E8F0'),
        bottom=Side(style='medium', color='0B1D3A'),
    )

# Outer border only on a range
from openpyxl.styles.borders import BORDER_THIN
# Apply outer borders manually (top/bottom/left/right edges)
min_row, max_row = 1, 20
min_col, max_col = 1, 6

for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
    for cell in row:
        border_kw = {}
        if cell.row == min_row:
            border_kw['top'] = Side(style='medium', color='0B1D3A')
        if cell.row == max_row:
            border_kw['bottom'] = Side(style='thin', color='E2E8F0')
        if cell.column == min_col:
            border_kw['left'] = Side(style='thin', color='E2E8F0')
        if cell.column == max_col:
            border_kw['right'] = Side(style='thin', color='E2E8F0')
        if border_kw:
            # Merge with existing inner borders
            existing = cell.border
            cell.border = Border(
                left=border_kw.get('left', existing.left),
                right=border_kw.get('right', existing.right),
                top=border_kw.get('top', existing.top),
                bottom=border_kw.get('bottom', existing.bottom),
            )

# Border style options: 'thin', 'medium', 'thick', 'double', 'dotted', 'dashed',
#   'dashDot', 'dashDotDot', 'hair', 'mediumDashed', 'mediumDashDot',
#   'mediumDashDotDot', 'slantDashDot'
```

---

## 7. Charts

openpyxl creates charts natively — no VBA needed.

### Bar Chart (Column)

```python
from openpyxl.chart import BarChart, Reference

chart = BarChart()
chart.type = 'col'  # 'col' for column, 'bar' for horizontal
chart.grouping = 'clustered'  # 'clustered', 'stacked', 'percentStacked'
chart.title = 'Monthly Revenue'
chart.y_axis.title = None
chart.x_axis.title = None

# Data reference (assumes data in sheet)
data = Reference(ws, min_col=2, min_row=1, max_row=13, max_col=3)
cats = Reference(ws, min_col=1, min_row=2, max_row=13)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

# Style
chart.style = 10  # Clean style
chart.width = 18  # cm
chart.height = 10

# Series colors
from openpyxl.chart.series import DataPoint
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
series = chart.series[0]
series.graphicalProperties.solidFill = '3B82F6'  # Blue

# Place chart
ws.add_chart(chart, 'F2')
```

### Line Chart

```python
from openpyxl.chart import LineChart, Reference

chart = LineChart()
chart.title = 'Monthly Trends'
chart.style = 10

data = Reference(ws, min_col=2, min_row=1, max_row=13, max_col=4)
cats = Reference(ws, min_col=1, min_row=2, max_row=13)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

# Style each series
colors = ['3B82F6', '8B5CF6', '10B981']
for i, s in enumerate(chart.series):
    s.graphicalProperties.line.solidFill = colors[i] if i < len(colors) else '000000'
    s.graphicalProperties.line.width = 25000  # EMUs (25000 = ~2pt)
    s.smooth = False

chart.width = 18
chart.height = 10
ws.add_chart(chart, 'F2')
```

### Pie Chart

```python
from openpyxl.chart import PieChart, Reference

chart = PieChart()
chart.title = 'Market Share'

data = Reference(ws, min_col=2, min_row=1, max_row=6)
cats = Reference(ws, min_col=1, min_row=2, max_row=6)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

# Slice colors
from openpyxl.chart.series import DataPoint
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
colors = ['3B82F6', '8B5CF6', '10B981', 'C9A84C', '64748B']
for i, color in enumerate(colors):
    pt = DataPoint(idx=i)
    pt.graphicalProperties.solidFill = color
    chart.series[0].data_points.append(pt)

# Data labels
chart.dataLabels = DataLabelList()
chart.dataLabels.showPercent = True
chart.dataLabels.showVal = False

chart.width = 12
chart.height = 10
ws.add_chart(chart, 'F2')
```

### Doughnut Chart

```python
from openpyxl.chart import DoughnutChart, Reference

chart = DoughnutChart()
chart.title = 'Revenue Split'

data = Reference(ws, min_col=2, min_row=1, max_row=6)
cats = Reference(ws, min_col=1, min_row=2, max_row=6)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

# Slice colors
colors = ['3B82F6', '8B5CF6', '10B981', 'C9A84C', '64748B']
for i, color in enumerate(colors):
    pt = DataPoint(idx=i)
    pt.graphicalProperties.solidFill = color
    chart.series[0].data_points.append(pt)

chart.width = 12
chart.height = 10
ws.add_chart(chart, 'F2')
```

### Area Chart

```python
from openpyxl.chart import AreaChart, Reference

chart = AreaChart()
chart.title = 'Revenue Over Time'
chart.grouping = 'standard'

data = Reference(ws, min_col=2, min_row=1, max_row=13, max_col=3)
cats = Reference(ws, min_col=1, min_row=2, max_row=13)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)

chart.series[0].graphicalProperties.solidFill = '3B82F6'
chart.width = 18
chart.height = 10
ws.add_chart(chart, 'F2')
```

### Scatter Chart

```python
from openpyxl.chart import ScatterChart, Reference, Series

chart = ScatterChart()
chart.title = 'Correlation'
chart.x_axis.title = 'X Values'
chart.y_axis.title = 'Y Values'

x_values = Reference(ws, min_col=1, min_row=2, max_row=50)
y_values = Reference(ws, min_col=2, min_row=2, max_row=50)
series = Series(y_values, x_values, title='Data Points')
chart.series.append(series)

chart.width = 15
chart.height = 10
ws.add_chart(chart, 'F2')
```

### Chart Styling Tips

```python
# Remove chart border
chart.plot_area.graphicalProperties = None  # or set no line

# Hide gridlines
chart.y_axis.delete = False
chart.y_axis.majorGridlines = None

# Legend position
chart.legend.position = 'b'  # 'b'=bottom, 'r'=right, 't'=top, 'l'=left

# Axis font size (via chart style)
chart.style = 10  # Built-in clean style

# No legend
chart.legend = None
```

---

## 8. Images

```python
from openpyxl.drawing.image import Image as XlImage

# Insert image
img = XlImage('/path/to/logo.png')

# Set size (in pixels, openpyxl converts internally)
img.width = 200
img.height = 80

# Add to sheet at specific cell
ws.add_image(img, 'B2')

# Image from PIL with aspect ratio enforcement
from PIL import Image as PILImage

def insert_image_fit(ws, path, anchor_cell, max_width, max_height):
    """Insert image maintaining aspect ratio within bounds."""
    pil_img = PILImage.open(path)
    w, h = pil_img.size
    scale = min(max_width / w, max_height / h)

    img = XlImage(path)
    img.width = int(w * scale)
    img.height = int(h * scale)
    ws.add_image(img, anchor_cell)
```

---

## 9. Conditional Formatting

openpyxl supports all conditional formatting types natively.

### Color Scale (Red -> Yellow -> Green)

```python
from openpyxl.formatting.rule import ColorScaleRule

ws.conditional_formatting.add('B2:B20',
    ColorScaleRule(
        start_type='min', start_color='F87171',     # Red
        mid_type='percentile', mid_value=50, mid_color='FDE047',  # Yellow
        end_type='max', end_color='4ADE80',          # Green
    )
)
```

### Data Bar

```python
from openpyxl.formatting.rule import DataBarRule

ws.conditional_formatting.add('C2:C20',
    DataBarRule(
        start_type='min', end_type='max',
        color='3B82F6',  # Blue
    )
)
```

### Icon Set

```python
from openpyxl.formatting.rule import IconSetRule

ws.conditional_formatting.add('D2:D20',
    IconSetRule(
        icon_style='3TrafficLights1',
        type='percent',
        values=[0, 33, 67],
    )
)
# Icon styles: '3TrafficLights1', '3Arrows', '3Symbols', '4Rating',
#   '5Quarters', '3Stars', etc.
```

### Cell Is Rule (Traffic Light Text)

```python
from openpyxl.formatting.rule import CellIsRule

# Red for < 50%
ws.conditional_formatting.add('D2:D20',
    CellIsRule(
        operator='lessThan',
        formula=['0.5'],
        fill=PatternFill(start_color='FEE2E2', end_color='FEE2E2', fill_type='solid'),
        font=Font(color='DC2626'),
    )
)

# Yellow for 50-80%
ws.conditional_formatting.add('D2:D20',
    CellIsRule(
        operator='between',
        formula=['0.5', '0.8'],
        fill=PatternFill(start_color='FEF9C3', end_color='FEF9C3', fill_type='solid'),
        font=Font(color='A16207'),
    )
)

# Green for > 80%
ws.conditional_formatting.add('D2:D20',
    CellIsRule(
        operator='greaterThan',
        formula=['0.8'],
        fill=PatternFill(start_color='DCFCE7', end_color='DCFCE7', fill_type='solid'),
        font=Font(color='16A34A'),
    )
)
```

### Formula-Based Rule

```python
from openpyxl.formatting.rule import FormulaRule

# Highlight positive values green
ws.conditional_formatting.add('F2:F20',
    FormulaRule(
        formula=['$F2>0'],
        font=Font(color='16A34A'),
    )
)

# Highlight negative values red
ws.conditional_formatting.add('F2:F20',
    FormulaRule(
        formula=['$F2<0'],
        font=Font(color='DC2626'),
    )
)
```

---

## 10. Named Ranges and Styles

### Named Ranges

```python
from openpyxl.workbook.defined_name import DefinedName

# Create workbook-level named range
ref = "Data!$A$1:$D$100"
defn = DefinedName('SalesData', attr_text=ref)
wb.defined_names.add(defn)

# Create named range for a single cell
ref = "Control!$B$1"
defn = DefinedName('ReportDate', attr_text=ref)
wb.defined_names.add(defn)

# Read named range value
for defn in wb.defined_names.definedName:
    if defn.name == 'SalesData':
        # Parse the destination
        for title, coord in defn.destinations:
            ws = wb[title]
            # Now read from ws using coord
```

### Named Styles

```python
from openpyxl.styles import NamedStyle

# Create a reusable style
header_style = NamedStyle(name='header')
header_style.font = Font(name='Montserrat', size=11, bold=True, color='FFFFFF')
header_style.fill = PatternFill(start_color='0B1D3A', end_color='0B1D3A', fill_type='solid')
header_style.alignment = Alignment(horizontal='center', vertical='center')
header_style.border = Border(
    bottom=Side(style='medium', color='C9A84C')
)

# Register style
wb.add_named_style(header_style)

# Apply to cells
for cell in ws[1]:
    cell.style = 'header'
```

---

## 11. Page Setup and Print

```python
# Orientation
ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE  # or ORIENTATION_PORTRAIT

# Paper size
ws.page_setup.paperSize = ws.PAPERSIZE_LETTER  # or PAPERSIZE_A4

# Fit to page
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0  # 0 = as many pages as needed
ws.sheet_properties.pageSetUpPr.fitToPage = True

# Margins (in inches)
ws.page_margins.left = 0.5
ws.page_margins.right = 0.5
ws.page_margins.top = 0.75
ws.page_margins.bottom = 0.75

# Print area
ws.print_area = 'A1:N50'

# Print titles (repeat rows/cols on every page)
ws.print_title_rows = '1:2'     # Repeat rows 1-2
ws.print_title_cols = 'A:A'     # Repeat column A

# Center on page
ws.page_setup.horizontalCentered = True

# Header/Footer
ws.oddHeader.center.text = '&"Montserrat,Bold"&14Dashboard Title'
ws.oddFooter.center.text = 'Page &P of &N'
ws.oddFooter.right.text = '&D'
```

---

## 12. View Settings

### Freeze Panes

```python
# Freeze row 1 (header) — freeze below A2
ws.freeze_panes = 'A2'

# Freeze rows 1-3 and column A — freeze below B4
ws.freeze_panes = 'B4'

# Unfreeze
ws.freeze_panes = None
```

### Hide Gridlines

```python
ws.sheet_view.showGridLines = False
```

### Zoom

```python
ws.sheet_view.zoomScale = 85  # 85%
```

### Set Active Cell / Selection

```python
ws.sheet_view.selection[0].activeCell = 'A1'
ws.sheet_view.selection[0].sqref = 'A1'
```

---

## 13. Helper Functions

### bulk_write_df — Write DataFrame with Formatting

```python
def bulk_write_df(ws, start_row, start_col, df, index=False, header=True):
    """Write a pandas DataFrame to a worksheet.

    Args:
        ws: openpyxl Worksheet
        start_row: starting row (1-based)
        start_col: starting column (1-based)
        df: pandas DataFrame
        index: whether to write the index column
        header: whether to write column headers
    """
    for r_idx, row in enumerate(dataframe_to_rows(df, index=index, header=header), start_row):
        for c_idx, value in enumerate(row, start_col):
            ws.cell(row=r_idx, column=c_idx, value=value)
```

### set_column_widths — Set Multiple Column Widths

```python
def set_column_widths(ws, widths):
    """Set column widths from a dict or list.

    Args:
        ws: openpyxl Worksheet
        widths: dict like {'A': 25, 'B': 15} or list [25, 15, 15]
    """
    if isinstance(widths, dict):
        for col, width in widths.items():
            ws.column_dimensions[col].width = width
    elif isinstance(widths, list):
        for i, width in enumerate(widths):
            col_letter = get_column_letter(i + 1)
            ws.column_dimensions[col_letter].width = width
```

### quick_format_range — Apply Common Formatting

```python
def quick_format_range(ws, cell_range, font_name=None, font_size=None,
                        bold=None, font_color=None, bg_color=None,
                        h_align=None, v_align=None, wrap=False,
                        number_format=None):
    """Apply formatting to a range of cells.

    Args:
        ws: openpyxl Worksheet
        cell_range: string like 'A1:F1' or 'A1:F20'
        Other args: formatting properties (None = don't change)
    """
    for row in ws[cell_range]:
        if not hasattr(row, '__iter__'):
            row = [row]
        for cell in row:
            if font_name or font_size or bold is not None or font_color:
                cell.font = Font(
                    name=font_name or cell.font.name,
                    size=font_size or cell.font.size,
                    bold=bold if bold is not None else cell.font.bold,
                    color=font_color or cell.font.color,
                )
            if bg_color:
                cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type='solid')
            if h_align or v_align or wrap:
                cell.alignment = Alignment(
                    horizontal=h_align or cell.alignment.horizontal,
                    vertical=v_align or cell.alignment.vertical,
                    wrap_text=wrap,
                )
            if number_format:
                cell.number_format = number_format
```

### apply_table_style — Style a Data Table

```python
def apply_table_style(ws, start_row, end_row, start_col, end_col,
                       header_bg='0B1D3A', header_fg='FFFFFF',
                       alt_row_color='F1F5F9', border_color='E2E8F0',
                       header_font='Montserrat', body_font='Calibri'):
    """Apply professional table styling to a range."""
    # Header row
    for c in range(start_col, end_col + 1):
        cell = ws.cell(row=start_row, column=c)
        cell.font = Font(name=header_font, size=11, bold=True, color=header_fg)
        cell.fill = PatternFill(start_color=header_bg, end_color=header_bg, fill_type='solid')
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = Border(bottom=Side(style='medium', color=header_bg))
    ws.row_dimensions[start_row].height = 35

    # Data rows
    for r in range(start_row + 1, end_row + 1):
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = Font(name=body_font, size=10, color='1E293B')
            # Banded rows
            if (r - start_row) % 2 == 0:
                cell.fill = PatternFill(start_color=alt_row_color, end_color=alt_row_color, fill_type='solid')
            # Borders
            cell.border = Border(
                left=Side(style='thin', color=border_color),
                right=Side(style='thin', color=border_color),
                top=Side(style='thin', color=border_color),
                bottom=Side(style='thin', color=border_color),
            )
        ws.row_dimensions[r].height = 22
```

### apply_number_formats — Bulk Format Columns

```python
def apply_number_formats(ws, start_row, end_row, col_formats):
    """Apply number formats to column ranges.

    Args:
        ws: openpyxl Worksheet
        start_row: first data row
        end_row: last data row
        col_formats: dict like {'B': '$#,##0.00', 'C': '#,##0', 'D': '0.0%'}
    """
    for col_letter, fmt in col_formats.items():
        col_idx = column_index_from_string(col_letter)
        for r in range(start_row, end_row + 1):
            ws.cell(row=r, column=col_idx).number_format = fmt
```

### rgb_hex — Convert RGB Tuple to Hex

```python
def rgb_hex(r, g, b):
    """Convert RGB tuple to hex string for openpyxl.

    Args:
        r, g, b: integers 0-255
    Returns:
        str: hex color like '0B1D3A'
    """
    return f'{r:02X}{g:02X}{b:02X}'

# Usage
color = rgb_hex(11, 29, 58)  # '0B1D3A'
font = Font(color=color)
fill = PatternFill(start_color=color, end_color=color, fill_type='solid')
```

### image_preflight — Check Image Dimensions

```python
def image_preflight(image_path, target_width=None, target_height=None):
    """Check image dimensions and compute sizing info."""
    from PIL import Image

    if not os.path.exists(image_path):
        raise FileNotFoundError(f"Image not found: {image_path}")

    img = Image.open(image_path)
    w, h = img.size
    ar = w / h

    result = {'width': w, 'height': h, 'aspect_ratio': round(ar, 3), 'format': img.format}

    if target_width and target_height:
        target_ar = target_width / target_height
        result['mode'] = 'FIT' if abs(ar - target_ar) < 0.1 else 'FILL'
        scale = min(target_width / w, target_height / h)
        result['scaled_width'] = int(w * scale)
        result['scaled_height'] = int(h * scale)

    return result
```

---

## 14. Opening in Excel When Done

After openpyxl finishes creating/editing the file, open it in Excel:

```python
import subprocess
import time

def open_in_excel(xlsx_path):
    """Open the file in Excel after openpyxl is done."""
    subprocess.run(['open', xlsx_path])
    time.sleep(2)

# Or with explicit Excel targeting:
def open_in_excel_app(xlsx_path):
    """Open file specifically in Microsoft Excel."""
    subprocess.run(['osascript', '-e',
        f'tell application "Microsoft Excel" to open "{xlsx_path}"'],
        capture_output=True)
    time.sleep(2)
```

### Typical Workflow

```python
from openpyxl import Workbook
import subprocess

# 1. Create workbook with openpyxl (Excel NOT running)
wb = Workbook()
ws = wb.active
ws.title = 'Dashboard'

# 2. Write data, format, add charts, images, borders...
# ... all openpyxl operations ...

# 3. Save
wb.save(xlsx_path)

# 4. Open in Excel
subprocess.run(['open', xlsx_path])

# 5. (Optional) Run VBA macros via AppleScript
import time
time.sleep(2)  # Wait for Excel to open
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to do Visual Basic "Call RefreshAllPivots"'])
```

### No "Clean Handoff" Needed

Unlike the old xlwings+openpyxl dual-engine approach, there is no handoff complexity. The workflow is simple:
1. openpyxl creates/edits the file (Excel is NOT running)
2. Open the file in Excel when ready
3. Run VBA macros if needed via AppleScript
4. Done

If you need to edit the file again with openpyxl after Excel has had it open, close Excel first:
```python
subprocess.run(['osascript', '-e',
    'tell application "Microsoft Excel" to close active workbook saving yes'],
    capture_output=True)
time.sleep(1)
# Now openpyxl can safely edit the file
wb = load_workbook(xlsx_path)
```
