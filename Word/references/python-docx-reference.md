# python-docx Complete Reference

## Table of Contents

1. [Standard Imports](#standard-imports)
2. [Opening and Saving](#opening-and-saving)
3. [Page Setup](#page-setup)
4. [Sections](#sections)
5. [Paragraphs](#paragraphs)
6. [Runs and Formatting](#runs-and-formatting)
7. [Paragraph Formatting](#paragraph-formatting)
8. [Multiple Paragraphs with Mixed Formatting](#multiple-paragraphs-with-mixed-formatting)
9. [Styles](#styles)
10. [Headings](#headings)
11. [Lists](#lists)
12. [Tables](#tables)
13. [Table Cell Formatting (lxml)](#table-cell-formatting-lxml)
14. [Images](#images)
15. [Headers and Footers](#headers-and-footers)
16. [Page Breaks](#page-breaks)
17. [Horizontal Rules (lxml)](#horizontal-rules-lxml)
18. [Paragraph Shading (lxml)](#paragraph-shading-lxml)
19. [Paragraph Borders (lxml)](#paragraph-borders-lxml)
20. [Page Borders (lxml)](#page-borders-lxml)
21. [Tab Stops](#tab-stops)
22. [Watermarks (lxml)](#watermarks-lxml)
23. [Table of Contents Placeholder (lxml)](#table-of-contents-placeholder-lxml)
24. [Reading Document Content (Audit)](#reading-document-content-audit)
25. [Embedded Helper Functions](#embedded-helper-functions)
26. [Unit Quick Reference](#unit-quick-reference)

---

## Standard Imports

```python
import os
from docx import Document
from docx.shared import Inches, Pt, Cm, Emu, RGBColor, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from lxml import etree
```

## Opening and Saving

```python
docx_path = os.environ['DOCX_PATH']

# Open existing
doc = Document(docx_path)

# OR create new blank
doc = Document()

# ALWAYS save at end:
doc.save(docx_path)
```

## Page Setup

```python
# Access via first section (or any section)
section = doc.sections[0]

# Page size — Letter (8.5" x 11")
section.page_width = Inches(8.5)
section.page_height = Inches(11)

# Page size — A4
section.page_width = Cm(21)
section.page_height = Cm(29.7)

# Margins
section.top_margin = Inches(1)
section.bottom_margin = Inches(1)
section.left_margin = Inches(1)
section.right_margin = Inches(1)

# Orientation
section.orientation = WD_ORIENT.PORTRAIT  # or LANDSCAPE

# Calculate content width
content_width = section.page_width - section.left_margin - section.right_margin
# Letter standard: 8.5" - 1" - 1" = 6.5" = Inches(6.5) = 5943600 EMU

# Header/footer distance from edge
section.header_distance = Inches(0.5)
section.footer_distance = Inches(0.5)

# Different first page header/footer
section.different_first_page_header_footer = True
```

## Sections

```python
# Add a new section (starts a new page with potentially different layout)
from docx.enum.section import WD_SECTION_START

new_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
# Options: NEW_PAGE, EVEN_PAGE, ODD_PAGE, CONTINUOUS, NEW_COLUMN

# Copy settings from previous section
prev = doc.sections[-2]
new_section.page_width = prev.page_width
new_section.page_height = prev.page_height
new_section.left_margin = prev.left_margin
new_section.right_margin = prev.right_margin

# Columns (via lxml — python-docx doesn't expose columns directly)
sectPr = new_section._sectPr
cols = sectPr.find(qn('w:cols'))
if cols is None:
    cols = parse_xml(f'<w:cols {nsdecls("w")} w:num="2" w:space="720"/>')
    sectPr.append(cols)
else:
    cols.set(qn('w:num'), '2')
    cols.set(qn('w:space'), '720')  # 720 twips = 0.5 inch gap
```

## Paragraphs

```python
# Add a paragraph
p = doc.add_paragraph('Hello, World!')

# Add paragraph with a style
p = doc.add_paragraph('Heading text', style='Heading 1')

# Access existing paragraphs
for para in doc.paragraphs:
    print(para.text)
    print(para.style.name)

# Clear all content from document
for element in doc.element.body:
    doc.element.body.remove(element)
```

## Runs and Formatting

```python
# A run is a contiguous piece of text with the same formatting
p = doc.add_paragraph()
run = p.add_run('Bold text')
run.bold = True
run.font.size = Pt(14)
run.font.name = 'Georgia'
run.font.color.rgb = RGBColor(0x0B, 0x1D, 0x3A)

# Multiple runs in one paragraph (mixed formatting)
p = doc.add_paragraph()
r1 = p.add_run('Revenue: ')
r1.font.size = Pt(12)
r1.font.color.rgb = RGBColor(0x33, 0x41, 0x55)

r2 = p.add_run('$14.8B')
r2.bold = True
r2.font.size = Pt(12)
r2.font.color.rgb = RGBColor(0xC9, 0xA8, 0x4C)

# Font properties
run.font.size = Pt(12)
run.font.name = 'Georgia'
run.font.bold = True
run.font.italic = True
run.font.underline = True
run.font.strike = True       # strikethrough
run.font.all_caps = True
run.font.small_caps = True
run.font.color.rgb = RGBColor(0xFF, 0x00, 0x00)
run.font.highlight_color = None  # WD_COLOR_INDEX values

# Superscript / subscript
run.font.superscript = True
run.font.subscript = True
```

## Paragraph Formatting

```python
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING

pf = p.paragraph_format

# Alignment
pf.alignment = WD_ALIGN_PARAGRAPH.LEFT      # LEFT, CENTER, RIGHT, JUSTIFY

# Spacing
pf.space_before = Pt(12)     # space above paragraph
pf.space_after = Pt(6)       # space below paragraph
pf.line_spacing = 1.15       # multiplier (1.0, 1.15, 1.5, 2.0)
# OR exact line spacing:
pf.line_spacing_rule = WD_LINE_SPACING.EXACTLY
pf.line_spacing = Pt(18)

# Indentation
pf.left_indent = Inches(0.5)
pf.right_indent = Inches(0.5)
pf.first_line_indent = Inches(0.25)  # first line of paragraph

# Keep with next (prevent orphan headings)
pf.keep_with_next = True

# Keep together (don't split across pages)
pf.keep_together = True

# Page break before
pf.page_break_before = True

# Widow/orphan control
pf.widow_control = True
```

## Multiple Paragraphs with Mixed Formatting

```python
# Title + subtitle in one text frame approach
p_title = doc.add_paragraph()
p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_title.space_after = Pt(4)
r = p_title.add_run('QUARTERLY REPORT')
r.bold = True
r.font.size = Pt(28)
r.font.name = 'Montserrat'
r.font.color.rgb = RGBColor(0x0B, 0x1D, 0x3A)

p_sub = doc.add_paragraph()
p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
p_sub.space_before = Pt(4)
r2 = p_sub.add_run('Q4 2024 Financial Summary')
r2.font.size = Pt(14)
r2.font.name = 'Georgia'
r2.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)
```

## Styles

```python
# Use built-in styles
p = doc.add_paragraph('Title', style='Title')
p = doc.add_paragraph('Heading', style='Heading 1')
p = doc.add_paragraph('Body text', style='Normal')
p = doc.add_paragraph('Quote text', style='Quote')
p = doc.add_paragraph('Intense quote', style='Intense Quote')

# Modify an existing style
style = doc.styles['Normal']
style.font.name = 'Georgia'
style.font.size = Pt(11)
style.font.color.rgb = RGBColor(0x1A, 0x20, 0x2C)
style.paragraph_format.space_after = Pt(8)
style.paragraph_format.line_spacing = 1.15

# Modify heading styles
h1 = doc.styles['Heading 1']
h1.font.name = 'Montserrat'
h1.font.size = Pt(24)
h1.font.bold = True
h1.font.color.rgb = RGBColor(0x0B, 0x1D, 0x3A)
h1.paragraph_format.space_before = Pt(24)
h1.paragraph_format.space_after = Pt(8)
h1.paragraph_format.keep_with_next = True

h2 = doc.styles['Heading 2']
h2.font.name = 'Montserrat'
h2.font.size = Pt(18)
h2.font.bold = True
h2.font.color.rgb = RGBColor(0x1E, 0x3A, 0x5F)
h2.paragraph_format.space_before = Pt(18)
h2.paragraph_format.space_after = Pt(6)

h3 = doc.styles['Heading 3']
h3.font.name = 'Montserrat'
h3.font.size = Pt(14)
h3.font.bold = True
h3.font.color.rgb = RGBColor(0x1E, 0x3A, 0x5F)

# Create a custom character style (for inline accent)
from docx.enum.style import WD_STYLE_TYPE
accent_style = doc.styles.add_style('AccentText', WD_STYLE_TYPE.CHARACTER)
accent_style.font.color.rgb = RGBColor(0xC9, 0xA8, 0x4C)
accent_style.font.bold = True
# Use it: run = p.add_run('gold text', style='AccentText')
```

## Headings

```python
# Built-in heading levels (1-9)
doc.add_heading('Main Title', level=1)
doc.add_heading('Sub Section', level=2)
doc.add_heading('Detail', level=3)

# Heading with custom formatting (override style)
h = doc.add_heading('', level=1)
run = h.add_run('EXECUTIVE SUMMARY')
run.font.size = Pt(24)
run.font.name = 'Montserrat'
run.font.color.rgb = RGBColor(0x0B, 0x1D, 0x3A)
# Add letter spacing via lxml
rPr = run._element.get_or_add_rPr()
spacing = parse_xml(f'<w:spacing {nsdecls("w")} w:val="60"/>')  # 60 = 3pt in half-points
rPr.append(spacing)
```

## Lists

```python
# Bullet list
doc.add_paragraph('First item', style='List Bullet')
doc.add_paragraph('Second item', style='List Bullet')
doc.add_paragraph('Third item', style='List Bullet')

# Numbered list
doc.add_paragraph('Step one', style='List Number')
doc.add_paragraph('Step two', style='List Number')
doc.add_paragraph('Step three', style='List Number')

# Nested bullets (List Bullet 2, List Bullet 3)
doc.add_paragraph('Main point', style='List Bullet')
doc.add_paragraph('Sub-point', style='List Bullet 2')
doc.add_paragraph('Sub-sub-point', style='List Bullet 3')

# Nested numbers
doc.add_paragraph('Main step', style='List Number')
doc.add_paragraph('Sub-step', style='List Number 2')
```

## Tables

```python
# Create a table
rows, cols = 4, 3
table = doc.add_table(rows=rows, cols=cols)
table.style = 'Table Grid'  # visible borders
table.alignment = WD_TABLE_ALIGNMENT.CENTER

# Pre-calculate column widths (MUST sum to content width)
content_width = Inches(6.5)  # Letter, 1" margins
col_widths = [Inches(2.5), Inches(2.0), Inches(2.0)]  # sum = 6.5"

# Set column widths
for i, width in enumerate(col_widths):
    for row in table.rows:
        row.cells[i].width = width

# Populate header row
headers = ['Metric', 'Q3 2024', 'Q4 2024']
for i, header in enumerate(headers):
    cell = table.cell(0, i)
    cell.text = header
    # Format header text
    for paragraph in cell.paragraphs:
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        for run in paragraph.runs:
            run.bold = True
            run.font.size = Pt(11)
            run.font.name = 'Georgia'
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    # Header background (via lxml)
    shading_elm = parse_xml(
        f'<w:shd {nsdecls("w")} w:fill="0B1D3A" w:val="clear"/>'
    )
    cell._tc.get_or_add_tcPr().append(shading_elm)

# Populate data rows with alternating colors
data = [
    ['Revenue', '$12.4B', '$14.8B'],
    ['Margin', '23.1%', '25.4%'],
    ['Users', '1.1M', '1.3M'],
]
for row_idx, row_data in enumerate(data):
    for col_idx, value in enumerate(row_data):
        cell = table.cell(row_idx + 1, col_idx)
        cell.text = value
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(11)
                run.font.name = 'Georgia'
        # Alternating row shading
        if row_idx % 2 == 0:
            shading = parse_xml(
                f'<w:shd {nsdecls("w")} w:fill="F1F5F9" w:val="clear"/>'
            )
            cell._tc.get_or_add_tcPr().append(shading)

# Cell vertical alignment
cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

# Cell margins/padding (via lxml)
tc = cell._tc
tcPr = tc.get_or_add_tcPr()
tcMar = parse_xml(
    f'<w:tcMar {nsdecls("w")}>'
    f'  <w:top w:w="60" w:type="dxa"/>'
    f'  <w:start w:w="80" w:type="dxa"/>'
    f'  <w:bottom w:w="60" w:type="dxa"/>'
    f'  <w:end w:w="80" w:type="dxa"/>'
    f'</w:tcMar>'
)
tcPr.append(tcMar)

# Merge cells
table.cell(0, 0).merge(table.cell(0, 2))  # merge header across 3 cols

# Remove all borders from table (clean/minimal look)
# See Table Cell Formatting section for border control
```

## Table Cell Formatting (lxml)

```python
# Cell shading (background color)
def set_cell_shading(cell, hex_color):
    """Set cell background color."""
    shading = parse_xml(
        f'<w:shd {nsdecls("w")} w:fill="{hex_color}" w:val="clear"/>'
    )
    cell._tc.get_or_add_tcPr().append(shading)

# Cell borders
def set_cell_border(cell, **kwargs):
    """Set cell borders. kwargs: top, bottom, start, end, insideH, insideV.
    Each value is a dict: {'sz': '4', 'val': 'single', 'color': '000000', 'space': '0'}"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
        tcPr.append(tcBorders)
    for edge, attrs in kwargs.items():
        element = tcBorders.find(qn(f'w:{edge}'))
        if element is None:
            element = parse_xml(f'<w:{edge} {nsdecls("w")}/>')
            tcBorders.append(element)
        for key, val in attrs.items():
            element.set(qn(f'w:{key}'), val)

# Example: thick accent top border on header cell
set_cell_border(header_cell,
    top={'sz': '24', 'val': 'single', 'color': 'C9A84C', 'space': '0'})

# Remove all borders from a cell
def remove_cell_borders(cell):
    """Remove all borders from a cell."""
    none_border = {'sz': '0', 'val': 'none', 'color': 'auto', 'space': '0'}
    set_cell_border(cell,
        top=none_border, bottom=none_border,
        start=none_border, end=none_border)

# Remove all borders from entire table
def remove_table_borders(table):
    """Remove all borders from a table (minimal/clean look)."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
        tbl.insert(0, tblPr)
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'  <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'</w:tblBorders>'
    )
    existing = tblPr.find(qn('w:tblBorders'))
    if existing is not None:
        tblPr.remove(existing)
    tblPr.append(borders)

# Set table width to 100% of content area
def set_table_full_width(table):
    """Make table span full content width."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
        tbl.insert(0, tblPr)
    tblW = tblPr.find(qn('w:tblW'))
    if tblW is None:
        tblW = parse_xml(f'<w:tblW {nsdecls("w")} w:type="pct" w:w="5000"/>')
        tblPr.append(tblW)
    else:
        tblW.set(qn('w:type'), 'pct')
        tblW.set(qn('w:w'), '5000')  # 5000 = 100%
```

## Images

**CRITICAL: Never distort images. Always preserve aspect ratio. Read source dimensions before inserting.**

```python
from docx.shared import Inches
from PIL import Image as PILImage

# ═══════════════════════════════════════════════════════════════
# SAFE IMAGE INSERTION (USE THESE — never raw add_picture with both w+h)
# ═══════════════════════════════════════════════════════════════

def calc_image_size(img_path, max_width_inches=6.5, max_height_inches=4.0):
    """Calculate image display size preserving aspect ratio.
    ALWAYS use this before inserting any image."""
    img = PILImage.open(img_path)
    w_px, h_px = img.size
    ratio = w_px / h_px
    width = min(max_width_inches, max_height_inches * ratio)
    height = width / ratio
    if height > max_height_inches:
        height = max_height_inches
        width = height * ratio
    return Inches(width), Inches(height)


def safe_add_picture(doc, img_path, max_width_inches=6.5, max_height_inches=9.0):
    """Insert image with LOCKED aspect ratio. The safe way to add images.
    Reads source dimensions, calculates proportional size, inserts correctly.
    NEVER specify both width and height manually — use this instead."""
    w, h = calc_image_size(img_path, max_width_inches, max_height_inches)
    return doc.add_picture(img_path, width=w, height=h)


def safe_add_picture_to_cell(cell, img_path, max_width_inches=2.0, max_height_inches=3.0):
    """Insert image into a table cell with LOCKED aspect ratio."""
    w, h = calc_image_size(img_path, max_width_inches, max_height_inches)
    paragraph = cell.paragraphs[0]
    run = paragraph.add_run()
    return run.add_picture(img_path, width=w, height=h)


# ═══════════════════════════════════════════════════════════════
# CROP-TO-FIT (when image must fill a specific area without distortion)
# ═══════════════════════════════════════════════════════════════

def crop_to_aspect(img_path, target_width, target_height, output_path=None):
    """Center-crop an image to match a target aspect ratio WITHOUT distortion.
    Use this when an image must fill an exact W×H area (e.g., cover banner).
    Crops the source to the target ratio, then the caller inserts at target size.
    Returns the path to the cropped image."""
    img = PILImage.open(img_path)
    src_w, src_h = img.size
    target_ratio = target_width / target_height
    src_ratio = src_w / src_h

    if src_ratio > target_ratio:
        # Source is wider — crop sides
        new_w = int(src_h * target_ratio)
        offset = (src_w - new_w) // 2
        crop_box = (offset, 0, offset + new_w, src_h)
    else:
        # Source is taller — crop top/bottom
        new_h = int(src_w / target_ratio)
        offset = (src_h - new_h) // 2
        crop_box = (0, offset, src_w, offset + new_h)

    cropped = img.crop(crop_box)
    if output_path is None:
        import os
        base, ext = os.path.splitext(img_path)
        output_path = f"{base}_cropped{ext}"
    cropped.save(output_path, quality=95)
    return output_path


# ═══════════════════════════════════════════════════════════════
# ROUNDED CORNERS (visual polish — modern, professional feel)
# ═══════════════════════════════════════════════════════════════

def round_corners(img_path, radius=25, output_path=None):
    """Add rounded corners to an image. Saves as PNG to preserve transparency.
    radius: corner radius in pixels. Scale with image size:
      - Small images (< 500px): radius 15-20
      - Medium images (500-1500px): radius 25-35
      - Large images (> 1500px): radius 40-60
    Returns path to the rounded image."""
    img = PILImage.open(img_path).convert('RGBA')
    w, h = img.size

    # Create rounded mask
    from PIL import ImageDraw
    mask = PILImage.new('L', (w, h), 0)
    draw = ImageDraw.Draw(mask)
    draw.rounded_rectangle([(0, 0), (w, h)], radius=radius, fill=255)

    # Apply mask
    result = PILImage.new('RGBA', (w, h), (255, 255, 255, 0))
    result.paste(img, mask=mask)

    if output_path is None:
        import os
        base, _ = os.path.splitext(img_path)
        output_path = f"{base}_rounded.png"
    result.save(output_path, 'PNG')
    return output_path


def prepare_image(img_path, max_width_inches=6.5, max_height_inches=9.0,
                  corner_radius=25, crop_ratio=None):
    """All-in-one image preparation: optional crop, rounded corners, safe sizing.
    This is the recommended entry point for all image insertions.

    Args:
        img_path: Source image path
        max_width_inches: Maximum display width
        max_height_inches: Maximum display height
        corner_radius: Rounded corner radius in px (0 to skip rounding)
        crop_ratio: Target W/H ratio to crop to (None to skip cropping)

    Returns: (processed_path, width_inches, height_inches)
    """
    import os
    processed = img_path

    # Step 1: Crop to target ratio if specified
    if crop_ratio is not None:
        processed = crop_to_aspect(processed, crop_ratio, 1.0,
                                   output_path=os.path.splitext(img_path)[0] + '_prepped_crop.png')

    # Step 2: Round corners if radius > 0
    if corner_radius > 0:
        processed = round_corners(processed, radius=corner_radius,
                                  output_path=os.path.splitext(img_path)[0] + '_prepped_final.png')

    # Step 3: Calculate safe display size
    w, h = calc_image_size(processed, max_width_inches, max_height_inches)
    return processed, w, h


# ═══════════════════════════════════════════════════════════════
# BASIC INSERTION (only for simple cases — prefer safe_ helpers above)
# ═══════════════════════════════════════════════════════════════

# Inline image — ONLY specify width, let height auto-calculate
doc.add_picture('path/to/image.png', width=Inches(6.5))
# Height is auto-calculated to preserve aspect ratio ✓

# ⚠️ DANGER — specifying both width AND height risks distortion:
# doc.add_picture('path/to/image.png', width=Inches(4), height=Inches(3))
# Only do this if you've VERIFIED the ratio matches the source image!

# Image inside a table cell (use safe_add_picture_to_cell instead)
cell = table.cell(0, 0)
paragraph = cell.paragraphs[0]
run = paragraph.add_run()
run.add_picture('path/to/image.png', width=Inches(2))
# Height auto-calculated ✓

# Center an image
last_paragraph = doc.paragraphs[-1]
last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

# ═══════════════════════════════════════════════════════════════
# POST-INSERTION AUDIT (verify no distortion)
# ═══════════════════════════════════════════════════════════════

def audit_image_ratios(doc):
    """Check all inline images for aspect ratio consistency.
    Run this after building any document with images."""
    for i, shape in enumerate(doc.inline_shapes):
        w, h = shape.width, shape.height
        ratio = w / h if h else 0
        print(f"  Image {i+1}: {w}x{h} EMU, ratio={ratio:.3f}")
    # Compare against expected ratios from source files
```

## Headers and Footers

```python
# Access header/footer via section
section = doc.sections[0]

# Header
header = section.header
header.is_linked_to_previous = False  # independent from previous section
hp = header.paragraphs[0]
hp.text = 'Company Name'
hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
for run in hp.runs:
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x64, 0x74, 0x8B)

# Footer
footer = section.footer
footer.is_linked_to_previous = False
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Add page number to footer (via lxml — requires field code)
run = fp.add_run()
fldChar_begin = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
run._r.append(fldChar_begin)

run2 = fp.add_run()
instrText = parse_xml(f'<w:instrText {nsdecls("w")} xml:space="preserve"> PAGE </w:instrText>')
run2._r.append(instrText)

run3 = fp.add_run()
fldChar_end = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
run3._r.append(fldChar_end)

# Different first page (no header/footer on cover page)
section.different_first_page_header_footer = True
# The first page header/footer is accessed via:
first_header = section.first_page_header
first_footer = section.first_page_footer
# Leave them empty for a clean cover page

# Add a horizontal line in header/footer
# Use paragraph bottom border (see Paragraph Borders section)
```

## Page Breaks

```python
# Simple page break
doc.add_page_break()

# Page break via paragraph format (before a heading)
p = doc.add_paragraph('New Section', style='Heading 1')
p.paragraph_format.page_break_before = True

# Section break (for changing layout — margins, orientation, columns)
from docx.enum.section import WD_SECTION_START
new_section = doc.add_section(WD_SECTION_START.NEW_PAGE)
```

## Horizontal Rules (lxml)

```python
def add_horizontal_rule(doc, color='C9A84C', thickness_pt=1.5, width_pct=100):
    """Add a horizontal rule as a paragraph with a bottom border.
    color: hex without #. thickness_pt: line thickness. width_pct: ignored (full width)."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    # Convert thickness to eighths of a point
    sz = str(int(thickness_pt * 8))
    pPr = p._element.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="single" w:sz="{sz}" w:space="1" w:color="{color}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)
    return p


def add_accent_underline(doc, color='C9A84C', thickness_pt=2):
    """Add a short accent underline (like the pptx-design-agent accent bar).
    This is a thin paragraph border, visually separating title from content."""
    return add_horizontal_rule(doc, color, thickness_pt)
```

## Paragraph Shading (lxml)

```python
def set_paragraph_shading(paragraph, hex_color):
    """Add background shading to a paragraph."""
    pPr = paragraph._element.get_or_add_pPr()
    shd = parse_xml(
        f'<w:shd {nsdecls("w")} w:fill="{hex_color}" w:val="clear"/>'
    )
    pPr.append(shd)


# Example: callout paragraph with shading + left indent
p = doc.add_paragraph()
p.paragraph_format.left_indent = Inches(0.15)
p.paragraph_format.right_indent = Inches(0.15)
p.paragraph_format.space_before = Pt(8)
p.paragraph_format.space_after = Pt(8)
set_paragraph_shading(p, 'F8FAFC')
run = p.add_run('Key insight: This is an important callout.')
run.font.size = Pt(11)
run.font.name = 'Georgia'
```

## Paragraph Borders (lxml)

```python
def set_paragraph_border(paragraph, side='left', color='C9A84C', sz='24', val='single'):
    """Add a border to one side of a paragraph.
    side: top, bottom, left, right.
    sz: thickness in eighths of a point (24 = 3pt).
    val: single, double, dotted, dashed, thick, etc."""
    pPr = paragraph._element.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}/>')
        pPr.append(pBdr)
    border = parse_xml(
        f'<w:{side} {nsdecls("w")} w:val="{val}" w:sz="{sz}" '
        f'w:space="4" w:color="{color}"/>'
    )
    existing = pBdr.find(qn(f'w:{side}'))
    if existing is not None:
        pBdr.remove(existing)
    pBdr.append(border)


# Example: callout box = left border + shading
p = doc.add_paragraph()
p.paragraph_format.left_indent = Inches(0.25)
p.paragraph_format.space_before = Pt(10)
p.paragraph_format.space_after = Pt(10)
set_paragraph_shading(p, 'F8FAFC')
set_paragraph_border(p, 'left', 'C9A84C', '24')  # 3pt gold left border
run = p.add_run('Important: ')
run.bold = True
run.font.color.rgb = RGBColor(0xC9, 0xA8, 0x4C)
run2 = p.add_run('This is a key takeaway for the quarter.')
run2.font.size = Pt(11)
```

## Page Borders (lxml)

```python
def add_page_border(section, color='C9A84C', sz='12', val='single', space='24'):
    """Add a border around the entire page.
    sz: eighths of a point. space: distance from text in points."""
    sectPr = section._sectPr
    pgBorders = parse_xml(
        f'<w:pgBorders {nsdecls("w")} w:offsetFrom="page">'
        f'  <w:top w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'
        f'  <w:left w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'
        f'  <w:bottom w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'
        f'  <w:right w:val="{val}" w:sz="{sz}" w:space="{space}" w:color="{color}"/>'
        f'</w:pgBorders>'
    )
    existing = sectPr.find(qn('w:pgBorders'))
    if existing is not None:
        sectPr.remove(existing)
    # Insert after pgSz/pgMar if they exist
    ref = sectPr.find(qn('w:pgMar'))
    if ref is not None:
        ref.addnext(pgBorders)
    else:
        sectPr.append(pgBorders)
```

## Tab Stops

```python
from docx.enum.text import WD_TAB_ALIGNMENT, WD_TAB_LEADER

pf = p.paragraph_format
tab_stops = pf.tab_stops

# Right-aligned tab at 6.5" (for right-aligned date on same line as left text)
tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)

# Dot leader tab (for TOC-style entries)
tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)

# Center tab
tab_stops.add_tab_stop(Inches(3.25), WD_TAB_ALIGNMENT.CENTER)

# Use tab in text
run = p.add_run('Left text\tRight text')
```

## Watermarks (lxml)

```python
def add_text_watermark(section, text='DRAFT', color='C0C0C0'):
    """Add a text watermark to the document via header.
    Note: This adds a shape to the default header."""
    header = section.header
    header.is_linked_to_previous = False
    # Watermark is implemented as a VML shape in the header
    # This is complex XML — simplified version:
    p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    pPr = p._element.get_or_add_pPr()
    # The full watermark implementation requires VML XML
    # For production, consider using python-docx-template or direct XML manipulation
    # Simpler alternative: add faded centered text to header
    run = p.add_run(text)
    run.font.size = Pt(72)
    run.font.color.rgb = RGBColor(
        int(color[0:2], 16), int(color[2:4], 16), int(color[4:6], 16))
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
```

## Table of Contents Placeholder (lxml)

```python
def add_toc_placeholder(doc, title='Table of Contents'):
    """Add a TOC placeholder. Word/AppleScript must update fields to populate it."""
    # Add TOC heading
    doc.add_heading(title, level=1)

    # Add TOC field code
    p = doc.add_paragraph()
    run = p.add_run()
    r = run._r

    fldChar_begin = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="begin"/>')
    r.append(fldChar_begin)

    run2 = p.add_run()
    instrText = parse_xml(
        f'<w:instrText {nsdecls("w")} xml:space="preserve">'
        f' TOC \\o "1-3" \\h \\z \\u </w:instrText>'
    )
    run2._r.append(instrText)

    run3 = p.add_run()
    fldChar_separate = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="separate"/>')
    run3._r.append(fldChar_separate)

    run4 = p.add_run('[Update fields to populate Table of Contents]')
    run4.font.color.rgb = RGBColor(0x94, 0xA3, 0xB8)
    run4.font.size = Pt(10)

    run5 = p.add_run()
    fldChar_end = parse_xml(f'<w:fldChar {nsdecls("w")} w:fldCharType="end"/>')
    run5._r.append(fldChar_end)

    doc.add_page_break()
```

## Reading Document Content (Audit)

```python
def audit_document(doc):
    """Print a structured audit of the entire document."""
    print(f"Sections: {len(doc.sections)}")
    for i, section in enumerate(doc.sections):
        print(f"\n--- Section {i+1} ---")
        print(f"  Page: {section.page_width}x{section.page_height}")
        print(f"  Margins: L={section.left_margin} R={section.right_margin} "
              f"T={section.top_margin} B={section.bottom_margin}")
        print(f"  Orientation: {section.orientation}")

    print(f"\nParagraphs: {len(doc.paragraphs)}")
    for i, para in enumerate(doc.paragraphs):
        style = para.style.name
        text = para.text[:100] if para.text else ''
        if text or style != 'Normal':
            print(f"  [{i}] Style='{style}' | '{text}'")
            for run in para.runs:
                fn = run.font.name
                fs = run.font.size
                fb = run.font.bold
                try:
                    fc = str(run.font.color.rgb) if run.font.color and run.font.color.rgb else 'inherit'
                except:
                    fc = 'inherit'
                print(f"    Font: {fn}, Size: {fs}, Bold: {fb}, Color: {fc}")

    print(f"\nTables: {len(doc.tables)}")
    for i, table in enumerate(doc.tables):
        rows = len(table.rows)
        cols = len(table.columns)
        print(f"  Table {i+1}: {rows}x{cols}")
        for row_idx, row in enumerate(table.rows):
            cells_text = [cell.text[:30] for cell in row.cells]
            print(f"    Row {row_idx}: {cells_text}")

    print(f"\nInline Shapes (images): {len(doc.inline_shapes)}")
    for i, shape in enumerate(doc.inline_shapes):
        print(f"  Image {i+1}: W={shape.width}, H={shape.height}")
```

## Embedded Helper Functions

Copy-paste these at the top of every Python script. They cover 90% of common operations:

```python
import os
from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


def hex_to_rgb(hex_str):
    """Convert '#RRGGBB' or 'RRGGBB' to RGBColor."""
    h = hex_str.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_cell_shading(cell, hex_color):
    """Set cell background color."""
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color.lstrip("#")}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shd)


def set_paragraph_shading(paragraph, hex_color):
    """Set paragraph background shading."""
    pPr = paragraph._element.get_or_add_pPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color.lstrip("#")}" w:val="clear"/>')
    pPr.append(shd)


def set_paragraph_border(paragraph, side='left', color='C9A84C', sz='24', val='single'):
    """Add border to paragraph side. sz in eighths of a point."""
    pPr = paragraph._element.get_or_add_pPr()
    pBdr = pPr.find(qn('w:pBdr'))
    if pBdr is None:
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}/>')
        pPr.append(pBdr)
    border = parse_xml(
        f'<w:{side} {nsdecls("w")} w:val="{val}" w:sz="{sz}" '
        f'w:space="4" w:color="{color.lstrip("#")}"/>')
    existing = pBdr.find(qn(f'w:{side}'))
    if existing is not None:
        pBdr.remove(existing)
    pBdr.append(border)


def add_horizontal_rule(doc, color='C9A84C', thickness_pt=1.5):
    """Add a colored horizontal rule."""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    sz = str(int(thickness_pt * 8))
    pPr = p._element.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'  <w:bottom w:val="single" w:sz="{sz}" w:space="1" w:color="{color.lstrip("#")}"/>'
        f'</w:pBdr>')
    pPr.append(pBdr)
    return p


def add_callout_box(doc, title, body, pal, title_size=11, body_size=11):
    """Add a callout box (shaded paragraph with left accent border).
    Returns the paragraph for further customization."""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.2)
    p.paragraph_format.right_indent = Inches(0.2)
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(10)
    set_paragraph_shading(p, pal['callout_bg'])
    set_paragraph_border(p, 'left', pal['callout_border'], '24')
    if title:
        r_title = p.add_run(title + '  ')
        r_title.bold = True
        r_title.font.size = Pt(title_size)
        r_title.font.color.rgb = hex_to_rgb(pal['accent'])
        r_title.font.name = 'Georgia'
    r_body = p.add_run(body)
    r_body.font.size = Pt(body_size)
    r_body.font.color.rgb = hex_to_rgb(pal['body'])
    r_body.font.name = 'Georgia'
    return p


def add_pull_quote(doc, text, pal, size=16):
    """Add an indented pull quote in accent color."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.left_indent = Inches(0.75)
    p.paragraph_format.right_indent = Inches(0.75)
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(18)
    r = p.add_run(f'\u201c{text}\u201d')
    r.italic = True
    r.font.size = Pt(size)
    r.font.color.rgb = hex_to_rgb(pal['accent'])
    r.font.name = 'Georgia'
    return p


def make_kpi_row(doc, metrics, pal):
    """Create a KPI card row as a table.
    metrics: list of dicts: {'label': str, 'value': str, 'note': str}
    Returns the table."""
    n = len(metrics)
    table = doc.add_table(rows=1, cols=n)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    remove_table_borders(table)

    content_width = Inches(6.5)
    col_w = content_width / n
    for i, metric in enumerate(metrics):
        cell = table.cell(0, i)
        cell.width = int(col_w)
        # Accent top border
        set_cell_border_single(cell, 'top', pal['accent'], '18')
        # Padding
        set_cell_padding(cell, top=80, bottom=80, start=100, end=100)

        # Label
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.space_after = Pt(2)
        r = p.add_run(metric['label'])
        r.bold = True
        r.font.size = Pt(9)
        r.font.color.rgb = hex_to_rgb(pal['accent'])
        r.font.name = 'Georgia'

        # Value
        p2 = cell.add_paragraph()
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.space_after = Pt(2)
        r2 = p2.add_run(metric['value'])
        r2.bold = True
        r2.font.size = Pt(22)
        r2.font.color.rgb = hex_to_rgb(pal['heading'])
        r2.font.name = 'Georgia'

        # Note
        p3 = cell.add_paragraph()
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r3 = p3.add_run(metric['note'])
        r3.font.size = Pt(9)
        color = pal['positive'] if '+' in metric['note'] or '\u25b2' in metric['note'] else (
            pal['negative'] if '-' in metric['note'] or '\u25bc' in metric['note'] else pal['muted'])
        r3.font.color.rgb = hex_to_rgb(color)
        r3.font.name = 'Georgia'

    return table


def remove_table_borders(table):
    """Remove all borders from a table."""
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}/>')
        tbl.insert(0, tblPr)
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'  <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>'
        f'</w:tblBorders>')
    existing = tblPr.find(qn('w:tblBorders'))
    if existing is not None:
        tblPr.remove(existing)
    tblPr.append(borders)


def set_cell_border_single(cell, side, color, sz='12', val='single'):
    """Set a single border on a cell side."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.find(qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
        tcPr.append(tcBorders)
    border = parse_xml(
        f'<w:{side} {nsdecls("w")} w:val="{val}" w:sz="{sz}" '
        f'w:space="0" w:color="{color.lstrip("#")}"/>')
    existing = tcBorders.find(qn(f'w:{side}'))
    if existing is not None:
        tcBorders.remove(existing)
    tcBorders.append(border)


def set_cell_padding(cell, top=60, bottom=60, start=80, end=80):
    """Set cell padding in DXA (twips). 72 twips = 1pt."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'  <w:top w:w="{top}" w:type="dxa"/>'
        f'  <w:start w:w="{start}" w:type="dxa"/>'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'  <w:end w:w="{end}" w:type="dxa"/>'
        f'</w:tcMar>')
    existing = tcPr.find(qn('w:tcMar'))
    if existing is not None:
        tcPr.remove(existing)
    tcPr.append(tcMar)


def setup_styles(doc, pal, heading_font='Montserrat', body_font='Georgia'):
    """Configure document styles to match palette."""
    # Normal (body)
    style = doc.styles['Normal']
    style.font.name = body_font
    style.font.size = Pt(11)
    style.font.color.rgb = hex_to_rgb(pal['body'])
    style.paragraph_format.space_after = Pt(8)
    style.paragraph_format.line_spacing = 1.15

    # Headings
    for level, size, before, after in [(1, 24, 24, 8), (2, 18, 18, 6), (3, 14, 14, 4)]:
        h = doc.styles[f'Heading {level}']
        h.font.name = heading_font
        h.font.size = Pt(size)
        h.font.bold = True
        h.font.color.rgb = hex_to_rgb(pal['heading'])
        h.paragraph_format.space_before = Pt(before)
        h.paragraph_format.space_after = Pt(after)
        h.paragraph_format.keep_with_next = True


def make_styled_table(doc, headers, rows, pal, col_widths=None):
    """Create a professionally styled table.
    headers: list of str. rows: list of list of str.
    col_widths: list of Inches values (must sum to content width)."""
    n_cols = len(headers)
    n_rows = len(rows) + 1
    table = doc.add_table(rows=n_rows, cols=n_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set column widths
    if col_widths:
        for i, w in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = w

    # Header row
    for i, header in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(header)
        r.bold = True
        r.font.size = Pt(10)
        r.font.name = 'Georgia'
        r.font.color.rgb = hex_to_rgb(pal['table_header_tx'])
        set_cell_shading(cell, pal['table_header_bg'])
        set_cell_padding(cell)

    # Data rows
    for row_idx, row_data in enumerate(rows):
        for col_idx, value in enumerate(row_data):
            cell = table.cell(row_idx + 1, col_idx)
            cell.text = ''
            p = cell.paragraphs[0]
            r = p.add_run(str(value))
            r.font.size = Pt(10)
            r.font.name = 'Georgia'
            r.font.color.rgb = hex_to_rgb(pal['body'])
            if row_idx % 2 == 0:
                set_cell_shading(cell, pal['table_alt_row'])
            set_cell_padding(cell)

    return table
```

## Unit Quick Reference

| Measurement | Inches | Points | Twips (DXA) | EMU |
|-------------|--------|--------|-------------|-----|
| 1 inch | 1 | 72 | 1440 | 914400 |
| 0.5 inch | 0.5 | 36 | 720 | 457200 |
| Letter width (8.5") | 8.5 | 612 | 12240 | 7772400 |
| Letter height (11") | 11 | 792 | 15840 | 10058400 |
| A4 width (8.27") | 8.27 | 595.4 | 11907 | 7560310 |
| A4 height (11.69") | 11.69 | 841.7 | 16839 | 10692130 |
| Content width (1" margins) | 6.5 | 468 | 9360 | 5943600 |
| 1 point | 0.0139 | 1 | 20 | 12700 |
| 1 cm | 0.394 | 28.35 | 567 | 360000 |

**python-docx units:**
- `Inches(1)` = 914400 EMU
- `Pt(12)` = 152400 EMU
- `Cm(1)` = 360000 EMU
- `Emu(914400)` = 1 inch
- `Twips(1440)` = 1 inch

**OOXML XML units:**
- `w:w`, `w:sz` for borders: eighths of a point (8 = 1pt, 24 = 3pt)
- `w:w` for spacing/indent: twips (1440 = 1 inch, 720 = 0.5 inch)
- `w:type="dxa"`: value is in twips
- `w:type="pct"`: value is in fiftieths of a percent (5000 = 100%)
