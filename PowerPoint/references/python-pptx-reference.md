# python-pptx Complete Reference

## Table of Contents

1. [Standard Imports](#standard-imports)
2. [Opening and Saving](#opening-and-saving)
3. [Adding Slides](#adding-slides)
4. [Slide Background — Gradient (lxml)](#slide-background--gradient-lxml)
5. [Slide Background — Solid](#slide-background--solid)
6. [Clearing All Shapes](#clearing-all-shapes)
7. [Adding Shapes](#adding-shapes)
8. [Common Shape Types](#common-shape-types)
9. [Adding Text Boxes](#adding-text-boxes)
10. [Multiple Paragraphs](#multiple-paragraphs)
11. [Mixed Formatting in One Paragraph](#mixed-formatting-in-one-paragraph)
12. [Letter Spacing (lxml)](#letter-spacing-lxml)
13. [Transparency (lxml)](#transparency-lxml)
14. [Rounded Rectangle Corner Radius (lxml)](#rounded-rectangle-corner-radius-lxml)
15. [Remove Shape Outline (lxml)](#remove-shape-outline-lxml)
16. [Adding Tables](#adding-tables)
17. [Adding Charts](#adding-charts)
18. [Adding Images](#adding-images)
19. [Reading All Slide Content (Audit)](#reading-all-slide-content-audit)
20. [Overlap and Bounds Checker](#overlap-and-bounds-checker)
21. [Text Frame Sizing](#text-frame-sizing)
22. [Embedded Helper Functions](#embedded-helper-functions)
23. [EMU Quick Reference](#emu-quick-reference)

---

## Standard Imports

```python
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
```

## Opening and Saving

```python
pptx_path = os.environ['PPTX_PATH']

# Open existing
prs = Presentation(pptx_path)

# OR create new blank
prs = Presentation()
prs.slide_width = Emu(12192000)
prs.slide_height = Emu(6858000)

# ALWAYS save at end:
prs.save(pptx_path)
```

## Adding Slides

```python
layout = prs.slide_layouts[6]  # 6 = Blank (fully custom)
slide = prs.slides.add_slide(layout)
```

## Slide Background — Gradient (lxml)

Most reliable method for gradients:

```python
bg_elem = slide.background._element
bgPr = bg_elem.find(qn('p:bgPr'))
if bgPr is None:
    bgPr = etree.SubElement(bg_elem, qn('p:bgPr'))

for child in list(bgPr):
    if child.tag != qn('a:effectLst'):
        bgPr.remove(child)

gradFill = etree.SubElement(bgPr, qn('a:gradFill'))
gsLst = etree.SubElement(gradFill, qn('a:gsLst'))

gs1 = etree.SubElement(gsLst, qn('a:gs')); gs1.set('pos', '0')
srgb1 = etree.SubElement(gs1, qn('a:srgbClr')); srgb1.set('val', '0B1D3A')

gs2 = etree.SubElement(gsLst, qn('a:gs')); gs2.set('pos', '100000')
srgb2 = etree.SubElement(gs2, qn('a:srgbClr')); srgb2.set('val', '162D50')

lin = etree.SubElement(gradFill, qn('a:lin'))
lin.set('ang', '5400000')  # top-to-bottom
lin.set('scaled', '1')

if bgPr.find(qn('a:effectLst')) is None:
    etree.SubElement(bgPr, qn('a:effectLst'))
```

## Slide Background — Solid

```python
bg = slide.background; fill = bg.fill; fill.solid()
fill.fore_color.rgb = RGBColor(0xF7, 0xF8, 0xFA)
```

## Clearing All Shapes

```python
for sp in list(slide.shapes):
    slide.shapes._spTree.remove(sp._element)
```

## Adding Shapes

```python
shape = slide.shapes.add_shape(
    MSO_SHAPE.ROUNDED_RECTANGLE,
    Emu(457200), Emu(1500000), Emu(5400000), Emu(4800000)
)
shape.fill.solid()
shape.fill.fore_color.rgb = RGBColor(0x1A, 0x33, 0x58)
shape.line.color.rgb = RGBColor(0x2A, 0x4A, 0x78)
shape.line.width = Pt(0.75)

# Remove outline:
shape.line.fill.background()
```

## Common Shape Types

```python
MSO_SHAPE.RECTANGLE           MSO_SHAPE.ROUNDED_RECTANGLE
MSO_SHAPE.OVAL                MSO_SHAPE.DIAMOND
MSO_SHAPE.RIGHT_TRIANGLE      MSO_SHAPE.ISOSCELES_TRIANGLE
MSO_SHAPE.RIGHT_ARROW         MSO_SHAPE.CHEVRON
MSO_SHAPE.STAR_5_POINT        MSO_SHAPE.HEART
MSO_SHAPE.LIGHTNING_BOLT      MSO_SHAPE.CROSS
MSO_SHAPE.DONUT
```

## Adding Text Boxes

```python
txBox = slide.shapes.add_textbox(
    Emu(457200), Emu(228600), Emu(10000000), Emu(600000)
)
tf = txBox.text_frame
tf.word_wrap = True
p = tf.paragraphs[0]
p.alignment = PP_ALIGN.LEFT
run = p.add_run()
run.text = "EXECUTIVE SUMMARY"
run.font.size = Pt(28)
run.font.bold = True
run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
run.font.name = "Georgia"
```

## Multiple Paragraphs

```python
p1 = tf.paragraphs[0]
run1 = p1.add_run(); run1.text = "Title Line"
run1.font.size = Pt(20); run1.font.bold = True
run1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
run1.font.name = "Georgia"

p2 = tf.add_paragraph(); p2.space_before = Pt(6)
run2 = p2.add_run(); run2.text = "Body text here"
run2.font.size = Pt(16)
run2.font.color.rgb = RGBColor(0x88, 0x99, 0xAA)
run2.font.name = "Georgia"
```

## Mixed Formatting in One Paragraph

```python
run1 = p.add_run(); run1.text = "Revenue: "
run1.font.size = Pt(18)
run1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

run2 = p.add_run(); run2.text = "$14.8B"
run2.font.size = Pt(18); run2.font.bold = True
run2.font.color.rgb = RGBColor(0xC9, 0xA8, 0x4C)
```

## Letter Spacing (lxml)

```python
rPr = run._r.get_or_add_rPr()
rPr.set('spc', '300')  # 300 = 3pt tracking
```

## Transparency (lxml)

```python
spPr = shape._element.spPr
solidFill = spPr.find(qn('a:solidFill'))
srgbClr = solidFill.find(qn('a:srgbClr'))
alpha = etree.SubElement(srgbClr, qn('a:alpha'))
alpha.set('val', '15000')  # 15% opacity (15000/1000)
```

### KNOWN PITFALL: Gradient fills on shapes via lxml

**DO NOT** add `<a:gradFill>` directly to `spPr` via `etree.SubElement` after `add_shape()`. This produces a light blue (theme-default) rectangle instead of your gradient because:

1. `add_shape()` creates shapes with theme style fills (`<p:style>`), not explicit `spPr` fills. Adding `gradFill` to `spPr` doesn't override the theme — the theme fill shows through.
2. `etree.SubElement` appends at the end, placing `gradFill` AFTER `<a:ln>`, which violates OOXML schema order. PowerPoint silently ignores out-of-order elements.

**The fix:** Always call `shape.fill.solid()` first (creates explicit fill that overrides theme), then replace that `solidFill` with `gradFill` via lxml, inserting BEFORE `<a:ln>`. See the `add_gradient_shape()` helper in the Embedded Helper Functions section.

## Rounded Rectangle Corner Radius (lxml)

```python
spPr = shape._element.spPr
prstGeom = spPr.find(qn('a:prstGeom'))
avLst = prstGeom.find(qn('a:avLst'))
if avLst is None:
    avLst = etree.SubElement(prstGeom, qn('a:avLst'))
else:
    for child in list(avLst):
        avLst.remove(child)
gd = etree.SubElement(avLst, qn('a:gd'))
gd.set('name', 'adj')
gd.set('fmla', 'val 5000')  # lower = less rounded
```

## Remove Shape Outline (lxml)

```python
spPr = shape._element.spPr
ln = spPr.find(qn('a:ln'))
if ln is None:
    ln = etree.SubElement(spPr, qn('a:ln'))
noFill = etree.SubElement(ln, qn('a:noFill'))
```

## Adding Tables

```python
rows, cols = 4, 5
table_shape = slide.shapes.add_table(
    rows, cols, Emu(457200), Emu(1500000), Emu(11200000), Emu(2000000)
)
table = table_shape.table
table.columns[0].width = Emu(3000000)

# Style header row
for col_idx in range(cols):
    cell = table.cell(0, col_idx)
    cell.text = ["Metric", "2022", "2023", "2024", "Change"][col_idx]
    for paragraph in cell.text_frame.paragraphs:
        paragraph.alignment = PP_ALIGN.CENTER
        for run in paragraph.runs:
            run.font.size = Pt(14); run.font.bold = True
            run.font.color.rgb = RGBColor(0x0B, 0x1D, 0x3A)
            run.font.name = "Georgia"
    cell.fill.solid()
    cell.fill.fore_color.rgb = RGBColor(0xC9, 0xA8, 0x4C)

# Alternating row colors
for row_idx in range(1, rows):
    for col_idx in range(cols):
        cell = table.cell(row_idx, col_idx)
        cell.text = "..."
        for paragraph in cell.text_frame.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(14)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                run.font.name = "Georgia"
        cell.fill.solid()
        if row_idx % 2 == 0:
            cell.fill.fore_color.rgb = RGBColor(0x15, 0x2C, 0x4D)
        else:
            cell.fill.fore_color.rgb = RGBColor(0x1A, 0x33, 0x58)
```

## Adding Charts

```python
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

chart_data = CategoryChartData()
chart_data.categories = ['Q1', 'Q2', 'Q3', 'Q4']
chart_data.add_series('Revenue', (100, 150, 120, 180))

chart_frame = slide.shapes.add_chart(
    XL_CHART_TYPE.COLUMN_CLUSTERED,
    Emu(1270000), Emu(1500000), Emu(9600000), Emu(4800000),
    chart_data
)
chart = chart_frame.chart
chart.has_legend = True
chart.legend.include_in_layout = False
```

## Adding Images

```python
from pptx.util import Emu

# Full-slide background image (16:9)
pic = slide.shapes.add_picture(
    'path/to/background.png',
    Emu(0), Emu(0), Emu(12192000), Emu(6858000)
)

# Positioned content image
pic = slide.shapes.add_picture(
    'path/to/photo.png',
    Emu(457200), Emu(1500000), Emu(5000000), Emu(3500000)
)

# Image inside a card area
pic = slide.shapes.add_picture(
    'path/to/image.png',
    Emu(card_left + 100000), Emu(card_top + 100000),
    Emu(card_width - 200000), Emu(2000000)
)
```

## Reading All Slide Content (Audit)

```python
for slide_idx, slide in enumerate(prs.slides):
    print(f"\n{'='*60}")
    print(f"--- Slide {slide_idx + 1} ---")
    for shape in slide.shapes:
        print(f"  Shape: type={shape.shape_type}, Name='{shape.name}'")
        print(f"    Pos: L={shape.left}, T={shape.top}, W={shape.width}, H={shape.height}")
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                txt = para.text.strip()
                if txt:
                    print(f"    Text: '{txt[:120]}'")
                    for run in para.runs:
                        fn = run.font.name
                        fs = run.font.size
                        fb = run.font.bold
                        try:
                            fc = str(run.font.color.rgb) if run.font.color and run.font.color.rgb else 'inherit'
                        except:
                            fc = 'inherit'
                        print(f"      Font: {fn}, Size: {fs}, Bold: {fb}, Color: {fc}")
        if hasattr(shape, 'image'):
            try:
                print(f"    [IMAGE] type={shape.image.content_type}, size={len(shape.image.blob)} bytes")
            except:
                pass
```

## Overlap and Bounds Checker

```python
def check_overlaps(slide):
    shapes_info = []
    for s in slide.shapes:
        shapes_info.append({
            'name': s.name,
            'left': s.left, 'top': s.top,
            'right': s.left + s.width,
            'bottom': s.top + s.height
        })
    for i, a in enumerate(shapes_info):
        for j, b in enumerate(shapes_info):
            if i >= j: continue
            if any(kw in a['name'] for kw in ['BG', 'Bar', 'Oval', 'accent']): continue
            if any(kw in b['name'] for kw in ['BG', 'Bar', 'Oval', 'accent']): continue
            if (a['left'] < b['right'] and a['right'] > b['left'] and
                a['top'] < b['bottom'] and a['bottom'] > b['top']):
                print(f"  OVERLAP: {a['name']} <-> {b['name']}")
    for s in shapes_info:
        if s['right'] > 12192000 or s['bottom'] > 6858000:
            print(f"  OUT OF BOUNDS: {s['name']}")
```

## Text Frame Sizing

**CRITICAL: Never guess text frame dimensions.** Always calculate width per-paragraph (sum ALL runs), compute wrapped line count against frame width, then derive height.

**The bug this prevents:** A paragraph with multiple runs (e.g., bold + normal text on the same line) has a rendered width equal to the SUM of all runs. Calculating per-run width makes each look fine, but the combined width exceeds the frame — causing wrapping the height doesn't account for, leading to text overflow.

### Sizing Helper Functions

```python
import math

PT = 12700  # 1 point in EMU

def run_width_emu(text, font_size, bold=False, spacing_hundredths=0):
    """Width of a single text run in EMU.
    Georgia avg char width: bold ~0.60*fs, regular ~0.55*fs.
    spacing_hundredths: letter spacing in hundredths of a point (e.g., 300 = 3pt)."""
    factor = 0.60 if bold else 0.55
    sp_pt = spacing_hundredths / 100.0
    return len(text) * (factor * font_size + sp_pt) * PT * 1.12  # 12% safety


def para_width_emu(runs):
    """Total width of a paragraph = SUM of all run widths.
    runs: list of (text, font_size, bold, spacing_hundredths) tuples.
    THIS IS THE KEY INSIGHT — always sum all runs, never measure individually."""
    return sum(run_width_emu(t, fs, b, sp) for t, fs, b, sp in runs)


def frame_dims(paragraphs, max_width):
    """Calculate text frame (width, height) accounting for line wrapping.
    paragraphs: list of dicts:
        {'runs': [(text, fs, bold, sp), ...],
         'max_fs': int,  # largest font size in this paragraph
         'space_before_pt': float,  # optional, default 0
         'space_after_pt': float}   # optional, default 0
    max_width: maximum allowed frame width in EMU.
    Returns (width_emu, height_emu)."""
    total_h_pt = 0
    for para in paragraphs:
        pw = para_width_emu(para['runs'])
        n_lines = max(1, math.ceil(pw / max_width))
        line_h_pt = para['max_fs'] * 1.35
        para_h_pt = (n_lines * line_h_pt
                     + para.get('space_before_pt', 0)
                     + para.get('space_after_pt', 0))
        total_h_pt += para_h_pt
    h = int(total_h_pt * PT * 1.18)  # 18% height safety
    return max_width, h


def single_line_dims(text, font_size, bold=False, spacing=0):
    """For text that must NOT wrap — returns (width, height) sized exactly.
    Use with word_wrap=False on the text frame."""
    w = int(run_width_emu(text, font_size, bold, spacing))
    h = int(font_size * 1.35 * PT * 1.18)
    return w, h
```

### Usage Rules

1. **Per-paragraph, not per-run**: A paragraph's width = sum of ALL its runs.
2. **Wrapped lines = `ceil(para_width / frame_width)`**: Always compute this.
3. **Single-line elements**: Use `word_wrap=False` and `single_line_dims()`.
4. **Multi-run paragraphs**: Use `frame_dims()` with all runs listed per paragraph.
5. **Safety margins**: 12% on width, 18% on height (Georgia renders slightly wider than calculated).

### Example

```python
# WRONG — measures runs separately, misses that they share one line:
w1 = run_width_emu("Bold part", 16, True)   # looks fine alone
w2 = run_width_emu(" — normal part", 15)     # looks fine alone
# But paragraph renders as w1+w2 on ONE line — may exceed frame!

# RIGHT — measure the full paragraph:
paras = [
    {'runs': [("Bold part", 16, True, 0), (" — normal part", 15, False, 0)],
     'max_fs': 16, 'space_after_pt': 8},
    {'runs': [("Second paragraph here.", 15, False, 0)],
     'max_fs': 15, 'space_before_pt': 8},
]
frame_w, frame_h = frame_dims(paras, max_width=5300000)
# frame_h now correctly accounts for any wrapping
```

---

## Embedded Helper Functions

Copy-paste these at the top of every Python script. They cover 90% of common operations:

```python
import os
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree


def hex_to_rgb(hex_str):
    """Convert '#RRGGBB' or 'RRGGBB' to RGBColor."""
    h = hex_str.lstrip('#')
    return RGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))


def set_gradient_bg(slide, hex_start, hex_end, angle=5400000):
    """Set slide gradient background via lxml."""
    bg_elem = slide.background._element
    bgPr = bg_elem.find(qn('p:bgPr'))
    if bgPr is None:
        bgPr = etree.SubElement(bg_elem, qn('p:bgPr'))
    for child in list(bgPr):
        if child.tag != qn('a:effectLst'):
            bgPr.remove(child)
    gradFill = etree.SubElement(bgPr, qn('a:gradFill'))
    gsLst = etree.SubElement(gradFill, qn('a:gsLst'))
    gs1 = etree.SubElement(gsLst, qn('a:gs')); gs1.set('pos', '0')
    srgb1 = etree.SubElement(gs1, qn('a:srgbClr')); srgb1.set('val', hex_start.lstrip('#'))
    gs2 = etree.SubElement(gsLst, qn('a:gs')); gs2.set('pos', '100000')
    srgb2 = etree.SubElement(gs2, qn('a:srgbClr')); srgb2.set('val', hex_end.lstrip('#'))
    lin = etree.SubElement(gradFill, qn('a:lin'))
    lin.set('ang', str(angle)); lin.set('scaled', '1')
    if bgPr.find(qn('a:effectLst')) is None:
        etree.SubElement(bgPr, qn('a:effectLst'))


def add_rect(slide, left, top, width, height, fill_hex, border_hex=None):
    """Add a rectangle shape with optional border."""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   Emu(left), Emu(top), Emu(width), Emu(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    if border_hex:
        shape.line.color.rgb = hex_to_rgb(border_hex)
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    return shape


def add_rounded_rect(slide, left, top, width, height, fill_hex,
                     border_hex=None, radius=5000):
    """Add a rounded rectangle with configurable corner radius."""
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                   Emu(left), Emu(top), Emu(width), Emu(height))
    shape.fill.solid()
    shape.fill.fore_color.rgb = hex_to_rgb(fill_hex)
    if border_hex:
        shape.line.color.rgb = hex_to_rgb(border_hex)
        shape.line.width = Pt(0.75)
    else:
        shape.line.fill.background()
    spPr = shape._element.spPr
    prstGeom = spPr.find(qn('a:prstGeom'))
    avLst = prstGeom.find(qn('a:avLst'))
    if avLst is None:
        avLst = etree.SubElement(prstGeom, qn('a:avLst'))
    else:
        for child in list(avLst):
            avLst.remove(child)
    gd = etree.SubElement(avLst, qn('a:gd'))
    gd.set('name', 'adj')
    gd.set('fmla', f'val {radius}')
    return shape


def add_textbox(slide, left, top, width, height, text,
                font_name="Georgia", font_size=16, bold=False,
                color_hex="FFFFFF", alignment=PP_ALIGN.LEFT, spacing=0):
    """Add a text box with formatted text."""
    txBox = slide.shapes.add_textbox(Emu(left), Emu(top), Emu(width), Emu(height))
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = hex_to_rgb(color_hex)
    run.font.name = font_name
    if spacing > 0:
        rPr = run._r.get_or_add_rPr()
        rPr.set('spc', str(spacing))
    return txBox


def set_transparency(shape, alpha_percent):
    """Set shape fill transparency. alpha_percent: 0=invisible, 100=opaque."""
    spPr = shape._element.spPr
    solidFill = spPr.find(qn('a:solidFill'))
    if solidFill is None:
        return
    srgbClr = solidFill.find(qn('a:srgbClr'))
    if srgbClr is None:
        return
    for existing in srgbClr.findall(qn('a:alpha')):
        srgbClr.remove(existing)
    alpha = etree.SubElement(srgbClr, qn('a:alpha'))
    alpha.set('val', str(int(alpha_percent * 1000)))


def clear_slide(slide):
    """Remove all shapes from a slide."""
    for sp in list(slide.shapes):
        slide.shapes._spTree.remove(sp._element)


def add_slide_header(slide, title_text, pal, subtitle_text=None, font_name="Georgia"):
    """Standard slide header: left accent bar + title + underline + optional subtitle.
    Returns Y position where content should start below the header."""
    add_rect(slide, 0, 0, 57150, 6858000, pal['accent'])
    add_textbox(slide, 457200, 200000, 10000000, 550000, title_text,
                font_name=font_name, font_size=30, bold=True,
                color_hex=pal['text_primary'], spacing=350)
    add_rect(slide, 457200, 750000, 2200000, 32000, pal['accent'])
    y = 900000
    if subtitle_text:
        add_textbox(slide, 457200, 830000, 10000000, 350000, subtitle_text,
                    font_name=font_name, font_size=16, color_hex=pal['text_muted'])
        y = 1250000
    return y


def add_slide_footer(slide, text, pal, font_name="Georgia"):
    """Standard slide footer bar."""
    add_rect(slide, 0, 6500000, 12192000, 358000, pal.get('footer_bg', '#0A1628'))
    add_textbox(slide, 457200, 6530000, 11280000, 300000, text,
                font_name=font_name, font_size=14,
                color_hex=pal['text_muted'], spacing=150)


def make_kpi_card(slide, left, top, width, height, label, value, note, pal,
                  font_name="Georgia"):
    """Create a KPI metric card with label, big value, and note."""
    add_rounded_rect(slide, left, top, width, height,
                     pal['card_fill'], pal['card_border'], radius=4000)
    add_rect(slide, left + 150000, top, width - 300000, 28000, pal['accent'])

    tb = slide.shapes.add_textbox(
        Emu(left + 180000), Emu(top + 180000),
        Emu(width - 360000), Emu(height - 280000)
    )
    tf = tb.text_frame; tf.word_wrap = True

    p = tf.paragraphs[0]; p.space_after = Pt(4)
    r = p.add_run(); r.text = label
    r.font.size = Pt(14); r.font.bold = True
    r.font.color.rgb = hex_to_rgb(pal['accent'])
    r.font.name = font_name
    rPr = r._r.get_or_add_rPr(); rPr.set('spc', '200')

    p2 = tf.add_paragraph(); p2.space_before = Pt(8); p2.space_after = Pt(4)
    r2 = p2.add_run(); r2.text = value
    r2.font.size = Pt(28); r2.font.bold = True
    r2.font.color.rgb = hex_to_rgb(pal['text_primary'])
    r2.font.name = font_name

    p3 = tf.add_paragraph(); p3.space_before = Pt(2)
    r3 = p3.add_run(); r3.text = note
    r3.font.size = Pt(14)
    r3.font.color.rgb = hex_to_rgb(pal['text_muted'])
    r3.font.name = font_name
    return tb


def add_bg_image(slide, image_path, overlay_hex='#000000', overlay_opacity=40):
    """Add full-bleed background image with semi-transparent overlay."""
    slide.shapes.add_picture(
        image_path, Emu(0), Emu(0), Emu(12192000), Emu(6858000)
    )
    overlay = add_rect(slide, 0, 0, 12192000, 6858000, overlay_hex)
    set_transparency(overlay, overlay_opacity)
    return overlay


def add_gradient_shape(slide, left, top, width, height, hex_start, hex_end,
                       angle='5400000', alpha_start=None, alpha_end=None):
    """Add a rectangle with gradient fill, properly overriding theme defaults.

    CRITICAL: Do NOT use etree.SubElement to add gradFill directly to spPr
    after add_shape(). This causes two bugs:
    1. add_shape() applies a theme fill via <p:style>, not an explicit spPr fill.
       Simply adding gradFill to spPr won't override the theme — both render,
       and the theme fill (often light blue) shows through transparent areas.
    2. SubElement appends gradFill AFTER <a:ln>, violating OOXML schema order
       (fill must come before ln). PowerPoint ignores out-of-order elements.

    The fix:
    1. Call shape.fill.solid() first — this creates an explicit fill in spPr
       that properly overrides the theme style.
    2. Remove that solidFill and replace with gradFill via lxml.
    3. Insert gradFill BEFORE <a:ln> to respect OOXML schema order.

    Args:
        alpha_start/alpha_end: 0=invisible, 100=fully opaque (maps to OOXML val*1000)
    """
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                    Emu(left), Emu(top), Emu(width), Emu(height))
    shape.line.fill.background()

    # Step 1: Override theme fill with explicit solid fill
    shape.fill.solid()
    shape.fill.fore_color.rgb = RGBColor(0, 0, 0)

    # Step 2: Remove solidFill, replace with gradFill
    spPr = shape._element.spPr
    solidFill = spPr.find(qn('a:solidFill'))
    if solidFill is not None:
        spPr.remove(solidFill)

    # Step 3: Insert gradFill BEFORE <a:ln> (OOXML schema order)
    ln = spPr.find(qn('a:ln'))
    gradFill = etree.Element(qn('a:gradFill'))
    if ln is not None:
        spPr.insert(list(spPr).index(ln), gradFill)
    else:
        spPr.append(gradFill)

    gsLst = etree.SubElement(gradFill, qn('a:gsLst'))
    gs1 = etree.SubElement(gsLst, qn('a:gs')); gs1.set('pos', '0')
    srgb1 = etree.SubElement(gs1, qn('a:srgbClr'))
    srgb1.set('val', hex_start.lstrip('#'))
    if alpha_start is not None:
        a = etree.SubElement(srgb1, qn('a:alpha'))
        a.set('val', str(int(alpha_start * 1000)))

    gs2 = etree.SubElement(gsLst, qn('a:gs')); gs2.set('pos', '100000')
    srgb2 = etree.SubElement(gs2, qn('a:srgbClr'))
    srgb2.set('val', hex_end.lstrip('#'))
    if alpha_end is not None:
        a = etree.SubElement(srgb2, qn('a:alpha'))
        a.set('val', str(int(alpha_end * 1000)))

    lin = etree.SubElement(gradFill, qn('a:lin'))
    lin.set('ang', str(angle)); lin.set('scaled', '1')
    return shape
```

## EMU Quick Reference

| Measurement | EMU Value |
|-------------|-----------|
| 1 inch | 914400 |
| 0.5 inch | 457200 |
| Slide width (10") | 12192000 |
| Slide height (7.5") | 6858000 |
| 1 point | 12700 |
| 1 cm | 360000 |
