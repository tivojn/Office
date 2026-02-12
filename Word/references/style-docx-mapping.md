# Style → python-docx Implementation Mapping

Concrete python-docx values for each document design style. Read `references/design-styles-catalog.md` for full style descriptions — this file provides code-level implementation.

## Common Imports

```python
from docx.shared import Inches, Pt, Cm, Emu, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
```

---

## STYLE-01 — Strategy Consulting

```python
STYLE_01 = {
    "name": "Strategy Consulting",
    "fonts": {
        "heading": {"name": "Georgia", "size": Pt(24), "bold": True, "color": RGBColor(0x0B, 0x1D, 0x3A)},
        "subheading": {"name": "Georgia", "size": Pt(18), "bold": True, "color": RGBColor(0x0B, 0x1D, 0x3A)},
        "body": {"name": "Calibri", "size": Pt(11), "bold": False, "color": RGBColor(0x1A, 0x1A, 0x1A)},
        "caption": {"name": "Calibri", "size": Pt(9), "bold": False, "color": RGBColor(0x99, 0x99, 0x99)},
    },
    "palette": {
        "heading": "#0B1D3A",
        "body": "#1A1A1A",
        "muted": "#999999",
        "accent": "#003DA5",
        "accent2": "#0B1D3A",
        "table_header_bg": "#003DA5",
        "table_header_tx": "#FFFFFF",
        "table_alt_row": "#F5F5F5",
        "callout_bg": "#F8F9FA",
        "callout_border": "#003DA5",
        "cover_accent": "#003DA5",
        "border": "#CCCCCC",
        "positive": "#16A34A",
        "negative": "#DC2626",
    },
    "page": {"margins": Inches(1), "line_spacing": 1.15},
    "cover_pattern": "classic_corporate",
    "table_style": "banded",
    "accent_rule": {"color": "#003DA5", "width": Pt(1)},
    "design_notes": "No shadows, no 3D. Hairline borders. Max 2 colors per page. Data-first.",
}
```

---

## STYLE-02 — Executive Editorial

```python
STYLE_02 = {
    "name": "Executive Editorial",
    "fonts": {
        "heading": {"name": "Rockwell", "size": Pt(26), "bold": True, "color": RGBColor(0xC8, 0x10, 0x2E)},
        "subheading": {"name": "Georgia", "size": Pt(18), "bold": True, "color": RGBColor(0x33, 0x33, 0x33)},
        "body": {"name": "Georgia", "size": Pt(11), "bold": False, "color": RGBColor(0x33, 0x33, 0x33)},
        "pullquote": {"name": "Georgia", "size": Pt(16), "bold": True, "italic": True, "color": RGBColor(0xC8, 0x10, 0x2E)},
    },
    "palette": {
        "heading": "#C8102E",
        "body": "#333333",
        "muted": "#8C8279",
        "accent": "#C8102E",
        "accent2": "#8C8279",
        "table_header_bg": "#C8102E",
        "table_header_tx": "#FFFFFF",
        "table_alt_row": "#FAF7F2",
        "callout_bg": "#FAF7F2",
        "callout_border": "#C8102E",
        "cover_accent": "#C8102E",
        "border": "#B8B0A8",
        "positive": "#16A34A",
        "negative": "#DC2626",
    },
    "page": {"margins_lr": Inches(1.25), "margins_tb": Inches(1), "line_spacing": 1.3},
    "cover_pattern": "minimal_elegant",
    "table_style": "accent_top",
    "accent_rule": {"color": "#C8102E", "width": Pt(0.75)},
    "design_notes": "Generous whitespace. Crimson accents sparingly. Pull quotes as focal points.",
}
```

---

## STYLE-03 — Creative Brief

```python
STYLE_03 = {
    "name": "Creative Brief",
    "fonts": {
        "heading": {"name": "Comic Sans MS", "size": Pt(24), "bold": True, "color": RGBColor(0x3C, 0x3C, 0x3C)},
        "subheading": {"name": "Calibri", "size": Pt(16), "bold": True, "color": RGBColor(0x3C, 0x3C, 0x3C)},
        "body": {"name": "Calibri", "size": Pt(11), "bold": False, "color": RGBColor(0x3C, 0x3C, 0x3C)},
        "annotation": {"name": "Calibri", "size": Pt(10), "bold": False, "italic": True, "color": RGBColor(0xA0, 0xA0, 0xA0)},
    },
    "palette": {
        "heading": "#3C3C3C",
        "body": "#3C3C3C",
        "muted": "#A0A0A0",
        "accent": "#D46A3C",
        "accent2": "#2B4570",
        "table_header_bg": "#3C3C3C",
        "table_header_tx": "#FFFFFF",
        "table_alt_row": "#F5F0E1",
        "callout_bg": "#F5F0E1",
        "callout_border": "#D46A3C",
        "cover_accent": "#D46A3C",
        "border": "#A0A0A0",
        "positive": "#5A8C46",
        "negative": "#C44E3B",
    },
    "page": {"margins": Inches(0.75), "line_spacing": 1.2},
    "cover_pattern": "side_accent",
    "table_style": "thin_borders",
    "accent_rule": {"color": "#D46A3C", "width": Pt(1), "style": "dashed"},
    "design_notes": "Loose layout. Nothing perfectly aligned. Annotation-style labels. Dashed rules.",
}
```

---

## STYLE-04 — Playful / Kawaii

```python
STYLE_04 = {
    "name": "Playful / Kawaii",
    "fonts": {
        "heading": {"name": "Calibri", "size": Pt(26), "bold": True, "color": RGBColor(0x4A, 0x20, 0x40)},
        "subheading": {"name": "Calibri", "size": Pt(16), "bold": True, "color": RGBColor(0x4A, 0x20, 0x40)},
        "body": {"name": "Calibri", "size": Pt(12), "bold": False, "color": RGBColor(0x4A, 0x20, 0x40)},
    },
    "palette": {
        "heading": "#4A2040",
        "body": "#4A2040",
        "muted": "#A0A0A0",
        "accent": "#FFAB91",
        "accent2": "#81D4FA",
        "accent3": "#FFF59D",
        "bg_pink": "#FFD6E0",
        "bg_lavender": "#E8D5F5",
        "bg_mint": "#D5F5E3",
        "table_header_bg": "#FFD6E0",
        "table_header_tx": "#4A2040",
        "table_alt_row": "#FFF0F5",
        "callout_bg": "#E8D5F5",
        "callout_border": "#FFAB91",
        "cover_accent": "#FFAB91",
        "border": "#E8D5F5",
        "positive": "#66BB6A",
        "negative": "#EF5350",
    },
    "page": {"margins": Inches(1), "line_spacing": 1.3},
    "cover_pattern": "bold_banner",
    "table_style": "pastel_banded",
    "accent_rule": {"color": "#FFAB91", "width": Pt(2)},
    "design_notes": "Generous padding. Soft colors. Rotate pastel bgs across sections. Extra cell padding in tables.",
}
```

---

## STYLE-05 — Corporate Modern

```python
STYLE_05 = {
    "name": "Corporate Modern",
    "fonts": {
        "heading": {"name": "Calibri", "size": Pt(28), "bold": True, "color": RGBColor(0x1D, 0x1D, 0x1F)},
        "subheading": {"name": "Calibri", "size": Pt(18), "bold": False, "color": RGBColor(0x6E, 0x6E, 0x73)},
        "body": {"name": "Calibri", "size": Pt(11), "bold": False, "color": RGBColor(0x1D, 0x1D, 0x1F)},
        "metric": {"name": "Calibri", "size": Pt(36), "bold": True, "color": RGBColor(0x00, 0x71, 0xE3)},
    },
    "palette": {
        "heading": "#1D1D1F",
        "body": "#1D1D1F",
        "muted": "#6E6E73",
        "accent": "#0071E3",
        "accent2": "#6E6E73",
        "table_header_bg": "#0071E3",
        "table_header_tx": "#FFFFFF",
        "table_alt_row": "#F8F9FA",
        "callout_bg": "#FFFFFF",
        "callout_border": "#D1D1D6",
        "callout_accent": "#0071E3",
        "cover_accent": "#0071E3",
        "border": "#D1D1D6",
        "positive": "#16A34A",
        "negative": "#DC2626",
    },
    "page": {"margins": Inches(1), "line_spacing": 1.15},
    "cover_pattern": "classic_corporate",
    "table_style": "banded",
    "accent_rule": {"color": "#0071E3", "width": Pt(1)},
    "design_notes": "Card-based sections. Systematic spacing. Blue accent underlines. KPI cards. Subtle borders.",
}
```

---

## STYLE-06 — Bold Narrative

```python
STYLE_06 = {
    "name": "Bold Narrative",
    "fonts": {
        "heading": {"name": "Arial Black", "size": Pt(32), "bold": True, "color": RGBColor(0x1A, 0x1A, 0x40)},
        "subheading": {"name": "Calibri", "size": Pt(18), "bold": True, "color": RGBColor(0x1A, 0x1A, 0x40)},
        "body": {"name": "Calibri", "size": Pt(11), "bold": False, "color": RGBColor(0x1A, 0x1A, 0x40)},
        "pullquote": {"name": "Calibri", "size": Pt(20), "bold": True, "color": RGBColor(0xFF, 0xB3, 0x00)},
    },
    "palette": {
        "heading": "#1A1A40",
        "body": "#1A1A40",
        "muted": "#6E6E73",
        "accent": "#FFB300",
        "accent2": "#1E88E5",
        "accent3": "#E91E63",
        "table_header_bg": "#1A1A40",
        "table_header_tx": "#FFFFFF",
        "table_alt_row": "#F5F5F5",
        "callout_bg": "#1A1A40",
        "callout_tx": "#FFFFFF",
        "callout_border": "#FFB300",
        "cover_accent": "#FFB300",
        "border": "#1A1A40",
        "positive": "#66BB6A",
        "negative": "#EF5350",
    },
    "page": {"margins_lr": Inches(1.25), "margins_tb": Inches(1), "line_spacing": 1.2},
    "cover_pattern": "bold_banner",
    "table_style": "accent_top",
    "accent_rule": {"color": "#FFB300", "width": Pt(4)},
    "design_notes": "Large dramatic titles. Dark callout boxes (inverted). Thick accent bars. Cinematic whitespace.",
}
```

---

## STYLE-07 — Warm Organic

```python
STYLE_07 = {
    "name": "Warm Organic",
    "fonts": {
        "heading": {"name": "Cambria", "size": Pt(24), "bold": True, "color": RGBColor(0x2D, 0x1B, 0x0E)},
        "subheading": {"name": "Cambria", "size": Pt(16), "bold": True, "color": RGBColor(0x2D, 0x1B, 0x0E)},
        "body": {"name": "Calibri", "size": Pt(11), "bold": False, "color": RGBColor(0x44, 0x38, 0x2C)},
    },
    "palette": {
        "heading": "#2D1B0E",
        "body": "#44382C",
        "muted": "#8B7355",
        "accent": "#C2704E",
        "accent2": "#7C9A6E",
        "table_header_bg": "#C2704E",
        "table_header_tx": "#FFFFFF",
        "table_alt_row": "#FDF8F0",
        "callout_bg": "#FEF7ED",
        "callout_border": "#C2704E",
        "cover_accent": "#C2704E",
        "border": "#D4C5B0",
        "positive": "#7C9A6E",
        "negative": "#C2704E",
    },
    "page": {"margins": Inches(1), "line_spacing": 1.3},
    "cover_pattern": "minimal_elegant",
    "table_style": "accent_top",
    "accent_rule": {"color": "#C2704E", "width": Pt(1)},
    "design_notes": "Generous whitespace. Warm tones. Two-color accent (terracotta + sage). Rounded image corners.",
}
```

---

## STYLE-08 — Magazine Editorial

```python
STYLE_08 = {
    "name": "Magazine Editorial",
    "fonts": {
        "heading": {"name": "Georgia", "size": Pt(36), "bold": True, "color": RGBColor(0x0A, 0x0A, 0x0A)},
        "subheading": {"name": "Georgia", "size": Pt(20), "bold": True, "color": RGBColor(0x0A, 0x0A, 0x0A)},
        "body": {"name": "Georgia", "size": Pt(11), "bold": False, "color": RGBColor(0x0A, 0x0A, 0x0A)},
        "pullquote": {"name": "Georgia", "size": Pt(22), "bold": True, "italic": True, "color": RGBColor(0xFF, 0xD6, 0x00)},
        "caption": {"name": "Calibri", "size": Pt(9), "bold": False, "color": RGBColor(0x75, 0x75, 0x75)},
    },
    "palette": {
        "heading": "#0A0A0A",
        "body": "#0A0A0A",
        "muted": "#757575",
        "accent": "#FFD600",
        "accent2": "#E91E63",
        "accent3": "#1A237E",
        "table_header_bg": "#0A0A0A",
        "table_header_tx": "#FFFFFF",
        "table_alt_row": "#F5F5F5",
        "callout_bg": "#0A0A0A",
        "callout_tx": "#FFFFFF",
        "callout_border": "#FFD600",
        "cover_accent": "#FFD600",
        "border": "#0A0A0A",
        "positive": "#16A34A",
        "negative": "#DC2626",
    },
    "page": {"margins_left": Inches(1.5), "margins_right": Inches(1), "margins_tb": Inches(1), "line_spacing": 1.2},
    "cover_pattern": "bold_banner",
    "table_style": "minimal_rules",
    "accent_rule": {"color": "#0A0A0A", "width": Pt(3)},
    "design_notes": "Asymmetric margins. Massive headlines 36pt+. Thick divider rules. One bold accent per page. High contrast.",
}
```

---

## STYLE-09 — Technical Documentation

```python
STYLE_09 = {
    "name": "Technical Documentation",
    "fonts": {
        "heading": {"name": "Calibri", "size": Pt(22), "bold": True, "color": RGBColor(0x2C, 0x2C, 0x2C)},
        "subheading": {"name": "Calibri", "size": Pt(16), "bold": True, "color": RGBColor(0x2C, 0x2C, 0x2C)},
        "body": {"name": "Calibri", "size": Pt(11), "bold": False, "color": RGBColor(0x2C, 0x2C, 0x2C)},
        "code": {"name": "Consolas", "size": Pt(10), "bold": False, "color": RGBColor(0x2C, 0x2C, 0x2C)},
        "step_number": {"name": "Calibri", "size": Pt(14), "bold": True, "color": RGBColor(0xE5, 0x39, 0x35)},
    },
    "palette": {
        "heading": "#2C2C2C",
        "body": "#2C2C2C",
        "muted": "#757575",
        "accent": "#E53935",
        "accent2": "#FFEB3B",
        "code_bg": "#F0EDE8",
        "table_header_bg": "#2C2C2C",
        "table_header_tx": "#FFFFFF",
        "table_alt_row": "#F8F8F8",
        "callout_bg": "#EBF5FF",
        "callout_border": "#2563EB",
        "warning_bg": "#FFFDE7",
        "warning_border": "#FFEB3B",
        "error_bg": "#FFEBEE",
        "error_border": "#E53935",
        "cover_accent": "#E53935",
        "border": "#E0E0E0",
        "positive": "#16A34A",
        "negative": "#E53935",
    },
    "page": {"margins": Inches(1), "line_spacing": 1.15},
    "cover_pattern": "classic_corporate",
    "table_style": "banded",
    "accent_rule": {"color": "#E0E0E0", "width": Pt(0.5)},
    "design_notes": "Scannable. Deep heading hierarchy. Code blocks with warm-white bg. Numbered panels. Note/Warning/Error callouts.",
}
```

---

## STYLE-10 — Dashboard Report

```python
STYLE_10 = {
    "name": "Dashboard Report",
    "fonts": {
        "heading": {"name": "Calibri", "size": Pt(20), "bold": True, "color": RGBColor(0x1C, 0x1C, 0x1E)},
        "subheading": {"name": "Calibri", "size": Pt(14), "bold": True, "color": RGBColor(0x1C, 0x1C, 0x1E)},
        "body": {"name": "Calibri", "size": Pt(10), "bold": False, "color": RGBColor(0x1C, 0x1C, 0x1E)},
        "metric": {"name": "Calibri", "size": Pt(40), "bold": True, "color": RGBColor(0x00, 0x71, 0xE3)},
        "metric_label": {"name": "Calibri", "size": Pt(10), "bold": False, "color": RGBColor(0xA0, 0xA0, 0xA0)},
    },
    "palette": {
        "heading": "#1C1C1E",
        "body": "#1C1C1E",
        "muted": "#A0A0A0",
        "accent_blue": "#0071E3",
        "accent_green": "#34C759",
        "accent_purple": "#AF52DE",
        "accent_orange": "#FF9F0A",
        "card_bg": "#F5F5F7",
        "card_dark": "#2C2C2E",
        "table_header_bg": "#1C1C1E",
        "table_header_tx": "#FFFFFF",
        "table_alt_row": "#F5F5F7",
        "callout_bg": "#F5F5F7",
        "callout_border": "#0071E3",
        "cover_accent": "#0071E3",
        "border": "#D1D1D6",
        "positive": "#34C759",
        "negative": "#FF3B30",
    },
    "page": {"margins": Inches(0.75), "line_spacing": 1.1},
    "cover_pattern": "minimal_elegant",
    "table_style": "kpi_cards",
    "accent_rule": {"color": "#D1D1D6", "width": Pt(0.5)},
    "design_notes": "Metric-forward. KPI cards prominent. Color-coded metrics. Dense but organized. Narrow margins.",
}
```

---

## STYLE-11 — Portfolio / Gallery

```python
STYLE_11 = {
    "name": "Portfolio / Gallery",
    "fonts": {
        "heading": {"name": "Calibri", "size": Pt(18), "bold": False, "color": RGBColor(0x1A, 0x1A, 0x1A)},
        "body": {"name": "Calibri", "size": Pt(10), "bold": False, "color": RGBColor(0x75, 0x75, 0x75)},
        "caption": {"name": "Calibri", "size": Pt(9), "bold": False, "color": RGBColor(0x75, 0x75, 0x75)},
    },
    "palette": {
        "heading": "#1A1A1A",
        "body": "#757575",
        "muted": "#A0A0A0",
        "accent": "#E60023",
        "table_header_bg": "#1A1A1A",
        "table_header_tx": "#FFFFFF",
        "table_alt_row": "#FAFAFA",
        "callout_bg": "#FFFFFF",
        "callout_border": "#EAEAEA",
        "cover_accent": "#E60023",
        "border": "#EAEAEA",
        "positive": "#16A34A",
        "negative": "#DC2626",
    },
    "page": {"margins": Inches(0.75), "line_spacing": 1.15},
    "cover_pattern": "bold_banner",
    "table_style": "invisible",
    "accent_rule": {"color": "#EAEAEA", "width": Pt(0.5)},
    "design_notes": "Images are hero. 60%+ page area is images. Small elegant captions. Rounded corners on all images. Generous whitespace.",
}
```

---

## STYLE-12 — Retro / Vintage

```python
STYLE_12 = {
    "name": "Retro / Vintage",
    "fonts": {
        "heading": {"name": "Rockwell", "size": Pt(28), "bold": True, "color": RGBColor(0x1B, 0x28, 0x38)},
        "subheading": {"name": "Rockwell", "size": Pt(18), "bold": True, "color": RGBColor(0x1B, 0x28, 0x38)},
        "body": {"name": "Consolas", "size": Pt(10), "bold": False, "color": RGBColor(0x1B, 0x28, 0x38)},
    },
    "palette": {
        "heading": "#1B2838",
        "body": "#1B2838",
        "muted": "#5A6A7A",
        "accent": "#0078BF",
        "accent2": "#FF665E",
        "accent3": "#00A95C",
        "bg_paper": "#F2E8D5",
        "table_header_bg": "#1B2838",
        "table_header_tx": "#F2E8D5",
        "table_alt_row": "#F2E8D5",
        "callout_bg": "#F2E8D5",
        "callout_border": "#1B2838",
        "cover_accent": "#0078BF",
        "border": "#1B2838",
        "positive": "#00A95C",
        "negative": "#FF665E",
    },
    "page": {"margins": Inches(1), "line_spacing": 1.3},
    "cover_pattern": "side_accent",
    "table_style": "thick_borders",
    "accent_rule": {"color": "#1B2838", "width": Pt(2)},
    "design_notes": "Max 2-3 colors. No pure black — use dark navy #1B2838. Bold flat shapes. Typewriter body. Paper bg tint.",
}
```

---

## Style Application Workflow

When applying a style, follow this order:

1. **Set page margins** from style `page` config
2. **Define heading/body styles** with style fonts + colors
3. **Apply palette colors** to all decorative elements (rules, borders, callouts)
4. **Build cover page** using the style's `cover_pattern`
5. **Style all tables** using the style's `table_style` pattern
6. **Follow design notes** for style-specific layout decisions
7. **Run the mandatory audit** — style changes don't exempt from quality checks
