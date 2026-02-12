# Style → Excel Implementation Mapping

Concrete Python values for each Excel design style. Read `references/design-styles-catalog.md` for full style descriptions — this file provides code-level implementation.

## Common Imports & Helpers

```python
import pandas as pd
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

def hex_to_rgb(hex_color):
    """Convert '#RRGGBB' to (R, G, B) tuple."""
    h = hex_color.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def hex_to_openpyxl(hex_color):
    """Convert '#RRGGBB' to openpyxl color string 'RRGGBB'."""
    return hex_color.lstrip('#')

def rgb_to_hex(r, g, b):
    """Convert (R, G, B) tuple to openpyxl hex string 'RRGGBB'."""
    return f'{r:02X}{g:02X}{b:02X}'
```

---

## XSTYLE-01 — Consulting & Strategy

```python
XSTYLE_01 = {
    "name": "Consulting & Strategy",
    "fonts": {
        "title":     {"name": "Calibri", "size": 20, "bold": True},
        "header":    {"name": "Calibri", "size": 12, "bold": True},
        "body":      {"name": "Georgia", "size": 11, "bold": False},
        "kpi_value": {"name": "Calibri", "size": 32, "bold": True},
        "kpi_label": {"name": "Calibri", "size": 10, "bold": True},
        "caption":   {"name": "Calibri", "size": 9,  "bold": False},
    },
    "palette": {
        "header_bg":   (11, 29, 58),       # #0B1D3A deep navy
        "header_text": (255, 255, 255),     # White
        "accent":      (201, 168, 76),      # #C9A84C gold
        "accent2":     (11, 29, 58),        # Navy
        "text":        (26, 32, 44),        # #1A202C
        "muted":       (100, 116, 139),     # #64748B
        "alt_row":     (241, 245, 249),     # #F1F5F9
        "border":      (226, 232, 240),     # #E2E8F0
        "card_bg":     (248, 250, 252),     # #F8FAFC
        "positive":    (22, 163, 74),       # #16A34A
        "negative":    (220, 38, 38),       # #DC2626
        "chart_1":     (201, 168, 76),      # Gold
        "chart_2":     (11, 29, 58),        # Navy
        "chart_3":     (59, 130, 246),      # Blue
        "chart_4":     (100, 116, 139),     # Slate
        "kpi_bg":      (11, 29, 58),        # Navy
        "kpi_text":    (255, 255, 255),     # White
        "kpi_label":   (201, 168, 76),      # Gold
    },
    "tab_colors": {
        "Dashboard": (201, 168, 76),        # Gold
        "Data":      (11, 29, 58),          # Navy
        "Charts":    (59, 130, 246),        # Blue
        "Reference": (100, 116, 139),       # Slate
        "Audit":     (220, 38, 38),         # Red
    },
    "table_style": "banded",
    "kpi_style": "accent_top",              # Gold thick top border
    "cond_format": {
        "positive_bg":   (220, 252, 231),   # Light green
        "positive_text": (22, 163, 74),     # Green
        "negative_bg":   (254, 226, 226),   # Light red
        "negative_text": (220, 38, 38),     # Red
        "data_bar":      (201, 168, 76),    # Gold
    },
    "print": {"orientation": "landscape", "paper": "letter"},
    "design_notes": "Restrained. Max 2-3 colors per sheet. Gold accents sparingly. Hairline borders.",
}
```

---

## XSTYLE-02 — Executive Dashboard

```python
XSTYLE_02 = {
    "name": "Executive Dashboard",
    "fonts": {
        "title":     {"name": "Calibri", "size": 20, "bold": True},
        "header":    {"name": "Calibri", "size": 12, "bold": True},
        "body":      {"name": "Calibri", "size": 11, "bold": False},
        "kpi_value": {"name": "Calibri", "size": 36, "bold": True},
        "kpi_label": {"name": "Calibri", "size": 10, "bold": False},
        "caption":   {"name": "Calibri", "size": 9,  "bold": False},
    },
    "palette": {
        "header_bg":   (29, 29, 31),        # #1D1D1F near-black
        "header_text": (255, 255, 255),     # White
        "accent":      (0, 113, 227),       # #0071E3 electric blue
        "accent2":     (110, 110, 115),     # #6E6E73
        "text":        (29, 29, 31),        # #1D1D1F
        "muted":       (110, 110, 115),     # #6E6E73
        "alt_row":     (248, 249, 250),     # #F8F9FA
        "border":      (209, 209, 214),     # #D1D1D6
        "card_bg":     (255, 255, 255),     # White
        "positive":    (52, 199, 89),       # #34C759
        "negative":    (255, 59, 48),       # #FF3B30
        "chart_1":     (0, 113, 227),       # Blue
        "chart_2":     (110, 110, 115),     # Gray
        "chart_3":     (52, 199, 89),       # Green
        "chart_4":     (255, 159, 10),      # Amber
        "kpi_bg":      (0, 113, 227),       # Blue
        "kpi_text":    (255, 255, 255),     # White
        "kpi_label":   (219, 234, 254),     # Light blue
    },
    "tab_colors": {
        "Dashboard": (0, 113, 227),         # Blue
        "Data":      (110, 110, 115),       # Gray
        "Charts":    (52, 199, 89),         # Green
        "Reference": (160, 160, 160),       # Slate
        "Audit":     (255, 59, 48),         # Red
    },
    "table_style": "banded",
    "kpi_style": "colored_bg",              # Blue background panel
    "cond_format": {
        "positive_bg":   (220, 252, 231),   # Light green
        "positive_text": (52, 199, 89),     # Green
        "negative_bg":   (254, 226, 226),   # Light red
        "negative_text": (255, 59, 48),     # Red
        "data_bar":      (0, 113, 227),     # Blue
    },
    "print": {"orientation": "landscape", "paper": "letter"},
    "design_notes": "Card-based. Systematic spacing. Blue accent underlines. KPI cards as hero. Subtle borders.",
}
```

---

## XSTYLE-03 — Corporate Report

```python
XSTYLE_03 = {
    "name": "Corporate Report",
    "fonts": {
        "title":     {"name": "Calibri", "size": 20, "bold": True},
        "header":    {"name": "Calibri", "size": 12, "bold": True},
        "body":      {"name": "Cambria", "size": 11, "bold": False},
        "kpi_value": {"name": "Calibri", "size": 32, "bold": True},
        "kpi_label": {"name": "Calibri", "size": 10, "bold": True},
        "caption":   {"name": "Calibri", "size": 9,  "bold": False},
    },
    "palette": {
        "header_bg":   (30, 58, 95),        # #1E3A5F dark blue
        "header_text": (255, 255, 255),     # White
        "accent":      (37, 99, 235),       # #2563EB royal blue
        "accent2":     (5, 150, 105),       # #059669 emerald
        "text":        (45, 55, 72),        # #2D3748
        "muted":       (113, 128, 150),     # #718096
        "alt_row":     (240, 244, 248),     # #F0F4F8
        "border":      (203, 213, 225),     # #CBD5E1
        "card_bg":     (255, 255, 255),     # White
        "positive":    (5, 150, 105),       # Emerald
        "negative":    (220, 38, 38),       # Red
        "chart_1":     (37, 99, 235),       # Royal blue
        "chart_2":     (5, 150, 105),       # Emerald
        "chart_3":     (30, 58, 95),        # Dark blue
        "chart_4":     (113, 128, 150),     # Gray
        "kpi_bg":      (30, 58, 95),        # Dark blue
        "kpi_text":    (255, 255, 255),     # White
        "kpi_label":   (147, 197, 253),     # Light blue
    },
    "tab_colors": {
        "Dashboard": (37, 99, 235),         # Royal blue
        "Data":      (30, 58, 95),          # Dark blue
        "Charts":    (5, 150, 105),         # Emerald
        "Reference": (113, 128, 150),       # Gray
        "Audit":     (220, 38, 38),         # Red
    },
    "table_style": "banded",
    "kpi_style": "full_box",                # Full border box, dark bg
    "cond_format": {
        "positive_bg":   (209, 250, 229),   # Light emerald
        "positive_text": (5, 150, 105),     # Emerald
        "negative_bg":   (254, 226, 226),   # Light red
        "negative_text": (220, 38, 38),     # Red
        "data_bar":      (37, 99, 235),     # Royal blue
    },
    "print": {"orientation": "portrait", "paper": "letter"},
    "design_notes": "Traditional. Clear hierarchy. Two accents (blue + green). Print-optimized. Repeating headers.",
}
```

---

## XSTYLE-04 — Data Science & Technical

```python
XSTYLE_04 = {
    "name": "Data Science & Technical",
    "fonts": {
        "title":     {"name": "Calibri", "size": 18, "bold": True},
        "header":    {"name": "Calibri", "size": 11, "bold": True},
        "body":      {"name": "Consolas", "size": 10, "bold": False},
        "kpi_value": {"name": "Consolas", "size": 28, "bold": True},
        "kpi_label": {"name": "Calibri", "size": 10, "bold": False},
        "caption":   {"name": "Calibri", "size": 9,  "bold": False},
    },
    "palette": {
        "header_bg":   (30, 30, 46),        # #1E1E2E dark
        "header_text": (255, 255, 255),     # White
        "accent":      (0, 188, 212),       # #00BCD4 cyan
        "accent2":     (224, 64, 251),      # #E040FB magenta
        "accent3":     (255, 235, 59),      # #FFEB3B yellow
        "text":        (44, 44, 44),        # #2C2C2C
        "muted":       (117, 117, 117),     # #757575
        "alt_row":     (245, 245, 245),     # #F5F5F5
        "border":      (224, 224, 224),     # #E0E0E0
        "card_bg":     (250, 250, 250),     # #FAFAFA
        "positive":    (0, 188, 212),       # Cyan
        "negative":    (224, 64, 251),      # Magenta
        "chart_1":     (0, 188, 212),       # Cyan
        "chart_2":     (224, 64, 251),      # Magenta
        "chart_3":     (255, 235, 59),      # Yellow
        "chart_4":     (158, 158, 158),     # Gray
        "kpi_bg":      (30, 30, 46),        # Dark
        "kpi_text":    (0, 188, 212),       # Cyan value
        "kpi_label":   (160, 160, 160),     # Muted label
    },
    "tab_colors": {
        "Dashboard": (0, 188, 212),         # Cyan
        "Data":      (30, 30, 46),          # Dark
        "Charts":    (224, 64, 251),        # Magenta
        "Reference": (158, 158, 158),       # Gray
        "Audit":     (255, 235, 59),        # Yellow
    },
    "table_style": "minimal",               # No banding, thin borders, grid visible
    "kpi_style": "accent_top",              # Cyan thin top border
    "cond_format": {
        "positive_bg":   (224, 247, 250),   # Light cyan
        "positive_text": (0, 151, 167),     # Dark cyan
        "negative_bg":   (252, 228, 236),   # Light magenta
        "negative_text": (173, 20, 87),     # Dark magenta
        "data_bar":      (0, 188, 212),     # Cyan
        "highlight":     (255, 235, 59),    # Yellow (for outliers)
    },
    "print": {"orientation": "landscape", "paper": "letter"},
    "design_notes": "Monospace for data. Dark headers. Grid visible in data tables. Minimal decoration. Precision-first.",
}
```

---

## XSTYLE-05 — Sales & Pipeline

```python
XSTYLE_05 = {
    "name": "Sales & Pipeline",
    "fonts": {
        "title":     {"name": "Calibri", "size": 20, "bold": True},
        "header":    {"name": "Calibri", "size": 12, "bold": True},
        "body":      {"name": "Calibri", "size": 11, "bold": False},
        "kpi_value": {"name": "Calibri", "size": 36, "bold": True},
        "kpi_label": {"name": "Calibri", "size": 10, "bold": True},
        "caption":   {"name": "Calibri", "size": 9,  "bold": False},
    },
    "palette": {
        "header_bg":   (15, 15, 26),        # #0F0F1A near-black
        "header_text": (255, 255, 255),     # White
        "accent":      (255, 107, 107),     # #FF6B6B coral
        "accent2":     (78, 205, 196),      # #4ECDC4 teal
        "accent3":     (255, 230, 109),     # #FFE66D yellow
        "text":        (26, 26, 46),        # #1A1A2E
        "muted":       (107, 114, 128),     # #6B7280
        "alt_row":     (250, 250, 250),     # #FAFAFA
        "border":      (229, 231, 235),     # #E5E7EB
        "card_bg":     (255, 255, 255),     # White
        "positive":    (78, 205, 196),      # Teal
        "negative":    (255, 107, 107),     # Coral
        "chart_1":     (255, 107, 107),     # Coral
        "chart_2":     (78, 205, 196),      # Teal
        "chart_3":     (255, 230, 109),     # Yellow
        "chart_4":     (153, 153, 187),     # Lavender
        "kpi_bg":      (15, 15, 26),        # Dark
        "kpi_text":    (255, 255, 255),     # White
        "kpi_label":   (255, 107, 107),     # Coral
    },
    "tab_colors": {
        "Dashboard": (255, 107, 107),       # Coral
        "Data":      (15, 15, 26),          # Dark
        "Charts":    (78, 205, 196),        # Teal
        "Pipeline":  (255, 230, 109),       # Yellow
        "Audit":     (220, 38, 38),         # Red
    },
    "table_style": "banded",
    "kpi_style": "colored_bg",              # Dark bg with colored values
    "cond_format": {
        "positive_bg":   (209, 250, 244),   # Light teal
        "positive_text": (78, 205, 196),    # Teal
        "negative_bg":   (255, 230, 230),   # Light coral
        "negative_text": (255, 107, 107),   # Coral
        "warning_bg":    (255, 249, 219),   # Light yellow
        "warning_text":  (161, 142, 7),     # Dark yellow
        "data_bar":      (255, 107, 107),   # Coral
    },
    "print": {"orientation": "landscape", "paper": "letter"},
    "design_notes": "Bold numbers. Color-coded status. Teal=good, coral=attention, yellow=caution. Dense but punchy.",
}
```

---

## XSTYLE-06 — Finance & Accounting

```python
XSTYLE_06 = {
    "name": "Finance & Accounting",
    "fonts": {
        "title":     {"name": "Calibri", "size": 18, "bold": True},
        "header":    {"name": "Calibri", "size": 11, "bold": True},
        "body":      {"name": "Calibri", "size": 11, "bold": False},
        "kpi_value": {"name": "Calibri", "size": 28, "bold": True},
        "kpi_label": {"name": "Calibri", "size": 10, "bold": True},
        "input":     {"name": "Calibri", "size": 11, "bold": False},  # Blue font
        "caption":   {"name": "Calibri", "size": 9,  "bold": False},
    },
    "palette": {
        "header_bg":   (26, 60, 42),        # #1A3C2A dark forest
        "header_text": (255, 255, 255),     # White
        "accent":      (45, 106, 79),       # #2D6A4F forest green
        "accent2":     (212, 168, 67),      # #D4A843 gold
        "text":        (45, 55, 72),        # #2D3748
        "muted":       (113, 128, 150),     # #718096
        "alt_row":     (232, 245, 233),     # #E8F5E9
        "border":      (200, 214, 192),     # #C8D6C0
        "card_bg":     (255, 255, 255),     # White
        "positive":    (45, 106, 79),       # Forest green
        "negative":    (220, 38, 38),       # Red
        "input_font":  (0, 102, 204),       # #0066CC blue (industry standard)
        "link_font":   (22, 163, 74),       # Green (cross-sheet links)
        "input_bg":    (219, 234, 254),     # Light blue input zone
        "locked_bg":   (229, 231, 235),     # Gray locked zone
        "assumption_bg": (254, 249, 195),   # Yellow key assumption
        "chart_1":     (45, 106, 79),       # Forest green
        "chart_2":     (212, 168, 67),      # Gold
        "chart_3":     (26, 60, 42),        # Dark green
        "chart_4":     (113, 128, 150),     # Gray
        "kpi_bg":      (45, 106, 79),       # Forest green
        "kpi_text":    (255, 255, 255),     # White
        "kpi_label":   (232, 245, 233),     # Light green
    },
    "tab_colors": {
        "Dashboard":     (45, 106, 79),     # Green
        "Data":          (26, 60, 42),       # Dark green
        "Assumptions":   (212, 168, 67),    # Gold
        "P&L":           (45, 106, 79),     # Green
        "Balance Sheet": (26, 60, 42),       # Dark green
        "Audit":         (220, 38, 38),     # Red
    },
    "table_style": "banded",
    "kpi_style": "accent_top",              # Green thick top border
    "cond_format": {
        "positive_bg":   (220, 252, 231),   # Light green
        "positive_text": (45, 106, 79),     # Forest green
        "negative_bg":   (254, 226, 226),   # Light red
        "negative_text": (220, 38, 38),     # Red
        "data_bar":      (45, 106, 79),     # Forest green
    },
    "financial_conventions": {
        "input_font_color":   (0, 102, 204),    # Blue = editable input
        "formula_font_color": (0, 0, 0),         # Black = formula
        "link_font_color":    (22, 163, 74),     # Green = cross-sheet link
        "external_font_color": (220, 38, 38),    # Red = external data
    },
    "print": {"orientation": "landscape", "paper": "letter"},
    "design_notes": "Industry conventions. Blue=input, black=formula, green=link. Yellow highlight for key assumptions. Print-optimized.",
}
```

---

## XSTYLE-07 — Marketing & Creative

```python
XSTYLE_07 = {
    "name": "Marketing & Creative",
    "fonts": {
        "title":     {"name": "Calibri", "size": 20, "bold": True},
        "header":    {"name": "Calibri", "size": 12, "bold": True},
        "body":      {"name": "Calibri", "size": 11, "bold": False},
        "kpi_value": {"name": "Calibri", "size": 32, "bold": True},
        "kpi_label": {"name": "Calibri", "size": 10, "bold": True},
        "caption":   {"name": "Calibri", "size": 9,  "bold": False},
    },
    "palette": {
        "header_bg":   (194, 112, 78),      # #C2704E terracotta
        "header_text": (255, 255, 255),     # White
        "accent":      (194, 112, 78),      # Terracotta
        "accent2":     (124, 154, 110),     # #7C9A6E sage
        "text":        (45, 27, 14),        # #2D1B0E dark brown
        "muted":       (139, 115, 85),      # #8B7355
        "alt_row":     (253, 248, 240),     # #FDF8F0 cream
        "border":      (232, 221, 208),     # #E8DDD0
        "card_bg":     (255, 251, 245),     # #FFFBF5 warm white
        "positive":    (124, 154, 110),     # Sage green
        "negative":    (194, 112, 78),      # Terracotta
        "chart_1":     (194, 112, 78),      # Terracotta
        "chart_2":     (124, 154, 110),     # Sage
        "chart_3":     (139, 115, 85),      # Brown
        "chart_4":     (217, 178, 147),     # Light terracotta
        "kpi_bg":      (194, 112, 78),      # Terracotta
        "kpi_text":    (255, 255, 255),     # White
        "kpi_label":   (253, 248, 240),     # Cream
    },
    "tab_colors": {
        "Dashboard": (194, 112, 78),        # Terracotta
        "Data":      (139, 115, 85),        # Brown
        "Charts":    (124, 154, 110),       # Sage
        "Campaigns": (253, 248, 240),       # Cream
        "Audit":     (220, 38, 38),         # Red
    },
    "table_style": "accent_top",            # Terracotta top border, cream banding
    "kpi_style": "colored_bg",              # Terracotta bg
    "cond_format": {
        "positive_bg":   (232, 245, 233),   # Light sage
        "positive_text": (124, 154, 110),   # Sage
        "negative_bg":   (253, 232, 222),   # Light terracotta
        "negative_text": (194, 112, 78),    # Terracotta
        "data_bar":      (194, 112, 78),    # Terracotta
    },
    "print": {"orientation": "landscape", "paper": "letter"},
    "design_notes": "Warm tones. Terracotta + sage dual accent. Cream backgrounds. No harsh blacks — use dark brown #2D1B0E.",
}
```

---

## XSTYLE-08 — Operations & Logistics

```python
XSTYLE_08 = {
    "name": "Operations & Logistics",
    "fonts": {
        "title":     {"name": "Calibri", "size": 20, "bold": True},
        "header":    {"name": "Calibri", "size": 12, "bold": True},
        "body":      {"name": "Calibri", "size": 11, "bold": False},
        "kpi_value": {"name": "Calibri", "size": 32, "bold": True},
        "kpi_label": {"name": "Calibri", "size": 10, "bold": True},
        "caption":   {"name": "Calibri", "size": 9,  "bold": False},
    },
    "palette": {
        "header_bg":   (45, 49, 66),        # #2D3142 charcoal
        "header_text": (255, 255, 255),     # White
        "accent":      (41, 98, 255),       # #2962FF cobalt
        "accent2":     (255, 143, 0),       # #FF8F00 amber
        "text":        (45, 49, 66),        # #2D3142
        "muted":       (158, 158, 158),     # #9E9E9E
        "alt_row":     (245, 245, 245),     # #F5F5F5
        "border":      (224, 224, 224),     # #E0E0E0
        "card_bg":     (255, 255, 255),     # White
        "positive":    (41, 98, 255),       # Cobalt (on-track)
        "negative":    (211, 47, 47),       # #D32F2F red (critical)
        "warning":     (255, 143, 0),       # Amber (delayed)
        "chart_1":     (41, 98, 255),       # Cobalt
        "chart_2":     (255, 143, 0),       # Amber
        "chart_3":     (45, 49, 66),        # Charcoal
        "chart_4":     (189, 189, 189),     # Light gray
        "kpi_bg":      (45, 49, 66),        # Charcoal
        "kpi_text":    (255, 255, 255),     # White
        "kpi_label":   (41, 98, 255),       # Cobalt
    },
    "tab_colors": {
        "Dashboard": (41, 98, 255),         # Cobalt
        "Data":      (45, 49, 66),          # Charcoal
        "Inventory": (255, 143, 0),         # Amber
        "Tracking":  (41, 98, 255),         # Cobalt
        "Audit":     (211, 47, 47),         # Red
    },
    "table_style": "banded",
    "kpi_style": "accent_top",              # Cobalt thick top border
    "cond_format": {
        "positive_bg":   (227, 242, 253),   # Light cobalt
        "positive_text": (41, 98, 255),     # Cobalt
        "negative_bg":   (255, 235, 238),   # Light red
        "negative_text": (211, 47, 47),     # Red
        "warning_bg":    (255, 243, 224),   # Light amber
        "warning_text":  (230, 126, 0),     # Dark amber
        "data_bar":      (41, 98, 255),     # Cobalt
    },
    "print": {"orientation": "landscape", "paper": "letter"},
    "design_notes": "Status-driven. Cobalt=on-track, amber=delayed, red=critical. Visible gridlines in data tables. Landscape print.",
}
```

---

## Style Application Workflow

When applying a style to a workbook, follow this order:

1. **Load style dict** from this file
2. **Set tab colors** on all sheets using `ws.sheet_properties.tabColor = rgb_to_hex(*color)`
3. **Apply title bar** — merged cells, dark bg, large bold white text
4. **Style all tables** — header row colors, banded rows, borders, totals row
5. **Build KPI panels** — style-appropriate bg, value size, label color, accent border
6. **Theme all charts** — series colors from `chart_1` through `chart_4`, clean styling
7. **Apply conditional formatting** — positive/negative colors from `cond_format` dict
8. **Set print layout** — orientation, margins, fit to page, repeat rows
9. **Follow design notes** — style-specific rules (e.g., XSTYLE-06 financial conventions)
10. **Run the mandatory audit** — style changes don't exempt from quality checks

## Custom Style Template

If none of the 8 styles fit, create a custom style dict following this structure:

```python
XSTYLE_CUSTOM = {
    "name": "Custom Style Name",
    "fonts": {
        "title":     {"name": "FontName", "size": 20, "bold": True},
        "header":    {"name": "FontName", "size": 12, "bold": True},
        "body":      {"name": "FontName", "size": 11, "bold": False},
        "kpi_value": {"name": "FontName", "size": 32, "bold": True},
        "kpi_label": {"name": "FontName", "size": 10, "bold": True},
        "caption":   {"name": "FontName", "size": 9,  "bold": False},
    },
    "palette": {
        # Core colors (required)
        "header_bg": (R, G, B), "header_text": (R, G, B),
        "accent": (R, G, B), "accent2": (R, G, B),
        "text": (R, G, B), "muted": (R, G, B),
        "alt_row": (R, G, B), "border": (R, G, B), "card_bg": (R, G, B),
        "positive": (R, G, B), "negative": (R, G, B),
        # Chart series (4 minimum)
        "chart_1": (R, G, B), "chart_2": (R, G, B),
        "chart_3": (R, G, B), "chart_4": (R, G, B),
        # KPI panel
        "kpi_bg": (R, G, B), "kpi_text": (R, G, B), "kpi_label": (R, G, B),
    },
    "tab_colors": {"Dashboard": (R,G,B), "Data": (R,G,B), ...},
    "table_style": "banded",        # banded | accent_top | minimal
    "kpi_style": "accent_top",      # accent_top | colored_bg | full_box
    "cond_format": {
        "positive_bg": (R,G,B), "positive_text": (R,G,B),
        "negative_bg": (R,G,B), "negative_text": (R,G,B),
        "data_bar": (R,G,B),
    },
    "print": {"orientation": "landscape", "paper": "letter"},
    "design_notes": "Brief guidance for style-specific decisions.",
}
```
