# Style → python-pptx Implementation Mapping

Translate each Design System style into concrete python-pptx values. Read `references/design_styles.md` for full style descriptions — this file provides the code-level implementation.

## Common Imports

```python
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
```

---

## STYLE-01 │ Strategy Consulting

```python
STYLE_01 = {
    "name": "Strategy Consulting",
    "slide_bg": RGBColor(0xFF, 0xFF, 0xFF),
    "fonts": {
        "title": {"name": "Georgia", "size": Pt(28), "bold": True, "color": RGBColor(0x1A, 0x1A, 0x1A)},
        "subtitle": {"name": "Calibri", "size": Pt(18), "bold": False, "color": RGBColor(0x66, 0x66, 0x66)},
        "body": {"name": "Calibri", "size": Pt(14), "bold": False, "color": RGBColor(0x1A, 0x1A, 0x1A)},
        "label": {"name": "Calibri", "size": Pt(12), "bold": False, "color": RGBColor(0x99, 0x99, 0x99)},
    },
    "palette": {
        "primary": RGBColor(0x00, 0x3D, 0xA5),   # Deep Royal Blue
        "text": RGBColor(0x1A, 0x1A, 0x1A),
        "secondary": RGBColor(0x66, 0x66, 0x66),
        "tertiary": RGBColor(0x99, 0x99, 0x99),
        "border": RGBColor(0xCC, 0xCC, 0xCC),
        "bg": RGBColor(0xFF, 0xFF, 0xFF),
    },
    "table": {
        "border_pt": 0.5,
        "header_fill": RGBColor(0x00, 0x3D, 0xA5),
        "header_text": RGBColor(0xFF, 0xFF, 0xFF),
        "alt_row_fill": RGBColor(0xF5, 0xF5, 0xF5),
    },
    "accent_bar": {"color": RGBColor(0x00, 0x3D, 0xA5), "height": Inches(0.04)},
    "design_notes": "No shadows, no 3D. Hairline borders. Monochrome icons only. Data-first.",
}
```

---

## STYLE-02 │ Executive Editorial

```python
STYLE_02 = {
    "name": "Executive Editorial",
    "slide_bg": RGBColor(0xFA, 0xF7, 0xF2),  # Warm cream
    "fonts": {
        "title": {"name": "Rockwell", "size": Pt(32), "bold": True, "color": RGBColor(0xC8, 0x10, 0x2E)},
        "subtitle": {"name": "Georgia", "size": Pt(20), "bold": False, "color": RGBColor(0x33, 0x33, 0x33)},
        "body": {"name": "Georgia", "size": Pt(14), "bold": False, "color": RGBColor(0x33, 0x33, 0x33)},
        "pullquote": {"name": "Georgia", "size": Pt(22), "bold": True, "italic": True, "color": RGBColor(0xC8, 0x10, 0x2E)},
    },
    "palette": {
        "primary": RGBColor(0xC8, 0x10, 0x2E),   # HBR Crimson
        "text": RGBColor(0x33, 0x33, 0x33),
        "secondary": RGBColor(0x8C, 0x82, 0x79),
        "tertiary": RGBColor(0xB8, 0xB0, 0xA8),
        "bg": RGBColor(0xFA, 0xF7, 0xF2),
        "bg_white": RGBColor(0xFF, 0xFF, 0xFF),
    },
    "accent_bar": {"color": RGBColor(0xC8, 0x10, 0x2E), "height": Inches(0.03)},
    "design_notes": "Generous white space. Crimson accents sparingly. Conceptual diagrams as focal points. Sidebar callouts for key takeaways.",
}
```

---

## STYLE-03 │ Sketch / Hand-Drawn

```python
STYLE_03 = {
    "name": "Sketch / Hand-Drawn",
    "slide_bg": RGBColor(0xF5, 0xF0, 0xE1),  # Kraft paper tone
    "fonts": {
        "title": {"name": "Comic Sans MS", "size": Pt(28), "bold": True, "color": RGBColor(0x3C, 0x3C, 0x3C)},
        # Fallback: Comic Sans MS is widely available; for hand-lettered feel
        "body": {"name": "Calibri", "size": Pt(14), "bold": False, "color": RGBColor(0x3C, 0x3C, 0x3C)},
        "annotation": {"name": "Calibri", "size": Pt(12), "bold": False, "italic": True, "color": RGBColor(0xA0, 0xA0, 0xA0)},
    },
    "palette": {
        "primary": RGBColor(0x3C, 0x3C, 0x3C),   # Dark pencil grey
        "ink": RGBColor(0x2B, 0x45, 0x70),         # Ink blue
        "accent": RGBColor(0xD4, 0x6A, 0x3C),      # Rust orange
        "secondary": RGBColor(0xA0, 0xA0, 0xA0),   # Light pencil grey
        "bg": RGBColor(0xF5, 0xF0, 0xE1),
    },
    "accent_bar": {"color": RGBColor(0xD4, 0x6A, 0x3C), "height": Inches(0.04)},
    "design_notes": "Loose layout. Slightly irregular shapes (use thin borders, no fill). No perfect geometry. Annotation-style labels.",
}
```

---

## STYLE-04 │ Kawaii / Cute Illustration

```python
STYLE_04 = {
    "name": "Kawaii / Cute",
    "slide_bg": RGBColor(0xFF, 0xD6, 0xE0),  # Baby pink
    "fonts": {
        "title": {"name": "Calibri", "size": Pt(30), "bold": True, "color": RGBColor(0x4A, 0x20, 0x40)},
        "body": {"name": "Calibri", "size": Pt(16), "bold": False, "color": RGBColor(0x4A, 0x20, 0x40)},
    },
    "palette": {
        "primary": RGBColor(0x4A, 0x20, 0x40),    # Deep plum
        "bg_pink": RGBColor(0xFF, 0xD6, 0xE0),
        "bg_lavender": RGBColor(0xE8, 0xD5, 0xF5),
        "bg_mint": RGBColor(0xD5, 0xF5, 0xE3),
        "accent_peach": RGBColor(0xFF, 0xAB, 0x91),
        "accent_sky": RGBColor(0x81, 0xD4, 0xFA),
        "accent_lemon": RGBColor(0xFF, 0xF5, 0x9D),
        "white": RGBColor(0xFF, 0xFF, 0xFF),
    },
    "shapes": {"corner_radius": Inches(0.2)},  # All corners rounded
    "design_notes": "Generous padding. Soft colors. Rounded everything. Alternate pastel bg colors across slides.",
}
```

---

## STYLE-05 │ Professional / Corporate Modern

```python
STYLE_05 = {
    "name": "Professional / Corporate Modern",
    "slide_bg": RGBColor(0xFF, 0xFF, 0xFF),
    "fonts": {
        "title": {"name": "Calibri", "size": Pt(36), "bold": True, "color": RGBColor(0x1D, 0x1D, 0x1F)},
        "subtitle": {"name": "Calibri", "size": Pt(18), "bold": False, "color": RGBColor(0x6E, 0x6E, 0x73)},
        "body": {"name": "Calibri", "size": Pt(14), "bold": False, "color": RGBColor(0x1D, 0x1D, 0x1F)},
        "metric": {"name": "Calibri", "size": Pt(48), "bold": True, "color": RGBColor(0x00, 0x71, 0xE3)},
    },
    "palette": {
        "primary": RGBColor(0x00, 0x71, 0xE3),    # Electric Blue
        "text": RGBColor(0x1D, 0x1D, 0x1F),
        "secondary": RGBColor(0x6E, 0x6E, 0x73),
        "border": RGBColor(0xD1, 0xD1, 0xD6),
        "bg": RGBColor(0xFF, 0xFF, 0xFF),
        "bg_alt": RGBColor(0xF8, 0xF9, 0xFA),
    },
    "card": {"corner_radius": Inches(0.06), "shadow": True, "border_color": RGBColor(0xD1, 0xD1, 0xD6)},
    "accent_bar": {"color": RGBColor(0x00, 0x71, 0xE3), "height": Inches(0.04)},
    "design_notes": "12-column grid. Card-based sections. Hero metrics large at top. Subtle shadows. Systematic spacing.",
}
```

---

## STYLE-06 │ Anime / Manga Illustration

```python
STYLE_06 = {
    "name": "Anime / Manga",
    "slide_bg": RGBColor(0x1A, 0x1A, 0x40),  # Deep twilight
    "fonts": {
        "title": {"name": "Arial Black", "size": Pt(40), "bold": True, "color": RGBColor(0xFF, 0xF8, 0xE1)},
        "body": {"name": "Calibri", "size": Pt(16), "bold": False, "color": RGBColor(0xFF, 0xFF, 0xFF)},
    },
    "palette": {
        "bg_dark": RGBColor(0x1A, 0x1A, 0x40),
        "cerulean": RGBColor(0x1E, 0x88, 0xE5),
        "rose": RGBColor(0xE9, 0x1E, 0x63),
        "amber": RGBColor(0xFF, 0xB3, 0x00),
        "highlight": RGBColor(0xFF, 0xF8, 0xE1),
        "shadow": RGBColor(0x1A, 0x1A, 0x2E),
    },
    "design_notes": "Wide cinematic framing. Rich atmospheric colors. Large dramatic titles. Minimal supporting text. Dark backgrounds.",
}
```

---

## STYLE-07 │ 3D Clay / Claymation

```python
STYLE_07 = {
    "name": "3D Clay / Claymation",
    "slide_bg": RGBColor(0xFF, 0xEC, 0xD2),  # Pale peach
    "fonts": {
        "title": {"name": "Calibri", "size": Pt(32), "bold": True, "color": RGBColor(0xE0, 0x7A, 0x5F)},
        "body": {"name": "Calibri", "size": Pt(16), "bold": False, "color": RGBColor(0x4A, 0x3A, 0x2E)},
    },
    "palette": {
        "bg_peach": RGBColor(0xFF, 0xEC, 0xD2),
        "bg_lavender": RGBColor(0xE8, 0xDE, 0xF8),
        "bg_sky": RGBColor(0xDC, 0xEE, 0xFB),
        "terracotta": RGBColor(0xE0, 0x7A, 0x5F),
        "ocean": RGBColor(0x3D, 0x85, 0xC6),
        "leaf": RGBColor(0x6B, 0xAF, 0x6B),
        "butter": RGBColor(0xFF, 0xE0, 0x82),
        "coral": RGBColor(0xFF, 0x8A, 0x65),
    },
    "shapes": {"corner_radius": Inches(0.15)},  # Chunky rounded corners
    "design_notes": "Generous negative space. Centered subjects. Warm, soft tones. Saturated but not neon. Rounded corners everywhere.",
}
```

---

## STYLE-08 │ Editorial / Magazine Spread

```python
STYLE_08 = {
    "name": "Editorial / Magazine Spread",
    "slide_bg": RGBColor(0xFF, 0xFF, 0xFF),
    "fonts": {
        "title": {"name": "Georgia", "size": Pt(48), "bold": True, "color": RGBColor(0x0A, 0x0A, 0x0A)},
        "body": {"name": "Georgia", "size": Pt(14), "bold": False, "color": RGBColor(0x0A, 0x0A, 0x0A)},
        "pullquote": {"name": "Georgia", "size": Pt(28), "bold": True, "italic": True, "color": RGBColor(0xFF, 0xD6, 0x00)},
        "caption": {"name": "Calibri", "size": Pt(10), "bold": False, "color": RGBColor(0x75, 0x75, 0x75)},
    },
    "palette": {
        "bg_white": RGBColor(0xFF, 0xFF, 0xFF),
        "bg_black": RGBColor(0x0A, 0x0A, 0x0A),
        "text_dark": RGBColor(0x0A, 0x0A, 0x0A),
        "text_light": RGBColor(0xFF, 0xFF, 0xFF),
        "accent_yellow": RGBColor(0xFF, 0xD6, 0x00),
        "accent_magenta": RGBColor(0xE9, 0x1E, 0x63),
        "accent_cobalt": RGBColor(0x1A, 0x23, 0x7E),
        "caption": RGBColor(0x75, 0x75, 0x75),
    },
    "design_notes": "Asymmetric grid. Massive headlines (48pt+). High contrast. Bold statement accent per slide. Thick divider rules.",
}
```

---

## STYLE-09 │ Storyboard / Sequential

```python
STYLE_09 = {
    "name": "Storyboard / Sequential",
    "slide_bg": RGBColor(0xFF, 0xFF, 0xFF),
    "fonts": {
        "title": {"name": "Calibri", "size": Pt(24), "bold": True, "color": RGBColor(0x2C, 0x2C, 0x2C)},
        "annotation": {"name": "Consolas", "size": Pt(11), "bold": False, "color": RGBColor(0x2C, 0x2C, 0x2C)},
        "step_number": {"name": "Calibri", "size": Pt(16), "bold": True, "color": RGBColor(0xE5, 0x39, 0x35)},
    },
    "palette": {
        "bg": RGBColor(0xFF, 0xFF, 0xFF),
        "bg_warm": RGBColor(0xF0, 0xED, 0xE8),
        "panel_border": RGBColor(0x2C, 0x2C, 0x2C),
        "panel_fill": RGBColor(0xF8, 0xF8, 0xF8),
        "accent_red": RGBColor(0xE5, 0x39, 0x35),
        "accent_yellow": RGBColor(0xFF, 0xEB, 0x3B),
    },
    "panel": {"border_pt": 1.5, "border_color": RGBColor(0x2C, 0x2C, 0x2C), "corner_radius": Inches(0.03)},
    "design_notes": "2x3 or 3x3 panel grid. Numbered panels. Annotation strip below each. Greyscale interiors. Single highlight color for key actions.",
}
```

---

## STYLE-10 │ Bento Grid

```python
STYLE_10 = {
    "name": "Bento Grid",
    "slide_bg": RGBColor(0x1C, 0x1C, 0x1E),  # Dark mode default
    "fonts": {
        "title": {"name": "Calibri", "size": Pt(20), "bold": True, "color": RGBColor(0xFF, 0xFF, 0xFF)},
        "supporting": {"name": "Calibri", "size": Pt(12), "bold": False, "color": RGBColor(0xA0, 0xA0, 0xA0)},
        "metric": {"name": "Calibri", "size": Pt(40), "bold": True, "color": RGBColor(0xFF, 0xFF, 0xFF)},
    },
    "palette": {
        "bg_dark": RGBColor(0x1C, 0x1C, 0x1E),
        "bg_light": RGBColor(0xF5, 0xF5, 0xF7),
        "tile_dark": RGBColor(0x2C, 0x2C, 0x2E),
        "tile_accent_blue": RGBColor(0x00, 0x71, 0xE3),
        "tile_accent_green": RGBColor(0x34, 0xC7, 0x59),
        "tile_accent_purple": RGBColor(0xAF, 0x52, 0xDE),
        "tile_accent_orange": RGBColor(0xFF, 0x9F, 0x0A),
        "text_light": RGBColor(0xFF, 0xFF, 0xFF),
        "text_dark": RGBColor(0x1C, 0x1C, 0x1E),
    },
    "tile": {"corner_radius": Inches(0.12), "gap": Inches(0.1)},
    "design_notes": "Mixed tile sizes (1x1, 1x2, 2x2). One idea per tile. Dark mode preferred. Each tile can have its own accent color. Rounded corners 12-16px.",
}
```

---

## STYLE-11 │ Bricks / Masonry

```python
STYLE_11 = {
    "name": "Bricks / Masonry",
    "slide_bg": RGBColor(0xEA, 0xEA, 0xEA),  # Light grey
    "fonts": {
        "caption": {"name": "Calibri", "size": Pt(11), "bold": False, "color": RGBColor(0x75, 0x75, 0x75)},
        "tag": {"name": "Calibri", "size": Pt(9), "bold": True, "color": RGBColor(0x75, 0x75, 0x75)},
        "title": {"name": "Calibri", "size": Pt(16), "bold": True, "color": RGBColor(0x1A, 0x1A, 0x1A)},
    },
    "palette": {
        "bg_light": RGBColor(0xEA, 0xEA, 0xEA),
        "bg_warm": RGBColor(0xFA, 0xF9, 0xF6),
        "bg_dark": RGBColor(0x1A, 0x1A, 0x1A),
        "card_fill": RGBColor(0xFF, 0xFF, 0xFF),
        "caption": RGBColor(0x75, 0x75, 0x75),
        "accent": RGBColor(0xE6, 0x00, 0x23),  # Pinterest red
    },
    "card": {"corner_radius": Inches(0.06), "gap": Inches(0.08)},
    "design_notes": "3-5 columns. Variable card heights. Content-driven card colors. Minimal borders. Full-bleed images in cards. High visual density.",
}
```

---

## STYLE-12 │ Retro / Risograph

```python
STYLE_12 = {
    "name": "Retro / Risograph",
    "slide_bg": RGBColor(0xF2, 0xE8, 0xD5),  # Uncoated paper
    "fonts": {
        "title": {"name": "Rockwell", "size": Pt(36), "bold": True, "color": RGBColor(0x1B, 0x28, 0x38)},
        "body": {"name": "Consolas", "size": Pt(14), "bold": False, "color": RGBColor(0x1B, 0x28, 0x38)},
    },
    "palette": {
        "bg_paper": RGBColor(0xF2, 0xE8, 0xD5),
        "riso_blue": RGBColor(0x00, 0x78, 0xBF),
        "riso_red": RGBColor(0xFF, 0x66, 0x5E),
        "riso_green": RGBColor(0x00, 0xA9, 0x5C),
        "dark_navy": RGBColor(0x1B, 0x28, 0x38),  # No pure black
    },
    "design_notes": "2-3 color max. No pure black — use dark navy. Bold flat shapes. Overlap elements for overprint effect. Paper texture tone bg. Typewriter font for body.",
}
```

---

## Style Application Workflow

When applying a style, use this order:

1. **Set slide background** from `slide_bg`
2. **Apply font settings** for each text level (title, body, etc.)
3. **Apply palette colors** to shapes, fills, borders
4. **Apply style-specific elements** (accent bars, card styles, tile layouts)
5. **Follow design notes** for style-specific layout decisions
6. **Run the mandatory audit** — style changes don't exempt from quality checks
