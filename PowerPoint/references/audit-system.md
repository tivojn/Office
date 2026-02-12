# Post-Generation Audit System

## Table of Contents

1. [Cascading Fix Problem](#cascading-fix-problem)
2. [Checks 1-11](#check-1-bounds)
3. [Word-Wrap Simulation](#word-wrap-simulation)
4. [Iterative Fix Loop](#iterative-fix-loop)
5. [Fix Strategies](#fix-strategies)
6. [Bullet List Layout Algorithm](#bullet-list-layout-algorithm)
7. [False Positive Avoidance](#false-positive-avoidance)
8. [Output Format](#output-format)
9. [Key Lessons Learned](#key-lessons-learned)

---

## Cascading Fix Problem

Fixing one issue often creates another. This is the #1 reason audits fail:
- Widening a text box to fix word-wrap → breaks container alignment (CHECK 4)
- Widening a container → may push it off-slide (CHECK 1)
- Resizing containers → may cause overlap with adjacent elements (CHECK 6)
- Fixing bullet text height → changes spacing for all items below (CHECK 5)

**The iterative loop is NON-NEGOTIABLE. A single-pass audit is useless.**

---

## CHECK 1: BOUNDS
```
For every shape on every slide:
  - shape.left >= 0
  - shape.top >= 0
  - shape.left + shape.width <= slide_width  (tolerance: 15000 EMU / ~0.016")
  - shape.top + shape.height <= slide_height  (tolerance: 15000 EMU / ~0.016")
```

## CHECK 2: TEXT CLIPPING (vertical overflow)
```
For every text frame:
  - Simulate word-wrap to get actual line count (see WRAP SIMULATION below)
  - Estimated height = num_lines × font_size_emu × 1.35
  - FLAG if estimated_height > text_frame.height × 1.1
```

## CHECK 3: WORD-WRAP QUALITY (horizontal — #1 cause of ugly slides)
```
For every text frame:
  - Find the longest single word in the text
  - Estimate its rendered width: len(word) × font_size_emu × char_width_factor × safety_margin
  - char_width_factor (EMU-based, multiply by font_size_emu):
      - 0.62 for bold slab-serif/bold (Rockwell Bold, Georgia Bold)
      - 0.57 for regular serif (Georgia, Times New Roman)
      - 0.58 for sans-serif (Arial, Calibri)
      - 0.60 for monospaced (Consolas) — fixed-width, wider average
      - 0.63 for heavy sans (Arial Black, Comic Sans MS Bold)
      - 0.59 for casual/hand (Comic Sans MS Regular)
      - 0.50 for condensed fonts
  - safety_margin: 1.12
  - FLAG CRITICAL if word_width > text_frame.width
  - FLAG WARNING if word_width > text_frame.width × 0.96 and len(word) > 5
```

**Common long-word offenders**: "Optimization", "Development", "Infrastructure", "Transformation", "Implementation", "Organization", "Acceleration", "Sustainability", "Self-optimizing", "cross-functional"

## CHECK 4: CONTAINER-TEXT SYNC ⚠️ CRITICAL FOR DIAGRAMS

The **#1 bug after applying fixes**: a text box gets resized but its parent container stays the same size, causing text to visually overflow its card/node.

```
For every parent-child pair (container shape + text box inside it):
  - Identify pairs: text_box center is inside container bounds
  - text_box.left >= container.left + padding (min 0.04")
  - text_box.top >= container.top + padding
  - text_box.left + text_box.width <= container.left + container.width - padding
  - text_box.top + text_box.height <= container.top + container.height - padding

Fix strategy (in order):
  1. Grow container to wrap text box + padding
  2. If container would go off-slide → shrink text box font (min 14pt)
  3. If font at minimum → re-center text box inside container, accept slight text reduction
```

## CHECK 5: BULLET/LIST ALIGNMENT ⚠️ CRITICAL FOR SIDE PANELS

Bullet lists implemented as dot-shape + text-box pairs have THREE common bugs:
1. **Dot not aligned with first line of text** — dot should center on first line's vertical center
2. **Text boxes oversized** — height set for 3 lines when text only wraps to 2, creating invisible overlap
3. **Spacing is "cramped"** — gaps between items are tiny because text box height includes unused space

```
Detection:
  - Find groups of small shapes (W < 0.15" and H < 0.15") near text boxes
  - Filter to only shapes within the SAME container/panel

For each bullet item (dot + text pair):
  1. Simulate wrap to get TRUE line count
  2. text_box.height = lines × font_size × 1.35  (EXACT, no padding)
  3. Stack items sequentially:
     item[n+1].top = item[n].top + item[n].height + gap
     gap = 0.14" to 0.20" (MUST be consistent)
  4. Dot positioning:
     dot.center_y = text.top + (font_size × 1.35 × 0.42)
     dot.left = consistent across all bullets (within 5000 EMU)
  5. Text box left = consistent across all bullets (within 5000 EMU)
  6. Resize parent container to fit: last_item.bottom + bottom_padding
```

## CHECK 6: OVERLAP CLASSIFICATION
```
For overlapping shape pairs, classify:
  - INTENTIONAL parent-child: text inside container shape → SKIP
  - INTENTIONAL layered UI: footer bar + footer text, accent bar → SKIP
  - INTENTIONAL image-bg stack: full-slide image + overlay + text → SKIP
  - UNINTENTIONAL: two independent text shapes colliding → FLAG

Detection: for each pair of shapes, check if bounding boxes overlap by >10%.
Skip pairs where one is clearly inside the other (parent-child).
Skip full-slide-sized shapes (image backgrounds) and their overlay shapes.
```

## CHECK 7: Z-ORDER
```
Verify no opaque fill shape is layered above a text shape it covers by >30% area.
```

## CHECK 8: FONT COMPLIANCE
```
  - All runs must have font.size >= Pt(14) (captions/labels may be Pt(10)+ in styled decks)
  - All runs must have explicit font.name set (theme defaults are unreliable)
  - All runs must have explicit font.size set

  STYLE-AWARE: If a style is active, also verify:
  - Title runs use the style's title font name
  - Body runs use the style's body font name
  - Font sizes match the style's hierarchy (load from style-pptx-mapping.md)
  - FLAG WARNING if a run uses a font name not in the active style's font dict
```

## CHECK 9: SPACING CONSISTENCY
```
  - Primary left margins should be consistent across slides (within ~100000 EMU)
  - Card gaps should be uniform where cards are in a row/column
  - Bullet item gaps should be uniform within each list
```

## CHECK 10: COLOR/FILL INTEGRITY
```
  - Verify all shape fills use the deck's intended palette
  - Check transparent overlays aren't accidentally 0% or 100% opacity

  STYLE-AWARE: If a style is active, also verify:
  - All shape fill colors exist in the active style's palette dict
  - FLAG WARNING for any RGBColor not in the style's palette (tolerance: ±10 per channel)
```

## CHECK 11: STYLE COMPLIANCE (only when a design style is active)

```
Skip this check entirely if no design style was specified.

Load the active style dict from references/style-pptx-mapping.md.

11a — BACKGROUND:
  For every slide:
    - slide.background.fill.fore_color.rgb must match style["slide_bg"]
    - Exception: title/section slides may use an alternate bg from the palette
    - FLAG CRITICAL if background is default white when style specifies a colored bg

11b — ACCENT ELEMENTS:
  If style defines accent_bar:
    - Verify accent bar shapes exist on content slides
    - accent_bar color matches style["accent_bar"]["color"]
    - accent_bar height ≈ style["accent_bar"]["height"] (tolerance: ±5000 EMU)
  If style has NO accent_bar defined:
    - No stray accent bars should be present

11c — FONT FAMILY CONSISTENCY:
  Collect all unique font.name values across the deck:
    - Every font name must appear in the active style's fonts dict values
    - FLAG WARNING for each font name NOT in the style

11d — COLOR PALETTE COHERENCE:
  Collect all unique RGBColor values from shape fills, font colors, and line colors:
    - Each color must match a value in the style's palette dict (tolerance: ±10 per RGB channel)
    - FLAG WARNING for off-palette colors
    - Exception: pure white (#FFFFFF) and pure black (#000000) are always allowed
    - Exception: semi-transparent overlays and shadows are excluded

11e — STYLE-SPECIFIC LAYOUT RULES:
  STYLE-09 (Storyboard): Verify panel grid exists
  STYLE-10 (Bento): Verify tile layout with uniform gaps and rounded corners
  STYLE-04 (Kawaii): Verify all shape corners are rounded
  STYLE-07 (Clay): Verify rounded corners on all container shapes
  STYLE-01 (Strategy): Verify no drop shadows or 3D effects
  STYLE-08 (Editorial): Verify at least one headline >= Pt(48) on non-title slides
  STYLE-12 (Retro): Verify no pure black (#000000) — must use dark navy (#1B2838)

11f — IMAGE BACKGROUND COMPLIANCE (only when images are used):
  For every slide with an image background:
    1. Image shape at (0,0) with size = slide dimensions
       FLAG CRITICAL if image doesn't cover full slide
    2. Overlay shape exists between image and text (z-order)
       FLAG CRITICAL if text sits directly on image with no overlay
    3. All text shapes within overlay bounds
       FLAG WARNING if text extends beyond overlay edges
    4. Text contrast against overlay color
       FLAG WARNING if low contrast detected
    5. Z-order: image (bottom) → overlay (middle) → text/shapes (top)
       FLAG CRITICAL if image or overlay is above any text shape
    6. Minimum font sizes on image backgrounds: Titles >= Pt(24), Body >= Pt(16)
       FLAG WARNING if below thresholds
```

### Fix Strategies for CHECK 11

```
11a fix: Set slide.background.fill.solid() and .fore_color.rgb = style["slide_bg"]
11b fix: Add/recolor accent bars per style spec; remove if style doesn't define them
11c fix: Replace non-style fonts with nearest style-equivalent
11d fix: Map off-palette colors to nearest palette color by Euclidean RGB distance
11e fix: Apply style-specific corrections (add rounded corners, remove shadows, etc.)
11f fix: Resize image to cover slide, add/fix overlay, fix z-order, bump font sizes

After any CHECK 11 fix → re-run CHECK 1 (bounds), CHECK 4 (container sync), CHECK 8 (font)
```

---

## Word-Wrap Simulation

python-pptx has **NO rendering engine**. Simulate PowerPoint's word-wrap to calculate line counts.

```python
def simulate_wrap(text, box_w_emu, font_size_pt, font='Georgia', bold=False):
    """Simulate PowerPoint word-wrap. Returns line count."""
    CHAR_WIDTHS = {
        ('Georgia', False): 6800,    ('Georgia', True): 7200,
        ('Rockwell', False): 7100,   ('Rockwell', True): 7500,
        ('Calibri', False): 6400,    ('Calibri', True): 6800,
        ('Consolas', False): 7000,   ('Consolas', True): 7000,
        ('Arial Black', False): 7600,('Arial Black', True): 7600,
        ('Comic Sans MS', False): 7200, ('Comic Sans MS', True): 7600,
    }
    avg_char_w = CHAR_WIDTHS.get((font, bold), 6800 if not bold else 7200)

    words = text.split()
    lines = 1
    current_w = 0
    space_w = avg_char_w * font_size_pt * 0.35
    usable_w = box_w_emu * 0.95

    for word in words:
        word_w = len(word) * avg_char_w * font_size_pt
        test_w = current_w + (space_w if current_w > 0 else 0) + word_w
        if current_w > 0 and test_w > usable_w:
            lines += 1
            current_w = word_w
        else:
            current_w = test_w
    return lines
```

---

## Iterative Fix Loop

```python
MAX_PASSES = 5

for pass_num in range(1, MAX_PASSES + 1):
    issues = run_all_checks(prs)  # Checks 1-11 (11 only if style active)
    critical = [i for i in issues if i.severity == 'CRITICAL']

    if not critical:
        print(f"✅ Clean after {pass_num - 1} fix passes")
        break

    for issue in issues:
        apply_fix(issue)

    prs.save(path)
    prs = Presentation(path)  # Reload to get clean state

    print(f"Pass {pass_num}: fixed {len(issues)} issues, re-auditing...")
else:
    print(f"⚠️ {len(critical)} critical issues remain after {MAX_PASSES} passes")
```

---

## Fix Strategies

**BAD-WORD-WRAP (CHECK 3):**
1. WIDEN text frame AND parent container (cascade to CHECK 4)
2. If would go off-slide → REDUCE font size (min 14pt)
3. If font at minimum → USE shorter synonym
4. **After any width change → re-run CHECK 4**

**TEXT-CLIP / VERTICAL OVERFLOW (CHECK 2):**
1. INCREASE text frame height AND parent container
2. If would push below slide bottom → REDUCE font size
3. If font at minimum → SPLIT content across shapes
4. **After any height change → re-run CHECK 5 if in a list**

**CONTAINER-TEXT DESYNC (CHECK 4):**
1. Grow container to wrap text box + padding (0.04" each side)
2. Re-center text box inside container
3. If container would go off-slide → shrink both proportionally
4. **After container resize → re-run CHECK 1 and CHECK 6**

**BULLET MISALIGNMENT (CHECK 5):**
1. Recalculate TRUE height using simulate_wrap()
2. Set text_box.height = exact needed height
3. Stack items sequentially with consistent gap (0.14"–0.20")
4. Align dots: dot.center_y = text.top + line_height × 0.42
5. Resize parent panel to fit
6. **After re-stacking → re-run CHECK 1, CHECK 4**

**BOUNDS OVERFLOW (CHECK 1):**
1. Reduce width/height to fit
2. Reposition shape inward
3. **After repositioning → re-run CHECK 4, CHECK 5, CHECK 6**

**OVERLAP — UNINTENTIONAL (CHECK 6):**
1. Move lower-priority shape to create gap
2. Reduce width of one shape
3. **After moving → re-run CHECK 1**

---

## Bullet List Layout Algorithm

```python
DOT_SIZE = Inches(0.055)
DOT_LEFT = panel.left + Inches(0.16)
TEXT_LEFT = DOT_LEFT + DOT_SIZE + Inches(0.10)
TEXT_W = (panel.left + panel.width) - TEXT_LEFT - Inches(0.12)
LINE_H = font_size_emu * 1.35
ITEM_GAP = Inches(0.16)

current_top = content_start_y

for each (dot, text_box, text_content):
    lines = simulate_wrap(text_content, TEXT_W, font_size_pt)
    text_h = lines * LINE_H

    text_box.left = TEXT_LEFT
    text_box.top = current_top
    text_box.width = TEXT_W
    text_box.height = text_h

    dot.left = DOT_LEFT
    dot.top = current_top + LINE_H * 0.42 - DOT_SIZE / 2
    dot.width = DOT_SIZE
    dot.height = DOT_SIZE

    current_top += text_h + ITEM_GAP

panel.height = current_top - ITEM_GAP + bottom_padding - panel.top
```

---

## False Positive Avoidance

1. **Bullet dot misalignment detecting wrong shapes**: Filter small shapes to same container/panel only.
2. **Title left margin inconsistency**: Exclude page numbers and footer text near slide edges.
3. **Intentional overlaps**: Accent bars, background rectangles, container-child pairs. Use center-point containment test.
4. **Tight word warnings at 95-96%**: Only flag CRITICAL at >100%, WARNING at >96%.

---

## Output Format

Per-slide report:
```
[S#] [SEVERITY] [CHECK#] — Description → Fix applied / Remaining
```

Final summary:
```
🔴 CRITICAL: N (must be 0 before delivery)
🟡 WARNING: N (should be 0, acceptable if tight-word at >96%)
🔵 INFO: N (advisory)

STYLE: STYLE-XX (Name) or "Default (no style)"
STYLE COMPLIANCE: ✅ All checks passed / ⚠️ N issues
PASSES: X until clean
TOTAL FIXES: N applied
```

---

## Key Lessons Learned

1. **python-pptx has NO rendering engine** — estimate using char-count × char-width-factor × font-size. Use `simulate_wrap()` for line counting and the more conservative char_width_factor × safety_margin for single-word overflow.

2. **Fixing one issue often creates another** — the iterative loop is essential; each fix must trigger re-checks on related checks.

3. **Text box height is the most common source of visual bugs** — oversized text boxes create invisible overlap. Always calculate TRUE height from simulated wrap.

4. **Bullet lists should NOT use python-pptx bullet properties** — use explicit dot shapes + text boxes for pixel-level control. Audit with CHECK 5.

5. **Always re-audit after applying fixes** — detect→fix→re-verify is the ONLY reliable approach.

6. **Container-text sync is the #1 missed bug** — when a text box is widened, the parent container MUST grow to match.

7. **False positives kill audit credibility** — filter bullet detection to same-container, exclude page numbers, use center-point containment.

8. **python-pptx auto-shapes have `has_text_frame=True` even when empty** — detect dots by SIZE (`< Inches(0.15)`) AND empty text, never by `has_text_frame`.

9. **The audit must be GENERIC, not hardcoded to shape names** — discover bullet panels dynamically.

10. **Accent bars on dark backgrounds are visual artifacts** — detect thin decorative bars on dark slides and remove/flag them.
