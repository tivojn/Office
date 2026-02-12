# Post-Generation Audit System for Word Documents

## Table of Contents

1. [Cascading Fix Problem](#cascading-fix-problem)
2. [Checks 1-10](#check-1-page-setup)
3. [Iterative Fix Loop](#iterative-fix-loop)
4. [Fix Strategies](#fix-strategies)
5. [False Positive Avoidance](#false-positive-avoidance)
6. [Output Format](#output-format)
7. [Key Lessons Learned](#key-lessons-learned)

---

## Cascading Fix Problem

Fixing one issue often creates another. This is the #1 reason audits fail:
- Widening a table column to fit text → may push total table width beyond page margins (CHECK 3)
- Increasing font size to fix minimum → changes paragraph count and page breaks (CHECK 5)
- Adding a callout box to fix style compliance → may create an orphaned heading (CHECK 5)
- Fixing image aspect ratio → changes image height, affecting page flow (CHECK 5)

**The iterative loop is NON-NEGOTIABLE. A single-pass audit is useless.**

---

## CHECK 1: PAGE SETUP
```
For every section in the document:
  - left_margin + right_margin < page_width (content width > 0)
  - top_margin + bottom_margin < page_height (content height > 0)
  - page_width is standard: 8.5" (Letter) or 8.27" (A4) (tolerance: 0.1")
  - page_height is standard: 11" (Letter) or 11.69" (A4) (tolerance: 0.1")
  - All margins >= 0.5" (minimum printable area)

STYLE-AWARE: If a style is active, verify margins match style spec.
```

## CHECK 2: FONT COMPLIANCE
```
For every run in every paragraph:
  - run.font.size >= Pt(9) for captions/footnotes
  - run.font.size >= Pt(10) for body text
  - run.font.name is explicitly set (not None / inherited from theme)
  - run.font.size is explicitly set (not None / inherited from theme)

STYLE-AWARE: If a style is active, also verify:
  - Heading paragraphs use the style's heading font name
  - Body paragraphs use the style's body font name
  - Font sizes match the style's hierarchy
  - FLAG WARNING if a run uses a font not in the active style
```

## CHECK 3: TABLE INTEGRITY
```
For every table in the document:
  - Sum of column widths <= available content width (page width - margins)
    Tolerance: 50000 EMU (~0.05")
  - All cells have content or are intentionally empty
  - Header row exists and is formatted distinctly from body rows
  - No cell has text that would require < 9pt to fit (min font check)
  - Cell margins/padding are set (not default 0)

STYLE-AWARE: If a style is active:
  - Header row colors match style's table_header_bg / table_header_tx
  - Alt-row shading matches style's table_alt_row
  - Table border style matches style spec (banded / accent_top / minimal / etc.)
```

## CHECK 4: HEADING HIERARCHY
```
Collect all heading paragraphs (style name starts with "Heading"):
  - Heading levels must not skip: H1 → H2 → H3 (no H1 → H3)
  - At least one H1 exists (document title or first major section)
  - No duplicate H1 text (usually indicates copy-paste error)
  - Heading font properties are consistent within each level

STYLE-AWARE: If a style is active:
  - H1 uses style heading font at style heading size
  - H2 uses style subheading font at style subheading size
  - Heading colors match style palette heading color
```

## CHECK 5: PAGE FLOW & BREAKS
```
For every page break point:
  - No orphaned headings: heading is not the last element on a page
    (heading must have at least 2 lines of content following it on same page)
  - No widow lines: single line of a paragraph at top of new page
  - Cover page ends with a page break or section break
  - TOC (if present) ends with a page break
  - Tables are not split across pages if they fit on one page
    (if table height < remaining page space, no break should precede it)

Detection approach: Estimate page breaks by accumulating paragraph heights.
  - Body paragraph height ≈ ceil(text_length / chars_per_line) × line_height
  - chars_per_line ≈ content_width / (font_size × 0.55)
  - line_height = font_size × line_spacing_factor
  - Table height = rows × estimated_row_height
  - Image height = as specified in add_picture()
```

## CHECK 6: IMAGE INTEGRITY
```
For every inline shape (image) in the document:
  - Image exists and is not a broken reference
  - Aspect ratio is preserved: abs(width/height - original_ratio) < 0.02
    (read original dimensions from image blob or stored metadata)
  - Image width <= content width (not wider than margins allow)
  - Image is not microscopic: width >= Inches(0.5) AND height >= Inches(0.5)
  - Image is not oversized: width <= page_width

FLAG CRITICAL if aspect ratio is distorted by > 2%
FLAG WARNING if image width is within 0.1" of content width (may touch margins)
```

## CHECK 7: COLOR & FILL INTEGRITY
```
Collect all unique colors from:
  - run.font.color.rgb (text colors)
  - paragraph shading (via lxml)
  - table cell shading
  - table border colors

For each color:
  - Not accidentally default blue (#4472C4) from theme
  - Not accidentally default black when style specifies a different dark color

STYLE-AWARE: If a style is active:
  - Every color must match a value in the style's palette dict
    Tolerance: ±15 per RGB channel
  - Exception: pure white (#FFFFFF) and pure black (#000000) always allowed
  - FLAG WARNING for off-palette colors
```

## CHECK 8: SPACING CONSISTENCY
```
For body paragraphs:
  - space_after is consistent (within ±2pt) across all body paragraphs
  - space_before on headings is consistent within each heading level
  - line_spacing is explicitly set (not inherited/default)
  - No empty paragraphs used for spacing (text is empty AND space_before/after are both 0)
    Exception: deliberate spacer paragraphs with space_before/after set

For callout boxes / tables:
  - Cell padding is consistent across similar elements
  - Callout box internal spacing matches style spec
```

## CHECK 9: STYLE COMPLIANCE (only when a design style is active)

```
Skip this check entirely if no design style was specified.

Load the active style dict from references/style-docx-mapping.md.

9a — PAGE SETUP:
  Margins match style spec (tolerance: 0.05")

9b — COVER PAGE:
  Cover page pattern matches style's cover_pattern
  Accent colors on cover match style's cover_accent
  Cover fonts match style's heading font

9c — FONT FAMILIES:
  Collect all unique font names across the document:
    Every font must appear in the active style's fonts dict values
    FLAG WARNING for each font not in the style

9d — COLOR PALETTE:
  Collect all unique colors from text, shading, borders:
    Each color must match a style palette value (tolerance: ±15 per RGB channel)
    FLAG WARNING for off-palette colors
    Exception: white, black always allowed

9e — TABLE STYLING:
  All tables follow the style's table_style pattern
  Header colors match style table_header_bg/tx

9f — DECORATIVE ELEMENTS:
  Accent rules use style's accent_rule color and width
  Callout boxes use style's callout_bg and callout_border
  Pull quotes (if any) use style's pullquote font

9g — STYLE-SPECIFIC RULES:
  STYLE-01 (Strategy): No shadows, no decorative images, max 2 colors per page
  STYLE-03 (Creative): Dashed rules, annotation-style labels present
  STYLE-04 (Kawaii): Rotating pastel backgrounds across sections
  STYLE-06 (Bold Narrative): Dark inverted callouts present on content pages
  STYLE-08 (Magazine): At least one heading >= 36pt, asymmetric margins
  STYLE-09 (Technical): Code blocks have warm-white bg, Note/Warning callout types
  STYLE-11 (Portfolio): 60%+ page area is images on gallery pages
  STYLE-12 (Retro): No pure black — must use dark navy #1B2838
```

## CHECK 10: DOCUMENT STRUCTURE
```
  - Document has at least one heading paragraph
  - If document has > 3 pages, a TOC placeholder should exist
  - Header/footer are set (not empty) on content pages
  - Section breaks are used appropriately (cover page separate from body)
  - No more than 3 consecutive body paragraphs without a visual break
    (callout, table, image, horizontal rule, or heading)
```

---

## Iterative Fix Loop

```python
MAX_PASSES = 5

for pass_num in range(1, MAX_PASSES + 1):
    issues = run_all_checks(doc)  # Checks 1-10 (9 only if style active)
    critical = [i for i in issues if i.severity == 'CRITICAL']

    if not critical:
        print(f"✅ Clean after {pass_num - 1} fix passes")
        break

    for issue in issues:
        apply_fix(issue)

    doc.save(path)
    doc = Document(path)  # Reload to get clean state

    print(f"Pass {pass_num}: fixed {len(issues)} issues, re-auditing...")
else:
    print(f"⚠️ {len(critical)} critical issues remain after {MAX_PASSES} passes")
```

---

## Fix Strategies

**FONT COMPLIANCE (CHECK 2):**
1. Set missing font.name to style's body font (or "Calibri" if no style)
2. Set missing font.size to style's body size (or Pt(11) if no style)
3. If font size < minimum, increase to minimum
4. **After font fix → re-run CHECK 5 (page flow may change)**

**TABLE OVERFLOW (CHECK 3):**
1. Recalculate column widths to fit content width
2. Proportionally reduce oversized columns
3. If text won't fit, reduce font size (min 9pt in cells)
4. **After width fix → re-run CHECK 3 to verify total width**

**HEADING HIERARCHY (CHECK 4):**
1. If H3 follows H1, insert H2 level heading or promote the H3
2. If no H1 exists, promote the first heading to H1
3. **After heading changes → re-run CHECK 5 (page flow)**

**PAGE FLOW (CHECK 5):**
1. Orphaned heading: add page break BEFORE the heading (push to next page)
2. Widow line: adjust paragraph's widow_control and orphan_control properties
3. Split table: add page break before table if it fits on next page
4. **After page break changes → re-run CHECK 10 (structure)**

**IMAGE DISTORTION (CHECK 6):**
1. Read original dimensions from image blob
2. Recalculate correct height from width (or vice versa) preserving ratio
3. Resize inline shape to correct dimensions
4. **After image resize → re-run CHECK 5 (page flow)**

**COLOR COMPLIANCE (CHECK 7 / CHECK 9d):**
1. Map off-palette colors to nearest palette color by Euclidean RGB distance
2. Replace in-place via run.font.color.rgb or lxml shading attributes
3. **After color fix → no cascading checks needed**

**SPACING (CHECK 8):**
1. Set consistent space_after on all body paragraphs
2. Remove empty spacer paragraphs and replace with proper space_before/after
3. **After spacing fix → re-run CHECK 5 (page flow)**

---

## False Positive Avoidance

1. **Empty paragraphs for spacing:** Only flag if BOTH space_before AND space_after are 0. Intentional spacers have non-zero spacing.
2. **Theme-inherited fonts:** Some runs inherit font from paragraph style. Only flag if the paragraph style itself has no font set.
3. **Cover page heading hierarchy:** Cover page title may not follow normal heading hierarchy — exclude from CHECK 4.
4. **Table cell colors:** Table cells may use palette colors not in the main text palette. Check against full palette including table-specific colors.
5. **Default colors from python-docx:** A run with `font.color.rgb = None` means "inherited from style" — this is OK if the style has the right color. Only flag if the style itself is missing color.

---

## Output Format

Per-section report:
```
[P#] [SEVERITY] [CHECK#] — Description → Fix applied / Remaining
```
(P# = page number estimate, or S# = section number)

Final summary:
```
🔴 CRITICAL: N (must be 0 before delivery)
🟡 WARNING: N (should be 0, acceptable if tight)
🔵 INFO: N (advisory)

STYLE: STYLE-XX (Name) or "Default (no style)"
STYLE COMPLIANCE: ✅ All checks passed / ⚠️ N issues
PASSES: X until clean
TOTAL FIXES: N applied
```

---

## Key Lessons Learned

1. **python-docx has NO layout engine** — it writes XML, Word renders it. You cannot know exact page breaks without Word. Estimate conservatively.

2. **Fixing one issue often creates another** — the iterative loop is essential. Font size changes affect page flow. Table width changes affect margins.

3. **Image distortion is the #1 visual bug** — always read source dimensions with Pillow before inserting. Never set both width AND height.

4. **Empty paragraphs are the #1 spacing anti-pattern** — use space_before/space_after instead. But don't false-positive on intentional spacers.

5. **Table column widths must sum correctly** — if they don't, Word auto-adjusts unpredictably. Always pre-calculate.

6. **Font inheritance is tricky** — a run with font.name=None inherits from paragraph style, which inherits from document default. Audit the full chain.

7. **Heading hierarchy matters for TOC** — skipped levels create broken TOC entries. Always check H1→H2→H3 sequence.

8. **Cover pages need section breaks** — a page break is not enough if you want different headers/footers on the cover.

9. **The audit must work generically** — discover structure dynamically, don't hardcode section names or paragraph counts.

10. **Color drift is subtle** — off-palette colors creep in when copy-pasting from other code. Always audit all colors against the active palette.
