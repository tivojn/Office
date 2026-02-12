# Design System Reference

## Table of Contents

1. [Typography](#typography)
2. [Color Palettes](#color-palettes)
3. [Page Layout Rules](#page-layout-rules)
4. [Visual Hierarchy](#visual-hierarchy)
5. [Decorative Elements](#decorative-elements)
6. [Document Structure Patterns](#document-structure-patterns)
7. [Cover Page Patterns](#cover-page-patterns)
8. [Table Design](#table-design)
9. [Icons & Bullets](#icons--bullets)
10. [Image Generation](#image-generation)
11. [Composition Planning](#composition-planning)
12. [Palette Template](#palette-template)

## Typography

- **Minimum body font size: 10pt.** Footnotes/captions: 9pt minimum. Table cells: 9pt minimum.
- Preferred body text: 11-12pt. Headings: H1=24-28pt, H2=18-22pt, H3=14-16pt.
- Always use a deliberate font pairing:
  - Georgia + Source Sans Pro (classic + modern)
  - Playfair Display + Lato (editorial elegance)
  - Montserrat + Open Sans (clean corporate)
  - Raleway + Merriweather (modern + readable)
  - Calibri + Cambria (safe cross-platform)
- Add letter spacing (tracking) to major titles: 1-3pt on display headings via lxml `w:spacing w:val`.
- Line spacing: body text at 1.15-1.5 (single is too tight, double wastes space). Headings at exact or 1.0.
- Paragraph spacing: body 6-12pt after; headings 18-24pt before, 6-12pt after.
- **Never use empty paragraphs for vertical spacing.** Use `space_before` / `space_after` on paragraph format.
- Never use default Calibri for both heading and body. Always differentiate heading vs body fonts.

## Color Palettes

Define a complete palette before building any document.

### Dark Premium (finance, executive reports, luxury)
- Page background: White `#FFFFFF` (documents are always printed — keep pages white/light)
- Heading color: `#0B1D3A` (deep navy)
- Body text: `#1A202C` (near-black)
- Muted text: `#64748B` (slate gray)
- Accent: `#C9A84C` (gold)
- Accent 2: `#0B1D3A` (navy for borders/rules)
- Table header BG: `#0B1D3A` | Table header text: `#FFFFFF`
- Table alt row: `#F1F5F9`
- Callout/sidebar BG: `#F8FAFC` | Callout border: `#C9A84C`
- Cover accent: `#C9A84C` (gold bar/border)
- Positive: `#16A34A` | Negative: `#DC2626`

### Light Clean (startups, tech, education)
- Heading color: `#1E293B`
- Body text: `#334155`
- Muted text: `#94A3B8`
- Accent: `#3B82F6` (blue)
- Accent 2: `#8B5CF6` (purple)
- Table header BG: `#3B82F6` | Table header text: `#FFFFFF`
- Table alt row: `#F0F9FF`
- Callout/sidebar BG: `#EFF6FF` | Callout border: `#3B82F6`
- Cover accent: `#3B82F6`

### Warm Earth (sustainability, nature, lifestyle)
- Heading color: `#2D1B0E`
- Body text: `#44382C`
- Muted text: `#8B7355`
- Accent: `#C2704E` (terracotta)
- Accent 2: `#7C9A6E` (sage green)
- Table header BG: `#C2704E` | Table header text: `#FFFFFF`
- Table alt row: `#FDF8F0`
- Callout/sidebar BG: `#FEF7ED` | Callout border: `#C2704E`
- Cover accent: `#C2704E`

### Bold Vibrant (creative, marketing, events)
- Heading color: `#0F0F1A`
- Body text: `#1A1A2E`
- Muted text: `#6B7280`
- Accent: `#FF6B6B` (coral red)
- Accent 2: `#4ECDC4` (teal)
- Accent 3: `#FFE66D` (yellow)
- Table header BG: `#0F0F1A` | Table header text: `#FFFFFF`
- Table alt row: `#FAFAFA`
- Callout/sidebar BG: `#FFF5F5` | Callout border: `#FF6B6B`
- Cover accent: `#FF6B6B`

### Corporate Blue (business reports, proposals, professional)
- Heading color: `#1E3A5F`
- Body text: `#2D3748`
- Muted text: `#718096`
- Accent: `#2563EB` (royal blue)
- Accent 2: `#059669` (emerald)
- Table header BG: `#1E3A5F` | Table header text: `#FFFFFF`
- Table alt row: `#F0F4F8`
- Callout/sidebar BG: `#EBF5FF` | Callout border: `#2563EB`
- Cover accent: `#2563EB`

## Page Layout Rules

### Page Dimensions
- **Letter**: 8.5" x 11" (21.59cm x 27.94cm) — US standard
- **A4**: 8.27" x 11.69" (21cm x 29.7cm) — international standard
- Default to Letter unless user specifies A4.

### Margins
- **Standard**: 1" all sides (professional, balanced)
- **Narrow**: 0.75" all sides (more content per page)
- **Wide**: 1.25" left/right, 1" top/bottom (elegant, plenty of whitespace)
- **Asymmetric**: 1.5" left, 1" right, 1" top/bottom (for binding)
- **Content width** = page width - left margin - right margin
  - Letter standard: 8.5" - 1" - 1" = **6.5"**
  - Letter narrow: 8.5" - 0.75" - 0.75" = **7.0"**
  - A4 standard: 8.27" - 1" - 1" = **6.27"**

### Spacing Conventions
- Gap between sections: 24-36pt before new heading
- Gap between body paragraphs: 6-12pt after
- Gap between heading and body: 6-12pt after heading
- Table cell padding: 0.05"-0.1" all sides
- Image spacing from text: 12-18pt before/after
- Callout box internal padding: 0.15"-0.25" all sides

### Pre-Calculation Rules
Before placing ANY element, calculate:
1. **Available content width** = page width - left margin - right margin
2. **Table column widths** must SUM to available content width (or less)
3. **Image width** should not exceed content width (scale proportionally)
4. **Two-column content**: each column = (content width - gutter) / 2, gutter = 0.25"-0.5"
5. **Sidebar layout**: sidebar = 30-35% of content width, main = 65-70%

## Visual Hierarchy

Layer order (most prominent to least):
1. **Headings** — largest font, bold, accent color, with decorative underline/border
2. **Subheadings** — smaller but still bold, differentiated from body
3. **Body text** — standard size, comfortable reading font
4. **Callout boxes / pull quotes** — stand out via background shading + border
5. **Tables** — structured data with clear header row
6. **Captions / footnotes** — smallest text, muted color
7. **Decorative accents** — horizontal rules, borders (support, don't compete)

## Decorative Elements

Every major section should have at least 2-3 of these:

- **Horizontal rule**: thin colored line (accent color) between sections. Use paragraph border-bottom or a thin table row.
- **Heading underline**: short colored bar under major headings (accent color, 1-2pt, width = 1.5"-2").
- **Callout box / card**: paragraph with background shading + left border (accent color, 2-4pt). Use `paragraph.paragraph_format.border_left` via lxml or a single-cell table with borders.
- **Pull quote**: large italic text in accent color, with decorative borders above and below. Stand out from body flow.
- **Sidebar panel**: narrow column (using a table) with accent background, containing supplementary info.
- **KPI card**: single-cell table with thick top border (accent), containing label + big number + note.
- **Numbered accent circles**: colored circles with white numbers for step-by-step content (via images or font tricks).
- **Page border**: subtle border on cover page or entire document (via lxml section properties).
- **Header/footer accent bar**: thin colored line in header or footer separating from content.
- **Table top accent**: colored top border on first row (accent color, 2-3pt).

### Callout Box Construction (python-docx)

The most reliable method is a **single-cell table** with custom borders and shading:

```python
# Callout as single-cell table
table = doc.add_table(rows=1, cols=1)
table.style = 'Table Grid'
cell = table.cell(0, 0)
# Add content paragraphs inside the cell
p = cell.paragraphs[0]
p.text = "Key Insight"
# Set cell shading and borders via lxml (see python-docx reference)
```

Alternative: paragraph-level shading + indent via lxml (less visual control but simpler).

### Pull Quote Construction

```python
# Pull quote = styled paragraph
pq = doc.add_paragraph()
pq.paragraph_format.left_indent = Inches(0.75)
pq.paragraph_format.right_indent = Inches(0.75)
pq.paragraph_format.space_before = Pt(18)
pq.paragraph_format.space_after = Pt(18)
run = pq.add_run('"The only way to do great work is to love what you do."')
run.font.size = Pt(16)
run.font.italic = True
run.font.color.rgb = RGBColor(0xC9, 0xA8, 0x4C)  # accent
```

## Document Structure Patterns

### 1. Simple Report
```
Cover Page (title, subtitle, date, author)
─────────────────────────────
Table of Contents
─────────────────────────────
Section 1: Introduction
  Body text
Section 2: Findings
  Body text + tables + callouts
Section 3: Conclusion
  Body text + recommendations
─────────────────────────────
Appendix (optional)
```

### 2. Executive Brief (1-2 pages)
```
Title + Date (top of page 1)
Executive Summary (callout box)
─── horizontal rule ───
Key Metrics (KPI cards in a row)
─── horizontal rule ───
Analysis (body text with pull quotes)
Recommendations (numbered list)
```

### 3. Proposal / Pitch
```
Cover Page (dramatic: full-bleed image + title overlay)
─────────────────────────────
Executive Summary
The Problem
Our Solution
  Features (icon + text grid)
Pricing (styled table)
Timeline (milestone table)
Call to Action
```

### 4. Technical Documentation
```
Cover Page (clean, logo, version)
─────────────────────────────
Table of Contents
─────────────────────────────
1. Overview
2. Architecture (with diagrams)
3. API Reference (styled code blocks)
4. Configuration (tables)
5. Troubleshooting (callout boxes)
─────────────────────────────
Changelog
```

### 5. Newsletter / Magazine Style
```
Banner header (full-width image + title)
─────────────────────────────
Two-column layout:
  Left: main article
  Right: sidebar (links, tips, highlights)
─────────────────────────────
Feature Story (full-width)
─────────────────────────────
Quick Bites (grid of small items)
```

### Layout Rhythm Across Sections

Monotonous layouts kill engagement. Vary visual elements:

- **Alternate** full-width and constrained content (e.g., full-width table, then indented body text)
- **Break up** long text runs with callout boxes, pull quotes, or images every 2-3 paragraphs
- **Use different decorative elements** per section (one section has a sidebar, next has KPI cards, next has a pull quote)
- **Rule: No more than 3 consecutive body-text-only paragraphs** without a visual break

## Cover Page Patterns

### 1. Classic Corporate
```
┌─────────────────────────────┐
│                              │
│    [Company Logo]            │
│                              │
│    ══════════════════        │  ← accent bar (gold/blue)
│    DOCUMENT TITLE            │  ← large, bold, heading font
│    Subtitle / Description    │  ← smaller, muted
│    ══════════════════        │  ← accent bar
│                              │
│                              │
│    Author Name               │
│    Date                      │
│    Version                   │
│                              │
└─────────────────────────────┘
```
- White background, centered text
- Accent bar = horizontal rule in palette accent color
- Build with: centered paragraphs + horizontal rule shapes + spacer paragraphs

### 2. Bold Banner
```
┌─────────────────────────────┐
│ ████████████████████████████ │  ← full-width accent block (cover image or solid color)
│ ████████████████████████████ │
│ ████████████████████████████ │
│ ████  DOCUMENT TITLE  ████ │  ← white text on dark bg
│ ████████████████████████████ │
│ ████████████████████████████ │
├─────────────────────────────┤
│                              │
│    Subtitle                  │
│    Author  ·  Date           │
│                              │
└─────────────────────────────┘
```
- Full-bleed image or dark solid block (via a full-width table with shading)
- White text overlaid (table cell with dark bg + white text)
- Clean info below

### 3. Minimal Elegant
```
┌─────────────────────────────┐
│                              │
│                              │
│                              │
│                              │
│         DOCUMENT             │  ← center-aligned, large
│         TITLE                │
│         ───────              │  ← short accent underline
│         Subtitle             │
│                              │
│                              │
│                              │
│    Author  ·  Date  ·  v1.0 │  ← bottom of page, muted
└─────────────────────────────┘
```
- Maximum whitespace, center-aligned
- Very little decoration — just the title and a thin accent line
- Author info at bottom margin

### 4. Side-Accent
```
┌────┬────────────────────────┐
│    │                         │
│ A  │    DOCUMENT TITLE       │  ← A = accent color bar
│ C  │    Subtitle             │     (full page height)
│ C  │                         │
│ E  │    ─────────────        │
│ N  │                         │
│ T  │    Author Name          │
│    │    Date                 │
│ B  │    Department           │
│ A  │                         │
│ R  │                         │
│    │                         │
└────┴────────────────────────┘
```
- Thin vertical accent bar on left side (via table column or page border)
- Text left-aligned with generous left margin
- Professional, understated

## Table Design

### Pre-Calculation

Before building ANY table:
1. **Count columns** and decide proportional widths
2. **Available width** = page width - left margin - right margin
3. **Sum of column widths MUST equal available width** (or table will auto-resize unpredictably)
4. **Minimum column width**: 0.75" for narrow data, 1.5" for text columns
5. **Header row height**: auto (let text dictate) but set minimum padding

### Table Style Patterns

#### Professional Banded
```
┌──────────┬──────────┬──────────┐
│ HEADER   │ HEADER   │ HEADER   │  ← dark bg + white text + bold
├──────────┼──────────┼──────────┤
│ Row 1    │ data     │ data     │  ← white bg
├──────────┼──────────┼──────────┤
│ Row 2    │ data     │ data     │  ← light gray/tinted bg
├──────────┼──────────┼──────────┤
│ Row 3    │ data     │ data     │  ← white bg
└──────────┴──────────┴──────────┘
```
- Header: palette `table_header_bg` + white text + bold
- Alternating rows: white and `table_alt_row`
- Thin borders: light gray (#E2E8F0)

#### Accent-Top Minimal
```
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━  ← thick accent-colored top border
  HEADER     HEADER     HEADER
──────────────────────────────  ← thin gray line
  Row 1      data       data
  Row 2      data       data
  Row 3      data       data
──────────────────────────────  ← thin gray bottom border
```
- No vertical borders, no cell borders
- Only horizontal: thick accent top, thin gray separator after header, thin gray bottom
- Clean, modern look

#### KPI Card Row (as table)
```
┌──────────┬──────────┬──────────┐
│  ━━━━━━  │  ━━━━━━  │  ━━━━━━  │  ← accent top bar per cell
│  METRIC  │  METRIC  │  METRIC  │  ← small label, accent color
│  $14.8B  │  23.4%   │  1,247   │  ← big number, dark text
│  +12.3%  │  ▲ 5.2%  │  ▼ 2.1%  │  ← note, green/red
└──────────┴──────────┴──────────┘
```
- Each cell is a KPI card: accent top border, label, big value, change indicator
- Build as 1-row table with 3+ cells, each cell formatted as a card

## Icons & Bullets

- **NEVER use emoji or Unicode symbols for icons.** They render inconsistently across platforms.
- **Preferred bullets**: standard bullet (•), en-dash (–), or custom numbered lists.
- **For step indicators**: use bold colored numbers: `1.` `2.` `3.` in accent color.
- **For status indicators**: use text labels in color — `[PASS]` in green, `[FAIL]` in red, `[PENDING]` in gray.
- **For icons in documents**: generate small AI images (32x32 to 64x64 px) or use inline shapes via lxml.

## Image Generation

Full AI image generation capability is a core design tool.

### When to Generate Images
- **Cover pages**: Hero background image or accent illustration
- **Section headers**: Atmospheric images representing each section's theme
- **Content illustrations**: Relevant visuals alongside text
- **Diagrams**: Conceptual illustrations, process flows, architecture diagrams
- **Background textures**: Subtle patterns for callout boxes or cover pages

### Image Generation Workflow
1. Compose a detailed prompt — specific about style, composition, color palette, mood, subject. Align with document's design language.
2. Generate the image via `generate_image` tool.
3. Save to known path — alongside .docx or in `./images/`. Descriptive filenames: `cover_hero.png`, `section2_illustration.png`.
4. Insert into document via `doc.add_picture()` with calculated dimensions.

### Image Prompt Best Practices
- **Always specify aspect ratio**: "wide 16:9" for cover banners, "square 1:1" for inline, "tall 2:3" for sidebar.
- **Always specify style**: "photorealistic", "watercolor", "minimal flat design", "line illustration".
- **Always match the palette**: Reference exact hex colors from the chosen palette.
- **Always specify composition**: "centered subject with clean edges for document layout".
- **Exclude text**: Always include "no text, no words, no letters, no watermarks".
- **Specify quality**: "ultra high quality, 8K, professional" for photorealistic.

### Image Sizing in Documents
- **Full-width inline**: width = content width (6.5" for Letter standard margins)
- **Half-width**: 3.0"-3.25" (for side-by-side with text)
- **Thumbnail**: 1.5"-2.0" wide
- **Cover banner**: full page width (8.5" for Letter, account for margins in positioning)
- **Always preserve aspect ratio** — calculate height from width proportionally.

### Image Integrity Rules (MANDATORY — prevents distortion)

**These rules are non-negotiable. A distorted image destroys the entire document's credibility.**

1. **NEVER specify both width and height** when inserting an image unless you've mathematically verified the aspect ratio matches the source. Specify width only — let python-docx auto-calculate height.
2. **Always read source dimensions first** with `PILImage.open(path).size` before any sizing decision.
3. **Use `safe_add_picture()`** (see python-docx reference) for every image insertion. This reads the source, calculates proportional dimensions, and inserts safely.
4. **Crop-to-fit, never stretch-to-fit.** When an image must fill a specific rectangular area:
   - Calculate source ratio (Sw/Sh) and target ratio (Tw/Th)
   - If ratios differ, center-crop the source to the target ratio using Pillow BEFORE inserting
   - Use `crop_to_aspect()` helper for this
   - Insert the cropped image at the target width (height auto-calculated)
5. **Common distortion traps to avoid:**
   - Cover banner: generated image is 1:1, target area is 16:9 → crop first
   - Table cell image: forcing width AND height to match cell → only set width
   - Logo: scaling to fit without reading original dimensions → always read first

### Rounded Corners (Visual Polish)

Rounded corners transform images from raw screenshots into designed elements. They add a modern, polished, editorial feel.

- **Default to rounding** all images unless they are full-bleed backgrounds, tiny icons, or already masked.
- **Radius scale:**
  - Small images (< 500px wide): 15-20px
  - Medium images (500-1500px): 25-35px
  - Large/cover images (> 1500px): 40-60px
- **Always save as PNG** after rounding to preserve corner transparency.
- **Use `round_corners()`** helper (see python-docx reference) or the all-in-one `prepare_image()`.
- **Rounded corners look best when** the image has generous spacing from surrounding text (12-18pt minimum).

### Image Strategy Decision Tree
```
Visual/creative document (marketing, portfolio)?
  -> YES -> Cover image + section illustrations + inline images
  -> NO -> Data-heavy (report, analysis)?
    -> YES -> Cover image only; use charts/tables for data
    -> NO -> Informational/educational?
      -> YES -> Cover + 2-3 relevant illustrations
      -> NO -> Minimal images, focus on typography
```

## Composition Planning

**Core Principle:** Document layout and content are ONE design. Plan the full structure — where images go, where text flows, how tables break across pages, where callouts create visual interest — before generating any content.

### The Composition-First Workflow

1. **Decide document structure first.** How many sections? What types of content (text, tables, images, callouts)?
2. **Choose a structure pattern** from the catalog above. This determines section order and visual elements.
3. **Choose a palette** that matches the document's purpose and tone.
4. **Plan visual rhythm.** Alternate between text-heavy and visual-heavy sections. Break up long text with callouts, pull quotes, or images every 2-3 paragraphs.
5. **Pre-calculate all dimensions.** Page width, content width, table column widths, image sizes — compute everything before writing code.
6. **Build incrementally.** One section per tool call. Verify page breaks and layout after each section.

### Page Break Strategy

- **Cover page**: always ends with a page break (section break preferred for different header/footer)
- **TOC**: always ends with a page break
- **Before major sections**: use `doc.add_page_break()` or section break for layout changes
- **Tables**: if a table won't fit on the current page, add a page break before it rather than splitting
- **Never split a callout box / KPI card row across pages** — add a page break before if needed

### Content Density Guidelines

| Document Type | Words per Page | Visual Elements per Page |
|--------------|---------------|------------------------|
| Executive Brief | 200-300 | 2-3 (KPI cards, callout, chart) |
| Detailed Report | 400-500 | 1-2 (table, callout, or image) |
| Technical Doc | 350-450 | 1-2 (code block, diagram, table) |
| Proposal | 250-350 | 2-3 (image, table, callout) |
| Newsletter | 300-400 | 3-4 (images, sidebars, quotes) |

## Palette Template

Copy and customize for each document:

```python
pal = {
    'heading':        '#0B1D3A',
    'body':           '#1A202C',
    'muted':          '#64748B',
    'accent':         '#C9A84C',
    'accent2':        '#0B1D3A',
    'table_header_bg':'#0B1D3A',
    'table_header_tx':'#FFFFFF',
    'table_alt_row':  '#F1F5F9',
    'callout_bg':     '#F8FAFC',
    'callout_border': '#C9A84C',
    'cover_accent':   '#C9A84C',
    'positive':       '#16A34A',
    'negative':       '#DC2626',
    'border':         '#E2E8F0',
}
```
