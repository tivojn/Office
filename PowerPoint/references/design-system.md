# Design System Reference

## Table of Contents

1. [Typography](#typography)
2. [Color Palettes](#color-palettes)
3. [Layout Rules](#layout-rules)
4. [Visual Hierarchy](#visual-hierarchy)
5. [Decorative Elements](#decorative-elements)
6. [Icons](#icons)
7. [Image Generation](#image-generation)
8. [Composition Planning](#composition-planning) — **Image + overlay as one cohesive design**
9. [Palette Template](#palette-template)

## Typography

- **Minimum font size: 14pt on any element.** No exceptions.
- Preferred body text: 16-18pt. Titles: 24-32pt.
- Always use a deliberate font pairing:
  - Georgia + Source Sans Pro
  - Playfair Display + Lato
  - Montserrat + Open Sans
  - Raleway + Merriweather
- Add letter spacing (tracking) to titles: `spc="200"` to `spc="400"` on `<a:rPr>`.
- Never use default Calibri for both heading and body.

## Color Palettes

Define a complete palette before building any deck.

### Dark Premium (finance, corporate)
- Background gradient: `#0B1D3A` -> `#162D50`
- Card/panel: `#1A3358` / border `#2A4A78`
- Text: `#FFFFFF` / muted `#8899AA`
- Accent: `#C9A84C` | Positive: `#4ADE80` | Negative: `#F87171`

### Light Clean (startups, tech, education)
- Background: `#F7F8FA` -> `#EDF0F5`
- Card: `#FFFFFF` / border `#E2E8F0`
- Text: `#1A202C` / muted `#718096`
- Accent: `#3B82F6` | Secondary: `#8B5CF6`

### Warm Earth (sustainability, nature, lifestyle)
- Background: `#FDF8F0` -> `#F5ECE0`
- Card: `#FFFFFF` / border `#E8DDD0`
- Text: `#2D1B0E` / muted `#8B7355`
- Accent: `#C2704E` | Secondary: `#7C9A6E`

### Bold Vibrant (creative, marketing, events)
- Background: `#0F0F1A` -> `#1A1A2E`
- Card: `#1A1A2E` / border `#2A2A4E`
- Text: `#FFFFFF` / muted `#9999BB`
- Accents: `#FF6B6B` / `#4ECDC4` / `#FFE66D`

### Tropical Dark (travel, vacation, island themes)
- Background gradient: `#0B1D3A` -> `#162D50`
- Card/panel: `#142B4D` / border `#1E3A63`
- Text: `#FFFFFF` / muted `#8CA0BC` / soft `#A8D8E0`
- Accents: `#E8A838` (gold) / `#4ECDC4` (teal) / `#E8653A` (coral) / `#F0A830` (amber)

## Layout Rules

- Slide dimensions: **12192000 x 6858000 EMU** (10" x 7.5")
- Minimum margin from edge: **457200 EMU** (0.5 inch)
- Gap between cards/panels: **150000-250000 EMU**
- Every shape MUST have explicit left, top, width, height.
- Prefer more slides with less content. Max 3-4 key points per slide.
- **Decorative elements need clear separation from content.** Never place decorative elements (slide numbers, accent icons) at similar positions to titles — they will visually crowd. Ensure no horizontal or vertical overlap. Place decorative counters in the opposite corner or a clearly distinct zone from the title.

## Visual Hierarchy

Layer order (back to front):

1. **Background image** (full bleed or sectioned)
2. **Overlay** (semi-transparent)
3. **Decorative accents** (bars, lines, shapes)
4. **Content cards/panels**
5. **Text and data**

## Decorative Elements

Every slide should have at least 2-3 of these:

- **Accent bar**: thin vertical/horizontal stripe in accent color
- **Title underline**: short horizontal bar under the title
- **Card panels**: rounded rectangles with subtle border. Use `adj=10000` for moderate rounded corners (default). Scale: 3000=barely visible, 10000=moderate, 16667=PowerPoint default, 50000=pill shape (too extreme for content cards)
- **Top-bar on cards**: thin accent line at top of each card
- **Transparency accents**: background shapes at 6-15% opacity (circles, rectangles)
- **Gradient backgrounds**: never flat solid color (unless an image background is used)
- **AI-generated images**: hero backgrounds, section illustrations, thematic visuals

## Icons

- **NEVER use emoji or Unicode symbols.**
- **Preferred**: Generate small AI images as custom icons for high-quality decks.
- **Fallback**: Geometric shapes (Ellipse, Diamond, Star5, etc.) or colored circles with single-character labels.
- Consistent size: 30-50pt shape-based, 300000-500000 EMU image-based.

## Image Generation

Full AI image generation capability is a core design tool.

### When to Generate Images
- **Title slides**: Cinematic hero background image
- **Section dividers**: Atmospheric images representing each section's theme
- **Content slides**: Relevant illustrations, visual metaphors alongside text
- **Icon replacements**: Small themed illustrations instead of geometric shapes
- **Background textures**: Subtle gradient textures, patterns, atmospheric backgrounds
- **Data visualization enhancement**: Contextual imagery to complement charts and tables

### Image Generation Workflow
1. Compose a detailed prompt — specific about style, composition, lighting, color palette, mood, subject. Align with deck's design language.
2. Generate the image via `generate_image` tool.
3. Save to known path — alongside .pptx or in `./images/`. Descriptive filenames: `slide1_hero_bg.png`, `maui_sunset_wide.png`.
4. Insert into slide via `slide.shapes.add_picture()` with precise positioning.

### Image Prompt Best Practices
- **Always specify aspect ratio**: "wide 16:9" for backgrounds, "square 1:1" for cards, "tall 2:3" for side panels.
- **Always specify style**: "photorealistic", "watercolor illustration", "minimal flat design", "cinematic photography", etc.
- **Always match the deck palette**: Reference exact hex colors.
- **Always specify composition**: "centered subject with negative space on the right for text overlay", etc.
- **Exclude text from images**: Always include "no text, no words, no letters, no watermarks".
- **Specify quality keywords**: "ultra high quality, 8K, professional photography" for photorealistic; "detailed illustration, clean lines" for illustrated styles.

### Image Prompt Templates

**Hero Background (Title Slide):**
```
Ultra wide 16:9 cinematic photograph of [subject/scene], [lighting], [mood], dominant color palette of [hex colors], professional photography, 8K resolution, shallow depth of field, no text no words no watermarks. Composition leaves [left/center/right] area with negative space for title text overlay.
```

**Section Background:**
```
Wide 16:9 [style] of [subject], soft [lighting], [mood] atmosphere, color palette harmonizing with [palette colors], subtle as slide background with text overlay, slightly blurred/desaturated for readability, no text no words no watermarks.
```

**Inline Content Image:**
```
[Aspect ratio] [style] of [specific subject], [composition details], vibrant and detailed, [mood], colors complementing [palette], clean background, professional quality, no text no words no watermarks.
```

**Texture/Pattern Background:**
```
Seamless abstract [texture type] pattern, [color palette], subtle and elegant, suitable as presentation background, soft gradient, no text no words no watermarks.
```

### Layering Strategy with Images

When using a generated image as background:
1. Add the image first (back of z-order).
2. **If image was composed with intentional negative space** (see [Composition Planning](#composition-planning)): skip overlay entirely, or add minimal 5-10% tint.
3. **If image is busy everywhere**: add a **targeted overlay** only where text goes (not full-bleed). Use gradient overlays that fade to transparent.
4. Then add content cards, text, and decorative elements on top.

Preferred order: **Image -> Targeted Overlay (if needed) -> Content Cards -> Text**

**Do NOT double up overlays.** When using semi-transparent content cards (e.g., rounded rectangles at 50-70% opacity), the cards already provide sufficient text contrast. Adding an additional heavy panel overlay (e.g., left-side 25-50% scrim) hides the background image unnecessarily. One layer of contrast is enough — either a panel overlay OR content cards, not both.

**IMPORTANT:** See [Composition Planning](#composition-planning) for the full system of coordinating image generation with overlay placement. The best slides need NO overlay because the image was generated with the right negative space from the start.

### Image Strategy Decision Tree

```
Visual/creative topic (travel, lifestyle, marketing, portfolio)?
  -> YES -> Generate images for EVERY slide (backgrounds + content)
  -> NO -> Data-heavy (finance, analytics, quarterly review)?
    -> YES -> Images for title + section dividers only; shapes/charts for data
    -> NO -> Educational or informational?
      -> YES -> Relevant illustrations for 50% of slides
      -> NO -> Minimal images, focus on typography and shapes
```

### Image Count Guidelines

| Deck Type | Title | Sections | Content | Total (10 slides) |
|-----------|-------|----------|---------|-------------------|
| Travel/Lifestyle | Hero BG | Full BG each | Inline images | 8-12 |
| Corporate/Finance | Subtle BG | Accent only | Shapes/charts | 2-4 |
| Marketing/Creative | Bold BG | Full BG each | BG + inline mix | 6-10 |
| Educational | Themed BG | Illustrations | Diagrams | 4-8 |
| Portfolio/Design | Full BG | Full BG | Full BG + inline | 10-15 |

### Image File Management
- Store in same directory as .pptx or `./images/` subfolder.
- Descriptive names: `slide01_title_bg.png`, `slide03_maui_hero.png`.
- After save, images are embedded in .pptx — source files can be cleaned up.
- Prefer PNG for quality. JPEG for photographic backgrounds where file size matters.

## Composition Planning

**Core Principle: Background image and overlay content are ONE design.** Never design them separately. Before generating any image or placing any text, plan the full composition — where the image focal point lives, where text goes, and how they complement each other.

### The Composition-First Workflow

1. **Decide content zones first.** What text, stats, or cards go on this slide? How much space do they need?
2. **Choose a layout pattern** from the catalog below. This determines where content zones and image focal zones sit.
3. **Generate the image WITH the layout in mind.** The image prompt must specify where negative space / dark areas / blurred regions should be — matching exactly where your content will go.
4. **Place content in the planned zones.** Because the image was designed for this layout, heavy overlays are unnecessary. Use minimal or no overlay.
5. **Verify the marriage.** Image focal points should be VISIBLE and UNBLOCKED. Text should sit in areas the image naturally leaves open.

### Layout Rhythm Across Slides

Monotonous layouts kill engagement. Vary layouts across slides using these rhythm patterns:

**Alternating (default for most decks):**
- Odd slides: content LEFT, image focal RIGHT
- Even slides: content RIGHT, image focal LEFT
- Creates a natural visual zigzag that guides the eye

**Progressive:**
- Start with full-bleed hero images (title/intro)
- Transition to split layouts (body content)
- End with centered/symmetrical layouts (conclusion/CTA)
- Mirrors the narrative arc of a presentation

**Grouped:**
- Same layout for related slides within a section
- Switch layout when section changes
- Creates visual chapters within the deck

**Rule: Never use the same layout for more than 3 consecutive slides.**

### Creative Layout Catalog

#### 1. Split Left-Right (Classic)
```
┌─────────┬─────────┐
│  TEXT    │  IMAGE  │
│  ZONE   │  FOCAL  │
│         │  POINT  │
└─────────┴─────────┘
```
- Content on left ~45%, image focal on right ~55% (or flipped)
- Image prompt: "subject on the right side, dark/blurred left area for text"
- Best for: stats, bullet points, player profiles

#### 2. Center Stage
```
┌───┬───────────┬───┐
│   │  IMAGE    │   │
│ T │  FOCAL    │ T │
│ E │  CENTER   │ E │
│ X │           │ X │
│ T │           │ T │
└───┴───────────┴───┘
```
- Image focal point centered, text in vertical side panels (left + right)
- Side panels: semi-transparent vertical strips or naturally dark image edges
- Image prompt: "centered subject, dark/faded edges on both sides"
- Best for: hero moments, single-subject focus, dramatic reveals

#### 3. Top-Bottom Split
```
┌─────────────────┐
│   IMAGE FOCAL   │
│   (top 60%)     │
├─────────────────┤
│   TEXT CONTENT   │
│   (bottom 40%)   │
└─────────────────┘
```
- Image dominates upper portion, content in lower band
- Lower band: dark gradient fade from image, or image has natural ground/dark bottom
- Image prompt: "subject in upper portion, dark ground/shadow at bottom for text"
- Best for: landscapes, architecture, cinematic establishing shots

#### 4. Bottom-Up (Inverted)
```
┌─────────────────┐
│   TEXT CONTENT   │
│   (top 35%)     │
├─────────────────┤
│   IMAGE FOCAL   │
│   (bottom 65%)   │
└─────────────────┘
```
- Content in upper band, image dominates below
- Image prompt: "subject in lower 2/3, dark/sky area at top for text"
- Best for: products rising up, growth metaphors, skyline shots

#### 5. Diagonal / Asymmetric
```
┌──────────────────┐
│ TEXT ╲            │
│       ╲  IMAGE   │
│        ╲ FOCAL   │
│         ╲        │
└──────────────────┘
```
- Content in upper-left triangle, image focal in lower-right (or vice versa)
- Use angled accent shapes or gradient masks to create the diagonal divide
- Image prompt: "subject in lower-right, dark upper-left corner fading diagonally"
- Best for: dynamic/energetic topics, sports, action, innovation

#### 6. Floating Card
```
┌─────────────────┐
│                  │
│   ┌──────┐      │
│   │ CARD │ IMG  │
│   │ TEXT │ ALL  │
│   └──────┘      │
│                  │
└─────────────────┘
```
- Full-bleed image with a single floating content card (semi-transparent, rounded)
- Card positioned where image has less visual activity
- Image prompt: "full scene, area of lower visual complexity on [left/center/right] for overlay card"
- Best for: quotes, key takeaways, single KPI highlights, minimal elegant slides

#### 7. Grid / Mosaic
```
┌────┬────┬────┬────┐
│IMG │TEXT│IMG │TEXT│
├────┼────┼────┼────┤
│TEXT│IMG │TEXT│IMG │
└────┴────┴────┴────┘
```
- Alternating image tiles and text tiles in a grid
- Each image tile is a cropped portion or separate generated image
- Best for: team pages, multi-product showcases, comparison slides

#### 8. Panoramic Strip
```
┌─────────────────────┐
│    TITLE TEXT        │
├─────────────────────┤
│  ▓▓ IMAGE STRIP ▓▓  │ (full-width, ~40% height)
├─────────────────────┤
│    BODY TEXT         │
└─────────────────────┘
```
- Cinematic horizontal image strip sandwiched between text zones
- Image prompt: "wide panoramic, 16:4 aspect, cinematic"
- Best for: timeline slides, journey narratives, before-after comparisons

#### 9. Bleed with Gradient Fade
```
┌─────────────────┐
│                  │
│  IMAGE COVERS   │
│  ENTIRE SLIDE   │
│                  │
│ ▒▒▒ gradient ▒▒▒│ <- gradient overlay fades to dark
│ TEXT ON DARK     │
└─────────────────┘
```
- Image covers full slide; gradient overlay fades from transparent to dark (any direction)
- Text sits in the darkened gradient zone
- Image prompt: "full scene, all areas visually interesting" (gradient does the work)
- Best for: immersive/cinematic feels, title slides, mood-setting

#### 10. L-Shape / Corner
```
┌──────┬──────────┐
│      │          │
│ TEXT │  IMAGE   │
│      │  FOCAL   │
├──────┤          │
│ TEXT │          │
└──────┴──────────┘
```
- Content wraps around one corner (L-shaped), image fills the rest
- Image prompt: "subject fills right and bottom, dark upper-left corner for text"
- Best for: data-heavy slides with a strong visual, dashboards with hero image

### Image-Overlay Coordination Rules

1. **Match negative space to content zones.** When generating an image, specify EXACTLY where you need dark/empty/blurred areas. These must align with where text will be placed.

2. **Dark-on-dark, light-on-light.** If the image has naturally dark areas, place WHITE text there — no overlay needed. If the image is bright everywhere, use a dark semi-transparent overlay ONLY where text goes.

3. **Overlay intensity scale:**
   - **None (0%):** Image has natural dark zones perfectly matching text placement. Ideal.
   - **Subtle (5-10%):** Image is mostly cooperative but needs slight contrast boost.
   - **Light (15-25%):** Image is busy; need a gentle scrim for readability.
   - **Medium (30-50%):** Image is very bright/complex; use targeted panel overlay, not full-bleed.
   - **Heavy (50%+):** AVOID. If you need this much overlay, the image composition is wrong. Regenerate the image with better negative space.

4. **Never overlay the focal point.** The whole point of the image is its subject. If your text blocks the subject, change the layout — don't darken the subject with an overlay.

5. **Targeted > full-bleed overlays.** A semi-transparent rectangle covering just the text zone is better than darkening the entire slide. It preserves more of the image.

6. **Gradient overlays > solid overlays.** A gradient that fades from 40% opacity to 0% looks more natural than a hard-edged solid rectangle.

### Theme Pairing: Image + Content Styling

| Image Mood | BG Tone | Text Color | Accent Color | Overlay Style |
|------------|---------|------------|-------------|---------------|
| Dark/moody (night, dramatic) | Dark | White/light | Gold, amber, neon | Minimal or none |
| Bright/airy (daylight, clean) | Light | Dark/charcoal | Blue, teal, coral | Dark gradient where text goes |
| Warm (sunset, earth, cozy) | Warm | Cream/white | Orange, terracotta | Warm-tinted semi-transparent |
| Cool (ocean, tech, modern) | Cool | White/ice blue | Cyan, electric blue | Cool-tinted gradient |
| Vibrant (neon, pop art, bold) | Dark | White/bright | Multi-accent neon | Minimal dark panels |

### Composition Prompt Engineering

When generating images for slides, ALWAYS include composition directives in the prompt:

**Template:**
```
[Subject/scene description], [style], [lighting/mood].
COMPOSITION: [focal point position] with [negative space location] for text overlay.
Color palette: [matching deck palette colors].
No text, no words, no letters, no watermarks.
```

**Examples by layout:**

Split left-right:
```
NBA basketball players in dynamic poses, cinematic photography, dramatic rim lighting.
COMPOSITION: Players grouped on the RIGHT side of the frame, LEFT 40% is dark empty court floor with dramatic shadows, suitable for text overlay.
Color palette: dark navy, gold accents, white highlights.
No text, no words, no letters, no watermarks.
```

Center stage:
```
A single golden trophy on a pedestal, dramatic spotlight from above, dark background.
COMPOSITION: Trophy CENTERED in frame, both LEFT and RIGHT edges fade to deep black, creating natural text zones on both sides.
No text, no words, no letters, no watermarks.
```

Diagonal:
```
Racing car speeding on a track, motion blur, dramatic angle.
COMPOSITION: Car in LOWER-RIGHT, motion trails sweeping diagonally, UPPER-LEFT fades to dark sky, creating diagonal text zone.
No text, no words, no letters, no watermarks.
```

### Iterative Composition Refinement (Practical Lessons)

AI image generators rarely nail spatial composition on the first try. Plan for 2-3 iterations:

**The iteration loop:**
1. Generate image with composition directives → place overlay → check result
2. If focal points are blocked by the overlay → adjust (see options below)
3. Repeat until image and overlay work together cleanly

**When the image doesn't match the layout:**

| Problem | Fix |
|---------|-----|
| Subject too close to text zone | **Regenerate with reference**: upload the current image and ask to push subjects further away from the text zone |
| Not enough negative space | **Exaggerate the percentages**: if you asked for "40% empty center", try "60% empty center". AI generators under-deliver on spatial instructions |
| Subject doesn't look right | **Regenerate with more identity cues**: specify jersey numbers, physical features, distinctive traits — not just names |
| Overlay covers too much image | **Shrink overlay width**: a 6" overlay on a 16" slide covers 37.5% — try 5" or narrower. Also check if overlay can be more transparent |

**Key practical rules:**

1. **Exaggerate spatial instructions by 1.5x.** If you need 40% empty space, ask for 60%. AI image generators compress toward center and under-deliver on edge placement.

2. **Use reference-based iteration.** When the style/mood is right but composition is wrong, upload the image as reference and ask for specific spatial adjustments. This preserves quality while fixing layout.

3. **Co-design overlay width and image composition.** Calculate what percentage of the slide your overlay covers BEFORE generating the image. The image's negative space must be at least that wide, plus margins.

4. **Identity cues matter for recognizable subjects.** Names alone aren't enough — specify physical features (shaved head, beard, headband), jersey numbers, team colors, and distinctive traits.

5. **Check the marriage at actual slide size.** A thumbnail can look fine but at full 16:9 the overlay might clip a subject's face by just a few pixels. Always verify at presentation scale.

## Palette Template

Copy and customize for each deck:

```python
pal = {
    'bg_start':     '#0B1D3A',
    'bg_end':       '#162D50',
    'accent':       '#E8A838',
    'accent2':      '#4ECDC4',
    'accent3':      '#E8653A',
    'card_fill':    '#142B4D',
    'card_border':  '#1E3A63',
    'text_primary': 'FFFFFF',
    'text_muted':   '8CA0BC',
    'text_soft':    'A8D8E0',
    'positive':     '4ADE80',
    'negative':     'F87171',
    'footer_bg':    '#0A1628',
}
```
