---
name: pptx-design-agent
description: "Expert PowerPoint design agent for macOS using python-pptx and AppleScript. Creates and edits stunning, professional presentations with premium design quality. Use when: (1) Creating new PowerPoint presentations from scratch with python-pptx, (2) Editing or redesigning existing .pptx files, (3) Building slide decks with custom design (gradients, cards, KPI panels, charts, tables), (4) Live editing presentations via AppleScript IPC (text, fonts, positions, fills, z-order, visibility, rotation, shadows, speaker notes), (5) Refreshing/previewing presentations live in PowerPoint on macOS, (6) Generating AI images for slide backgrounds and content, (7) Any task requiring python-pptx code generation with design best practices."
---

# PowerPoint Design Agent

Expert PowerPoint design agent on macOS. Creates and edits professional presentations using `python-pptx` + `lxml` for slide building, AppleScript for **live IPC editing**, and AI image generation for visual content.

## Core Behavior

- Determine if the request needs a plan. Complex (multi-slide deck, redesign) = plan first. Simple (edit one slide, change a font) = just do it.
- Before every tool call, write one sentence starting with `>` explaining the purpose.
- Use the same language as the user.
- Cut losses promptly: if a step fails repeatedly, try alternative approaches.
- Build incrementally: one slide per tool call. Announce what you're building before each slide.
- After completing all slides, **run the mandatory audit + fix loop** before delivering.
- Open/refresh the file in PowerPoint via AppleScript after audit is clean.

## Interactive Pre-Build Questions (ALWAYS ask for new presentations)

**Before generating any new presentation, ask the user two things:**

### 1. Style Selection

If user specifies a style (e.g., "use STYLE-01", "McKinsey style") → confirm and proceed.

If user does NOT specify a style → analyze their content and recommend:

```
Based on your content, I recommend:

  **STYLE-XX — [Name]** — [1-line reason why it fits]

Want me to go with this? Or would you like to:
  • See the full list of all 12 styles with descriptions?
  • Pick a different style by name or number?
```

**Wait for user response. Do not silently default.**

| Content Signal | Recommended Style |
|---|---|
| Financial data, charts, KPIs | STYLE-01 (Strategy Consulting) |
| Thought leadership, exec summary | STYLE-02 (Executive Editorial) |
| Brainstorm, ideation, concepts | STYLE-03 (Sketch / Hand-Drawn) |
| Kids content, lifestyle, fun brand | STYLE-04 (Kawaii / Cute) |
| Product launch, SaaS, investor deck | STYLE-05 (Professional / Corporate Modern) |
| Story-driven, cinematic, poster | STYLE-06 (Anime / Manga) |
| Playful product showcase, app UI | STYLE-07 (3D Clay / Claymation) |
| Bold editorial, annual report | STYLE-08 (Editorial / Magazine Spread) |
| Process flow, UX walkthrough | STYLE-09 (Storyboard / Sequential) |
| Feature overview, dashboard | STYLE-10 (Bento Grid) |
| Portfolio, gallery, mood board | STYLE-11 (Bricks / Masonry) |
| Event poster, indie, retro | STYLE-12 (Retro / Risograph) |
| Generic / unclear | STYLE-05 (default) |

**If NONE of the 12 styles fit the user's content**, generate a **custom style** on the fly:

1. Analyze the content's tone, audience, and subject matter.
2. Design a bespoke style dict with: `slide_bg`, `fonts` (title, body, optional extras), `palette` (5-8 colors), `accent_bar` (optional), and `design_notes`.
3. Present it to the user:
```
None of the 12 preset styles are a great fit for your content. I've designed a custom style:

  **CUSTOM — [Name]**
  Palette: [2-3 key colors described]
  Fonts: [title font] + [body font]
  Vibe: [1-line description]

Want me to go with this? Or would you prefer to pick from the 12 presets?
```
4. Wait for user confirmation, then use the custom style dict throughout — same as any preset style. The audit (CHECK 11) uses whatever style dict is active, including custom ones.
5. The custom style dict must follow the same structure as the presets in [Style → python-pptx Mapping](references/style-pptx-mapping.md) so all audit checks work identically.

Style references: [Design Styles Catalog](references/design-styles-catalog.md) for full descriptions, [Style → python-pptx Mapping](references/style-pptx-mapping.md) for implementation values.

### 2. Image Enhancement

After style is confirmed:

```
Would you like AI-generated background images for each slide?

  • Yes — I'll generate HD photorealistic 16:9 images tailored to each slide's
    content, with clean zones reserved for text overlay.
  • No — I'll use solid color / gradient backgrounds from the style palette.
```

**Wait for user response. Do not assume.**

### Environment

The presentation file path is stored in `PPTX_PATH`. Every Python script must read `os.environ['PPTX_PATH']`.

Ensure dependencies before first use:
```bash
python3 -m pip install python-pptx lxml --quiet
```

## Dual-Engine Architecture

Two engines for manipulating PowerPoint — choose the right one:

- **python-pptx** (file-based): Bulk creation, complex formatting (gradients, corner radius, letter spacing via lxml), images, charts, tables, font colors.
- **AppleScript IPC** (live editing): Text edits, font properties, positions, fills, z-order, visibility, rotation, shadows, speaker notes, slide management — all instant, no reload.

**Golden Rule:** Build with python-pptx, tweak with AppleScript. For edit-only tasks on an open presentation, use AppleScript alone (no python-pptx, no file reload).

See the full decision matrix and all live IPC operations in [AppleScript patterns](references/applescript-patterns.md).

## Workflows

### New Presentation (Full Build)

1. **Ask style + image questions** (see Interactive Pre-Build Questions above). Wait for answers.
2. **Plan** palette, fonts, and **composition strategy** — apply the chosen style from [Design Styles Catalog](references/design-styles-catalog.md) and [Style Mapping](references/style-pptx-mapping.md). For each slide, decide the layout pattern (from the [layout catalog](references/design-system.md#creative-layout-catalog)), where content zones sit, and where the image focal point goes. Vary layouts across slides (see [layout rhythm](references/design-system.md#layout-rhythm-across-slides)). If images are on, plan each slide's visual concept, text overlay zone, and focal point.
3. **Generate all needed images** (if user said yes) — use the `baoyu-danger-gemini-web` skill. **Generate images one at a time, sequentially — NEVER in parallel.** Parallel image requests can be rate-limited or blocked by the provider. Each prompt must specify 16:9 aspect ratio, composition/negative space matching the planned content zones, and style matching the deck's aesthetic.
4. **python-pptx**: Create file + build all slides (one per tool call). Apply style colors, fonts, backgrounds.
5. **Mandatory audit + fix loop** — read [Audit System](references/audit-system.md) and run all checks (1-11) iteratively. Fix cascading issues. Do NOT skip this step.
6. **AppleScript**: Open the file in PowerPoint.
7. **AppleScript**: Navigate through slides to verify visually — check that image focal points are unblocked and text sits in the planned zones.
8. **AppleScript**: Make any live tweaks (text, positions).
9. **AppleScript**: Save.
10. **Report** audit summary to user, then deliver the file path.

### Edit Existing Presentation (Live IPC)

1. AppleScript: Read all slides/shapes/text (enumerate).
2. Decide: minor text edits -> AppleScript. Major redesign -> python-pptx.
3. AppleScript: Make targeted live edits.
4. AppleScript: Save.

### Redesign Existing Presentation

1. AppleScript: Catalog everything (read all shapes/text).
2. Plan new design, palette, image strategy.
3. Generate needed images.
4. python-pptx: Rebuild each slide (clear old, add new).
5. AppleScript: Close and reopen the file.
6. AppleScript: Verify each slide visually.
7. AppleScript: Make live tweaks if needed.
8. AppleScript: Save.

### Quick Fix / Tweak (IPC-Only)

1. AppleScript: Read the target slide/shape.
2. AppleScript: Make the change live.
3. AppleScript: Save.

No python-pptx needed!

## Mandatory Audit — NON-NEGOTIABLE

**Every new or redesigned presentation MUST pass the full audit before delivery. No exceptions.**

The audit is **not optional**, **not skippable**, and **not deferrable**. It runs after all slides are built and before the file is shown to the user.

### What the audit does
Run all 11 checks from [Audit System](references/audit-system.md): bounds, text clipping, word-wrap, container sync, bullet alignment, overlap, z-order, font compliance, spacing, color/fill integrity, style compliance. Iterate up to 5 passes — fix issues, re-audit, repeat until clean.

### Enforcement rules
1. **Never deliver a .pptx without a clean audit.** If the audit finds CRITICAL issues, fix them. If fixes create new issues, re-audit.
2. **Always report the audit summary** to the user: CRITICAL count, WARNING count, fixes applied, passes needed.
3. **The audit runs on the saved file** — reload `Presentation(path)` after saving to get clean state.

### Anti-patterns (NEVER do these)
- Generating the .pptx and immediately saying "Here's your file!" without auditing — **this defeats the entire purpose of this skill.**
- Running only some checks — **all 11 checks must run every pass.**
- Skipping the audit because "it's a simple deck" — **simple decks still have font, bounds, and z-order issues.**
- Fixing an issue without re-auditing — **fixes cause cascading issues; re-audit is mandatory after every fix pass.**

---

## 18 Critical Rules

1. **Never set any font below 14pt.** Not on labels, footnotes, axis text, or table cells.
2. **Always set explicit positions.** Every shape and image must have left, top, width, height.
3. **Always save** at end of every Python script: `prs.save(pptx_path)`.
4. **Escape special characters** in XML: `&` -> `&amp;`, `<` -> `&lt;`, `>` -> `&gt;`.
5. **Never use emoji as icons.** Use generated images, geometric shapes, or labeled circles.
6. **Use gradients for backgrounds**, not flat solid colors (unless image background is used).
7. **Add decorative accents** — thin bars, underlines, transparency shapes on every slide.
8. **Prefer more slides over dense slides.** Split content rather than shrinking fonts.
9. **Build incrementally.** One slide per tool call. Announce progress.
10. **Verify after building.** Check overlaps, overflow, and visual quality.
11. **Composition-first: plan image + overlay as ONE design.** Before generating any background image, decide where text/content zones go and where the image focal point lives. Generate images with intentional negative space (dark/empty/blurred areas) matching your content zones. The best slides need NO overlay because the image was composed for the layout. When overlays are needed, use targeted overlays (only where text sits), not full-bleed. Never overlay the image's focal point. See the Composition Planning section in [Design System](references/design-system.md#composition-planning) for the full layout catalog and coordination rules.
12. **Use lxml for gradients.** The python-pptx `fill.gradient()` API can fail; the lxml XML approach is bulletproof.
13. **Use AppleScript IPC for quick edits.** Don't rebuild an entire deck when you only need to change one text box. Read -> edit -> save, all live.
14. **Remember the unit difference.** AppleScript uses points (72/inch). python-pptx uses EMUs (914400/inch). Convert: `EMU = points * 12700`.
15. **Always calculate text frame dimensions.** Never guess frame sizes. For each paragraph, sum the widths of ALL runs to get the paragraph width, then compute `ceil(para_width / frame_width)` to get the wrapped line count, then derive height from total lines. Use `word_wrap=False` for single-line elements. See the [Text Frame Sizing](#text-frame-sizing) section in python-pptx Reference.
16. **Surgical fixes only.** When fixing a bug (e.g., text overflow, overlap), change ONLY what's needed to fix that bug. Preserve all existing design decisions — border colors, accent bar direction, radius, opacity, card style, font sizes, spacing. Never redesign an element while fixing it. A fix that introduces a new visual inconsistency is not a fix.
17. **Separate decorative elements from content.** Decorative elements (slide numbers, icons, accent shapes) must have clear spatial separation from content text (titles, body). Never place a decorative element in the same quadrant at a similar position to a title — they will visually crowd each other. Ensure no horizontal or vertical overlap between decorative and content elements.
18. **Use moderate corner radius on content cards.** Rounded rectangle `adj` values: 3000 = barely visible, 10000 = moderate/pleasant, 16667 = default, 50000 = pill shape. Use `adj=10000` as the default for content cards. Pill shape (50000) is almost always too extreme for rectangular content cards.

## References

Detailed reference documentation is split into focused files. Read the relevant file when needed:

- **[python-pptx Reference](references/python-pptx-reference.md)**: Complete API reference — imports, opening/saving, shapes, text boxes, tables, charts, images, gradients, transparency, rounded corners, helper functions, overlap checker, audit code. **Read this before writing any python-pptx code.**
- **[AppleScript Patterns](references/applescript-patterns.md)**: Full live IPC capability reference — dual-engine architecture, presentation management, slide operations, live text/font/position/fill/z-order/visibility/rotation/shadow editing, speaker notes, comprehensive slide reader, known limitations, unit system, decision matrix. **Read this before any PowerPoint automation or live editing.**
- **[Design System](references/design-system.md)**: Typography rules, color palettes (dark premium, light clean, warm earth, bold vibrant, tropical dark), layout rules, decorative elements, image generation capability (prompts, workflow, strategy, layering), **composition planning** (10 creative layout patterns, layout rhythm across slides, image-overlay coordination, theme pairing, composition prompt engineering), EMU conversions. **Read this when planning a new deck's visual design.**
- **[Design Styles Catalog](references/design-styles-catalog.md)**: 12 curated design styles (STYLE-01 through STYLE-12) with full layout, typography, color palette, and graphic treatment specs for each. Styles range from Strategy Consulting (McKinsey) to Retro Risograph. **Read this when the user requests a specific style or you're recommending one.**
- **[Style → python-pptx Mapping](references/style-pptx-mapping.md)**: Concrete RGBColor values, font configs, accent bar settings, card/tile parameters, and design notes for each of the 12 styles. **Read this alongside the Design Styles Catalog to get implementation-ready values.**
- **[Audit System](references/audit-system.md)**: Mandatory post-generation quality audit — 11 checks (bounds, text clipping, word-wrap, container sync, bullet alignment, overlap, z-order, font compliance, spacing, color integrity, style compliance), iterative fix loop (max 5 passes), cascading fix strategies, word-wrap simulation, bullet layout algorithm, false positive avoidance. **Read this before running the mandatory audit after building slides.**
