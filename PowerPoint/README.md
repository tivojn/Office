# pptx-design-agent

A Claude Code skill for creating and editing professional PowerPoint presentations on macOS with premium design quality.

## Architecture

**Dual-engine approach** for macOS PowerPoint automation:

| Engine | Technology | Role |
|--------|-----------|------|
| **python-pptx** | python-pptx + lxml (file-based) | Bulk creation, gradients, corner radius, letter spacing, images, charts, tables |
| **AppleScript IPC** | osascript (live editing) | Text edits, font properties, positions, fills, z-order, visibility, rotation, shadows, speaker notes, slide management |

**Golden Rule:** Build with python-pptx, tweak with AppleScript. For edit-only tasks on an open presentation, use AppleScript alone (no python-pptx, no file reload).

**No stale display issue:** Unlike the xlsx skill, python-pptx writes files without PowerPoint open, so AppleScript's `open POSIX file` always loads a fresh copy from disk.

## Features

- **Create presentations from scratch** with premium design quality (gradients, cards, KPI panels, charts, tables)
- **Live-edit open presentations** via AppleScript IPC (text, fonts, positions, fills, z-order, shadows, speaker notes)
- **Composition-first design** — plan image + overlay as one design with intentional negative space
- **AI image generation** for slide backgrounds and content illustrations
- **5 built-in color palettes**: Dark Premium, Light Clean, Warm Earth, Bold Vibrant, Tropical Dark
- **10 creative layout patterns** with layout rhythm across slides
- **18 critical rules** for consistent, professional output

## Requirements

- **macOS** (AppleScript IPC requires Microsoft PowerPoint for Mac)
- **Python 3** with `python-pptx` and `lxml`
- **Microsoft PowerPoint** installed

## Installation

1. Copy the `pptx-design-agent/` folder into your Claude Code skills directory:

```bash
cp -r pptx-design-agent ~/.claude/skills/
```

2. Install Python dependencies:

```bash
python3 -m pip install python-pptx lxml
```

## Skill Structure

```
pptx-design-agent/
├── README.md                              # This file
├── SKILL.md                               # Main skill configuration & 18 critical rules
└── references/
    ├── python-pptx-reference.md           # python-pptx API reference, helpers, overlap checker
    ├── applescript-patterns.md            # Full live IPC capability reference & decision matrix
    └── design-system.md                   # Palettes, layouts, composition planning, image generation
```

## Workflows

### New Presentation (Full Build)

```
1. Plan:        Palette, fonts, composition strategy, layout rhythm
2. Generate:    AI images with intentional negative space for content zones
3. python-pptx: Build slides one at a time (one per tool call)
4. AppleScript: Open file in PowerPoint
5. AppleScript: Navigate & verify each slide visually
6. AppleScript: Make live tweaks (text, positions)
7. AppleScript: Save
```

### Edit Existing (Live IPC)

```
1. AppleScript: Read all slides/shapes/text (enumerate)
2. AppleScript: Make targeted live edits
3. AppleScript: Save
   (No python-pptx needed!)
```

### Redesign

```
1. AppleScript: Catalog everything (read all shapes/text)
2. Plan:        New design, palette, image strategy
3. Generate:    New images
4. python-pptx: Rebuild each slide
5. AppleScript: Close and reopen the file
6. AppleScript: Verify, tweak, save
```

### Quick Fix

```
AppleScript: Read -> edit -> save (no python-pptx needed)
```

## Usage

Once installed, Claude Code will automatically use this skill when you ask it to create or edit PowerPoint presentations. Examples:

- "Create a 10-slide pitch deck for my startup"
- "Build a quarterly business review presentation"
- "Redesign this presentation with a dark premium theme"
- "Change the title on slide 3 to 'Q4 Results'"
- "Add a background image to the title slide"

## License

MIT
