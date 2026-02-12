# docx-design-agent

A Claude Code skill for creating and editing professional Word documents on macOS with premium design quality.

## Architecture

**Dual-engine approach** for macOS Word automation:

| Engine | Technology | Role |
|--------|-----------|------|
| **python-docx** | python-docx + lxml (file-based) | Bulk creation, styles, tables, images, headers/footers, sections, accent borders, callout cards |
| **AppleScript IPC** | osascript (live editing) | Text edits, find/replace, field updates, TOC refresh, PDF export, formatting tweaks |

**Golden Rule:** Build with python-docx, finalize with AppleScript. For edit-only tasks on an open document, use AppleScript alone (no python-docx, no file reload).

## Features

- **Create documents from scratch** with premium design quality (styled cover pages, accent borders, callout cards, KPI tables, pull quotes, sidebar panels)
- **Live-edit open documents** via AppleScript IPC (text, fonts, find/replace, fields, TOC)
- **Finalize with AppleScript** — update TOC, refresh fields, export to PDF
- **AI image generation** for cover art, section illustrations, and inline visuals
- **5 built-in color palettes**: Dark Premium, Light Clean, Warm Earth, Bold Vibrant, Corporate Blue
- **Dual-engine architecture**: python-docx for building, AppleScript for live tweaks and finalization

## Requirements

- **macOS** (AppleScript IPC requires Microsoft Word for Mac)
- **Python 3** with `python-docx`, `lxml`, and `Pillow`
- **Microsoft Word** installed

## Installation

1. Copy the `docx-design-agent/` folder into your Claude Code skills directory:

```bash
cp -r docx-design-agent ~/.claude/skills/
```

2. Install Python dependencies:

```bash
python3 -m pip install python-docx lxml Pillow
```

## Skill Structure

```
docx-design-agent/
├── README.md                              # This file
├── SKILL.md                               # Main skill configuration & critical rules
└── references/
    ├── python-docx-reference.md           # python-docx API reference & helper functions
    ├── applescript-patterns.md            # AppleScript live IPC patterns & decision matrix
    └── design-system.md                   # Typography, palettes, layout rules, decorative elements
```

## Workflows

### New Document (Full Build)

```
1. Plan:        Design palette, fonts, layout, image strategy
2. Generate:    AI images for cover art and illustrations
3. python-docx: Build document section by section
4. AppleScript: Open in Word, update fields & TOC
5. AppleScript: Verify visually, make live tweaks
6. AppleScript: Save / Export PDF
```

### Edit Existing (Live IPC)

```
1. AppleScript: Read content from open document
2. AppleScript: Make targeted live edits (text, fonts, find/replace)
3. AppleScript: Update fields & TOC
4. AppleScript: Save
   (No python-docx needed!)
```

### Redesign

```
1. AppleScript: Catalog all content
2. Plan:        New design, palette, image strategy
3. python-docx: Rebuild entire document
4. AppleScript: Open, verify, tweak, save
```

### Quick Fix

```
AppleScript: Read -> edit -> save (no python-docx needed)
```

## Usage

Once installed, Claude Code will automatically use this skill when you ask it to create or edit Word documents. Examples:

- "Create a professional quarterly report"
- "Build a styled proposal document with cover page"
- "Redesign this document with a dark premium theme"
- "Update the table of contents and export to PDF"
- "Change the heading style across the entire document"

## License

MIT
