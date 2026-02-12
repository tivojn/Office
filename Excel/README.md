# xlsx-design-agent

A Claude Code skill for creating and editing professional Excel spreadsheets on macOS with industrial-grade reliability.

## Architecture

**Dual-engine approach** optimized for macOS Excel automation:

| Engine | Technology | Role |
|--------|-----------|------|
| **Python** | pandas, numpy, Pillow | Data transformation, validation, image preflight |
| **openpyxl** | openpyxl (file-based) | ALL file creation and editing — data I/O, formatting, borders, charts, images, tab colors, conditional formatting. Works without Excel running. |
| **VBA** | Excel-native macros | Pivot table refresh, recalculation, complex exports — operations requiring a live Excel instance |
| **AppleScript** | osascript | Orchestration (open/close/save Excel, run VBA macros) and recovery (stuck/modal/unfocused Excel) |

**Key insight:** openpyxl works directly on `.xlsx` files without needing Excel to be running. It handles everything — data, formatting, borders, charts, images, conditional formatting. AppleScript opens the file in Excel when ready and runs VBA macros for live-instance operations (pivot refresh, recalculation, exports). No appscript bridge, no complex handoff protocol.

## Features

- **File-based creation** — openpyxl builds the entire workbook without Excel running
- **Full formatting support** — borders, charts, images, conditional formatting all work natively
- **Simple workflow** — create file with openpyxl, open in Excel, done
- **VBA macros via AppleScript** — `do Visual Basic` for pivot refresh, exports
- **Template-first approach** — `.xlsm` templates with prebuilt structure + macros
- **Transactional runs** — copy template -> write data -> verify -> export
- **Structured auditing** — verify every build before export
- **8 curated design styles** — XSTYLE-01 through XSTYLE-08
- **Image insertion** — via openpyxl with aspect ratio enforcement

## Requirements

- **macOS** with **Microsoft Excel** installed (for VBA macros and final viewing)
- **Python 3** with `openpyxl`, `Pillow`, `pandas`

## Installation

1. Copy the `xlsx-design-agent/` folder into your Claude Code skills directory:

```bash
cp -r xlsx-design-agent ~/.claude/skills/
```

2. Install Python dependencies:

```bash
python3 -m pip install openpyxl Pillow pandas
```

## Skill Structure

```
xlsx-design-agent/
├── README.md                                  # This file
├── SKILL.md                                   # Main skill configuration (21 critical rules)
└── references/
    ├── openpyxl-reference.md                  # File-based I/O, formatting, charts, images, borders
    ├── vba-macros-reference.md                # VBA macro library (pivots, exports, auditing)
    ├── applescript-patterns.md                # Orchestration (open/save/run VBA) + recovery patterns
    ├── template-design.md                     # Template conventions & structure
    ├── design-system.md                       # Palettes, layouts, typography, chart/table design
    ├── design-styles-catalog.md               # 8 curated design styles (XSTYLE-01 through XSTYLE-08)
    ├── style-xlsx-mapping.md                  # Implementation-ready values for each style
    └── audit-system.md                        # Mandatory 10-check quality audit
```

## Workflow

### Typical Build

```
1. Python:   Transform data -> DataFrames
2. openpyxl: Create workbook, write data, format, borders, charts, images
3. Save:     wb.save(path)
4. Open:     subprocess.run(['open', path])  <- Excel reads the file
5. (Optional) AppleScript: Run VBA macros if needed (pivot refresh, etc.)
6. Done:     User sees the finished workbook in Excel
```

### Simple Tasks (openpyxl Only)

```
openpyxl: Create -> write data -> format -> borders -> charts -> save -> open
```

No Excel needed during creation. No handoff protocol. No engine switching.

## 21 Critical Rules

See `SKILL.md` for the full list. Highlights:

1. Never write cell-by-cell in loops when bulk patterns exist
2. openpyxl handles ALL formatting — borders, charts, images, conditional formatting
3. VBA macros called via AppleScript for live-instance operations
4. Default to `.xlsx` unless VBA macros are needed (then `.xlsm`)

## License

MIT
