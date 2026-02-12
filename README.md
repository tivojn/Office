# Office Design Agent Skills

Claude Code skills for creating professional Microsoft Office documents on macOS. Each skill uses a dual-engine architecture — Python libraries for file-based creation and AppleScript for live Excel/PowerPoint/Word orchestration.

## Skills

| Skill | Engine | What It Does |
|-------|--------|-------------|
| **[Excel](Excel/)** | openpyxl + AppleScript | Dashboards, KPI panels, styled tables, charts, conditional formatting, VBA macros |
| **[PowerPoint](PowerPoint/)** | python-pptx + AppleScript | Presentations with gradients, cards, charts, images, speaker notes |
| **[Word](Word/)** | python-docx + AppleScript | Documents with cover pages, accent borders, callout cards, TOC, PDF export |

## 8 Design Styles

All three skills share the same curated design system. Specify a style by name or let the agent recommend one based on your content.

| Style | Name | Palette |
|-------|------|---------|
| XSTYLE-01 | Consulting & Strategy | Navy + Gold |
| XSTYLE-02 | Executive Dashboard | Electric Blue + Dark |
| XSTYLE-03 | Corporate Report | Royal Blue + Emerald |
| XSTYLE-04 | Data Science & Technical | Dark + Cyan/Magenta |
| XSTYLE-05 | Sales & Pipeline | Coral + Teal |
| XSTYLE-06 | Finance & Accounting | Forest Green + Gold |
| XSTYLE-07 | Marketing & Creative | Terracotta + Sage |
| XSTYLE-08 | Operations & Logistics | Cobalt + Amber |

Custom styles are also supported — the agent will design a bespoke palette if none of the 8 presets fit.

## Installation

### Quick Install (npx)

```bash
# Install all 3 skills at once
npx skills add tivojn/Office -a claude-code

# Or install individually
npx skills add tivojn/Office --skill xlsx-design-agent -a claude-code
npx skills add tivojn/Office --skill pptx-design-agent -a claude-code
npx skills add tivojn/Office --skill docx-design-agent -a claude-code
```

### Manual Install

Copy any skill folder into `~/.claude/skills/`:

```bash
# Excel
cp -R Excel ~/.claude/skills/xlsx-design-agent

# PowerPoint
cp -R PowerPoint ~/.claude/skills/pptx-design-agent

# Word
cp -R Word ~/.claude/skills/docx-design-agent
```

### Dependencies

```bash
# Excel
pip install openpyxl Pillow pandas

# PowerPoint
pip install python-pptx Pillow lxml

# Word
pip install python-docx Pillow lxml
```

Requires macOS with Microsoft Office installed for AppleScript orchestration.

## How It Works

Each skill follows the same pattern:

1. **Python creates the file** — data, formatting, charts, images — all without Office running
2. **AppleScript opens the file** in the corresponding Office app
3. **VBA/AppleScript finalizes** — refresh pivots, update TOC, export PDF
4. **Mandatory audit** — 10-point quality check with iterative fix loop before delivery

## License

MIT
