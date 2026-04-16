# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Setup

**Requirements**: Node.js 18+, Python 3.10+

```bash
# Unix/macOS
bash shared/scripts/setup_deps.sh

# Windows (PowerShell)
.\shared\scripts\setup_deps.ps1
```

This installs `pptxgenjs`, `uuid` (Node), `python-pptx`, `Pillow`, and `markitdown` (Python) into `shared/scripts/node_modules/`.

**Verify:**
```bash
node -e "require('pptxgenjs')"
python -c "import pptx"
python -c "import markitdown"
```

## Architecture Overview

**Slide Agent** is an AI-powered PowerPoint generation system with three chainable phases:

```
Any Document  → [/markitdown]    → content.md (Markdown extract)
                                              ↓
PPTX Template → [/slide-analyze] → context.json + guideline.md + sample_code.js
                                              ↓
Topic/Slides  → [/slide-generate] → outline.md + preamble.js + code.js → output.pptx
                                              ↓
Edit Request  → [/slide-edit]    → updated code.js + code.js.bak → output.pptx
```

### Key Scripts

- [shared/scripts/extract_template.py](shared/scripts/extract_template.py) — Parses PPTX using `python-pptx`, outputs `context.json` (fonts, colors, positions, layouts, images) and extracted `images/` directory.
- [shared/scripts/run_pptxgenjs.js](shared/scripts/run_pptxgenjs.js) — Executes AI-generated PPTXGenJS code. Auto-fixes known PPTX spec bugs (phantom Content_Types, empty line elements, table styling). Injects the `images` variable mapping filenames to absolute paths.
- [shared/scripts/convert_to_markdown.py](shared/scripts/convert_to_markdown.py) — Converts documents (PDF, Word, Excel, PPTX, HTML, CSV, images…) to Markdown using `markitdown`. Output is saved to `slide-workspace/sources/{name}/content.md`.

### Workspace Layout (gitignored)

```
slide-workspace/
  sources/{name}/         ← converted source documents (markitdown output)
    content.md
  templates/{name}/       ← analyzed template artifacts
    original.pptx, context.json, guideline.md, sample_code.js, images/
  presentations/{name}/   ← generated presentation artifacts
    outline.md, preamble.js, code.js, code.js.bak, output.pptx
```

### Skills & Prompts

Skills in [.claude/skills/](.claude/skills/) define multi-step agent workflows triggered by `/markitdown`, `/slide-analyze`, `/slide-generate`, `/slide-edit`. Prompts in [shared/prompts/](shared/prompts/) provide LLM instructions for each generation phase. GitHub Copilot equivalents live in [.github/agents/](.github/agents/) and [.github/prompts/](.github/prompts/).


When **adding** a new skill:
When **modifying** a skill (workflow steps, naming conventions, rules, etc.):
1. Update `.claude/skills/{name}/SKILL.md`
2. Apply the same change to `.github/agents/{name}.agent.md`

When **deleting** a skill:
1. Remove `.claude/skills/{name}/SKILL.md`
2. Remove `.github/agents/{name}.agent.md`
3. Remove the row from the inventory table above

## Critical PPTXGenJS Constraints

These rules apply whenever writing or reviewing generated `code.js`:

- **Text valign**: Always set `valign: 'top'` explicitly on every `addText()` call — there is no safe default.
- **Shape references**: Use `pptx.shapes.RECTANGLE` (uppercase enum), never string literals.
- **Text arrays**: Rich text must be `[{ text: '...', options: {...} }]`, never a plain string.
- **Layout images**: Every slide must include ALL images from `slide_layouts[N].images` — these are layout-level decorations that must be manually added.
- **Backgrounds**: Set via `slide.background = { color: '...' }`, never via `addShape`.
- **Multiple masters**: Each master requires its own `defineSlideMaster()` call; never merge masters.
- **Helper functions**: Functions like `addLayout0Images()` must be defined outside IIFEs to avoid hoisting issues.
- **Table styling**: Apply fills at cell level; never use white borders on white backgrounds.

See [shared/docs/pitfalls.md](shared/docs/pitfalls.md) for the full list and [shared/docs/pptxgenjs-api.md](shared/docs/pptxgenjs-api.md) for the API reference.

## Script Exit Codes

All scripts return meaningful exit codes:
- `0` — success
- `2` — file not found
- `3` — parse/execution error
- `4` — (runner only) no `.pptx` produced

## context.json Key Fields

```json
{
  "presentation": { "slide_width_inches": 10, "slide_height_inches": 5.625 },
  "theme": { "color_scheme": { "accent1": "#FF5733" } },
  "slide_masters": [{ "name": "MASTER_0", "background_color": "#FFF", "images": [] }],
  "slide_layouts": [{ "name": "Layout 0", "master_index": 0, "images": [] }],
  "fonts_summary": { "families_used": ["Arial"] },
  "images_manifest": [{ "id": "...", "filename": "master0_Logo.png" }]
}
```

## Naming Conventions

All workspace output folders include a `YYYYMMDD-HHmmss` timestamp suffix so multiple runs never overwrite each other.

- Template names: `{slug}-{timestamp}` from PPTX filename (`My Corp.pptx` → `my-corp-20260407-143022`)
- Presentation names: `{slug}-{timestamp}` from topic (`Q4 Sales` → `q4-sales-20260407-143022`)
- Source names: lowercase hyphenated derived from input filename (`Q4-Report.xlsx` → `q4-report`) — no timestamp (sources are one-time conversions)
- Edit outputs: `output-{timestamp}.pptx` inside the existing presentation folder
- Image filenames: preserved as extracted (`master0_Logo.png`, `layout0_bg.jpg`)
