# Slide Agent — AI-Powered PowerPoint Generation from Templates

Slide Agent is a set of skills integrated into **Claude Code** and **GitHub Copilot** (VS Code) that lets you create, edit, and manage complete PowerPoint presentations from your existing templates — using only natural language.

---

## Overview

The system works in three phases:

```
PPTX Template   ──[slide-analyze]──▶  Extract design guidelines
                                              │
Topic / Content ──[slide-generate]──▶  Generate presentation
                                              │
Edit request    ──[slide-edit]──────▶  Modify & regenerate
```

**Core strength:** Every generated presentation faithfully follows the original template's design — fonts, colors, layouts, and embedded images are all preserved automatically.

---

## Installation

### Requirements

| Component | Minimum Version |
|---|---|
| [Claude Code](https://claude.ai/code) **or** [GitHub Copilot](https://github.com/features/copilot) (VS Code) | Latest |
| Node.js | 18+ |
| Python | 3.10+ |

### Install Dependencies

**Unix / macOS:**
```bash
bash shared/scripts/setup_deps.sh
```

**Windows (PowerShell):**
```powershell
.\shared\scripts\setup_deps.ps1
```

This installs:
- `pptxgenjs` ^4.0.1 and `uuid` ^10.0.0 (Node.js)
- `python-pptx` ≥1.0.0 and `Pillow` ≥10.0.0 (Python)

---

## Skills

### `/slide-analyze` — Analyze a Template

Analyzes a PPTX template file to extract its complete design system: fonts, colors, slide layouts, embedded images, and visual rules.

**Input:** Path to a `.pptx` template file

**Output** (saved to `slide-workspace/templates/{template-name}/`):
- `context.json` — Full structured data of the template
- `guideline.md` — Design rules, typography, color palette
- `sample_code.js` — Example PPTXGenJS code in the template's style
- `images/` — Extracted embedded images

**Example prompt:**
```
/slide-analyze
Template: slide-workspace/templates/my-company-template/original.pptx
```

---

### `/slide-generate` — Generate a Presentation

Creates a complete presentation from a previously analyzed template, based on your topic and content requirements.

**Prerequisite:** Template must be analyzed first with `/slide-analyze`

**Output** (saved to `slide-workspace/presentations/{presentation-name}/`):
- `outline.md` — Slide-by-slide content outline
- `preamble.js` — PPTXGenJS initialization code
- `code.js` — Full generated code for all slides
- `output.pptx` — The final PowerPoint file

**Example prompt:**
```
/slide-generate
Template: my-company-template
Name: team-weekly-report-w15
Topic: Backend Team Weekly Progress Report — Week 15
Slides: 8
```

---

### `/slide-edit` — Edit a Presentation

Edits an existing generated presentation using natural language. The system updates the code and regenerates the PPTX automatically.

**Prerequisite:** Presentation must be generated first with `/slide-generate`

**Safety:** Automatically creates a `code.js.bak` backup before any changes

**Example prompt:**
```
/slide-edit
Presentation: team-weekly-report-w15
Changes: Change slide 3 background to navy blue, update the metrics on slide 5
```

---

## Example Prompts

### Weekly progress report

```
/slide-generate
Template: corporate-weekly-report
Name: backend-team-weekly-w16
Topic: Backend Team Sprint Progress — Week 16
Slides: 10
Content:
- Sprint overview: 18/22 story points completed
- API authentication module: 100% done
- Database migration: in progress, 70%
- Blockers to escalate: bottleneck in review process
- Next week plan: finalize migration, start caching layer
```

### Training / onboarding deck

```
/slide-generate
Template: training-template
Name: training-docker-basics
Topic: Docker for Developers — From Zero to Production
Slides: 15
Style: Professional, include code examples and architecture diagrams
```

### Analyze a new template

```
/slide-analyze
I have my company's PowerPoint template at:
slide-workspace/templates/company-brand/original.pptx
Please analyze it and generate a guideline so I can create branded slides.
```

### Quick edits after generation

```
/slide-edit
Presentation: training-docker-basics
Changes:
1. Slide 1: Update date to April 15, 2026
2. Slide 7: Add a note "Requires Docker Desktop 4.x+"
3. Last slide: Add a QR code placeholder and contact information
```

### Full workflow — new topic on existing template

```
/slide-generate
Template: corporate-weekly-report
Name: infra-incident-postmortem
Topic: Infrastructure Incident Postmortem — April 2026
Slides: 12
Content:
- Incident summary and timeline
- Root cause analysis
- Impact assessment (users affected, downtime duration)
- Immediate remediation steps taken
- Long-term preventive measures
- Action items with owners and due dates
```

---

## Workspace Structure

```
slide-workspace/
├── templates/                    # Analyzed templates (reusable)
│   └── {template-name}/
│       ├── original.pptx         # Original template file
│       ├── context.json          # Structured template data
│       ├── guideline.md          # Design guidelines
│       ├── sample_code.js        # Example PPTXGenJS code
│       └── images/               # Extracted embedded images
│
└── presentations/                # Generated presentations
    └── {presentation-name}/
        ├── outline.md            # Content outline
        ├── preamble.js           # Initialization code
        ├── code.js               # Full slide code
        ├── code.js.bak           # Backup before last edit
        └── output.pptx           # Final PowerPoint file
```

---

## Why Use Templates?

Instead of building slides from scratch, Slide Agent deeply analyzes your template so that:

1. **Brand identity is preserved** — Logos, colors, and fonts are applied exactly as designed
2. **Layouts are reused correctly** — Slide masters and layout patterns are detected automatically
3. **Visual consistency is guaranteed** — Every slide from first to last follows the same style rules
4. **Design time is eliminated** — Just describe the content; the AI handles the presentation layer

A template only needs to be analyzed once. After that, you can generate any number of presentations from it.

---

## Recommended Workflow

```
1. Place your PPTX template in slide-workspace/templates/{name}/original.pptx
         ↓
2. /slide-analyze  →  Analyze once, reuse forever
         ↓
3. /slide-generate  →  Generate a presentation on any topic
         ↓
4. /slide-edit  →  Refine based on feedback
         ↓
5. Open output.pptx in PowerPoint or Google Slides
```

---

## Troubleshooting

| Issue | Solution |
|---|---|
| `python-pptx not found` | Re-run `setup_deps.sh` or `pip install python-pptx Pillow` |
| `pptxgenjs not found` | Run `npm install` inside `shared/scripts/` |
| PPTX file fails to open | Re-run `/slide-generate` — the runner auto-fixes common PPTX XML issues |
| Output doesn't match template style | Ensure `/slide-analyze` was run first; check `guideline.md` in the template folder |
| `/slide-edit` breaks the layout | Restore from `code.js.bak`, then retry with a more specific description |

---

## Tech Stack

- **[PPTXGenJS](https://gitbrent.github.io/PptxGenJS/)** — JavaScript library for generating PowerPoint files
- **[python-pptx](https://python-pptx.readthedocs.io/)** — Python library for parsing and extracting PPTX data
- **Claude AI / GitHub Copilot** — Design analysis, content generation, PPTXGenJS code synthesis
- **Agent Skills** — Workflow integration into Claude Code CLI and GitHub Copilot (VS Code) environments
