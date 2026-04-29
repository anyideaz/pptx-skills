# Slide Agent — AI-Powered PowerPoint Generation

Slide Agent generates complete PowerPoint presentations from your existing templates using natural language. Provide a `.pptx` template, describe the content, and the system produces a finished deck that preserves your brand's fonts, colors, layouts, and embedded imagery.

---

## Who This Is For

This guide is intended for **non-technical users** — professionals who want to use AI to accelerate routine slide work without writing code:

- Project managers producing weekly or sprint reports
- Internal trainers preparing onboarding and training materials
- Sales teams creating proposals and pitch decks
- Marketing teams building campaign and event presentations
- Anyone who regularly produces slides from a corporate template

No programming knowledge is required. You interact with the system entirely through natural-language instructions.

---

## What You Can Do

| Capability | Description |
|---|---|
| **Generate from template** | Produce a complete deck from a `.pptx` template plus a topic description. |
| **Convert documents to slides** | Turn an existing Word, PDF, or Excel file into structured slide content. |
| **Edit by instruction** | Modify finished decks with plain-language requests (e.g. *"change slide 3 background to navy"*). |
| **Reuse templates indefinitely** | Analyze a template once; generate any number of presentations from it afterward. |

---

## Prerequisites

You will need **one** of the following AI environments installed:

- **[Claude Code](https://claude.ai/code)** by Anthropic
- **[GitHub Copilot](https://github.com/features/copilot)** in VS Code

After installing the AI tool, complete the [first-time setup](#first-time-setup) below. This is required only once per machine.

---

## Workflow Overview

The system operates in three phases:

```
   Phase 1                Phase 2                Phase 3
┌──────────┐          ┌──────────┐          ┌──────────┐
│ ANALYZE  │  ────►  │ GENERATE │  ────►  │   EDIT   │
│ TEMPLATE │          │   DECK   │          │  (opt.)  │
└──────────┘          └──────────┘          └──────────┘
 (once per          (once per             (as needed)
  template)          presentation)
```

---

## Step-by-Step Guide

### Step 1 — Analyze a Template

Point the system at any `.pptx` file on your computer — Desktop, Downloads, a shared drive, or anywhere else. There is no need to copy it into the project folder.

In Claude Code or Copilot Chat, describe the request in plain language and reference the file path:

```
/slide-analyze
Please analyze template: ~/Desktop/AcmeCorp_Brand_Template.pptx
```

The system extracts the template's design system — color palette, typography, layouts, slide masters, and embedded images — and stores the result for reuse.

> Each template is analyzed **once**. Subsequent presentations using that template skip this step.

---

### Step 2 — Generate a Presentation

Describe what you need conversationally. The system parses your intent and produces a finished deck.

```
/slide-generate
Using the AcmeCorp template, create an 8-slide weekly report for the
backend team covering Sprint 15. We completed 18 of 22 story points,
finished the API authentication module, and made 70% progress on the
database migration. Highlight the review-queue bottleneck as a blocker,
and close with next week's plan to finalize the migration and begin the
caching layer.
```

The output is `output.pptx`, ready to open in PowerPoint, Keynote, or Google Slides.

---

### Step 3 — Edit (Optional)

Describe revisions in natural language:

```
/slide-edit
In the weekly report deck, please update the date on slide 1 to
April 29, 2026, replace the bar chart on slide 5 with a pie chart,
and add contact information plus a QR code placeholder to the final slide.
```

The prior version is automatically backed up before any edit is applied.

---

## Best Practices

### Be specific in content descriptions

Vague: *"Create a marketing presentation."*

Specific: *"Create a 6-slide pitch introducing Product X to enterprise customers, covering: problem, solution, demo, pricing, case study, call to action."*

The quality of the output is directly proportional to the specificity of the input.

### Reuse existing documents

If your content already exists as a Word, Excel, or PDF document, use `/markitdown` to convert it into slide-ready material rather than re-typing:

```
/markitdown
File: q1-financial-report.xlsx
```

You can then reference the converted content when invoking `/slide-generate`.

### Invest in template quality

Output quality scales with template quality. Templates with well-defined slide masters, varied layouts, and clean branding produce significantly better results. Choose or prepare your template carefully before the first analysis.

### One template, many presentations

After `/slide-analyze`, a single template can generate unlimited presentations. This is well-suited for:

- Recurring weekly, monthly, or quarterly reports
- Standardized training materials
- Pitch decks tailored per client

---

## Common Use Cases

### Weekly project report

```
/slide-generate
Using the internal report template, prepare a 10-slide Week 17 update
for the data team. Cover the week's highlights, status of major projects,
KPI metrics, risks and blockers, and the plan for next week.
```

### Training and onboarding

```
/slide-generate
Using the training template, create a 15-slide deck on IT procedures
for new hires. Keep the tone professional and concise, and include
supporting iconography where appropriate.
```

### Incident postmortem

```
/slide-generate
Using the internal report template, prepare a 12-slide postmortem on
the April 2026 infrastructure incident. Include the incident summary
and timeline, root cause analysis, impact assessment (users affected
and downtime duration), immediate remediation steps, long-term
preventive measures, and a list of action items with owners and due dates.
```

---

## Frequently Asked Questions

**Do I need programming experience?**
No. All interaction is through natural-language commands.

**Will the AI alter my template's design?**
No. Slide Agent is designed to follow the template's design system precisely — typography, color palette, and brand assets are preserved.

**Will new presentations overwrite existing ones?**
No. Each output folder is timestamped (e.g. `weekly-report-w15-20260429-143022`), so prior versions are never overwritten.

**Are the output files compatible with PowerPoint and Google Slides?**
Yes. Output is standard `.pptx`, compatible with Microsoft PowerPoint, Apple Keynote, Google Slides, and LibreOffice Impress.

**The output does not match my template's style. What should I check?**
Confirm that `/slide-analyze` was run for that template. If the issue persists, review the generated `guideline.md` to see how the system interpreted the template, then provide corrective feedback.

**What happens if an edit produces an unwanted result?**
Each edit creates a `code.js.bak` backup of the prior version. You can restore from backup or regenerate from scratch with `/slide-generate`.

**Is there a limit on slide count?**
There is no hard limit, though decks longer than 20 slides are best produced in segments for optimal quality.

---

## First-Time Setup

Required once per machine.

### System Requirements

| Component | Minimum Version |
|---|---|
| [Claude Code](https://claude.ai/code) **or** [GitHub Copilot](https://github.com/features/copilot) (VS Code) | Latest |
| [Node.js](https://nodejs.org/) | 18 or higher |
| [Python](https://www.python.org/downloads/) | 3.10 or higher |

### Install Dependencies

Open a terminal in the project directory and run:

**macOS / Linux:**
```bash
bash shared/scripts/setup_deps.sh
```

**Windows (PowerShell):**
```powershell
.\shared\scripts\setup_deps.ps1
```

Setup typically completes in one to two minutes.

### Verify Installation

```bash
node -e "require('pptxgenjs')"
python -c "import pptx"
```

If neither command produces an error, the system is ready.

---

## Workspace Structure

All generated artifacts are stored under `slide-workspace/`:

```
slide-workspace/
├── templates/                    Analyzed templates (reusable)
│   └── {template-name}/
│       ├── original.pptx         Original template file
│       ├── guideline.md          Extracted design rules
│       └── images/               Embedded images
│
└── presentations/                Generated presentations
    └── {presentation-name}/
        ├── outline.md            Slide-by-slide outline
        ├── output.pptx           Final PowerPoint file
        └── code.js.bak           Backup from prior edit
```

The primary output is `output.pptx`.

---

## Troubleshooting

| Issue | Resolution |
|---|---|
| Slash commands are not recognized | Ensure Claude Code or VS Code is opened from the `pptx-skills` project root. |
| `python-pptx not found` | Re-run the setup script, or install manually: `pip install python-pptx Pillow`. |
| `pptxgenjs not found` | Run `npm install` inside `shared/scripts/`. |
| Output `.pptx` fails to open | Re-run `/slide-generate`; the runtime auto-corrects common PPTX XML issues. |
| Output style does not match the template | Confirm `/slide-analyze` was executed; review `guideline.md` for misinterpretations. |
| `/slide-edit` produces a broken layout | Restore from `code.js.bak` and retry with a more precise instruction. |

For unresolved issues, describe the problem to Claude Code or Copilot directly — the AI can assist with diagnosis.

---

## Technical Reference

For implementation details and contribution guidelines:

- **Architecture and conventions:** [CLAUDE.md](CLAUDE.md)
- **PPTXGenJS API reference:** [shared/docs/pptxgenjs-api.md](shared/docs/pptxgenjs-api.md)
- **Known constraints and pitfalls:** [shared/docs/pitfalls.md](shared/docs/pitfalls.md)
- **Skill definitions:** [.claude/skills/](.claude/skills/) (Claude Code) and [.github/agents/](.github/agents/) (GitHub Copilot)

**Stack:** [PPTXGenJS](https://gitbrent.github.io/PptxGenJS/), [python-pptx](https://python-pptx.readthedocs.io/), [markitdown](https://github.com/microsoft/markitdown), Claude Agent Skills, GitHub Copilot Custom Agents.
