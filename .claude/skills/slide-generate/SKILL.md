---
name: slide-generate
description: Generate a complete PowerPoint presentation from an analyzed template, matching the template's style and layout patterns.
triggers:
  - generate presentation
  - create presentation
  - generate slides
  - make presentation
  - create slides from template
  - /slide-generate
compatibility: Requires Node.js 18+ with pptxgenjs installed. Run .claude/scripts/setup_deps.sh (Unix) or .claude/scripts/setup_deps.ps1 (Windows) to install dependencies. Template must be analyzed first with slide-analyze.
metadata:
  author: slide-deck-utils
  version: 1.0.0
---

# Skill: slide-generate

Generate a complete PowerPoint presentation from an analyzed template. The presentation will match the template's fonts, colors, layout patterns, and visual style.

**Prerequisite**: The template must be analyzed first using the `slide-analyze` skill.

## Workflow

### Step 1 — Accept Input

Gather from the user (ask if not provided):
- **Topic**: What the presentation is about
- **Template name**: Which analyzed template to use (list available templates in `slide-workspace/templates/` if unsure)
- **Slide count** (optional): Desired number of slides (default: 8-10)

### Step 2 — Verify Template Workspace

Check that the template workspace exists:
```
slide-workspace/templates/{template-name}/
  context.json     ← required
  guideline.md     ← required
  sample_code.js   ← required
  images/          ← may be empty
```

If missing, inform the user to run `slide-analyze` on the template first.

### Step 3 — Derive Presentation Name

Derive a filesystem-safe presentation name from the topic:
- Replace spaces and special characters with hyphens
- Convert to lowercase
- Example: `"Q4 Sales Review 2024"` → `q4-sales-review-2024`

### Step 4 — Create Presentation Workspace

Create the presentation workspace directory:
```
slide-workspace/presentations/{presentation-name}/
```

### Step 5 — Read Template Context

Read from the template workspace:
- `context.json` — full structured template data
- `guideline.md` — design guidelines
- `sample_code.js` — sample code for reference
- List all files in `images/` directory to get available image filenames

### Step 6 — Generate Outline

Read prompt rules from `.claude/prompts/outline-rules.md`.

Using the outline rules as system instructions and the template guideline + user's topic as context, generate a presentation outline:
- Numbered list of slide titles with bullet point content
- Match slide types from the template's Layout Patterns
- Include title/cover slide and thank-you/closing slide
- Keep to the requested slide count

Save the outline to:
```
slide-workspace/presentations/{presentation-name}/outline.md
```

### Step 7 — Generate Preamble Code

Read prompt rules from:
- `.claude/prompts/preamble-rules.md`
- `.claude/prompts/shared-pptxgenjs-rules.md`

Read the API reference from `.claude/docs/pptxgenjs-api.md`.

Using preamble-rules.md as system instructions, generate the PptxGenJS initialization code:
- `const pptx = new PptxGenJS();`
- `pptx.defineLayout(...)` if template uses custom dimensions
- All `pptx.defineSlideMaster(...)` calls — one per master in context.json

Do NOT include any slide content or `writeFile` call.

Save to:
```
slide-workspace/presentations/{presentation-name}/preamble.js
```

### Step 8 — Generate Per-Slide Code

Parse the outline into individual slides (each numbered item is a slide).

Read prompt rules from `.claude/prompts/slide-code-rules.md`.

For each slide in the outline:
- Use slide-code-rules.md as system instructions
- Provide: slide title + content, preamble.js for reference, template context, guideline, sample_code.js, available image filenames
- Generate code for ONE slide: `let slide = pptx.addSlide('MASTER_X')` + content
- Each slide starts with `// Slide N: <title>`

Collect all per-slide code blocks.

### Step 9 — Combine into Final Script

Combine the code using this structure:
```js
// Preamble
{preamble code}

// Hoisted helper functions (any `function name() {}` extracted from slide blocks)
{hoisted functions}

// Per-slide code (each wrapped in IIFE to prevent variable conflicts)
(function() {
  // Slide 1: ...
})();

(function() {
  // Slide 2: ...
})();

// ... more slides ...

pptx.writeFile({ fileName: 'output.pptx' });
```

**Important — function hoisting**: If any slide code defines top-level named functions (e.g., `function addLayout0Images(slide) {...}`), extract them from the slide block and place them before the IIFEs so they are accessible across all slides.

Save to:
```
slide-workspace/presentations/{presentation-name}/code.js
```

### Step 10 — Check Dependencies

Check that pptxgenjs is available:
```bash
node -e "require('pptxgenjs')" 2>/dev/null
```

If it fails, run the setup script:
- Unix/macOS: `bash .claude/scripts/setup_deps.sh`
- Windows: `powershell -ExecutionPolicy Bypass -File .claude/scripts/setup_deps.ps1`

### Step 11 — Run PPTXGenJS

Execute the code:
```bash
node .claude/scripts/run_pptxgenjs.js \
  "slide-workspace/presentations/{presentation-name}/code.js" \
  "slide-workspace/templates/{template-name}/images" \
  "slide-workspace/presentations/{presentation-name}/output.pptx"
```

### Step 12 — Report Result

Report to the user:
- Output file: `slide-workspace/presentations/{presentation-name}/output.pptx`
- Slide count generated
- Template used
- Any warnings from the runner

## Output Structure

```
slide-workspace/
  presentations/
    {presentation-name}/
      outline.md      ← Presentation outline
      preamble.js     ← PptxGenJS init + slide master definitions
      code.js         ← Complete combined script
      output.pptx     ← Generated presentation file
```

## Error Handling

- **Template not found**: Prompt user to run `slide-analyze` first
- **Runner exits with code 3**: JavaScript execution error — check code.js for syntax issues
- **Runner exits with code 4**: Script ran but no .pptx produced — check that `pptx.writeFile()` is called
- **Output too large**: Reduce slide count or simplify content

## Notes

- The `images` variable is pre-injected by run_pptxgenjs.js as `{ filename: absolutePath }` — never re-declare it in code.js
- Each slide MUST set `slide.background = { color: '<background_color>' }` — never rely on master inheritance
- Layout images (from `slide_layouts[layout_index].images`) must be added to every slide using that layout
- Multiple slide masters must never be collapsed into one
- Helper functions (e.g., `addLayout0Images`) must be hoisted to top level, outside IIFEs
