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
compatibility: Requires Node.js 18+ with pptxgenjs installed. Run shared/scripts/setup_deps.sh (Unix) or shared/scripts/setup_deps.ps1 (Windows) to install dependencies. Template must be analyzed first with slide-analyze.
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
  guideline.md     ← required
  sample_code.js   ← required
  images/          ← may be empty
```

If missing, inform the user to run `slide-analyze` on the template first.

### Step 3 — Derive Presentation Name

Derive a filesystem-safe presentation name from the topic with a timestamp suffix:
- Replace spaces and special characters with hyphens
- Convert to lowercase
- Append a timestamp in `YYYYMMDD-HHmmss` format (use the current date/time)
- Example: `"Q4 Sales Review 2024"` → `q4-sales-review-2024-20260407-143022`

The timestamp ensures each run produces a unique folder so outputs are never overwritten.

### Step 4 — Create Presentation Workspace

Create the presentation workspace directory:
```
slide-workspace/presentations/{presentation-name}/
```

### Step 5 — Read Template Context

Read from the template workspace:
- `guideline.md` — design guidelines
- `sample_code.js` — primary reference: contains master definitions, layout patterns, background colors, image usage, and positioning examples
- List all files in `images/` directory to get available image filenames

### Step 6 — Generate Outline

Read prompt rules from `shared/prompts/outline-rules.md`.

Using the outline rules as system instructions and the template guideline + user's topic as context, generate a presentation outline:
- Numbered list of slide titles with bullet point content
- Match slide types from the template's Layout Patterns
- Include title/cover slide and thank-you/closing slide
- Keep to the requested slide count

Save the outline to:
```
slide-workspace/presentations/{presentation-name}/outline.md
```

### Step 7 — Extract Preamble from sample_code.js

Extract the preamble directly from `sample_code.js` — no LLM generation needed.

The preamble consists of everything in `sample_code.js` **before** the first slide block, which includes:
- `const pptx = new PptxGenJS();`
- `pptx.defineLayout(...)` (if present)
- All `pptx.defineSlideMaster(...)` calls
- Any top-level helper functions (e.g. `function addLayout0Images(slide) {...}`)

A "slide block" begins when the first `pptx.addSlide` call appears. Stop extraction at that point.

Do NOT include any `addSlide` calls or `pptx.writeFile(...)`.

Save the extracted section to:
```
slide-workspace/presentations/{presentation-name}/preamble.js
```

### Step 8 — Generate Per-Slide Code

Parse the outline into individual slides (each numbered item is a slide).

Read prompt rules from `shared/prompts/slide-code-rules.md`.

For each slide in the outline:
- Use slide-code-rules.md as system instructions
- Provide: slide title + content, preamble.js for reference, guideline.md, sample_code.js (as style/pattern reference), available image filenames
- Generate code for ONE slide: `let slide = pptx.addSlide('MASTER_X')` + content
- Each slide starts with `// Slide N: <title>`
- Derive master name, background color, layout images, and positioning from `sample_code.js` patterns — do NOT use context.json

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
- Unix/macOS: `bash shared/scripts/setup_deps.sh`
- Windows: `powershell -ExecutionPolicy Bypass -File shared/scripts/setup_deps.ps1`

### Step 11 — Run PPTXGenJS

Execute the code:
```bash
node shared/scripts/run_pptxgenjs.js \
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
- `context.json` is NOT used during generation — all style/layout information is derived from `sample_code.js` and `guideline.md`
- The preamble is extracted verbatim from `sample_code.js` (not LLM-generated)
- Each slide MUST set `slide.background` matching the pattern in `sample_code.js`
- Layout helper functions (e.g., `addLayout0Images`) already defined in the preamble must be called as the FIRST statement on every slide
- Multiple slide masters must never be collapsed into one
- Helper functions must be hoisted to top level, outside IIFEs
