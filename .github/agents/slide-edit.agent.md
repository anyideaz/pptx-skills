---
description: Edit an existing generated presentation by modifying its PPTXGenJS code based on user instructions, then regenerate the PPTX file.
handoffs:
  - label: Analyze Template
    agent: slide-analyze
    prompt: Analyze a PPTX template
  - label: Generate New Presentation
    agent: slide-generate
    prompt: Generate a new presentation from template
---

# Skill: slide-edit

Edit an existing generated presentation. Reads the current PPTXGenJS code, applies the requested changes, and regenerates the PPTX file.

**Prerequisite**: The presentation must be generated first using the `slide-generate` skill.

## Workflow

### Step 1 — Accept Input

Gather from the user (ask if not provided):
- **Presentation name**: Which presentation to edit (list available presentations in `slide-workspace/presentations/` if unsure)
- **Edit description**: What changes to make (e.g., "change title font to red", "add a chart on slide 3", "remove slide 5")

### Step 2 — Verify Presentation Workspace

Check that the presentation workspace exists and contains the required files:
```
slide-workspace/presentations/{presentation-name}/
  code.js      ← required (current code to edit)
  outline.md   ← reference for slide structure
```

If missing, inform the user to run `slide-generate` first.

### Step 3 — Identify Template Used

Determine which template was used for this presentation:
1. Check `preamble.js` (if exists) for comments like `// Template: template-name`
2. Check `code.js` for `defineSlideMaster` calls or image references that match a template name
3. If still unclear, ask the user which template was used

The template workspace is at: `slide-workspace/templates/{template-name}/`

### Step 4 — Read Current State

Read from the presentation workspace:
- `code.js` — the current complete PPTXGenJS code
- `outline.md` — slide structure reference

Read from the template workspace:
- `context.json` — template structure data
- `guideline.md` — design rules
- `sample_code.js` — style reference

List all files in `slide-workspace/templates/{template-name}/images/` to get available image filenames.

### Step 5 — Read Prompt Rules

Read prompt rules from:
- `.claude/prompts/edit-rules.md` — editing-specific rules
- `.claude/prompts/shared-pptxgenjs-rules.md` — shared generation rules

Read the API reference from `.claude/docs/pptxgenjs-api.md`.

### Step 6 — Apply Edit

Using edit-rules.md as system instructions, generate the COMPLETE updated PPTXGenJS code:

Provide as context:
- Current `code.js` content
- User's edit description
- Template `guideline.md`
- Template `context.json` (with background_color fields)
- `sample_code.js` reference
- Available image filenames

Key editing rules:
- Output the COMPLETE updated code (not a diff or partial code)
- Preserve the overall template style unless user asks to change it
- Make ONLY the requested changes
- Keep `pptx.writeFile({ fileName: 'output.pptx' })` at the end
- Preserve `fill.transparency` values
- Never re-declare the `images` variable

### Step 7 — Save Updated Code

Back up the old code first:
```
slide-workspace/presentations/{presentation-name}/code.js.bak
```

Write the updated code to:
```
slide-workspace/presentations/{presentation-name}/code.js
```

### Step 8 — Check Dependencies

Check that pptxgenjs is available:
```bash
node -e "require('pptxgenjs')" 2>/dev/null
```

If it fails, run the setup script:
- Unix/macOS: `bash .claude/scripts/setup_deps.sh`
- Windows: `powershell -ExecutionPolicy Bypass -File .claude/scripts/setup_deps.ps1`

### Step 9 — Regenerate PPTX

Execute the updated code:
```bash
node .claude/scripts/run_pptxgenjs.js \
  "slide-workspace/presentations/{presentation-name}/code.js" \
  "slide-workspace/templates/{template-name}/images" \
  "slide-workspace/presentations/{presentation-name}/output.pptx"
```

### Step 10 — Report Result

Report to the user:
- Changes made (brief summary)
- Output file: `slide-workspace/presentations/{presentation-name}/output.pptx`
- Any warnings from the runner
- Note: previous code backed up to `code.js.bak`

## Error Handling

- **Presentation not found**: Prompt user to run `slide-generate` first
- **Template not identifiable**: Ask user which template was used
- **Runner exits with code 3**: JavaScript error in generated code — inspect the error and offer to fix
- **Runner exits with code 4**: No .pptx produced — check that `pptx.writeFile()` is present
- **Revert needed**: User can restore from `code.js.bak` if the edit made things worse

## Multiple Edits

For multiple sequential edits:
- Each edit reads the CURRENT `code.js` (already incorporating previous edits)
- The backup (`code.js.bak`) always reflects the PREVIOUS state before the latest edit
- The `outline.md` is not automatically updated — if slide structure changes significantly, update it manually or ask the user if they want the outline updated

## Notes

- The `images` variable is pre-injected by run_pptxgenjs.js — never re-declare it
- Each slide's `background_color` from context.json is the definitive background — always set it explicitly
- Layout images from `slide_layouts[layout_index].images` must be present on every slide using that layout
- Preserve `fill.transparency` values — do not drop transparency when editing shapes
