---
name: slide-analyze
description: Analyze a PPTX template to extract style guidelines, sample code, and cataloged images into the slide workspace.
triggers:
  - analyze template
  - analyze pptx
  - extract template style
  - pptx analysis
  - analyze slide template
  - /slide-analyze
compatibility: Requires Python 3.10+ with python-pptx and Pillow installed. Run .claude/scripts/setup_deps.sh (Unix) or .claude/scripts/setup_deps.ps1 (Windows) to install dependencies.
metadata:
  author: slide-deck-utils
  version: 1.0.0
---

# Skill: slide-analyze

Analyze a PPTX template file to extract fonts, colors, layouts, slide masters, and embedded images. Produces a structured `context.json`, a `guideline.md` with design rules, and `sample_code.js` demonstrating the template style.

## Workflow

### Step 1 — Accept Input

Ask the user for the PPTX file path if not provided. Confirm the file exists.

### Step 2 — Derive Template Name

Derive a filesystem-safe template name from the filename:
- Take the filename without extension (e.g., `my-corp-template.pptx` → `my-corp-template`)
- Replace spaces and special characters with hyphens
- Convert to lowercase

### Step 3 — Create Workspace

Create the template workspace directory:
```
slide-workspace/templates/{template-name}/
```

Copy the original PPTX file to:
```
slide-workspace/templates/{template-name}/original.pptx
```

### Step 4 — Check Dependencies

Check that python-pptx is available:
```bash
python -c "import pptx" 2>/dev/null || python3 -c "import pptx" 2>/dev/null
```

If the check fails, run the appropriate setup script:
- Unix/macOS: `bash .claude/scripts/setup_deps.sh`
- Windows: `powershell -ExecutionPolicy Bypass -File .claude/scripts/setup_deps.ps1`

### Step 5 — Extract Template Context

Run the extraction script:
```bash
python .claude/scripts/extract_template.py \
  "slide-workspace/templates/{template-name}/original.pptx" \
  "slide-workspace/templates/{template-name}/"
```

Or on Windows:
```bash
python .claude/scripts/extract_template.py "slide-workspace/templates/{template-name}/original.pptx" "slide-workspace/templates/{template-name}/"
```

This produces:
- `slide-workspace/templates/{template-name}/context.json` — full structured context
- `slide-workspace/templates/{template-name}/images/` — extracted image files

If exit code is non-zero, report the error and stop.

### Step 6 — Read Extracted Context

Read `context.json` and note:
- `presentation.slide_width_inches` and `slide_height_inches`
- `fonts_summary.families_used` — font names used
- `theme.color_scheme` — theme colors
- `slide_masters` — number of masters and their names
- `slide_layouts` — number of layouts and their names
- `images_manifest` — list of extracted images

List the files in `images/` directory to confirm images were extracted.

### Step 7 — Read Prompt Rules

Read the analysis prompt from `.claude/prompts/analyze-rules.md`.

Read the PptxGenJS API reference from `.claude/docs/pptxgenjs-api.md`.

### Step 8 — Generate Guideline and Sample Code

Using the extracted context.json content, images metadata, and the prompt rules from analyze-rules.md as your system instructions, generate two outputs:

**guideline.md** — A markdown document covering:
- Typography (exact font families, sizes, weights from context)
- Color palette (exact hex codes from theme and shape fills)
- Layout patterns (slide types detected, with descriptions)
- Spacing and alignment conventions
- Imagery style guidelines
- Brand/style rules

**sample_code.js** — A complete, runnable PPTXGenJS script that:
- Recreates the template structure
- Uses exact font names, colors, and positions from context
- References images via `images['filename']`
- Includes one example slide per layout type
- Follows all rules in analyze-rules.md

The JSON response format from analyze-rules.md is `{ "guideline": "...", "pptxgenjs_code": "..." }`.

### Step 9 — Save Outputs

Write the generated content to:
- `slide-workspace/templates/{template-name}/guideline.md`
- `slide-workspace/templates/{template-name}/sample_code.js`

### Step 10 — Report Summary

Report to the user:
- Template name and workspace path
- Fonts found: list from `fonts_summary.families_used`
- Theme colors: list from `theme.color_scheme`
- Masters: count and names
- Layouts: count and names
- Images extracted: count from `images_manifest`
- Files created: guideline.md, sample_code.js, context.json, images/

## Output Structure

```
slide-workspace/
  templates/
    {template-name}/
      original.pptx        ← Copy of input file
      context.json         ← Structured extraction (fonts, colors, layouts, masters)
      guideline.md         ← Design guidelines in Markdown
      sample_code.js       ← Sample PPTXGenJS code demonstrating the template
      images/              ← Extracted image files
        master0_Logo.png
        layout0_bg.jpg
        ...
```

## Error Handling

- **File not found**: Ask user to verify the path
- **Not a .pptx file**: Inform user only .pptx files are supported
- **extract_template.py exits with code 3**: PPTX parse error — file may be corrupted
- **No images directory**: Template has no embedded images (normal for simple templates)
- **LLM output not valid JSON**: Retry or ask user to try again

## Notes

- The `images` variable in generated code is pre-injected by the runner; never re-declare it
- `effective_background` on each slide is the definitive background — use it, not the raw `background` field
- Multiple slide masters are common in corporate templates — each gets its own `MASTER_N` name
