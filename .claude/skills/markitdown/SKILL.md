---
name: markitdown
description: Convert documents (Excel, PDF, Word, PPTX, HTML, CSV, images...) to Markdown using markitdown, ready as input for slide-generate.
triggers:
  - convert to markdown
  - convert file to markdown
  - markitdown
  - extract content from file
  - convert pdf to markdown
  - convert excel to markdown
  - convert word to markdown
  - convert docx to markdown
  - convert pptx to markdown
  - /markitdown
compatibility: Requires Python 3.10+ with markitdown installed. Run shared/scripts/setup_deps.sh (Unix) or shared/scripts/setup_deps.ps1 (Windows) to install dependencies.
metadata:
  author: slide-deck-utils
  version: 1.0.0
---

# Skill: markitdown

Convert documents of various formats into clean Markdown text using the [markitdown](https://github.com/microsoft/markitdown) library. The resulting Markdown can be used as content input for `slide-generate`.

**Supported formats**: PDF, Word (.docx), Excel (.xlsx/.xls), PowerPoint (.pptx), HTML, CSV, JSON, XML, images (JPEG/PNG/GIF/WebP), audio (MP3/WAV — requires optional deps), ZIP archives, EPub, and more.

## Workflow

### Step 1 — Accept Input

Ask the user for the input file path if not provided. Confirm the file exists.

Optionally ask:
- **Output name** (optional): name for the source workspace (default: derived from filename)

### Step 2 — Derive Source Name

Derive a filesystem-safe source name from the filename:
- Take the filename without extension (e.g., `Q4-Sales-Report.xlsx` → `q4-sales-report`)
- Replace spaces and special characters with hyphens
- Convert to lowercase

### Step 3 — Create Workspace

Create the source workspace directory:
```
slide-workspace/sources/{source-name}/
```

### Step 4 — Check Dependencies

Check that markitdown is available:
```bash
python -c "import markitdown" 2>/dev/null || python3 -c "import markitdown" 2>/dev/null
```

If the check fails, run the appropriate setup script:
- Unix/macOS: `bash shared/scripts/setup_deps.sh`
- Windows: `powershell -ExecutionPolicy Bypass -File shared/scripts/setup_deps.ps1`

### Step 5 — Run Conversion

Run the conversion script:
```bash
python shared/scripts/convert_to_markdown.py \
  "{input_file_path}" \
  "slide-workspace/sources/{source-name}/content.md"
```

On Windows (single line):
```bash
python shared/scripts/convert_to_markdown.py "{input_file_path}" "slide-workspace/sources/{source-name}/content.md"
```

If the exit code is non-zero, report the error and stop:
- Exit code `2` — input file not found
- Exit code `3` — conversion error (unsupported format or corrupted file)

### Step 6 — Read and Summarize Output

Read the first 100 lines of `slide-workspace/sources/{source-name}/content.md`.

Provide the user with a brief summary:
- File type converted
- Approximate content length (lines / characters)
- Key sections or headings detected in the Markdown
- Output file path

### Step 7 — Suggest Next Steps

Suggest the user can now run `/slide-generate` using this Markdown as the topic/content input:

> The converted content is saved at `slide-workspace/sources/{source-name}/content.md`.
> You can use it as input for `/slide-generate` — just reference this file or paste relevant sections as the presentation topic.

## Output Structure

```
slide-workspace/
  sources/
    {source-name}/
      content.md     ← Converted Markdown content
```

## Error Handling

- **File not found**: Ask user to verify the path
- **Unsupported format**: markitdown supports most common formats; report the extension and suggest converting to a supported format
- **Exit code 3 (conversion error)**: File may be password-protected, corrupted, or require optional dependencies (e.g., for audio transcription, `markitdown[audio]` is needed)
- **Empty output**: File may contain only images or non-extractable content; inform the user

## Notes

- markitdown preserves tables from Excel/Word as Markdown tables — these work well as slide content
- For PPTX files, markitdown extracts text from all slides including speaker notes
- For PDF files, text extraction quality depends on whether the PDF is text-based (not scanned images)
- Scanned PDFs (image-only) will produce minimal output; consider using an OCR tool first
- The output `content.md` can be directly referenced when prompting `/slide-generate`
