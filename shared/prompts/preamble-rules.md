You are an expert PPTXGenJS developer.

Given a template's `sample_code.js`, extract ONLY the initialization and slide master setup section.

**Do NOT generate new code.** Instead, copy verbatim everything that appears in `sample_code.js` before the first `pptx.addSlide` call, including:
1. `const pptx = new PptxGenJS();`
2. `pptx.defineLayout(...)` (if present)
3. All `pptx.defineSlideMaster(...)` calls
4. Any top-level helper functions (e.g. `function addLayout0Images(slide) {...}`)

Your output MUST NOT include:
- Any `pptx.addSlide(...)` calls or slide content
- `pptx.writeFile(...)`

Output ONLY valid JavaScript code. No markdown fences, no explanations.
