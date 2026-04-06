# Shared PPTXGenJS Slide Generation Rules

These rules apply to ALL slide code generation tasks.

- Generate a presentation slide with rich layout (sections, table, chart, cards, icons), not bullet lists. Ensure visual hierarchy and varied components.
- Use the EXACT font names, hex color values, and positions (in inches) from the Template Structure JSON.
- For embedded images, reference via `images['filename.ext']` (pre-defined variable; values are absolute file paths).
- Do NOT re-declare the `images` variable — it is already defined by the runtime.
- Master slide images (logos, decorations listed in slide_masters.images) should appear on every relevant slide.
- Background images should be set via `slide.background = { path: images['...'] }`.
- CRITICAL — Background color:
  * Each slide has a `background_color` field — use it directly as `slide.background = { color: '<background_color>' }`.
  * NEVER use addShape/addRect to simulate backgrounds. NEVER omit slide.background.
  * NEVER set gray/dark/black backgrounds unless `background_color` explicitly shows it.
- ALWAYS use `pptx.shapes.SHAPE_NAME` (UPPERCASE): RECTANGLE, ROUNDED_RECTANGLE, LINE, OVAL, ISOSCELES_TRIANGLE, DIAMOND, etc.
  * There is NO "TRIANGLE" — use ISOSCELES_TRIANGLE.
  * NEVER use string literals like 'rect' or undefined variables.
- Shape fill transparency: when template has `fill.transparency`, include it: `fill: { color: '...', transparency: N }`.
- addText API:
  * Single text: `slide.addText('Hello', { x:1, y:1, w:8, h:1, fontSize:24 })`
  * Multi-run: `slide.addText([{ text: 'Bold', options: { bold: true } }, { text: ' Normal' }], { ... })`
  * NEVER pass an array of plain strings.
- CRITICAL — Text vertical alignment (valign):
  * You MUST explicitly set `valign: 'top'` on EVERY addText call unless the template context specifies otherwise.
  * "middle" → `valign: 'mid'`, "bottom" → `valign: 'bot'`, "top" or absent → `valign: 'top'`.
- IMPORTANT — Layout images: Look up `slide_layouts[layout_index].images` and add ALL images from that layout to the slide. This is in ADDITION to the slide's own images. NEVER omit layout images. Define a helper function (e.g. `addLayout0Images(slide)`) and call it as the FIRST statement on every slide using that layout.
- IMPORTANT — Table borders: NEVER use `color: 'FFFFFF'` on white backgrounds. Default to 'BFBFBF'.
- IMPORTANT — Table row arrays: Every row passed to `addTable` must be a FLAT array of cell objects.
  * CORRECT:   `Array(9).fill({ text: '' })` — a row of 9 cells
  * WRONG:    `[Array(9).fill({ text: '' })]` — a row with 1 cell that is itself an array
- Add comments explaining each slide.
- CRITICAL — Prevent text/element overlap:
  * Allocate enough `h`: `h ≈ (fontSize / 72) * lineCount * 1.4` inches minimum.
  * Next element's `y` must be ≥ previous element's `y + h` + 0.08 inch gap.
  * For repeating item groups, use a fixed `itemH` covering ALL lines. Compute y as `startY + index * itemH`.
  * Example — 3 items, each with title (fontSize 14, h=0.3) + detail (fontSize 12, h=0.25), gap 0.1:
    ```js
    const startY = 1.5, itemH = 0.7;
    items.forEach((item, i) => {
      const y = startY + i * itemH;
      slide.addText(item.title, { x: 1, y, w: 9, h: 0.3, fontSize: 14, bold: true, valign: 'top' });
      slide.addText(item.detail, { x: 1, y: y + 0.35, w: 9, h: 0.25, fontSize: 12, valign: 'top' });
    });
    ```
- CRITICAL — Prevent title text wrapping/overflow:
  * For fontSize ≥ 24, ALWAYS add `fit: 'shrink'` as a safety net.
  * Example: `slide.addText('Long Title', { x: 0.5, y: 1, w: 10, h: 0.8, fontSize: 36, bold: true, fit: 'shrink', valign: 'top' })`
