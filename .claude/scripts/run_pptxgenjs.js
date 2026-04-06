"use strict";
/**
 * run_pptxgenjs.js — Standalone CLI for executing PPTXGenJS code.
 *
 * Usage:
 *   node run_pptxgenjs.js <code_file> <images_dir> <output_file>
 *
 * Arguments:
 *   code_file    — Path to a .js file containing PPTXGenJS code
 *   images_dir   — Directory containing image files (all files become available
 *                  via the `images` map as { filename: absolutePath })
 *   output_file  — Where to write the generated .pptx file
 *
 * Exit codes:
 *   0 — Success
 *   1 — Missing or invalid arguments
 *   2 — Code file not found
 *   3 — Execution timeout or runtime error
 *   4 — No .pptx file produced
 */

const { v4: uuidv4 } = require("uuid");
const fs = require("fs");
const path = require("path");
const { execSync } = require("child_process");
const os = require("os");
const JSZip = require("jszip");

// ── PPTX Bug Fixes ────────────────────────────────────────────────────────────

/**
 * Post-process PptxGenJS output to fix known bugs that cause PowerPoint
 * to show a "needs repair" dialog.
 *
 * Fixes applied:
 * 1. Phantom [Content_Types].xml — PptxGenJS registers one slideMaster
 *    Override per slide, but only creates one actual slideMaster file per
 *    defineSlideMaster() call. Extra entries point to non-existent files.
 * 2a. Empty <a:ln></a:ln> elements — replaced with <a:ln><a:noFill/></a:ln>
 *     for explicit "no border" semantics.
 * 2b. Empty <a:tblPr/> — PptxGenJS omits the required <a:tableStyleId>
 *     child, causing PowerPoint to flag the file for repair. We inject the
 *     built-in "No Style, No Grid" table style GUID.
 * 3. Broken .rels references — remove Relationship entries pointing to
 *    non-existent targets.
 */
async function _fixPptxBugs(pptxBuffer, runId) {
  try {
    const zip = await JSZip.loadAsync(pptxBuffer);
    const existingFiles = new Set(Object.keys(zip.files));
    let totalFixes = 0;

    // --- Fix 1: Phantom [Content_Types].xml entries ---
    const ctFile = zip.file("[Content_Types].xml");
    if (ctFile) {
      let ct = await ctFile.async("string");
      let phantomCount = 0;
      ct = ct.replace(/<Override\s[^>]+>/g, (match) => {
        const m = match.match(/PartName="([^"]*)"/);
        if (!m) return match;
        const zipPath = m[1].replace(/^\//, "");
        if (!existingFiles.has(zipPath)) {
          phantomCount++;
          return "";
        }
        return match;
      });
      if (phantomCount > 0) {
        console.log(`[RUN ${runId}] Removed ${phantomCount} phantom Content_Types entries`);
        zip.file("[Content_Types].xml", ct);
        totalFixes += phantomCount;
      }
    }

    // --- Fix 2: Patch slide XML issues ---
    for (const filePath of Object.keys(zip.files)) {
      if (!filePath.endsWith(".xml") || !filePath.startsWith("ppt/")) continue;
      const entry = zip.file(filePath);
      if (!entry) continue;
      let xml = await entry.async("string");
      const origXml = xml;
      // Fix 2a: Empty <a:ln> → explicit noFill
      xml = xml.replace(/<a:ln>\s*<\/a:ln>/g, "<a:ln><a:noFill/></a:ln>");
      xml = xml.replace(/<a:ln\/>/g, "<a:ln><a:noFill/></a:ln>");
      // Fix 2b: Empty <a:tblPr/> → inject "No Style, No Grid" table style
      const tblStyleTag = '<a:tableStyleId>{2D5ABB26-0587-4C30-8999-92F81FD0307C}</a:tableStyleId>';
      xml = xml.replace(/<a:tblPr\/>/g, `<a:tblPr firstRow="1" bandRow="1">${tblStyleTag}</a:tblPr>`);
      xml = xml.replace(/<a:tblPr>\s*<\/a:tblPr>/g, `<a:tblPr firstRow="1" bandRow="1">${tblStyleTag}</a:tblPr>`);
      if (xml !== origXml) {
        zip.file(filePath, xml);
        totalFixes++;
      }
    }

    // --- Fix 3: Validate and fix rels referencing non-existent targets ---
    for (const filePath of Object.keys(zip.files)) {
      if (!filePath.endsWith(".rels")) continue;
      const entry = zip.file(filePath);
      if (!entry) continue;
      let xml = await entry.async("string");
      const origXml = xml;

      xml = xml.replace(/<Relationship\s[^>]+>/g, (match) => {
        const tm = match.match(/Target="([^"]*)"/);
        if (!tm) return match;
        const target = tm[1];
        if (target.startsWith("http")) return match;

        const relsDir = filePath.replace(/[^/]+$/, "");
        const parentDir = relsDir.replace(/_rels\/$/, "");
        // Absolute targets (starting with /) resolve from ZIP root
        const raw = target.startsWith("/") ? target : (parentDir + target);
        const parts = raw.split("/");
        const resolved = [];
        for (const p of parts) {
          if (p === "..") resolved.pop();
          else if (p !== "." && p !== "") resolved.push(p);
        }
        const resolvedPath = resolved.join("/");

        if (!existingFiles.has(resolvedPath)) {
          console.log(`[RUN ${runId}] Removed broken rel: ${filePath} -> ${target} (${resolvedPath})`);
          totalFixes++;
          return "";
        }
        return match;
      });

      if (xml !== origXml) {
        zip.file(filePath, xml);
      }
    }

    if (totalFixes === 0) return pptxBuffer;

    console.log(`[RUN ${runId}] Applied ${totalFixes} PPTX fixes total`);
    return await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  } catch (err) {
    console.error(`[RUN ${runId}] PPTX fix failed, using original:`, err.message);
    return pptxBuffer;
  }
}

// ── Code Sanitizer ────────────────────────────────────────────────────────────

function sanitizeCode(code) {
  // Remove LLM re-declarations of `images` that cause TDZ errors.
  return code
    .replace(/^\s*(?:const|let|var)\s+images\s*=\s*[^;]*;?\s*$/gm, '// [wrapper] images re-declaration removed')
    .replace(/^\s*(?:const|let|var)\s+images\s*=\s*global\.images[^;]*;?\s*$/gm, '// [wrapper] images re-declaration removed');
}

// ── Wrapper Builder ───────────────────────────────────────────────────────────

function buildWrapper(code, imageMap) {
  const sanitized = sanitizeCode(code);
  const scriptDir = path.dirname(require.resolve("pptxgenjs"));
  return `
const _OrigPptxGenJS = require("pptxgenjs");
const _path = require("path");
const fs = require("fs");

// Image map: filename -> absolute path on disk
// Use var so LLM code can re-declare 'images' without TDZ errors
var images = ${JSON.stringify(imageMap)};

// Pre-declare so LLM code can do:  __pptxInstance = new PptxGenJS();
var __pptxInstance;

// Build shape lookup with common aliases so LLM-generated names resolve correctly
const _rawShapes = new _OrigPptxGenJS().shapes;
const _shapeAliases = {
  TRIANGLE: _rawShapes.ISOSCELES_TRIANGLE,
  RECT: _rawShapes.RECTANGLE,
  ROUND_RECT: _rawShapes.ROUNDED_RECTANGLE,
  ROUNDED_RECT: _rawShapes.ROUNDED_RECTANGLE,
  ROUND_RECTANGLE: _rawShapes.ROUNDED_RECTANGLE,
  CIRCLE: _rawShapes.OVAL,
  ELLIPSE: _rawShapes.OVAL,
  ARROW: _rawShapes.RIGHT_ARROW,
  ARROW_RIGHT: _rawShapes.RIGHT_ARROW,
  ARROW_LEFT: _rawShapes.LEFT_ARROW,
  ARROW_UP: _rawShapes.UP_ARROW,
  ARROW_DOWN: _rawShapes.DOWN_ARROW,
  FREEFORM: _rawShapes.CUSTOM_GEOMETRY,
  // camelCase / lowercase aliases LLMs commonly generate
  rect: _rawShapes.RECTANGLE,
  rectangle: _rawShapes.RECTANGLE,
  roundRect: _rawShapes.ROUNDED_RECTANGLE,
  roundedRect: _rawShapes.ROUNDED_RECTANGLE,
  roundedRectangle: _rawShapes.ROUNDED_RECTANGLE,
  line: _rawShapes.LINE,
  oval: _rawShapes.OVAL,
  circle: _rawShapes.OVAL,
  triangle: _rawShapes.ISOSCELES_TRIANGLE,
  diamond: _rawShapes.DIAMOND,
  chevron: _rawShapes.CHEVRON,
  trapezoid: _rawShapes.TRAPEZOID,
  parallelogram: _rawShapes.PARALLELOGRAM,
  rightArrow: _rawShapes.RIGHT_ARROW,
  leftArrow: _rawShapes.LEFT_ARROW,
  upArrow: _rawShapes.UP_ARROW,
  downArrow: _rawShapes.DOWN_ARROW,
};
const _shapesProxy = new Proxy(_rawShapes, {
  get(target, prop) {
    if (prop in target) return target[prop];
    if (prop in _shapeAliases) return _shapeAliases[prop];
    const upper = String(prop).toUpperCase();
    if (upper in target) return target[upper];
    if (upper in _shapeAliases) return _shapeAliases[upper];
    return undefined;
  }
});

// Resolve a shape argument: handles shape objects, string names, and aliases
function _resolveShape(shapeArg) {
  if (shapeArg && typeof shapeArg === 'object') return shapeArg;
  if (typeof shapeArg === 'string') {
    const resolved = _shapesProxy[shapeArg];
    if (resolved) return resolved;
  }
  return shapeArg;
}

// Auto-capture last pptx instance and track writeFile calls
let __lastPptxInstance;
let __writeFileCalled = false;

// Normalize valign values to the set PptxGenJS maps correctly.
const _valignMap = { mid: 'middle', center: 'middle', ctr: 'middle' };
function _normalizeValign(opts) {
  if (!opts || typeof opts !== 'object') return opts;
  if (opts.valign && _valignMap[opts.valign]) {
    opts.valign = _valignMap[opts.valign];
  }
  return opts;
}

// Deep-normalize valign in table cell options
function _normalizeTableRows(rows) {
  if (!Array.isArray(rows)) return rows;
  return rows.map(row => {
    if (!Array.isArray(row)) return row;
    return row.map(cell => {
      if (cell && typeof cell === 'object' && cell.options) {
        _normalizeValign(cell.options);
      }
      return cell;
    });
  });
}

// Convert line master objects to rect (no fill) — PptxGenJS generates malformed
// connector XML for line shapes in slide layouts
function _sanitizeMasterObjects(objects) {
  if (!Array.isArray(objects)) return objects;
  return objects.map(obj => {
    if (obj.line && !obj.rect) {
      const ln = obj.line;
      return {
        rect: {
          x: ln.x, y: ln.y, w: ln.w, h: ln.h,
          fill: { type: 'none' },
          line: {
            color: (ln.line && ln.line.color) || '000000',
            width: (ln.line && (ln.line.width || ln.line.pt)) || 1,
            dashType: ln.line && ln.line.dashType,
          },
          rectRadius: 0,
        }
      };
    }
    return obj;
  });
}

function PptxGenJS() {
  const instance = new _OrigPptxGenJS();

  // Expose shapes via all common property names LLMs use
  for (const alias of ['shapes', 'ShapeType', 'ShapeTypes', 'Shapes', 'shapeType', 'shapeTypes']) {
    Object.defineProperty(instance, alias, { get: () => _shapesProxy, configurable: true });
  }

  // Patch defineSlideMaster to sanitize objects (line → rect)
  const origDefineSlideMaster = instance.defineSlideMaster.bind(instance);
  instance.defineSlideMaster = function(opts) {
    if (opts && opts.objects) {
      opts.objects = _sanitizeMasterObjects(opts.objects);
    }
    return origDefineSlideMaster(opts);
  };

  // Patch addSlide to return slides with fixed addShape and addText
  const origAddSlide = instance.addSlide.bind(instance);
  instance.addSlide = function(...args) {
    const slide = origAddSlide(...args);

    const origAddShape = slide.addShape.bind(slide);
    slide.addShape = function(shapeArg, opts) {
      return origAddShape(_resolveShape(shapeArg), opts);
    };

    const origAddText = slide.addText.bind(slide);
    slide.addText = function(textArg, opts) {
      if (Array.isArray(textArg)) {
        textArg = textArg.map(item => {
          if (typeof item === 'string') return { text: item };
          return item;
        });
      }
      return origAddText(textArg, _normalizeValign(opts));
    };

    const origAddTable = slide.addTable.bind(slide);
    slide.addTable = function(rows, opts) {
      return origAddTable(_normalizeTableRows(rows), _normalizeValign(opts));
    };

    return slide;
  };

  // Patch writeFile: use write('nodebuffer') + fs.writeFileSync for reliability
  const origWrite = instance.write.bind(instance);
  instance.writeFile = async function(opts) {
    const fileName = (opts && opts.fileName) || 'output.pptx';
    console.log('[wrapper] writeFile intercepted, fileName:', fileName);
    try {
      const buf = await origWrite({ outputType: 'nodebuffer' });
      const outPath = _path.resolve(process.cwd(), fileName);
      fs.writeFileSync(outPath, buf);
      console.log('[wrapper] Wrote', buf.length, 'bytes to', outPath);
      __writeFileCalled = true;
    } catch (e) {
      console.error('[wrapper] writeFile error:', e);
      throw e;
    }
    return fileName;
  };

  __lastPptxInstance = instance;
  __pptxInstance = instance;

  return instance;
}

// Global shape aliases for LLM code that uses bare variables
const shapes = _shapesProxy;
const ShapeType = _shapesProxy;
const ShapeTypes = _shapesProxy;
const Shapes = _shapesProxy;

// Hijack require: any require('pptxgenjs') in LLM code returns our patched constructor
const _origRequire = module.require.bind(module);
module.require = function(mod) {
  if (mod === 'pptxgenjs' || mod === 'pptxgenjs/dist/pptxgen.cjs') return PptxGenJS;
  return _origRequire(mod);
};
var pptxgen = PptxGenJS;
var PptxGen = PptxGenJS;
var pptxGen = PptxGenJS;
var PPTXGENJS = PptxGenJS;

(async () => {
  ${sanitized}

  // Fallback: ensure a .pptx file is written
  const _inst = __pptxInstance || __lastPptxInstance;
  const _hasPptx = fs.readdirSync(process.cwd()).some(f => f.endsWith('.pptx'));
  console.log('[wrapper] Fallback check: hasPptx=' + _hasPptx + ' writeFileCalled=' + __writeFileCalled + ' hasInstance=' + !!_inst);
  if (!_hasPptx && !__writeFileCalled && _inst) {
    console.log('[wrapper] Fallback: writing output.pptx via patched writeFile...');
    await _inst.writeFile({ fileName: "output.pptx" });
  }
})().then(() => {
  process.exit(0);
}).catch((err) => {
  console.error('[wrapper] Fatal error:', err);
  process.exit(1);
});
`;
}

// ── CLI Main ──────────────────────────────────────────────────────────────────

async function main() {
  const args = process.argv.slice(2);
  if (args.length !== 3) {
    console.error("Usage: node run_pptxgenjs.js <code_file> <images_dir> <output_file>");
    process.exit(1);
  }

  const [codeFile, imagesDir, outputFile] = args;

  if (!fs.existsSync(codeFile)) {
    console.error(`ERROR: Code file not found: ${codeFile}`);
    process.exit(2);
  }

  const code = fs.readFileSync(codeFile, "utf-8");

  // Build imageMap from images directory
  const imageMap = {};
  if (fs.existsSync(imagesDir) && fs.statSync(imagesDir).isDirectory()) {
    for (const file of fs.readdirSync(imagesDir)) {
      const filePath = path.resolve(imagesDir, file);
      if (fs.statSync(filePath).isFile()) {
        imageMap[file] = filePath;
      }
    }
  }

  const runId = uuidv4();
  const workDir = path.join(os.tmpdir(), "pptxgen", runId);
  fs.mkdirSync(workDir, { recursive: true });

  const scriptPath = path.join(workDir, "script.js");

  try {
    const wrapper = buildWrapper(code, imageMap);
    fs.writeFileSync(scriptPath, wrapper);

    console.log(`[RUN ${runId}] Executing script...`);

    let execResult;
    try {
      execResult = execSync(`node "${scriptPath}"`, {
        cwd: workDir,
        env: { ...process.env, NODE_PATH: path.join(__dirname, "node_modules") },
        timeout: 60000,
        stdio: ["ignore", "pipe", "pipe"],
      });
      console.log(`[RUN ${runId}] stdout:`, execResult?.toString() || "(empty)");
    } catch (execErr) {
      const stderr = execErr.stderr ? execErr.stderr.toString() : "";
      const stdout = execErr.stdout ? execErr.stdout.toString() : "";
      console.error(`[RUN ${runId}] Script FAILED`);
      console.error(`[RUN ${runId}] stderr:`, stderr);
      console.error(`[RUN ${runId}] stdout:`, stdout);
      process.exit(3);
    }

    const filesInWorkDir = fs.readdirSync(workDir);
    const pptxFile = filesInWorkDir.find((f) => f.endsWith(".pptx"));
    if (!pptxFile) {
      console.error(`[RUN ${runId}] No .pptx file produced. Files:`, filesInWorkDir);
      process.exit(4);
    }

    const rawBuffer = fs.readFileSync(path.join(workDir, pptxFile));
    const pptxBuffer = await _fixPptxBugs(rawBuffer, runId);

    const outputDir = path.dirname(path.resolve(outputFile));
    fs.mkdirSync(outputDir, { recursive: true });
    fs.writeFileSync(path.resolve(outputFile), pptxBuffer);

    console.log(`[RUN ${runId}] Success! Output: ${outputFile} (${pptxBuffer.length} bytes)`);
  } finally {
    try {
      fs.rmSync(workDir, { recursive: true, force: true });
    } catch (_) {}
  }
}

main().catch((err) => {
  console.error("Unexpected error:", err);
  process.exit(3);
});
