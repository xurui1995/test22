#!/usr/bin/env node
/**
 * html2ppt.js — Convert HTML slide files to an editable PowerPoint (.pptx) file.
 *
 * Strategy:
 *   1. Use headless Chromium (puppeteer-core) to render each HTML slide.
 *   2. Walk the DOM via page.evaluate() to extract text, tables, and chart
 *      containers with their exact bounding boxes (getBoundingClientRect).
 *   3. Screenshot only JS chart elements (D3.js <svg>, Chart.js <canvas>).
 *   4. Build a PPTX slide with:
 *        • Native text boxes for all text elements  (editable)
 *        • Native tables for <table> elements        (editable)
 *        • Embedded PNG images for JS charts         (pixel-perfect)
 *        • Slide background colour from CSS
 *
 * Coordinate mapping:
 *   The HTML slide-container is 1280 × 720 CSS pixels.
 *   PPTX LAYOUT_WIDE is 13.33 × 7.5 inches = 1280 × 720 px at 96 dpi.
 *   Therefore: x_inches = x_css_px / 96 (≡ x_css_px × 13.33 / 1280).
 *
 * CDN resources (D3.js, Chart.js, Font Awesome) are intercepted and served
 * from the local vendor/ directory so the converter works fully offline.
 *
 * Usage:
 *   node html2ppt.js [options] <file1.html> [file2.html ...]
 *   node html2ppt.js [options] --glob "*.html"
 *
 * Options:
 *   -o, --output <file>   Output .pptx filename  (default: output.pptx)
 *   -w, --width  <px>     Slide viewport width   (default: 1280)
 *   -h, --height <px>     Slide viewport height  (default: 720)
 *   --wait   <ms>         Extra ms to wait after page load (default: 2000)
 *   --glob   <pattern>    Glob pattern for input files
 *   --help                Show this help message
 */

'use strict';

const fs        = require('fs');
const path      = require('path');
const puppeteer = require('puppeteer-core');
const PptxGenJS = require('pptxgenjs');

// ─── coordinate constants ─────────────────────────────────────────────────────
// 96 CSS pixels = 1 inch (CSS specification).
// PPTX LAYOUT_WIDE = 13.33" × 7.5" which equals exactly 1280 × 720 CSS px.
const PX_TO_IN   = 1 / 96;          // CSS pixels → inches
const SLIDE_PX_W = 1280;            // expected slide-container width  (px)
const SLIDE_PX_H = 720;             // expected slide-container height (px)

// ─── default colour fallbacks ─────────────────────────────────────────────────
// Used when a computed CSS colour cannot be parsed (e.g. transparent or missing).
const DEFAULT_BG_COLOR         = '121212'; // near-black — matches the slide template
const DEFAULT_TEXT_COLOR       = 'ffffff'; // white  — readable on dark backgrounds
const DEFAULT_TABLE_CELL_BG    = '1e1e1e'; // very dark grey — dark-theme table fill
const DEFAULT_TABLE_BORDER_COLOR = '333333'; // subtle dark grey border

// ─── vendor asset map ─────────────────────────────────────────────────────────
const VENDOR_DIR = path.join(__dirname, 'vendor');

function buildVendorMap() {
  return [
    { match: /d3js\.org.*d3.*\.js/,                     file: path.join(VENDOR_DIR, 'd3.v7.min.js'),        mime: 'application/javascript' },
    { match: /cdn\.jsdelivr\.net.*chart\.js.*\.js/,     file: path.join(VENDOR_DIR, 'chart.3.9.1.min.js'),  mime: 'application/javascript' },
    { match: /cdnjs\.cloudflare\.com.*chart\.js.*\.js/, file: path.join(VENDOR_DIR, 'chart.3.9.1.min.js'),  mime: 'application/javascript' },
    { match: /font-awesome.*\/css\/all.*\.css/,         file: path.join(VENDOR_DIR, 'fontawesome', 'all.min.css'), mime: 'text/css' },
    { match: /fontawesome.*\/css\/all.*\.css/,          file: path.join(VENDOR_DIR, 'fontawesome', 'all.min.css'), mime: 'text/css' },
    { match: /fa-solid-900\.woff2/,                     file: path.join(VENDOR_DIR, 'fontawesome', 'webfonts', 'fa-solid-900.woff2'),    mime: 'font/woff2' },
    { match: /fa-regular-400\.woff2/,                   file: path.join(VENDOR_DIR, 'fontawesome', 'webfonts', 'fa-regular-400.woff2'),  mime: 'font/woff2' },
    { match: /fa-brands-400\.woff2/,                    file: path.join(VENDOR_DIR, 'fontawesome', 'webfonts', 'fa-brands-400.woff2'),   mime: 'font/woff2' },
    { match: /fa-v4compatibility\.woff2/,               file: path.join(VENDOR_DIR, 'fontawesome', 'webfonts', 'fa-v4compatibility.woff2'), mime: 'font/woff2' },
    { match: /fa-solid-900\.ttf/,                       file: path.join(VENDOR_DIR, 'fontawesome', 'webfonts', 'fa-solid-900.woff2'),    mime: 'font/ttf' },
    { match: /fonts\.googleapis\.com/,  body: '',  mime: 'text/css' },
    { match: /fonts\.gstatic\.com/,     body: '',  mime: 'font/woff2' },
  ];
}

// ─── helper: resolve glob ────────────────────────────────────────────────────
function expandGlob(pattern) {
  if (typeof fs.globSync === 'function') return fs.globSync(pattern);
  try { return require('glob').globSync(pattern); } catch (_) {
    return fs.existsSync(pattern) ? [pattern] : [];
  }
}

// ─── parse CLI args ───────────────────────────────────────────────────────────
function parseArgs(argv) {
  const opts = { output: 'output.pptx', width: 1280, height: 720, wait: 2000, files: [] };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--help' || a === '-?') {
      const header = fs.readFileSync(__filename, 'utf8').match(/\/\*[\s\S]*?\*\//)[0]
        .replace(/^\/\*\*?\n?|\n? *\*\/$/g, '').replace(/^ *\* ?/gm, '');
      console.log(header);
      process.exit(0);
    } else if ((a === '-o' || a === '--output') && argv[i + 1]) { opts.output = argv[++i]; }
    else if ((a === '-w' || a === '--width')  && argv[i + 1]) { opts.width  = parseInt(argv[++i], 10); }
    else if ((a === '-h' || a === '--height') && argv[i + 1]) { opts.height = parseInt(argv[++i], 10); }
    else if (a === '--wait' && argv[i + 1])                   { opts.wait   = parseInt(argv[++i], 10); }
    else if (a === '--glob' && argv[i + 1])                   { opts.files.push(...expandGlob(argv[++i])); }
    else if (!a.startsWith('-'))                               { opts.files.push(a); }
  }
  return opts;
}

// ─── set up request interception for a page ──────────────────────────────────
async function setupVendorInterception(page, vendorMap) {
  await page.setRequestInterception(true);
  page.on('request', req => {
    const url = req.url();
    if (!url.startsWith('http')) { req.continue(); return; }
    for (const entry of vendorMap) {
      if (!entry.match.test(url)) continue;
      if ('body' in entry) {
        req.respond({ status: 200, contentType: entry.mime, body: entry.body });
        return;
      }
      if (fs.existsSync(entry.file)) {
        req.respond({ status: 200, contentType: entry.mime, body: fs.readFileSync(entry.file) });
        return;
      }
      req.abort('failed');
      return;
    }
    req.abort('failed');
  });
}

// ─── DOM extraction ───────────────────────────────────────────────────────────
// Runs inside the browser context (page.evaluate).
// Walks the .slide-container DOM tree and returns:
//   • bgColor  – slide background hex colour
//   • items    – array of {type, x, y, w, h, ...} descriptors
//   • containerX/Y – viewport position of the container (for clip screenshots)
async function extractSlideData(page) {
  return page.evaluate(() => {
    let _nextId = 0;

    // Convert CSS rgb/rgba() string → 'RRGGBB' hex, or null if transparent.
    function rgbToHex(css) {
      const m = (css || '').match(/rgba?\(\s*(\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\s*\)/);
      if (!m) return null;
      if (m[4] !== undefined && parseFloat(m[4]) < 0.05) return null;
      return [m[1], m[2], m[3]].map(v => parseInt(v).toString(16).padStart(2, '0')).join('');
    }

    function relRect(el, cRect) {
      const r = el.getBoundingClientRect();
      return {
        x: Math.round(r.left - cRect.left),
        y: Math.round(r.top  - cRect.top),
        w: Math.round(r.width),
        h: Math.round(r.height),
      };
    }

    // Returns true if el has at least one block-level direct child.
    function hasBlockChild(el) {
      for (const child of el.children) {
        const d = window.getComputedStyle(child).display;
        if (['block','flex','grid','table','table-row','table-cell','list-item'].includes(d))
          return true;
      }
      return false;
    }

    function isHidden(el) {
      const s = window.getComputedStyle(el);
      return s.display === 'none' || s.visibility === 'hidden' || parseFloat(s.opacity) < 0.01;
    }

    const container = document.querySelector('.slide-container');
    if (!container) return null;

    const cRect  = container.getBoundingClientRect();
    const cStyle = window.getComputedStyle(container);
    const bgColor = rgbToHex(cStyle.backgroundColor); // null if unparseable → DEFAULT_BG_COLOR applied in buildPptxSlide
    const items   = [];

    function walk(el, depth) {
      if (depth > 20 || isHidden(el)) return;

      const tag  = el.tagName.toLowerCase();
      const rect = relRect(el, cRect);

      // Skip invisible or out-of-bounds elements
      if (rect.w < 2 || rect.h < 2) return;
      if (rect.x + rect.w < -10 || rect.y + rect.h < -10) return;
      if (rect.x > 1290 || rect.y > 730) return;

      // ── Chart.js <canvas> ──────────────────────────────────────────────────
      if (tag === 'canvas') {
        const selId = 'pptx-' + (_nextId++);
        el.setAttribute('data-pptx-sel', selId);
        items.push({ type: 'chart', ...rect, selId });
        return;
      }

      // ── D3.js <svg> (non-empty) ────────────────────────────────────────────
      if (tag === 'svg' && el.childElementCount > 0) {
        const selId = 'pptx-' + (_nextId++);
        el.setAttribute('data-pptx-sel', selId);
        items.push({ type: 'chart', ...rect, selId });
        return;
      }

      // ── <table> ────────────────────────────────────────────────────────────
      if (tag === 'table') {
        const rows = [];
        el.querySelectorAll('tr').forEach(tr => {
          const cells = [];
          tr.querySelectorAll('td, th').forEach(td => {
            const cs = window.getComputedStyle(td);
            cells.push({
              text:     (td.innerText || '').trim(),
              bold:     parseInt(cs.fontWeight) >= 600,
              color:    rgbToHex(cs.color) || 'ffffff',
              bg:       rgbToHex(cs.backgroundColor),
              align:    cs.textAlign === 'start' ? 'left' : (cs.textAlign || 'left'),
              fontSize: parseFloat(cs.fontSize) || 14,
            });
          });
          if (cells.length > 0) rows.push(cells);
        });
        if (rows.length > 0) items.push({ type: 'table', ...rect, rows });
        return;
      }

      // ── Text leaf ─────────────────────────────────────────────────────────
      // An element with rendered text and no block-level children.
      const text = (el.innerText || '').trim();
      if (text && !hasBlockChild(el)) {
        const s = window.getComputedStyle(el);
        // -webkit-text-fill-color: transparent signals CSS gradient text → default white
        let color = rgbToHex(s.color);
        if (!color || s.webkitTextFillColor === 'transparent') color = 'ffffff';
        const fontFace = (s.fontFamily || 'sans-serif').split(',')[0].replace(/['"]/g, '').trim();
        items.push({
          type:     'text',
          ...rect,
          text,
          fontSize: parseFloat(s.fontSize) || 14,
          fontFace: fontFace || 'Arial',
          bold:     parseInt(s.fontWeight) >= 600,
          italic:   s.fontStyle === 'italic',
          color,
          align:    s.textAlign === 'start' ? 'left' : (s.textAlign || 'left'),
        });
        return;
      }

      // Recurse into children
      for (const child of el.children) {
        walk(child, depth + 1);
      }
    }

    for (const child of container.children) {
      walk(child, 0);
    }

    return {
      bgColor,
      items,
      containerX: Math.round(cRect.left),
      containerY: Math.round(cRect.top),
    };
  });
}

// ─── screenshot chart elements ───────────────────────────────────────────────
// For each chart item, take a screenshot of only that element (canvas/svg)
// using the data-pptx-sel attribute set by extractSlideData().
async function screenshotChartElements(page, slideData) {
  if (!slideData || !slideData.items) return;
  for (const item of slideData.items) {
    if (item.type !== 'chart' || !item.selId) continue;
    try {
      const handle = await page.$(`[data-pptx-sel="${item.selId}"]`);
      if (handle) {
        const buffer = await handle.screenshot({ type: 'png' });
        item.imageData = buffer.toString('base64');
        await handle.dispose();
      }
    } catch (err) {
      console.warn(`\n  ⚠ Chart screenshot failed (${item.selId}): ${err.message}`);
    }
  }
}

// ─── build one PPTX slide from extracted data ─────────────────────────────────
// Coordinate conversion: CSS px ÷ 96 = inches.
// Font size conversion: CSS px × 0.75 = points  (1 pt = 1/72 in; 1 px = 1/96 in).
function buildPptxSlide(pptx, slideData) {
  const slide = pptx.addSlide();

  if (slideData && slideData.bgColor) {
    slide.background = { color: slideData.bgColor };
  } else {
    slide.background = { color: DEFAULT_BG_COLOR };
  }

  if (!slideData || !slideData.items) return slide;

  const slideW = SLIDE_PX_W * PX_TO_IN;  // 13.33 in
  const slideH = SLIDE_PX_H * PX_TO_IN;  // 7.5  in

  for (const item of slideData.items) {
    // Convert coordinates from CSS pixels to inches
    let x = item.x * PX_TO_IN;
    let y = item.y * PX_TO_IN;
    let w = item.w * PX_TO_IN;
    let h = item.h * PX_TO_IN;

    // Clamp to slide bounds
    if (x < 0) { w += x; x = 0; }
    if (y < 0) { h += y; y = 0; }
    if (x + w > slideW) w = slideW - x;
    if (y + h > slideH) h = slideH - y;
    if (w < 0.05 || h < 0.05) continue;

    // ── Chart image ────────────────────────────────────────────────────────
    if (item.type === 'chart') {
      if (!item.imageData) continue;
      slide.addImage({ data: `data:image/png;base64,${item.imageData}`, x, y, w, h });

    // ── Text box ──────────────────────────────────────────────────────────
    } else if (item.type === 'text') {
      if (!item.text) continue;
      const fontSize = Math.max(6, Math.round(item.fontSize * 0.75)); // px → pt
      try {
        slide.addText(item.text, {
          x, y, w, h,
          fontSize,
          fontFace: item.fontFace || 'Arial',
          bold:     item.bold   || false,
          italic:   item.italic || false,
          color:    item.color  || DEFAULT_TEXT_COLOR,
          align:    item.align  || 'left',
          valign:   'top',
          wrap:     true,
        });
      } catch (e) {
        console.warn(`\n  ⚠ Text element skipped: ${e.message}`);
      }

    // ── Table ─────────────────────────────────────────────────────────────
    } else if (item.type === 'table') {
      if (!item.rows || item.rows.length === 0) continue;
      const tableData = item.rows.map(row =>
        row.map(cell => ({
          text: cell.text || '',
          options: {
            bold:     cell.bold || false,
            color:    cell.color || DEFAULT_TEXT_COLOR,
            fill:     { color: cell.bg || DEFAULT_TABLE_CELL_BG },
            align:    cell.align || 'left',
            fontSize: Math.max(6, Math.round((cell.fontSize || 14) * 0.75)),
          },
        }))
      );
      try {
        slide.addTable(tableData, {
          x, y, w, h,
          border: { pt: 0.5, color: DEFAULT_TABLE_BORDER_COLOR },
          color:  DEFAULT_TEXT_COLOR,
        });
      } catch (e) {
        console.warn(`\n  ⚠ Table element skipped: ${e.message}`);
      }
    }
  }

  return slide;
}

// ─── main ─────────────────────────────────────────────────────────────────────
async function main() {
  const opts = parseArgs(process.argv.slice(2));

  if (opts.files.length === 0) {
    console.error('Error: no input files specified.\nRun with --help for usage.');
    process.exit(1);
  }

  const files = opts.files.map(f => path.resolve(f)).filter(f => {
    if (!fs.existsSync(f)) { console.warn(`Warning: file not found — ${f}`); return false; }
    return true;
  });

  if (files.length === 0) {
    console.error('Error: none of the specified files exist.');
    process.exit(1);
  }

  console.log(`Converting ${files.length} HTML file(s) → ${opts.output}`);

  // ── Launch headless Chromium ──────────────────────────────────────────────
  const chromePaths = [
    '/usr/bin/chromium-browser',
    '/usr/bin/chromium',
    '/usr/bin/google-chrome',
    '/usr/bin/google-chrome-stable',
  ];
  const executablePath = chromePaths.find(p => fs.existsSync(p));
  if (!executablePath) {
    console.error('Error: Chromium/Chrome not found. Install chromium-browser.');
    process.exit(1);
  }

  const browser = await puppeteer.launch({
    executablePath,
    headless: 'new',
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--disable-gpu',
      '--lang=zh-CN',
      '--font-render-hinting=none',
    ],
  });

  const vendorMap = buildVendorMap();
  const slides    = [];

  for (const file of files) {
    process.stdout.write(`  • Processing: ${path.basename(file)} ... `);
    const page = await browser.newPage();

    // 2× device pixel ratio → crisp chart screenshots
    await page.setViewport({ width: opts.width, height: opts.height, deviceScaleFactor: 2 });

    // Intercept CDN requests → serve local vendor files
    await setupVendorInterception(page, vendorMap);

    await page.goto(`file://${file}`, { waitUntil: 'domcontentloaded', timeout: 30000 });

    // Wait for JS charts (D3/Chart.js) to finish rendering
    await new Promise(r => setTimeout(r, opts.wait));

    // Extract DOM structure (text, tables, chart positions)
    const slideData = await extractSlideData(page);

    // Screenshot only the JS chart elements
    await screenshotChartElements(page, slideData);

    slides.push({ file, slideData });
    await page.close();
    console.log('done');
  }

  await browser.close();

  // ── Build PPTX ────────────────────────────────────────────────────────────
  console.log('Assembling PPTX…');
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';   // 13.33 × 7.5 in (standard 16:9)

  for (const { slideData } of slides) {
    buildPptxSlide(pptx, slideData);
  }

  await pptx.writeFile({ fileName: opts.output });
  console.log(`✓ Saved: ${opts.output}  (${slides.length} slide${slides.length !== 1 ? 's' : ''})`);
}

main().catch(err => {
  console.error('Fatal error:', err.message || err);
  process.exit(1);
});
