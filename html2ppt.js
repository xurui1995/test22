#!/usr/bin/env node
/**
 * html2ppt.js — Convert HTML slide files to a PowerPoint (.pptx) file.
 *
 * Strategy: Use headless Chromium (puppeteer-core) to render each HTML slide
 * at 1280×720 pixels, capture a full-quality PNG screenshot, then assemble
 * all screenshots into a PPTX file with PptxGenJS.  This approach preserves
 * pixel-perfect fidelity for any CSS / D3.js / Chart.js content.
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

// ─── vendor asset map ─────────────────────────────────────────────────────────
// Maps CDN URL substrings → local file paths (relative to this script's dir).
// Any CDN URL that contains a key is served from the corresponding local file.
const VENDOR_DIR = path.join(__dirname, 'vendor');

function buildVendorMap() {
  return [
    // D3.js  (all versions/variants → our bundled v7)
    { match: /d3js\.org.*d3.*\.js/,                     file: path.join(VENDOR_DIR, 'd3.v7.min.js'),        mime: 'application/javascript' },
    // Chart.js
    { match: /cdn\.jsdelivr\.net.*chart\.js.*\.js/,     file: path.join(VENDOR_DIR, 'chart.3.9.1.min.js'),  mime: 'application/javascript' },
    { match: /cdnjs\.cloudflare\.com.*chart\.js.*\.js/, file: path.join(VENDOR_DIR, 'chart.3.9.1.min.js'),  mime: 'application/javascript' },
    // Font Awesome CSS
    { match: /font-awesome.*\/css\/all.*\.css/,         file: path.join(VENDOR_DIR, 'fontawesome', 'all.min.css'), mime: 'text/css' },
    { match: /fontawesome.*\/css\/all.*\.css/,          file: path.join(VENDOR_DIR, 'fontawesome', 'all.min.css'), mime: 'text/css' },
    // Font Awesome webfonts
    { match: /fa-solid-900\.woff2/,                     file: path.join(VENDOR_DIR, 'fontawesome', 'webfonts', 'fa-solid-900.woff2'),    mime: 'font/woff2' },
    { match: /fa-regular-400\.woff2/,                   file: path.join(VENDOR_DIR, 'fontawesome', 'webfonts', 'fa-regular-400.woff2'),  mime: 'font/woff2' },
    { match: /fa-brands-400\.woff2/,                    file: path.join(VENDOR_DIR, 'fontawesome', 'webfonts', 'fa-brands-400.woff2'),   mime: 'font/woff2' },
    { match: /fa-v4compatibility\.woff2/,               file: path.join(VENDOR_DIR, 'fontawesome', 'webfonts', 'fa-v4compatibility.woff2'), mime: 'font/woff2' },
    { match: /fa-solid-900\.ttf/,                       file: path.join(VENDOR_DIR, 'fontawesome', 'webfonts', 'fa-solid-900.woff2'),    mime: 'font/ttf' },
    // Google Fonts — return empty CSS so the browser falls back to system fonts
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
    // Only intercept external http(s) resources
    if (!url.startsWith('http')) { req.continue(); return; }

    for (const entry of vendorMap) {
      if (!entry.match.test(url)) continue;

      // Serve from a fixed body string
      if ('body' in entry) {
        req.respond({ status: 200, contentType: entry.mime, body: entry.body });
        return;
      }

      // Serve from local file
      if (fs.existsSync(entry.file)) {
        req.respond({
          status: 200,
          contentType: entry.mime,
          body: fs.readFileSync(entry.file),
        });
        return;
      }

      // File not found locally — abort so we don't hang
      req.abort('failed');
      return;
    }

    // Unknown external URL — abort (offline mode)
    req.abort('failed');
  });
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

  // ── Capture screenshots ───────────────────────────────────────────────────
  const screenshots = [];

  for (const file of files) {
    process.stdout.write(`  • Rendering: ${path.basename(file)} ... `);
    const page = await browser.newPage();

    // 2× device pixel ratio → crisp hi-res screenshots
    await page.setViewport({ width: opts.width, height: opts.height, deviceScaleFactor: 2 });

    // Intercept CDN requests → serve local vendor files
    await setupVendorInterception(page, vendorMap);

    await page.goto(`file://${file}`, { waitUntil: 'domcontentloaded', timeout: 30000 });

    // Wait for JS charts (D3/Chart.js) to finish rendering
    await new Promise(r => setTimeout(r, opts.wait));

    // Clip exactly to the .slide-container bounds
    let clip;
    try {
      clip = await page.evaluate((w, h) => {
        const el = document.querySelector('.slide-container');
        if (!el) return null;
        const r = el.getBoundingClientRect();
        return { x: Math.round(r.left), y: Math.round(r.top), width: w, height: h };
      }, opts.width, opts.height);
    } catch (_) { clip = null; }

    const screenshotOpts = { type: 'png', omitBackground: false };
    if (clip) screenshotOpts.clip = clip;

    const buffer = await page.screenshot(screenshotOpts);
    screenshots.push({ file, buffer });
    await page.close();
    console.log('done');
  }

  await browser.close();

  // ── Build PPTX ────────────────────────────────────────────────────────────
  console.log('Assembling PPTX…');
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';   // 13.33 × 7.5 in (standard 16:9)

  for (const { buffer } of screenshots) {
    const slide = pptx.addSlide();
    slide.addImage({
      data: `data:image/png;base64,${buffer.toString('base64')}`,
      x: 0, y: 0, w: '100%', h: '100%',
    });
  }

  await pptx.writeFile({ fileName: opts.output });
  console.log(`✓ Saved: ${opts.output}  (${screenshots.length} slide${screenshots.length !== 1 ? 's' : ''})`);
}

main().catch(err => {
  console.error('Fatal error:', err.message || err);
  process.exit(1);
});
