# html2ppt — HTML Slides to PowerPoint Converter

Convert HTML presentation slides to an **editable** `.pptx` file using headless Chromium.

## How It Works

1. **Headless Chromium** (via `puppeteer-core`) renders each HTML slide at 1280×720 px.
2. All CDN resources (D3.js, Chart.js, Font Awesome) are **served locally** from the `vendor/` directory — no internet connection required at render time.
3. The DOM is walked via `page.evaluate()` — each element's exact position is computed with `getBoundingClientRect()` relative to the `.slide-container`.
4. **PptxGenJS** builds a native `.pptx` slide with:
   - **Native text boxes** for all text elements (editable in PowerPoint / Google Slides)
   - **Native tables** for `<table>` elements (editable)
   - **PNG screenshots** only for JS chart elements (`<canvas>` for Chart.js, `<svg>` for D3.js)
   - Slide background colour from the computed CSS

### Coordinate Mapping

The `.slide-container` is always 1280 × 720 CSS pixels. PPTX `LAYOUT_WIDE` is 13.33 × 7.5 inches, which is exactly 1280 × 720 px at 96 dpi. So:

```
x_inches = x_css_px / 96
y_inches = y_css_px / 96
font_pt  = font_px  × 0.75   (1 in = 96 px = 72 pt)
```

## Requirements

- **Node.js** ≥ 18
- **Chromium** (e.g. `chromium-browser` on Ubuntu/Debian)
- **CJK fonts** for Chinese/Japanese/Korean text rendering:
  ```bash
  sudo apt-get install fonts-noto-cjk
  ```

## Install

```bash
npm install
```

## Usage

```bash
# Convert one or more HTML files (files are added as slides in the given order)
node html2ppt.js slide1.html slide2.html slide3.html

# Specify output filename
node html2ppt.js -o presentation.pptx slide1.html slide2.html

# Use glob pattern
node html2ppt.js --glob "html-preview*.html" -o output.pptx

# Via npm script
npm run convert -- slide1.html slide2.html -o output.pptx
```

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `-o, --output <file>` | `output.pptx` | Output PPTX filename |
| `-w, --width <px>` | `1280` | Slide viewport width |
| `-h, --height <px>` | `720` | Slide viewport height |
| `--wait <ms>` | `2000` | Extra wait after page load (for JS charts to render) |
| `--glob <pattern>` | — | Glob pattern for input files |
| `--help` | — | Show help |

## Convert All Example Slides

```bash
node html2ppt.js -o output.pptx \
  "html-preview (2).html"  \
  "html-preview (3).html"  \
  "html-preview (4).html"  \
  "html-preview (5).html"  \
  "html-preview (6).html"  \
  "html-preview (7).html"  \
  "html-preview (8).html"  \
  "html-preview (9).html"  \
  "html-preview (10).html" \
  "html-preview (11).html"
```

## Vendor Assets

The `vendor/` directory contains locally bundled copies of CDN libraries:

| File | Source |
|------|--------|
| `vendor/d3.v7.min.js` | D3.js v7 (from `d3` npm package) |
| `vendor/chart.3.9.1.min.js` | Chart.js v3.9.1 UMD build (from `chart.js` npm package) |
| `vendor/fontawesome/` | Font Awesome Free (from `@fortawesome/fontawesome-free` npm package) |

CDN URLs in the HTML files are intercepted at render time and served from these local files, so the converter works fully offline.

## HTML Slide Guidelines

See [`html_guide.txt`](./html_guide.txt) for the full authoring specification. Key constraints:

- Slide container: `width: 1280px; min-height: 720px` with class `slide-container`
- Use D3.js v7 (`https://d3js.org/d3.v7.min.js`) or Chart.js 3.x for data visualizations
- Native CSS only — no Tailwind or other CSS frameworks
- No CSS animations or `position: absolute` on main containers

## Requirements

- **Node.js** ≥ 18
- **Chromium** (e.g. `chromium-browser` on Ubuntu/Debian)
- **CJK fonts** for Chinese/Japanese/Korean text rendering:
  ```bash
  sudo apt-get install fonts-noto-cjk
  ```

## Install

```bash
npm install
```

## Usage

```bash
# Convert one or more HTML files (files are added as slides in the given order)
node html2ppt.js slide1.html slide2.html slide3.html

# Specify output filename
node html2ppt.js -o presentation.pptx slide1.html slide2.html

# Use glob pattern
node html2ppt.js --glob "html-preview*.html" -o output.pptx

# Via npm script
npm run convert -- slide1.html slide2.html -o output.pptx
```

### Options

| Flag | Default | Description |
|------|---------|-------------|
| `-o, --output <file>` | `output.pptx` | Output PPTX filename |
| `-w, --width <px>` | `1280` | Slide viewport width |
| `-h, --height <px>` | `720` | Slide viewport height |
| `--wait <ms>` | `2000` | Extra wait after page load (for JS charts to render) |
| `--glob <pattern>` | — | Glob pattern for input files |
| `--help` | — | Show help |

## Convert All Example Slides

```bash
node html2ppt.js -o output.pptx \
  "html-preview (2).html"  \
  "html-preview (3).html"  \
  "html-preview (4).html"  \
  "html-preview (5).html"  \
  "html-preview (6).html"  \
  "html-preview (7).html"  \
  "html-preview (8).html"  \
  "html-preview (9).html"  \
  "html-preview (10).html" \
  "html-preview (11).html"
```

## Vendor Assets

The `vendor/` directory contains locally bundled copies of CDN libraries:

| File | Source |
|------|--------|
| `vendor/d3.v7.min.js` | D3.js v7 (from `d3` npm package) |
| `vendor/chart.3.9.1.min.js` | Chart.js v3.9.1 UMD build (from `chart.js` npm package) |
| `vendor/fontawesome/` | Font Awesome Free (from `@fortawesome/fontawesome-free` npm package) |

CDN URLs in the HTML files are intercepted at render time and served from these local files, so the converter works fully offline.

## HTML Slide Guidelines

See [`html_guide.txt`](./html_guide.txt) for the full authoring specification. Key constraints:

- Slide container: `width: 1280px; min-height: 720px` with class `slide-container`
- Use D3.js v7 (`https://d3js.org/d3.v7.min.js`) or Chart.js 3.x for data visualizations
- Native CSS only — no Tailwind or other CSS frameworks
- No CSS animations or `position: absolute` on main containers
