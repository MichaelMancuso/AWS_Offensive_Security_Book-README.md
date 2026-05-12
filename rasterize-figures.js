/*
 * rasterize-figures.js
 * Extracts inline <svg> blocks from AWS_Offensive_Security.html and writes
 * each one as figures/fig-NN.png using sharp.
 *
 * Usage:
 *   npm install sharp
 *   node rasterize-figures.js
 *
 * The figure numbering matches the order <figure> elements appear in the HTML.
 * build-docx.js reads the same HTML in the same order and embeds figures/fig-NN.png
 * wherever it finds the matching <figure> block.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const sharp = require('sharp');

const FILE = 'AWS_Offensive_Security.html';

// ---------- locate the source HTML ----------------------------------------
// Mirrors the lookup logic in build-docx.js so this script can run standalone.

function findHtml() {
  const scriptDir = __dirname;
  const home = os.homedir();

  const candidates = [
    path.join(scriptDir, FILE),
    path.join(home, 'AppData', 'Roaming', 'Claude', 'local-agent-mode-sessions',
      '38a063fa-b67a-4d5f-b927-0c4121abcf17',
      'b4fb3676-2204-418d-a217-3d3c163ce4a8',
      'local_a9e42691-adaa-44ca-8428-df1ce4fdd8d8', 'outputs', FILE),
  ];
  for (const c of candidates) if (fs.existsSync(c)) return c;

  const searchRoots = [
    path.join(home, 'AppData', 'Roaming', 'Claude', 'local-agent-mode-sessions'),
  ];
  const pkgParents = [path.join(home, 'AppData', 'Local', 'Packages')];

  if (process.platform === 'linux' && fs.existsSync('/mnt/c/Users')) {
    try {
      for (const u of fs.readdirSync('/mnt/c/Users')) {
        const winHome = path.join('/mnt/c/Users', u);
        searchRoots.push(path.join(winHome, 'AppData', 'Roaming', 'Claude', 'local-agent-mode-sessions'));
        pkgParents.push(path.join(winHome, 'AppData', 'Local', 'Packages'));
      }
    } catch {}
  }

  for (const pkgParent of pkgParents) {
    if (!fs.existsSync(pkgParent)) continue;
    try {
      for (const d of fs.readdirSync(pkgParent)) {
        if (/^Claude[_.]/.test(d)) {
          searchRoots.push(path.join(pkgParent, d, 'LocalCache', 'Roaming', 'Claude', 'local-agent-mode-sessions'));
        }
      }
    } catch {}
  }

  const hits = [];
  for (const root of searchRoots) {
    if (!fs.existsSync(root)) continue;
    (function walk(dir) {
      let entries;
      try { entries = fs.readdirSync(dir, { withFileTypes: true }); } catch { return; }
      for (const e of entries) {
        const full = path.join(dir, e.name);
        if (e.isDirectory()) walk(full);
        else if (e.isFile() && e.name === FILE) hits.push(full);
      }
    })(root);
  }
  if (hits.length) {
    hits.sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
    return hits[0];
  }

  throw new Error('Could not find ' + FILE + '. Place it beside this script and re-run.');
}

// ---------- figure extraction --------------------------------------------

function extractFigures(html) {
  const body = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html)?.[1] || html;
  const figures = [];
  const figRe = /<figure\b[^>]*>([\s\S]*?)<\/figure>/gi;
  let m;
  let index = 0;
  while ((m = figRe.exec(body)) !== null) {
    index++;
    const inner = m[1];
    const svgMatch = /<svg\b[\s\S]*?<\/svg>/i.exec(inner);
    figures.push({ index, svg: svgMatch ? svgMatch[0] : null });
  }
  return figures;
}

// ---------- rasterization -------------------------------------------------

async function rasterize(svg, outPath) {
  // density 200 gives a crisp render for a ~6-inch-wide print figure.
  // If the SVG has no width/height attrs, sharp uses the viewBox + density.
  await sharp(Buffer.from(svg), { density: 200 })
    .png({ compressionLevel: 9 })
    .toFile(outPath);
}

// ---------- main ----------------------------------------------------------

// Extract the cover-art SVG (first <svg> inside <section class="... cover" ...>).
function extractCoverSvg(html) {
  const body = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html)?.[1] || html;
  const cov = /<section[^>]*class="[^"]*cover[^"]*"[^>]*>([\s\S]*?)<\/section>/i.exec(body);
  if (!cov) return null;
  const svg = /<svg\b[\s\S]*?<\/svg>/i.exec(cov[1]);
  return svg ? svg[0] : null;
}

async function main() {
  const htmlPath = findHtml();
  console.log('Source:', htmlPath);

  const html = fs.readFileSync(htmlPath, 'utf8');
  const figures = extractFigures(html);
  console.log(`Found ${figures.length} <figure> blocks.`);

  const outDir = path.join(__dirname, 'figures');
  fs.mkdirSync(outDir, { recursive: true });

  // Clear any stale fig-NN.png and cover-art.png from a prior run.
  for (const f of fs.readdirSync(outDir)) {
    if (/^(fig-\d+|cover-art)\.png$/i.test(f)) {
      try { fs.unlinkSync(path.join(outDir, f)); } catch {}
    }
  }

  // Cover art first, if present.
  const coverSvg = extractCoverSvg(html);
  if (coverSvg) {
    const coverOut = path.join(outDir, 'cover-art.png');
    try {
      await sharp(Buffer.from(coverSvg), { density: 220 })
        .png({ compressionLevel: 9 })
        .toFile(coverOut);
      const size = fs.statSync(coverOut).size;
      console.log(`  cover-art.png (${size.toLocaleString()} bytes)`);
    } catch (err) {
      console.error(`  cover-art.png: FAILED — ${err.message}`);
    }
  } else {
    console.log('  cover-art.png: no <svg> inside cover section, skipped');
  }

  let ok = 0;
  let skipped = 0;
  let failed = 0;

  for (const { index, svg } of figures) {
    const tag = 'fig-' + String(index).padStart(2, '0') + '.png';
    const outPath = path.join(outDir, tag);
    if (!svg) {
      console.log(`  ${tag}: no <svg> inside this <figure>, skipped`);
      skipped++;
      continue;
    }
    try {
      await rasterize(svg, outPath);
      const size = fs.statSync(outPath).size;
      console.log(`  ${tag} (${size.toLocaleString()} bytes)`);
      ok++;
    } catch (err) {
      console.error(`  ${tag}: FAILED — ${err.message}`);
      failed++;
    }
  }

  console.log('');
  console.log(`Done. ${ok} rasterized, ${skipped} skipped, ${failed} failed.`);
  console.log(`Output: ${outDir}`);
  if (failed > 0) process.exit(2);
}

main().catch(err => {
  console.error('rasterize-figures.js failed:', err);
  process.exit(1);
});
