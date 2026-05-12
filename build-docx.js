/*
 * build-docx.js
 * Rebuilds AWS_Offensive_Security.html natively as a .docx using docx-js.
 *
 * Usage:
 *   npm install
 *   node build-docx.js
 *
 * Output: AWS_Offensive_Security.docx in the same directory as this script.
 *
 * Design notes:
 *   - No HTML-parsing dependency. The source HTML is pretty-printed with
 *     one block per line, so a line-oriented scanner is reliable.
 *   - Figures: if figures/fig-NN.png exists (produced by rasterize-figures.js),
 *     it is embedded as a centered ImageRun ahead of the caption. If no PNG
 *     exists for a given figure, only the italic <figcaption> is rendered —
 *     so the build degrades gracefully when the figures step is skipped.
 *   - Callouts render as a shaded paragraph with a bold label line.
 *   - Terminal/code blocks render in Consolas with a light shaded background.
 *   - Chapter-title headings trigger a page break before them.
 *   - A docx-js TableOfContents is emitted. Open the file in Word and
 *     press F9 on the TOC to populate page numbers.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const os = require('os');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType, PageOrientation, LevelFormat,
  ExternalHyperlink, TabStopType, TabStopPosition, HeadingLevel,
  BorderStyle, WidthType, ShadingType, PageNumber, PageBreak,
  TableOfContents, Bookmark, ImageRun, HeightRule, VerticalAlign,
} = require('docx');

// ---------- brand palette (matches AWS_Offensive_Security.html) -----------
const BRAND = {
  ink:         '0B0F14',   // cover / terminal background
  inkSoft:     '1A2432',   // cover panels
  paper:       'FFFFFF',
  accent:      'C64545',   // red — eyebrow, warning
  opNote:      '0B4A6A',   // operator blue
  defNote:     '2F6E3A',   // defender green
  cream:       'DCD3BE',   // cover subtitle
  muted:       'A6A499',   // cover org line
  panel:       'F4F1EB',   // neutral panel
  termFg:      'B7E4C7',   // terminal text (soft mint)
  termFgDim:   '9AA6B0',   // terminal secondary
};

// Where rasterize-figures.js drops its PNGs.
const FIGURE_DIR = path.join(__dirname, 'figures');

// ---------- locate the source HTML ----------------------------------------

const FILE = 'AWS_Offensive_Security.html';

function findHtml() {
  const scriptDir = __dirname;

  const home = os.homedir();
  const candidates = [
    path.join(scriptDir, FILE),
    // Direct Cowork session temp path (prior session)
    path.join(home, 'AppData', 'Roaming', 'Claude', 'local-agent-mode-sessions',
      '38a063fa-b67a-4d5f-b927-0c4121abcf17',
      'b4fb3676-2204-418d-a217-3d3c163ce4a8',
      'local_a9e42691-adaa-44ca-8428-df1ce4fdd8d8', 'outputs', FILE),
  ];

  for (const c of candidates) {
    if (fs.existsSync(c)) return c;
  }

  // Recursive search under BOTH Cowork roots.
  // 1. %AppData%\Roaming\Claude\local-agent-mode-sessions  (classic install)
  // 2. %AppData%\Local\Packages\Claude_*\LocalCache\Roaming\Claude\local-agent-mode-sessions  (MSIX / packaged app)
  // When running on WSL, also crawl /mnt/c/Users/*/AppData/... for the same paths.
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
    // Newest wins
    hits.sort((a, b) => fs.statSync(b).mtimeMs - fs.statSync(a).mtimeMs);
    return hits[0];
  }

  throw new Error('Could not find ' + FILE + '. Place it beside build-docx.js and re-run.');
}

// ---------- HTML → block stream ------------------------------------------

function decodeEntities(s) {
  return s
    .replace(/&nbsp;/g, '\u00A0')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&mdash;/g, '\u2014')
    .replace(/&ndash;/g, '\u2013')
    .replace(/&hellip;/g, '\u2026')
    .replace(/&copy;/g, '\u00A9')
    .replace(/&rsquo;/g, '\u2019')
    .replace(/&lsquo;/g, '\u2018')
    .replace(/&rdquo;/g, '\u201D')
    .replace(/&ldquo;/g, '\u201C')
    .replace(/&#x([0-9A-Fa-f]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#([0-9]+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)));
}

// Strip tags from an HTML fragment and return (plain, runs) where runs is an
// ordered list of {text, bold, italic, code, link} spans for use in TextRuns.
function fragmentToRuns(html) {
  // Normalize whitespace
  html = html.replace(/\s+/g, ' ').trim();
  const runs = [];
  // A very small state machine over a handful of inline tags.
  // We treat unknown tags as transparent (drop the tags, keep the content).
  const stack = [{ bold: false, italic: false, code: false, link: null }];
  const top = () => stack[stack.length - 1];

  let i = 0;
  while (i < html.length) {
    if (html[i] === '<') {
      const close = html.indexOf('>', i);
      if (close === -1) break;
      const tag = html.slice(i + 1, close);
      const isClose = tag.startsWith('/');
      const name = (isClose ? tag.slice(1) : tag).split(/[\s>]/)[0].toLowerCase();

      if (isClose) {
        // Pop a matching state (best-effort).
        stack.pop();
        if (stack.length === 0) stack.push({ bold: false, italic: false, code: false, link: null });
      } else {
        const attrs = tag.slice(name.length);
        const next = { ...top() };
        if (name === 'strong' || name === 'b') next.bold = true;
        else if (name === 'em' || name === 'i') next.italic = true;
        else if (name === 'code') next.code = true;
        else if (name === 'a') {
          const m = /href="([^"]*)"/.exec(attrs);
          if (m) next.link = m[1];
        } else if (name === 'br') {
          runs.push({ text: '\n', ...top() });
          i = close + 1;
          continue;
        }
        stack.push(next);
      }
      i = close + 1;
    } else {
      const next = html.indexOf('<', i);
      const chunk = decodeEntities(html.slice(i, next === -1 ? html.length : next));
      if (chunk) runs.push({ text: chunk, ...top() });
      i = next === -1 ? html.length : next;
    }
  }

  return runs.filter(r => r.text);
}

function runsToTextRuns(runs, opts = {}) {
  const { defaultFont, forceSize } = opts;
  const out = [];
  for (const r of runs) {
    const runOpts = {
      text: r.text,
      bold: r.bold || undefined,
      italics: r.italic || undefined,
      font: r.code ? { name: 'Consolas' } : (defaultFont ? { name: defaultFont } : undefined),
      size: forceSize || (r.code ? 20 : undefined), // 10pt for inline code
      color: r.code ? '2E2E2E' : undefined,
    };
    if (r.link) {
      out.push(new ExternalHyperlink({
        link: r.link,
        children: [new TextRun({ ...runOpts, style: 'Hyperlink' })],
      }));
    } else {
      out.push(new TextRun(runOpts));
    }
  }
  return out;
}

// ---------- Block classification -----------------------------------------

function stripTags(s) {
  return decodeEntities(s.replace(/<[^>]+>/g, ''));
}

function parseHtmlToBlocks(html) {
  // Isolate body
  const body = /<body[^>]*>([\s\S]*)<\/body>/i.exec(html)?.[1] || html;

  const lines = body.split(/\r?\n/);
  const blocks = [];

  const L = lines.length;
  let i = 0;

  // Helper to push a page break at part-title / chapter-title boundaries.
  let sawFirstHeading = false;
  // Counts <figure> blocks as we encounter them, so parseHtmlToBlocks and
  // rasterize-figures.js agree on the N in figures/fig-NN.png.
  let figureIndex = 0;
  // Sections whose non-heading content we drop entirely (the TableOfContents
  // field replaces the manual TOC, and we don't need the static list-of-figures).
  const SUPPRESSED_SECTION_IDS = new Set(['toc', 'list-of-figures']);
  let suppressingSection = false;

  while (i < L) {
    const raw = lines[i];
    const line = raw.trim();

    if (!line) { i++; continue; }

    // Skip <style>, <script>
    if (/^<(style|script)\b/i.test(line)) {
      while (i < L && !/<\/(style|script)>/i.test(lines[i])) i++;
      i++;
      continue;
    }

    // Track section id / suppression
    const sectionOpen = /^<section\b[^>]*id="([^"]+)"/i.exec(line);
    if (sectionOpen) {
      suppressingSection = SUPPRESSED_SECTION_IDS.has(sectionOpen[1]);
      i++;
      continue;
    }
    if (/^<\/section>/i.test(line)) {
      suppressingSection = false;
      i++;
      continue;
    }
    if (/^<section\b/i.test(line)) { i++; continue; }

    // Skip structural wrappers
    if (/^<\/?(body|html|div class="(?!callout|chapter-goal|kicker|name|org)[^"]*")/i.test(line) ||
        /^<\/?(main|article|aside|nav|header|footer)\b/i.test(line)) {
      i++;
      continue;
    }

    // Horizontal rule
    if (/^<hr\b/i.test(line)) { blocks.push({ kind: 'pageBreak' }); i++; continue; }

    // Skip pure closing wrappers we didn't specifically match
    if (/^<\/div>/i.test(line)) { i++; continue; }

    // Inside a suppressed section, emit nothing but headings.
    if (suppressingSection && !/^<h[1-4]\b/i.test(line)) {
      i++;
      continue;
    }

    // Headings ---------------------------------------------------------
    const h1 = /^<h1\b([^>]*)>([\s\S]*?)<\/h1>/i.exec(line);
    if (h1) {
      const attrs = h1[1];
      const inner = h1[2];
      const isPartTitle = /class="[^"]*\bpart-title\b/.test(attrs);
      const isChapterTitle = /class="[^"]*\bchapter-title\b/.test(attrs);
      const isCoverTitle = /class="[^"]*\btitle\b/.test(attrs) && !isPartTitle && !isChapterTitle;
      if (isCoverTitle) {
        blocks.push({ kind: 'coverTitle', html: inner });
      } else {
        if (sawFirstHeading) blocks.push({ kind: 'pageBreak' });
        sawFirstHeading = true;
        blocks.push({
          kind: 'heading',
          level: 1,
          html: inner,
          isPartTitle,
          isChapterTitle,
        });
        const plain = stripTags(inner);
        if (/^Table of Contents/i.test(plain)) {
          blocks.push({ kind: 'toc' });
        }
      }
      i++;
      continue;
    }
    const h2 = /^<h2\b([^>]*)>([\s\S]*?)<\/h2>/i.exec(line);
    if (h2) {
      const attrs = h2[1];
      const inner = h2[2];
      const isSubtitle = /class="[^"]*\bsubtitle\b/.test(attrs);
      if (isSubtitle) {
        blocks.push({ kind: 'coverSubtitle', html: inner });
      } else {
        blocks.push({ kind: 'heading', level: 2, html: inner });
      }
      i++;
      continue;
    }
    const h3 = /^<h3\b[^>]*>([\s\S]*?)<\/h3>/i.exec(line);
    if (h3) {
      blocks.push({ kind: 'heading', level: 3, html: h3[1] });
      i++;
      continue;
    }
    const h4 = /^<h4\b[^>]*>([\s\S]*?)<\/h4>/i.exec(line);
    if (h4) {
      blocks.push({ kind: 'heading', level: 4, html: h4[1] });
      i++;
      continue;
    }

    // Kicker (small caps above chapter titles)
    const kicker = /^<div class="kicker">([\s\S]*?)<\/div>/i.exec(line);
    if (kicker) {
      blocks.push({ kind: 'kicker', html: kicker[1] });
      i++;
      continue;
    }

    // Author block on cover
    const author = /^<div class="name">([\s\S]*?)<\/div>/i.exec(line);
    if (author) { blocks.push({ kind: 'author', html: author[1] }); i++; continue; }
    const org = /^<div class="org">([\s\S]*?)<\/div>/i.exec(line);
    if (org) { blocks.push({ kind: 'org', html: org[1] }); i++; continue; }

    // Chapter-goal (captures following <h4> + <p> inside a div)
    if (/^<div class="chapter-goal"/i.test(line)) {
      const inner = [];
      i++;
      while (i < L && !/^<\/div>/i.test(lines[i].trim())) {
        inner.push(lines[i]);
        i++;
      }
      i++; // skip </div>
      blocks.push({ kind: 'chapterGoal', lines: inner });
      continue;
    }

    // Callout (single-line or multi-line forms)
    const coSingle = /^<div class="callout ([^"]+)">([\s\S]*?)<\/div>\s*$/i.exec(line);
    if (coSingle) {
      blocks.push({ kind: 'callout', variant: coSingle[1], inner: coSingle[2] });
      i++;
      continue;
    }
    const coOpen = /^<div class="callout ([^"]+)">\s*$/i.exec(line);
    if (coOpen) {
      const variant = coOpen[1];
      const inner = [];
      i++;
      while (i < L && !/^<\/div>/i.test(lines[i].trim())) {
        inner.push(lines[i]);
        i++;
      }
      i++; // skip </div>
      blocks.push({ kind: 'callout', variant, innerLines: inner });
      continue;
    }

    // Figure: swallow the <svg>, emit figureImage (if PNG exists) + figcaption
    if (/^<figure\b/i.test(line)) {
      figureIndex++;
      let cap = null;
      i++;
      while (i < L && !/^<\/figure>/i.test(lines[i].trim())) {
        const m = /<figcaption[^>]*>([\s\S]*?)<\/figcaption>/i.exec(lines[i]);
        if (m) cap = m[1];
        i++;
      }
      i++; // skip </figure>
      const pngName = 'fig-' + String(figureIndex).padStart(2, '0') + '.png';
      const pngPath = path.join(FIGURE_DIR, pngName);
      if (fs.existsSync(pngPath)) {
        blocks.push({ kind: 'figureImage', pngPath, index: figureIndex });
      }
      if (cap) blocks.push({ kind: 'figcaption', html: cap, index: figureIndex });
      continue;
    }

    // <pre class="terminal">...</pre>  — may be single-line or multi-line
    if (/^<pre\b/i.test(line)) {
      if (/<\/pre>/i.test(line)) {
        const m = /<pre[^>]*>([\s\S]*?)<\/pre>/i.exec(line);
        blocks.push({ kind: 'code', text: stripTags(m[1]) });
        i++;
      } else {
        const acc = [line.replace(/^<pre[^>]*>/i, '')];
        i++;
        while (i < L && !/<\/pre>/i.test(lines[i])) { acc.push(lines[i]); i++; }
        if (i < L) { acc.push(lines[i].replace(/<\/pre>.*/i, '')); i++; }
        blocks.push({ kind: 'code', text: stripTags(acc.join('\n')) });
      }
      continue;
    }

    // <ul> or <ol>
    if (/^<(ul|ol)\b/i.test(line)) {
      const ordered = /^<ol\b/i.test(line);
      const items = [];
      if (/<\/(ul|ol)>/i.test(line)) {
        const m = /<(?:ul|ol)[^>]*>([\s\S]*?)<\/(?:ul|ol)>/i.exec(line);
        const liRe = /<li[^>]*>([\s\S]*?)<\/li>/gi;
        let li;
        while ((li = liRe.exec(m[1]))) items.push(li[1]);
        i++;
      } else {
        i++;
        while (i < L && !/^<\/(ul|ol)>/i.test(lines[i].trim())) {
          const li = /<li[^>]*>([\s\S]*?)<\/li>/i.exec(lines[i]);
          if (li) items.push(li[1]);
          else if (/^<li\b/i.test(lines[i].trim())) {
            const acc = [lines[i].replace(/^<li[^>]*>/i, '')];
            i++;
            while (i < L && !/<\/li>/i.test(lines[i])) { acc.push(lines[i]); i++; }
            if (i < L) acc.push(lines[i].replace(/<\/li>.*/i, ''));
            items.push(acc.join(' '));
          }
          i++;
        }
        i++; // skip </ul|ol>
      }
      blocks.push({ kind: 'list', ordered, items });
      continue;
    }

    // <table>
    if (/^<table\b/i.test(line)) {
      const acc = [line];
      i++;
      while (i < L && !/<\/table>/i.test(lines[i])) { acc.push(lines[i]); i++; }
      if (i < L) { acc.push(lines[i]); i++; }
      const blob = acc.join('\n');
      const caption = /<caption[^>]*>([\s\S]*?)<\/caption>/i.exec(blob)?.[1] || null;
      const theadRow = /<thead[\s\S]*?<tr[^>]*>([\s\S]*?)<\/tr>[\s\S]*?<\/thead>/i.exec(blob);
      const headers = [];
      if (theadRow) {
        const thRe = /<th[^>]*>([\s\S]*?)<\/th>/gi;
        let m;
        while ((m = thRe.exec(theadRow[1]))) headers.push(m[1]);
      }
      const rows = [];
      const bodyMatch = /<tbody[\s\S]*?>([\s\S]*?)<\/tbody>/i.exec(blob);
      const rowsBlob = bodyMatch ? bodyMatch[1] : blob;
      const trRe = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
      let tr;
      while ((tr = trRe.exec(rowsBlob))) {
        if (theadRow && tr.index < theadRow.index + theadRow[0].length) continue;
        const cells = [];
        const tdRe = /<td[^>]*>([\s\S]*?)<\/td>/gi;
        let td;
        while ((td = tdRe.exec(tr[1]))) cells.push(td[1]);
        if (cells.length) rows.push(cells);
      }
      blocks.push({ kind: 'table', caption, headers, rows });
      continue;
    }

    // Paragraph
    const pSingle = /^<p\b[^>]*>([\s\S]*?)<\/p>/i.exec(line);
    if (pSingle) {
      blocks.push({ kind: 'paragraph', html: pSingle[1] });
      i++;
      continue;
    }
    if (/^<p\b/i.test(line)) {
      const acc = [line.replace(/^<p[^>]*>/i, '')];
      i++;
      while (i < L && !/<\/p>/i.test(lines[i])) { acc.push(lines[i]); i++; }
      if (i < L) { acc.push(lines[i].replace(/<\/p>.*/i, '')); i++; }
      blocks.push({ kind: 'paragraph', html: acc.join(' ') });
      continue;
    }

    // Default: treat any remaining line as a raw paragraph (drop empty)
    const stripped = stripTags(line);
    if (stripped.trim()) {
      blocks.push({ kind: 'paragraph', html: line });
    }
    i++;
  }

  return blocks;
}

// ---------- Block → docx children ----------------------------------------

// Match the HTML callouts exactly: operator/defender/warning backgrounds
// and matching accent-bar colors pulled from the book palette.
const CALLOUT_COLORS = {
  operator:  { fill: 'EEF6FB', accent: BRAND.opNote },
  defender:  { fill: 'EEF7EF', accent: BRAND.defNote },
  warning:   { fill: 'FBEEEE', accent: BRAND.accent },
  tip:       { fill: 'FFF4CC', accent: '8A6D00' },
  note:      { fill: BRAND.panel, accent: '25303B' },
  danger:    { fill: 'FBEEEE', accent: '9B1C1C' },
};

function headingParagraph(block) {
  const level = block.level;
  const runs = runsToTextRuns(fragmentToRuns(block.html));
  const isPartTitle = block.isPartTitle;
  const isChapterTitle = block.isChapterTitle;
  const hl = [HeadingLevel.HEADING_1, HeadingLevel.HEADING_2, HeadingLevel.HEADING_3, HeadingLevel.HEADING_4][level - 1];

  // Part titles are framed by a thick top+bottom accent rule to evoke the
  // HTML's dark part-opener slab while staying fully editable as native text.
  if (isPartTitle) {
    const text = stripTags(block.html).replace(/\s+/g, ' ').trim();
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore: true,
      alignment: AlignmentType.CENTER,
      spacing: { before: 1600, after: 240, line: 480 },
      border: {
        top:    { style: BorderStyle.SINGLE, size: 24, color: BRAND.accent, space: 16 },
        bottom: { style: BorderStyle.SINGLE, size: 24, color: BRAND.accent, space: 16 },
      },
      children: [new TextRun({
        text,
        bold: true,
        size: 72,
        font: 'Arial',
        color: BRAND.ink,
      })],
    });
  }

  // Chapter titles: big, page break, underlined in brand accent red.
  if (isChapterTitle) {
    const text = stripTags(block.html).replace(/\s+/g, ' ').trim();
    return new Paragraph({
      heading: HeadingLevel.HEADING_1,
      pageBreakBefore: true,
      alignment: AlignmentType.LEFT,
      spacing: { before: 120, after: 240, line: 520 },
      border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: BRAND.accent, space: 8 } },
      children: [new TextRun({ text, bold: true, size: 56, font: 'Arial', color: BRAND.ink })],
    });
  }

  return new Paragraph({
    heading: hl,
    alignment: AlignmentType.LEFT,
    children: runs,
    spacing: { before: 240, after: 120 },
  });
}

function paragraph(block) {
  const runs = runsToTextRuns(fragmentToRuns(block.html));
  return new Paragraph({
    children: runs,
    spacing: { after: 160, line: 300 },
    alignment: AlignmentType.JUSTIFIED,
  });
}

function figureImage(block) {
  // Target ~5.3" wide so the figure fits within the 5.65" content width on the
  // 7.25" trim with breathing room. Derive the display height from the PNG's
  // intrinsic pixel aspect ratio so rasterized diagrams aren't squashed.
  const data = fs.readFileSync(block.pngPath);
  let widthPx = 1200, heightPx = 675; // fallback 16:9
  // Parse PNG IHDR (at byte offset 16, big-endian uint32 x2) for real dims.
  if (data.length > 24 && data[0] === 0x89 && data[1] === 0x50) {
    widthPx  = data.readUInt32BE(16);
    heightPx = data.readUInt32BE(20);
  }
  const displayWidthIn = 5.3;
  const displayHeightIn = displayWidthIn * (heightPx / widthPx);
  // docx-js uses pixels for width/height; 96px = 1in.
  const w = Math.round(displayWidthIn * 96);
  const h = Math.round(displayHeightIn * 96);
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 200, after: 80 },
    children: [new ImageRun({
      type: 'png',
      data,
      transformation: { width: w, height: h },
      altText: {
        title: 'Figure ' + block.index,
        description: 'Figure ' + block.index,
        name: 'fig-' + String(block.index).padStart(2, '0'),
      },
    })],
  });
}

function figcaption(block) {
  const runs = fragmentToRuns(block.html).map(r => ({ ...r, italic: true }));
  const num = typeof block.index === 'number' ? block.index : null;
  const label = 'Figure ' + (num !== null ? num : '') + (num !== null ? '. ' : '. ');
  return new Paragraph({
    children: [
      new TextRun({ text: label, bold: true, italics: true, size: 20, color: BRAND.accent }),
      ...runsToTextRuns(runs, { forceSize: 20 }),
    ],
    spacing: { before: 120, after: 280 },
    alignment: AlignmentType.CENTER,
  });
}

function codeBlock(block) {
  // Dark terminal theme to match HTML <pre class="terminal">.
  // Prompts ($ or #) and comments (#) get a dimmer color for readability.
  const paragraphs = [];
  const lines = block.text.replace(/\r/g, '').split('\n');
  lines.forEach((ln, idx) => {
    const isComment = /^\s*#/.test(ln) && !/^\s*#\s*\w+:/.test(ln);
    const fg = isComment ? BRAND.termFgDim : BRAND.termFg;
    paragraphs.push(new Paragraph({
      children: [new TextRun({ text: ln || ' ', font: { name: 'Consolas' }, size: 18, color: fg })],
      shading: { type: ShadingType.CLEAR, fill: BRAND.ink, color: 'auto' },
      spacing: {
        before: idx === 0 ? 180 : 0,
        after:  idx === lines.length - 1 ? 200 : 0,
        line: 260,
      },
      border: idx === 0
        ? { top:    { style: BorderStyle.SINGLE, size: 8, color: BRAND.accent, space: 0 } }
        : undefined,
    }));
  });
  // Bottom accent bar on the last line already covered via fill.
  return paragraphs;
}

function list(block) {
  const ref = block.ordered ? 'numbers' : 'bullets';
  return block.items.map(it => new Paragraph({
    numbering: { reference: ref, level: 0 },
    children: runsToTextRuns(fragmentToRuns(it)),
    spacing: { after: 80 },
  }));
}

function callout(block) {
  const { variant } = block;
  const colors = CALLOUT_COLORS[variant] || CALLOUT_COLORS.note;
  const paragraphs = [];

  let innerHtml;
  if (block.inner) innerHtml = block.inner;
  else innerHtml = (block.innerLines || []).join('\n');

  const h4 = /<h4[^>]*>([\s\S]*?)<\/h4>/i.exec(innerHtml);
  const label = h4 ? stripTags(h4[1]) : variant.toUpperCase();
  const body = innerHtml.replace(/<h4[^>]*>[\s\S]*?<\/h4>/i, '');

  paragraphs.push(new Paragraph({
    children: [new TextRun({ text: label, bold: true, color: colors.accent, size: 22 })],
    shading: { type: ShadingType.CLEAR, fill: colors.fill, color: 'auto' },
    spacing: { before: 160, after: 60 },
    border: { left: { style: BorderStyle.SINGLE, size: 24, color: colors.accent } },
  }));

  const pRe = /<p[^>]*>([\s\S]*?)<\/p>/gi;
  let m;
  let any = false;
  while ((m = pRe.exec(body))) {
    any = true;
    paragraphs.push(new Paragraph({
      children: runsToTextRuns(fragmentToRuns(m[1])),
      shading: { type: ShadingType.CLEAR, fill: colors.fill, color: 'auto' },
      spacing: { after: 80, line: 280 },
      border: { left: { style: BorderStyle.SINGLE, size: 24, color: colors.accent } },
    }));
  }
  if (!any && body.trim()) {
    paragraphs.push(new Paragraph({
      children: runsToTextRuns(fragmentToRuns(body)),
      shading: { type: ShadingType.CLEAR, fill: colors.fill, color: 'auto' },
      spacing: { after: 200, line: 280 },
      border: { left: { style: BorderStyle.SINGLE, size: 24, color: colors.accent } },
    }));
  }

  paragraphs.push(new Paragraph({ children: [], spacing: { after: 80 } }));
  return paragraphs;
}

function chapterGoal(block) {
  const paragraphs = [];
  const joined = block.lines.join('\n');
  const h4 = /<h4[^>]*>([\s\S]*?)<\/h4>/i.exec(joined);
  const label = h4 ? stripTags(h4[1]) : 'Chapter Goal';
  const body = joined.replace(/<h4[^>]*>[\s\S]*?<\/h4>/i, '');
  paragraphs.push(new Paragraph({
    children: [new TextRun({ text: label, bold: true, color: '1F4E79', size: 22 })],
    shading: { type: ShadingType.CLEAR, fill: 'F1F6FB', color: 'auto' },
    spacing: { before: 160, after: 60 },
    border: { left: { style: BorderStyle.SINGLE, size: 24, color: '1F4E79' } },
  }));
  const pRe = /<p[^>]*>([\s\S]*?)<\/p>/gi;
  let m;
  while ((m = pRe.exec(body))) {
    paragraphs.push(new Paragraph({
      children: runsToTextRuns(fragmentToRuns(m[1])),
      shading: { type: ShadingType.CLEAR, fill: 'F1F6FB', color: 'auto' },
      spacing: { after: 80, line: 280 },
      border: { left: { style: BorderStyle.SINGLE, size: 24, color: '1F4E79' } },
    }));
  }
  return paragraphs;
}

function bookTable(block) {
  // Content width for 7.25" trim with 0.8" side margins = 5.65" = 8136 DXA.
  const contentWidth = 8136;
  const cols = Math.max(
    block.headers.length,
    ...block.rows.map(r => r.length),
    1
  );
  const colWidth = Math.floor(contentWidth / cols);
  const columnWidths = new Array(cols).fill(colWidth);
  columnWidths[cols - 1] = contentWidth - colWidth * (cols - 1);

  const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
  const borders = { top: border, bottom: border, left: border, right: border };

  function cell(html, opts = {}) {
    const runs = runsToTextRuns(fragmentToRuns(html));
    const width = { size: columnWidths[opts.colIdx ?? 0], type: WidthType.DXA };
    return new TableCell({
      borders,
      width,
      shading: opts.header ? { type: ShadingType.CLEAR, fill: 'E6EEF7', color: 'auto' } : undefined,
      margins: { top: 80, bottom: 80, left: 120, right: 120 },
      children: [new Paragraph({
        children: runs.length ? runs : [new TextRun('')],
        spacing: { after: 0, line: 260 },
      })],
    });
  }

  const rows = [];
  if (block.headers.length) {
    rows.push(new TableRow({
      tableHeader: true,
      children: block.headers.map((h, idx) => cell(h, { header: true, colIdx: idx })),
    }));
  }
  for (const r of block.rows) {
    const padded = r.slice();
    while (padded.length < cols) padded.push('');
    rows.push(new TableRow({
      children: padded.map((c, idx) => cell(c, { colIdx: idx })),
    }));
  }

  const out = [];
  if (block.caption) {
    out.push(new Paragraph({
      children: [new TextRun({ text: 'Table: ', bold: true, italics: true, size: 20 }),
                 new TextRun({ text: stripTags(block.caption), italics: true, size: 20 })],
      spacing: { before: 160, after: 80 },
      alignment: AlignmentType.CENTER,
    }));
  }
  out.push(new Table({
    width: { size: contentWidth, type: WidthType.DXA },
    columnWidths,
    rows,
  }));
  out.push(new Paragraph({ children: [], spacing: { after: 200 } }));
  return out;
}

// ---------- Cover page ---------------------------------------------------

// Dark rectangle that fills the page. Everything inside is white/cream/red on
// near-black (BRAND.ink). We use a single-cell table with a forced height so
// the fill reliably covers the whole page area rather than relying on
// paragraph shading, which only colors line-height.
function coverPage(parts) {
  const contentWidth = 8136; // 7.25" trim with 0.8" side margins = 5.65"
  const cellHeight   = 10800; // safely inside 7.85" body height (9.25" - 2*0.7")
  const titleText = stripTags(parts.title).replace(/\s+/g, ' ').trim();
  // Split "AWS Offensive Security" into three stacked lines for the hero slab.
  const titleLines = titleText.split(/\s+/);
  const eyebrow = "A RED-TEAM PRACTITIONER'S HANDBOOK";

  const coverArtPath = path.join(FIGURE_DIR, 'cover-art.png');
  const haveCoverArt = fs.existsSync(coverArtPath);

  const p = (opts) => new Paragraph({
    alignment: opts.align || AlignmentType.LEFT,
    spacing: { before: opts.before || 0, after: opts.after || 0, line: opts.line || 300 },
    children: opts.children,
  });

  // Uniform dark-cell styling for every paragraph inside the cover.
  const onDark = (children, extra = {}) => new Paragraph({
    alignment: extra.align || AlignmentType.LEFT,
    spacing: { before: extra.before || 0, after: extra.after || 0, line: extra.line || 300 },
    shading: { type: ShadingType.CLEAR, fill: BRAND.ink, color: 'auto' },
    children,
  });

  const cellChildren = [];

  // Top air
  cellChildren.push(p({ children: [new TextRun({ text: ' ', color: BRAND.ink })], after: 600 }));

  // Eyebrow
  cellChildren.push(p({
    align: AlignmentType.LEFT,
    after: 480,
    children: [new TextRun({
      text: eyebrow,
      color: BRAND.accent,
      font: 'Arial',
      size: 20, // 10pt
      bold: true,
      characterSpacing: 100, // expanded tracking (in 1/20 pt)
    })],
  }));

  // Title — 3 stacked lines ("AWS" / "Offensive" / "Security")
  titleLines.forEach((word, idx) => {
    cellChildren.push(p({
      align: AlignmentType.LEFT,
      line: 1040, // 52pt line height for 48pt text
      after: idx === titleLines.length - 1 ? 480 : 0,
      children: [new TextRun({
        text: word,
        color: BRAND.paper,
        font: 'Arial',
        size: 96, // 48pt
        bold: true,
      })],
    }));
  });

  // Subtitle — italic cream, serif feel (use Georgia, universally installed)
  cellChildren.push(p({
    align: AlignmentType.LEFT,
    line: 360,
    after: 720,
    children: [new TextRun({
      text: stripTags(parts.subtitle),
      color: BRAND.cream,
      font: 'Georgia',
      size: 40, // 20pt
      italics: true,
    })],
  }));

  // Cover art image, if rasterized
  if (haveCoverArt) {
    const data = fs.readFileSync(coverArtPath);
    let wpx = 1400, hpx = 560;
    if (data.length > 24 && data[0] === 0x89 && data[1] === 0x50) {
      wpx = data.readUInt32BE(16);
      hpx = data.readUInt32BE(20);
    }
    const wIn = 4.8;
    const hIn = wIn * (hpx / wpx);
    cellChildren.push(new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 720 },
      children: [new ImageRun({
        type: 'png',
        data,
        transformation: { width: Math.round(wIn * 96), height: Math.round(hIn * 96) },
        altText: { title: 'Cover art', description: 'Adversary reaching through AWS services', name: 'cover-art' },
      })],
    }));
  }

  // Spacer that pushes author block toward the bottom.
  for (let k = 0; k < 3; k++) {
    cellChildren.push(p({ children: [new TextRun({ text: ' ', color: BRAND.ink })], line: 360 }));
  }

  // Author name — uppercase white
  if (parts.author) {
    cellChildren.push(p({
      after: 120,
      children: [new TextRun({
        text: parts.author.toUpperCase(),
        color: BRAND.paper,
        font: 'Arial',
        size: 32, // 16pt
        bold: true,
        characterSpacing: 40,
      })],
    }));
  }

  // Org — small muted caps with generous tracking
  if (parts.org) {
    cellChildren.push(p({
      children: [new TextRun({
        text: parts.org.toUpperCase(),
        color: BRAND.muted,
        font: 'Arial',
        size: 18, // 9pt
        characterSpacing: 120,
      })],
    }));
  }

  const coverTable = new Table({
    width: { size: contentWidth, type: WidthType.DXA },
    columnWidths: [contentWidth],
    borders: {
      top:    { style: BorderStyle.NONE, size: 0, color: BRAND.ink },
      bottom: { style: BorderStyle.NONE, size: 0, color: BRAND.ink },
      left:   { style: BorderStyle.NONE, size: 0, color: BRAND.ink },
      right:  { style: BorderStyle.NONE, size: 0, color: BRAND.ink },
      insideHorizontal: { style: BorderStyle.NONE, size: 0, color: BRAND.ink },
      insideVertical:   { style: BorderStyle.NONE, size: 0, color: BRAND.ink },
    },
    rows: [
      new TableRow({
        height: { value: cellHeight, rule: HeightRule.ATLEAST },
        children: [
          new TableCell({
            width: { size: contentWidth, type: WidthType.DXA },
            shading: { type: ShadingType.CLEAR, fill: BRAND.ink, color: 'auto' },
            margins: { top: 400, bottom: 400, left: 600, right: 600 },
            verticalAlign: VerticalAlign.TOP,
            children: cellChildren,
          }),
        ],
      }),
    ],
  });

  return [
    coverTable,
    new Paragraph({ children: [new PageBreak()] }),
  ];
}

// ---------- Build document -----------------------------------------------

function buildChildrenFromBlocks(blocks) {
  const children = [];

  const coverTitleBlock = blocks.find(b => b.kind === 'coverTitle');
  const coverSubBlock = blocks.find(b => b.kind === 'coverSubtitle');
  const authorBlock = blocks.find(b => b.kind === 'author');
  const orgBlock = blocks.find(b => b.kind === 'org');

  if (coverTitleBlock) {
    children.push(...coverPage({
      title: coverTitleBlock.html,
      subtitle: coverSubBlock ? coverSubBlock.html : '',
      author: authorBlock ? stripTags(authorBlock.html) : '',
      org: orgBlock ? stripTags(orgBlock.html) : '',
    }));
  }

  for (const b of blocks) {
    switch (b.kind) {
      case 'coverTitle':
      case 'coverSubtitle':
      case 'author':
      case 'org':
        break;
      case 'heading':
        children.push(headingParagraph(b));
        break;
      case 'kicker':
        // Brand-red eyebrow over chapter/part titles, uppercase with tracking.
        children.push(new Paragraph({
          alignment: AlignmentType.LEFT,
          spacing: { before: 240, after: 40 },
          children: [new TextRun({
            text: stripTags(b.html).toUpperCase(),
            size: 20,
            color: BRAND.accent,
            bold: true,
            font: 'Arial',
            characterSpacing: 120,
          })],
        }));
        break;
      case 'paragraph':
        children.push(paragraph(b));
        break;
      case 'figureImage':
        children.push(figureImage(b));
        break;
      case 'figcaption':
        children.push(figcaption(b));
        break;
      case 'code':
        children.push(...codeBlock(b));
        break;
      case 'list':
        children.push(...list(b));
        break;
      case 'callout':
        children.push(...callout(b));
        break;
      case 'chapterGoal':
        children.push(...chapterGoal(b));
        break;
      case 'table':
        children.push(...bookTable(b));
        break;
      case 'pageBreak':
        children.push(new Paragraph({ children: [new PageBreak()] }));
        break;
      case 'toc':
        children.push(new TableOfContents('Table of Contents', { hyperlink: true, headingStyleRange: '1-3' }));
        children.push(new Paragraph({ children: [new PageBreak()] }));
        break;
      default:
        break;
    }
  }

  return children;
}

function buildDocument(children) {
  return new Document({
    creator: 'Michael Mancuso',
    title: 'AWS Offensive Security',
    description: "A Red-Team Practitioner's Handbook",
    styles: {
      default: {
        document: { run: { font: 'Source Sans Pro', size: 22 } },
      },
      paragraphStyles: [
        {
          id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 40, bold: true, color: '1F4E79', font: 'Source Sans Pro' },
          paragraph: { spacing: { before: 360, after: 160 }, outlineLevel: 0 },
        },
        {
          id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 30, bold: true, color: '1F4E79', font: 'Source Sans Pro' },
          paragraph: { spacing: { before: 280, after: 120 }, outlineLevel: 1 },
        },
        {
          id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 26, bold: true, color: '1F4E79', font: 'Source Sans Pro' },
          paragraph: { spacing: { before: 220, after: 100 }, outlineLevel: 2 },
        },
        {
          id: 'Heading4', name: 'Heading 4', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 22, bold: true, italics: true, color: '1F4E79', font: 'Source Sans Pro' },
          paragraph: { spacing: { before: 180, after: 80 }, outlineLevel: 3 },
        },
      ],
    },
    numbering: {
      config: [
        {
          reference: 'bullets',
          levels: [{
            level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          }],
        },
        {
          reference: 'numbers',
          levels: [{
            level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: { left: 720, hanging: 360 } } },
          }],
        },
      ],
    },
    sections: [{
      properties: {
        page: {
          // HTML @page: 7.25in x 9.25in with 0.7in / 0.8in margins.
          // 1440 DXA = 1 inch.
          size: { width: 10440, height: 13320 },
          margin: { top: 1008, right: 1152, bottom: 1008, left: 1152 },
        },
      },
      headers: {
        default: new Header({ children: [new Paragraph({
          alignment: AlignmentType.RIGHT,
          children: [new TextRun({ text: 'AWS Offensive Security', size: 18, color: '888888' })],
        })] }),
      },
      footers: {
        default: new Footer({ children: [new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT], size: 18, color: '888888' })],
        })] }),
      },
      children,
    }],
  });
}

// ---------- main ---------------------------------------------------------

async function main() {
  const src = findHtml();
  console.log('Source:', src);
  const html = fs.readFileSync(src, 'utf8');
  console.log('Parsing HTML (' + html.length.toLocaleString() + ' bytes)...');

  const blocks = parseHtmlToBlocks(html);
  console.log('Extracted', blocks.length, 'blocks.');

  const children = buildChildrenFromBlocks(blocks);
  console.log('Rendered', children.length, 'top-level document children.');

  const doc = buildDocument(children);
  console.log('Packing .docx...');
  const buf = await Packer.toBuffer(doc);

  const outPath = path.join(__dirname, 'AWS_Offensive_Security.docx');
  fs.writeFileSync(outPath, buf);
  const size = fs.statSync(outPath).size;
  console.log('Wrote', outPath, '(' + size.toLocaleString() + ' bytes)');
  console.log('Open in Word and press F9 on the Table of Contents to populate page numbers.');
}

main().catch(err => {
  console.error('\nBuild failed:', err.message);
  console.error(err.stack);
  process.exit(1);
});
