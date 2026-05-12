# AWS Offensive Security — Book Build

This folder has everything you need to produce `AWS_Offensive_Security.docx` from the HTML source that was written earlier. Pick whichever build path fits your setup — you don't need both.

## Path A — docx-js (no Word required)

Cleanest output. Requires Node.js 18+.

```
cd "%USERPROFILE%\Desktop\aws-book-build"
npm install
node build-docx.js
```

Under 20 seconds. Writes `AWS_Offensive_Security.docx` next to `build-docx.js`. Open in Word, click the Table of Contents, press **F9** to populate page numbers.

The script finds the HTML automatically — it searches both Cowork data roots (`%AppData%\Roaming\Claude\...` and `%LocalAppData%\Packages\Claude_*\LocalCache\Roaming\Claude\...`). If it can't find it, copy `AWS_Offensive_Security.html` into this folder and re-run.

## Path B — Word COM (no Node required)

Faster if you already have Word installed. Uses Word's own HTML importer.

```
powershell -ExecutionPolicy Bypass -File .\Convert-BookToDocx.ps1
```

Takes 1–2 minutes while Word opens the HTML, converts, and saves. Same output filename. This path preserves the SVG diagrams natively, because Word can render inline SVG on import.

## Comparison

| | docx-js (A) | Word COM (B) |
|---|---|---|
| Needs Node | yes | no |
| Needs Word | no | yes |
| Build time | ~15s | 1–2 min |
| TOC | real field, F9 to populate | rebuilt from HTML |
| SVG figures | dropped, caption kept | preserved |
| Styling | fully custom | Word's HTML heuristics |
| Navigation pane | accurate (real Heading styles) | accurate |

If you want the figures, use Path B. If you want the tighter styling and don't care about the diagrams being there, use Path A.

## What's in the book

Cover, front matter, 20 chapters across 6 parts, 4 appendices, glossary, bibliography, index, colophon. Approximately 300 printed pages with 20 figures, 44 terminal captures, 20 callouts, 8 reference tables.

## Troubleshooting

- **`Could not find AWS_Offensive_Security.html`** — the HTML isn't in the script's search paths. Copy it into this folder from wherever you find it (the original write went to `%AppData%\Roaming\Claude\local-agent-mode-sessions\<session>\outputs\`) and re-run.
- **`Cannot find module 'docx'`** — you haven't run `npm install` yet.
- **TOC says "Error! No table of contents entries found"** — right-click the TOC in Word → **Update Field** (or click inside it and press F9). This is normal for any programmatically generated docx TOC.
- **PowerShell: "Word is not installed"** — use Path A instead.
