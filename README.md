# md-to-docx-local

Local Markdown to DOCX converter based on
[`vace/markdown-docx`](https://github.com/vace/markdown-docx).

## Install

```bash
npm install
```

## Convert One File

```bash
npm run convert -- /path/to/input.md --out-dir /path/to/output
```

or after building the binary:

```bash
./dist/md-to-docx-linux-x64 /path/to/input.md --out-dir /path/to/output
```

The default engine is `markdown-docx`:

```bash
md-to-docx input.md
md-to-docx input.md --engine markdown-docx
```

Pandoc can be used as an optional backend when it is installed on the system:

```bash
md-to-docx input.md --engine pandoc
md-to-docx input.md --engine pandoc --reference-doc reference.docx
```

The same academic post-processing is applied after both engines.

## Convert a Folder

```bash
npm run convert -- /path/to/docs --out-dir /path/to/docx
```

## Image Sizing

Images are scaled down to `560px` wide by default while preserving aspect ratio.
Change that limit with:

```bash
md-to-docx input.md --max-image-width 720
```

Disable automatic scaling with:

```bash
md-to-docx input.md --max-image-width 0
```

## Default Academic Style

The default config uses an academic document style:

- page margins: left `3 cm`, right `1 cm`, top `2 cm`, bottom `2 cm`;
- line spacing: `1.0`;
- body font size: `12 pt`;
- blank Markdown spacer paragraphs: `12 pt`;
- body first-line indent: `1.25 cm`;
- blank Markdown spacer paragraphs between ordinary text paragraphs are removed;
- display math paragraphs are centered, use zero first-line indent, and do not
  keep extra blank spacer paragraphs immediately before or after them;
- document title from a Pandoc title block: `16 pt`, bold;
- Markdown headings `#` through `######`: `12 pt`, bold, with zero
  before/after spacing and `1.25 cm` first-line indent;
- every Markdown heading is followed by one blank spacer paragraph;
- consecutive Markdown headings are separated by one blank spacer paragraph;
- numbered top-level headings such as `# 6. Proposed Algorithm` start on a new
  page;
- use `-np` for dissertation-style output where every top-level heading starts
  on a new page, or `-no-np` for article-style continuous sections;
- Markdown list items: marker starts at `1.25 cm`, with a readable `0.5 cm`
  gap between the marker and text;
- default text color: black;
- wide images: capped at `560 px` and image paragraphs use zero indent.

## CLI Style Overrides

The defaults above are saved in `config/default.json`. Any setting can be
changed per conversion without editing the config.

Common shortcuts:

```bash
# Choose the conversion engine.
md-to-docx paper.md -e markdown-docx
md-to-docx paper.md -e pandoc

# Use a Pandoc reference DOCX.
md-to-docx paper.md -e pandoc --reference-doc reference.docx

# Dissertation-style sections: every top-level heading starts on a new page.
md-to-docx paper.md -np

# Article-style sections: continue top-level headings on the same page.
md-to-docx paper.md -no-np

# Body line spacing.
md-to-docx paper.md -ls 1.5
md-to-docx paper.md -lineSpacing 1.5

# Margins in centimeters.
md-to-docx paper.md -ml 3 -mr 1 -mt 2 -mb 2

# First-line indent in centimeters.
md-to-docx paper.md -fl 1.25

# Image width cap in pixels.
md-to-docx paper.md -miw 640
```

For less common settings, use a direct config path:

```bash
md-to-docx paper.md -set style.headings.blankLinesAfter=0
md-to-docx paper.md -set style.headings.pageBreakBeforeNumberedTopLevel=false
md-to-docx paper.md -set captions.figurePrefix=Figure
```

Markdown list semantics are preserved: `1.` / `2.` produce numbered lists, while
`-`, `*`, or `+` produce bullet lists. Use numbered lists only where the order or
count is meaningful.

Figures get automatic captions below the image from Markdown alt text:

```markdown
![Round-10 mean fail percentage](figures/round10.png)
```

Output:

```text
Сурет 1. Round-10 mean fail percentage
```

Tables get automatic captions above the table. Add a title with an explicit
marker:

```markdown
<!-- table-caption: Round-10 comparison under the full matrix -->

| Scenario | AES | Proposed |
|---|---:|---:|
| ecb | 3.39 | 4.76 |
```

Output:

```text
Кесте 1
Round-10 comparison under the full matrix
```

## Use with SboxFiniteAutomata

From anywhere:

```bash
/home/sayat/projects/science/md-to-docx-local/dist/md-to-docx-linux-x64 \
  /home/sayat/projects/science/SboxFiniteAutomata/docs/algorithm_explanations \
  --out-dir /home/sayat/projects/science/SboxFiniteAutomata/generated/docx/algorithm_explanations
```

## Build Binary

```bash
npm run build
```

The binary is written to:

```text
dist/md-to-docx-linux-x64
```

## Release

Releases are tag-driven. The local command bumps `package.json`, commits the
version change, creates an annotated `vX.Y.Z` git tag, and pushes it:

```bash
pnpm release patch
pnpm release minor
pnpm release 0.2.0 --notes "First public binary release"
```

GitHub Actions watches tags matching `v*.*.*`. For every release tag it builds
Linux binaries on Ubuntu, macOS binaries on macOS runners, and the Windows
binary on Windows. It uploads:

- `md-to-docx-linux-x64`;
- `md-to-docx-linux-arm64`;
- `md-to-docx-macos-x64`;
- `md-to-docx-macos-arm64`;
- `md-to-docx-win-x64.exe`;
- source archives in `.tar.gz` and `.zip` formats;
- `SHA256SUMS.txt`.

Use the local multi-platform build directly with:

```bash
pnpm run build:all
node scripts/build-release.mjs linux-x64
```

## Notes

- Pandoc title blocks such as `% Title`, `% Author`, `% Date` are converted to
  normal DOCX title metadata and heading text.
- Local image paths are resolved relative to the Markdown file directory, so
  `docs/scopus/file.md` can reference `figures/chart.png`.
- Wide images are capped by `converter.image.maxWidthPx` from
  `config/default.json`.
- Academic margins, paragraph spacing, first-line indentation, and black text
  are applied from `style` in `config/default.json`.
- Math is rendered through KaTeX using `markdown-docx`.
- Default settings live in `config/default.json`.
