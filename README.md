# md-to-docx-local

Local Markdown to DOCX converter based on
[`vace/markdown-docx`](https://github.com/vace/markdown-docx).

## Quick Start

Most users do not need Node.js, pnpm, or a build step. Download the binary for
your operating system from the
[GitHub Releases](https://github.com/Sayyat/md-to-docx-local/releases) page,
put it in a directory that is listed in `PATH`, and run `md-to-docx`.
`PATH` is the operating system setting that tells the terminal where to search
for command-line programs.

Choose the release asset that matches your system:

| System | CPU | Release asset |
| --- | --- | --- |
| Linux | x64 / amd64 | `md-to-docx-linux-x64` |
| Linux | arm64 / aarch64 | `md-to-docx-linux-arm64` |
| macOS Intel | x64 | `md-to-docx-macos-x64` |
| macOS Apple Silicon | arm64 | `md-to-docx-macos-arm64` |
| Windows | x64 | `md-to-docx-win-x64.exe` |

If the Releases page does not have binaries yet, use the
[Development Setup](#development-setup) section and build locally.

Check your CPU on Linux or macOS with:

```bash
uname -m
```

Typical values:

- `x86_64` means x64 / amd64;
- `aarch64` or `arm64` means arm64.

## Install a Release Binary

### Linux

Download `md-to-docx-linux-x64` or `md-to-docx-linux-arm64`, then install it as
`md-to-docx`:

```bash
mkdir -p ~/.local/bin
cp ~/Downloads/md-to-docx-linux-x64 ~/.local/bin/md-to-docx
chmod +x ~/.local/bin/md-to-docx
```

Make sure `~/.local/bin` is in `PATH`:

```bash
echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.bashrc
source ~/.bashrc
```

If you use Zsh:

```bash
echo 'export PATH="$HOME/.local/bin:$PATH"' >> ~/.zshrc
source ~/.zshrc
```

### macOS

Download `md-to-docx-macos-x64` or `md-to-docx-macos-arm64`, then install it as
`md-to-docx`:

```bash
mkdir -p ~/bin
cp ~/Downloads/md-to-docx-macos-arm64 ~/bin/md-to-docx
chmod +x ~/bin/md-to-docx
```

Add `~/bin` to `PATH` for the default macOS Zsh shell:

```bash
echo 'export PATH="$HOME/bin:$PATH"' >> ~/.zshrc
source ~/.zshrc
```

If macOS blocks the downloaded binary because it came from the internet, remove
the quarantine flag:

```bash
xattr -d com.apple.quarantine ~/bin/md-to-docx 2>/dev/null || true
```

### Windows

Download `md-to-docx-win-x64.exe`, create a personal bin directory, and copy the
binary there as `md-to-docx.exe`:

```powershell
$bin = "$env:USERPROFILE\bin"
New-Item -ItemType Directory -Force $bin
Copy-Item "$env:USERPROFILE\Downloads\md-to-docx-win-x64.exe" "$bin\md-to-docx.exe" -Force
```

Add that directory to the user `PATH` from PowerShell:

```powershell
$bin = "$env:USERPROFILE\bin"
$userPath = [Environment]::GetEnvironmentVariable("Path", "User")
if (($userPath -split ";") -notcontains $bin) {
  [Environment]::SetEnvironmentVariable("Path", "$userPath;$bin", "User")
}
```

Close and reopen the terminal after changing `PATH`.

### Verify Installation

```bash
md-to-docx --help
```

If this command works from any folder, installation is complete.

## Convert One File

```bash
md-to-docx /path/to/input.md
md-to-docx /path/to/input.md --out-dir /path/to/output
```

The default output path is next to the input file:

```text
/path/to/input.md -> /path/to/input.docx
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
md-to-docx /path/to/docs --out-dir /path/to/docx
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

After `md-to-docx` is available in `PATH`, run it from the
`SboxFiniteAutomata` repository:

```bash
cd /home/sayat/projects/science/SboxFiniteAutomata
md-to-docx docs/algorithm_explanations --out-dir generated/docx/algorithm_explanations
```

## Development Setup

Advanced users can clone the repository and build the binary themselves:

```bash
git clone https://github.com/Sayyat/md-to-docx-local.git
cd md-to-docx-local
pnpm install
```

This project is maintained with `pnpm`. The npm lockfile is kept only as a
compatibility snapshot; development, builds, and releases use `pnpm`.

Run from source:

```bash
pnpm run convert -- /path/to/input.md
```

Build the local Linux x64 binary:

```bash
pnpm run build
```

The binary is written to:

```text
dist/md-to-docx-linux-x64
```

Build all release binaries:

```bash
pnpm run build:all
```

Local cross-platform builds are useful for checking the packaging flow. Official
release binaries are built by GitHub Actions on Linux, macOS, and Windows
runners.

Build only one release target:

```bash
pnpm run bundle
node scripts/build-release.mjs linux-x64
node scripts/build-release.mjs macos-arm64
node scripts/build-release.mjs win-x64
```

## Agent Guidance

This repository includes `AGENTS.md` for AI coding agents and maintainer
automation. Ordinary users can ignore it.

The maintainer keeps shared agent rules in:

```text
~/.gemini/GEMINI.md
~/.gemini/rules/md_to_docx_local.md
```

Those local files come from the public rules repository:

```text
https://github.com/Sayyat/antigravity-rules
```

Agents working on this repository should read those rules first, then follow the
repository-specific summary in `AGENTS.md`.

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
