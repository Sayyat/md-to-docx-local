#!/usr/bin/env node

const fs = require("node:fs/promises");
const os = require("node:os");
const http = require("node:http");
const https = require("node:https");
const path = require("node:path");
const { spawn } = require("node:child_process");
const { fileURLToPath } = require("node:url");
const { imageSize } = require("image-size");
const JSZip = require("jszip");
const { MarkdownDocx, Packer } = require("markdown-docx");

const PROJECT_ROOT = path.resolve(__dirname, "..");
const DEFAULT_CONFIG = path.join(PROJECT_ROOT, "config", "default.json");
const DEFAULT_CONFIG_DATA = {
  engine: "markdown-docx",
  document: {
    creator: "md-to-docx-local",
    description: "Generated locally from Markdown.",
  },
  style: {
    textColor: "000000",
    pageMarginsCm: {
      left: 3,
      right: 1,
      top: 2,
      bottom: 2,
    },
    paragraph: {
      lineSpacing: 1,
      firstLineCm: 1.25,
      spacingBeforePt: 0,
      spacingAfterPt: 0,
      removeBlankLinesBetweenParagraphs: true,
    },
    math: {
      align: "center",
      lineSpacing: 1,
      spacingBeforePt: 0,
      spacingAfterPt: 0,
      removeAdjacentBlankLines: true,
    },
    title: {
      fontSizePt: 16,
      lineSpacing: 1,
      spacingBeforePt: 0,
      spacingAfterPt: 0,
    },
    headings: {
      fontSizePt: 12,
      lineSpacing: 1,
      firstLineCm: 1.25,
      blankLinesAfter: 1,
      blankLinesBetweenConsecutive: 1,
      pageBreakBeforeTopLevel: false,
      pageBreakBeforeNumberedTopLevel: false,
      spacingBeforePt: 0,
      spacingAfterPt: 0,
    },
    listItem: {
      lineSpacing: 1,
      firstLineCm: 1.25,
      numberTextGapCm: 0.5,
      levelIndentCm: 0.75,
      spacingBeforePt: 0,
      spacingAfterPt: 0,
    },
    captions: {
      figureAlign: "center",
      tableAlign: "left",
      imageAlign: "left",
    },
    tables: {
      blankLinesAfter: 1,
    },
  },
  captions: {
    enabled: true,
    figurePrefix: "Сурет",
    tablePrefix: "Кесте",
  },
  converter: {
    gfm: true,
    ignoreImage: false,
    ignoreFootnote: false,
    ignoreHtml: false,
    math: {
      engine: "katex",
      libreOfficeCompat: true,
    },
    image: {
      maxWidthPx: 560,
    },
    theme: {
      heading1: "000000",
      heading2: "000000",
      heading3: "000000",
      heading4: "000000",
      heading5: "000000",
      heading6: "000000",
      link: "000000",
      code: "000000",
      tag: "000000",
      codespan: "000000",
      blockquote: "000000",
      html: "000000",
      del: "000000",
      hr: "000000",
      border: "000000",
      tableHeaderBackground: "D9D9D9",
      bodySize: 12,
      spaceSize: 12,
      heading1Size: 12,
      heading2Size: 12,
      heading3Size: 12,
      heading4Size: 12,
      heading5Size: 12,
      heading6Size: 12,
      lineSpacing: 1,
      linkUnderline: false,
    },
  },
  pandoc: {
    binary: "pandoc",
    from: "markdown+tex_math_dollars+pipe_tables+table_captions+fenced_code_blocks+strikeout+footnotes+yaml_metadata_block",
    standalone: true,
    referenceDoc: null,
    extraArgs: [],
  },
};
const IMAGE_TYPE_WHITELIST = new Set(["jpg", "jpeg", "png", "gif", "bmp"]);
const FIGURE_CAPTION_MARKER = "@@MD_TO_DOCX_FIGURE_CAPTION@@";
const TABLE_CAPTION_MARKER = "@@MD_TO_DOCX_TABLE_CAPTION@@";
const DOC_TITLE_MARKER = "@@MD_TO_DOCX_DOC_TITLE@@";
const DOC_SUBTITLE_MARKER = "@@MD_TO_DOCX_DOC_SUBTITLE@@";
const CLI_VALUE_OVERRIDES = createCliValueOverrides();

function printHelp() {
  console.log(`Usage:
  md-to-docx <input.md> [--output output.docx]
  md-to-docx <input.md|dir> [more inputs...] [--out-dir docx-output-dir]

Examples:
  md-to-docx docs/algorithm_explanations/angle_phi_subbytes_exact_kk.md
  md-to-docx docs/algorithm_explanations --out-dir docs/docx/algorithm_explanations

Options:
  -e, --engine <name>              Conversion engine: markdown-docx or pandoc.
  -o, --output <file>              Output DOCX path for one input file.
      --out-dir <dir>              Output directory for one or more inputs.
  -c, --config <file>              markdown-docx config JSON.
      --pandoc-bin <file>          Pandoc executable path. Default: pandoc.
      --reference-doc <file>       Pandoc reference DOCX.
      --keep-pandoc-title-block    Keep leading Pandoc % title block as text.
      --libreoffice-compat         Force simpler math output for LibreOffice.
      --ignore-images              Skip Markdown images.
      --max-image-width, -miw <px>  Scale wider images down to this width. Use 0 to disable.
      -np, -newPageSections        Start every top-level section on a new page.
      -no-np, -continueSections    Continue top-level sections on the same page.
      -ls, -lineSpacing <n>        Body line spacing.
      -fl, -firstLine <cm>         Body first-line indent in centimeters.
      -tc, -textColor <hex>        Default text color, for example 000000.
      -ml, -marginLeft <cm>        Left page margin in centimeters.
      -mr, -marginRight <cm>       Right page margin in centimeters.
      -mt, -marginTop <cm>         Top page margin in centimeters.
      -mb, -marginBottom <cm>      Bottom page margin in centimeters.
      -set <path=value>            Override any config value, for example:
                                   -set style.headings.blankLinesAfter=0
      --dry-run                    Print conversions without writing DOCX files.
  -h, --help                       Show this help.
`);
}

function createCliValueOverrides() {
  const aliases = new Map();
  const add = (names, pathSegments, type) => {
    for (const name of names) {
      aliases.set(name, { pathSegments, type });
    }
  };

  add(["-tc", "-textColor", "--text-color", "--textColor"], ["style", "textColor"], "string");
  add(["-ml", "-marginLeft", "--margin-left", "--marginLeft"], ["style", "pageMarginsCm", "left"], "number");
  add(["-mr", "-marginRight", "--margin-right", "--marginRight"], ["style", "pageMarginsCm", "right"], "number");
  add(["-mt", "-marginTop", "--margin-top", "--marginTop"], ["style", "pageMarginsCm", "top"], "number");
  add(["-mb", "-marginBottom", "--margin-bottom", "--marginBottom"], ["style", "pageMarginsCm", "bottom"], "number");

  add(["-ls", "-lineSpacing", "--line-spacing", "--lineSpacing"], ["style", "paragraph", "lineSpacing"], "number");
  add(["-fl", "-firstLine", "--first-line", "--firstLine"], ["style", "paragraph", "firstLineCm"], "number");
  add(["-pb", "-paragraphBefore", "--paragraph-before", "--paragraphBefore"], ["style", "paragraph", "spacingBeforePt"], "number");
  add(["-pa", "-paragraphAfter", "--paragraph-after", "--paragraphAfter"], ["style", "paragraph", "spacingAfterPt"], "number");
  add(["-pbl", "-removeParagraphBlankLines", "--remove-paragraph-blank-lines", "--removeParagraphBlankLines"], ["style", "paragraph", "removeBlankLinesBetweenParagraphs"], "boolean");

  add(["-ma", "-mathAlign", "--math-align", "--mathAlign"], ["style", "math", "align"], "string");
  add(["-mls", "-mathLineSpacing", "--math-line-spacing", "--mathLineSpacing"], ["style", "math", "lineSpacing"], "number");
  add(["-mbf", "-mathBefore", "--math-before", "--mathBefore"], ["style", "math", "spacingBeforePt"], "number");
  add(["-maf", "-mathAfter", "--math-after", "--mathAfter"], ["style", "math", "spacingAfterPt"], "number");
  add(["-mbl", "-removeMathBlankLines", "--remove-math-blank-lines", "--removeMathBlankLines"], ["style", "math", "removeAdjacentBlankLines"], "boolean");

  add(["-tfs", "-titleFontSize", "--title-font-size", "--titleFontSize"], ["style", "title", "fontSizePt"], "number");
  add(["-tls", "-titleLineSpacing", "--title-line-spacing", "--titleLineSpacing"], ["style", "title", "lineSpacing"], "number");
  add(["-tbf", "-titleBefore", "--title-before", "--titleBefore"], ["style", "title", "spacingBeforePt"], "number");
  add(["-taf", "-titleAfter", "--title-after", "--titleAfter"], ["style", "title", "spacingAfterPt"], "number");

  add(["-hfs", "-headingFontSize", "--heading-font-size", "--headingFontSize"], ["style", "headings", "fontSizePt"], "number");
  add(["-hls", "-headingLineSpacing", "--heading-line-spacing", "--headingLineSpacing"], ["style", "headings", "lineSpacing"], "number");
  add(["-hfl", "-headingFirstLine", "--heading-first-line", "--headingFirstLine"], ["style", "headings", "firstLineCm"], "number");
  add(["-hba", "-headingBlankAfter", "--heading-blank-after", "--headingBlankAfter"], ["style", "headings", "blankLinesAfter"], "integer");
  add(["-hbb", "-headingBlankBetween", "--heading-blank-between", "--headingBlankBetween"], ["style", "headings", "blankLinesBetweenConsecutive"], "integer");
  add(["-pbt", "-pageBreakTopLevel", "--page-break-top-level", "--pageBreakTopLevel"], ["style", "headings", "pageBreakBeforeTopLevel"], "boolean");
  add(["-pbn", "-pageBreakNumbered", "--page-break-numbered", "--pageBreakNumbered"], ["style", "headings", "pageBreakBeforeNumberedTopLevel"], "boolean");
  add(["-hbf", "-headingBefore", "--heading-before", "--headingBefore"], ["style", "headings", "spacingBeforePt"], "number");
  add(["-haf", "-headingAfter", "--heading-after", "--headingAfter"], ["style", "headings", "spacingAfterPt"], "number");

  add(["-lls", "-listLineSpacing", "--list-line-spacing", "--listLineSpacing"], ["style", "listItem", "lineSpacing"], "number");
  add(["-lfl", "-listFirstLine", "--list-first-line", "--listFirstLine"], ["style", "listItem", "firstLineCm"], "number");
  add(["-lg", "-listGap", "--list-gap", "--listGap"], ["style", "listItem", "numberTextGapCm"], "number");
  add(["-li", "-listIndent", "--list-indent", "--listIndent"], ["style", "listItem", "levelIndentCm"], "number");
  add(["-lbf", "-listBefore", "--list-before", "--listBefore"], ["style", "listItem", "spacingBeforePt"], "number");
  add(["-laf", "-listAfter", "--list-after", "--listAfter"], ["style", "listItem", "spacingAfterPt"], "number");

  add(["-fia", "-figureAlign", "--figure-align", "--figureAlign"], ["style", "captions", "figureAlign"], "string");
  add(["-tba", "-tableAlign", "--table-align", "--tableAlign"], ["style", "captions", "tableAlign"], "string");
  add(["-ia", "-imageAlign", "--image-align", "--imageAlign"], ["style", "captions", "imageAlign"], "string");
  add(["-cap", "-captions", "--captions"], ["captions", "enabled"], "boolean");
  add(["-fp", "-figurePrefix", "--figure-prefix", "--figurePrefix"], ["captions", "figurePrefix"], "string");
  add(["-tp", "-tablePrefix", "--table-prefix", "--tablePrefix"], ["captions", "tablePrefix"], "string");

  add(["-bs", "-bodySize", "--body-size", "--bodySize"], ["converter", "theme", "bodySize"], "number");
  add(["-ss", "-spaceSize", "--space-size", "--spaceSize"], ["converter", "theme", "spaceSize"], "number");
  add(["-thu", "-linkUnderline", "--link-underline", "--linkUnderline"], ["converter", "theme", "linkUnderline"], "boolean");
  add(["-thb", "-tableHeaderBackground", "--table-header-background", "--tableHeaderBackground"], ["converter", "theme", "tableHeaderBackground"], "string");

  return aliases;
}

function parseArgs(argv) {
  const options = {
    config: DEFAULT_CONFIG,
    dryRun: false,
    ignoreImages: false,
    keepPandocTitleBlock: false,
    libreOfficeCompat: false,
    maxImageWidth: null,
    output: null,
    outDir: null,
    configOverrides: {},
  };
  const inputs = [];

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    const override = CLI_VALUE_OVERRIDES.get(arg);
    if (override) {
      const rawValue = requireValue(argv, ++i, arg);
      setNestedOverride(
        options.configOverrides,
        override.pathSegments,
        parseOverrideValue(rawValue, override.type, arg),
      );
      continue;
    }

    switch (arg) {
      case "-h":
      case "--help":
        options.help = true;
        break;
      case "-e":
      case "-engine":
      case "--engine":
        setNestedOverride(options.configOverrides, ["engine"], normalizeEngineName(requireValue(argv, ++i, arg)));
        break;
      case "-o":
      case "--output":
        options.output = requireValue(argv, ++i, arg);
        break;
      case "--out-dir":
        options.outDir = requireValue(argv, ++i, arg);
        break;
      case "-c":
      case "--config":
        options.config = requireValue(argv, ++i, arg);
        break;
      case "-pandocBin":
      case "--pandoc-bin":
      case "--pandocBin":
        setNestedOverride(options.configOverrides, ["pandoc", "binary"], requireValue(argv, ++i, arg));
        break;
      case "-referenceDoc":
      case "--reference-doc":
      case "--referenceDoc":
        setNestedOverride(options.configOverrides, ["pandoc", "referenceDoc"], requireValue(argv, ++i, arg));
        break;
      case "--keep-pandoc-title-block":
        options.keepPandocTitleBlock = true;
        break;
      case "--libreoffice-compat":
        options.libreOfficeCompat = true;
        break;
      case "--ignore-images":
        options.ignoreImages = true;
        break;
      case "-miw":
      case "-maxImageWidth":
      case "--max-image-width":
        options.maxImageWidth = parseNonNegativeInteger(requireValue(argv, ++i, arg), arg);
        break;
      case "-np":
      case "-newPageSections":
      case "--new-page-sections":
      case "--newPageSections":
        setNestedOverride(options.configOverrides, ["style", "headings", "pageBreakBeforeTopLevel"], true);
        setNestedOverride(options.configOverrides, ["style", "headings", "pageBreakBeforeNumberedTopLevel"], true);
        break;
      case "-no-np":
      case "-continueSections":
      case "--continue-sections":
      case "--continueSections":
        setNestedOverride(options.configOverrides, ["style", "headings", "pageBreakBeforeTopLevel"], false);
        setNestedOverride(options.configOverrides, ["style", "headings", "pageBreakBeforeNumberedTopLevel"], false);
        break;
      case "-keepParagraphBlankLines":
      case "--keep-paragraph-blank-lines":
      case "--keepParagraphBlankLines":
        setNestedOverride(options.configOverrides, ["style", "paragraph", "removeBlankLinesBetweenParagraphs"], false);
        break;
      case "-no-mbl":
      case "-keepMathBlankLines":
      case "--keep-math-blank-lines":
      case "--keepMathBlankLines":
        setNestedOverride(options.configOverrides, ["style", "math", "removeAdjacentBlankLines"], false);
        break;
      case "-no-cap":
      case "-noCaptions":
      case "--no-captions":
      case "--noCaptions":
        setNestedOverride(options.configOverrides, ["captions", "enabled"], false);
        break;
      case "-set":
      case "--set":
        applyAssignmentOverride(options.configOverrides, requireValue(argv, ++i, arg), arg);
        break;
      case "--dry-run":
        options.dryRun = true;
        break;
      default:
        if (arg.startsWith("-")) {
          throw new Error(`Unknown option: ${arg}`);
        }
        inputs.push(arg);
    }
  }

  if (options.output && (options.outDir || inputs.length !== 1)) {
    throw new Error("--output can only be used with exactly one input file and without --out-dir.");
  }

  return { inputs, options };
}

function normalizeEngineName(rawEngine) {
  const engine = String(rawEngine ?? "").trim().toLowerCase();
  if (engine === "markdown-docx" || engine === "markdowndocx" || engine === "mdx") {
    return "markdown-docx";
  }
  if (engine === "pandoc") {
    return "pandoc";
  }
  throw new Error(`Unsupported engine: ${rawEngine}. Use markdown-docx or pandoc.`);
}

function setNestedOverride(root, pathSegments, value) {
  let target = root;
  for (let index = 0; index < pathSegments.length - 1; index += 1) {
    const segment = pathSegments[index];
    target[segment] ??= {};
    target = target[segment];
  }
  target[pathSegments[pathSegments.length - 1]] = value;
}

function applyAssignmentOverride(root, assignment, optionName) {
  const equalsIndex = assignment.indexOf("=");
  if (equalsIndex <= 0) {
    throw new Error(`${optionName} expects path=value.`);
  }

  const pathText = assignment.slice(0, equalsIndex).trim();
  const valueText = assignment.slice(equalsIndex + 1).trim();
  const pathSegments = pathText.split(".").map((segment) => segment.trim()).filter(Boolean);
  if (pathSegments.length === 0) {
    throw new Error(`${optionName} expects a non-empty config path.`);
  }

  setNestedOverride(root, pathSegments, parseInferredOverrideValue(valueText));
}

function parseOverrideValue(rawValue, type, optionName) {
  switch (type) {
    case "boolean":
      return parseBoolean(rawValue, optionName);
    case "integer":
      return parseNonNegativeInteger(rawValue, optionName);
    case "number":
      return parseFiniteNumber(rawValue, optionName);
    case "string":
      return String(rawValue);
    default:
      throw new Error(`Unsupported override type for ${optionName}: ${type}`);
  }
}

function parseInferredOverrideValue(rawValue) {
  if (/^(?:true|yes|on)$/i.test(rawValue)) {
    return true;
  }
  if (/^(?:false|no|off)$/i.test(rawValue)) {
    return false;
  }
  if (/^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(rawValue)) {
    return Number(rawValue);
  }
  return rawValue;
}

function parseBoolean(rawValue, optionName) {
  if (/^(?:true|yes|on|1)$/i.test(rawValue)) {
    return true;
  }
  if (/^(?:false|no|off|0)$/i.test(rawValue)) {
    return false;
  }
  throw new Error(`${optionName} expects a boolean value: true or false.`);
}

function requireValue(argv, index, optionName) {
  const value = argv[index];
  if (!value || value.startsWith("-")) {
    throw new Error(`${optionName} expects a value.`);
  }
  return value;
}

function parseFiniteNumber(rawValue, optionName) {
  const value = Number(rawValue);
  if (!Number.isFinite(value)) {
    throw new Error(`${optionName} expects a number.`);
  }
  return value;
}

function parseNonNegativeInteger(rawValue, optionName) {
  const value = Number(rawValue);
  if (!Number.isInteger(value) || value < 0) {
    throw new Error(`${optionName} expects a non-negative integer.`);
  }
  return value;
}

async function loadConfig(configPath) {
  const absoluteConfigPath = path.resolve(configPath);
  try {
    const raw = await fs.readFile(absoluteConfigPath, "utf8");
    const parsed = JSON.parse(raw);
    if (parsed.converter || parsed.document || parsed.style || parsed.captions || parsed.pandoc || parsed.engine) {
      return {
        engine: parsed.engine ?? DEFAULT_CONFIG_DATA.engine,
        converter: parsed.converter ?? {},
        document: parsed.document ?? {},
        style: parsed.style ?? {},
        captions: parsed.captions ?? {},
        pandoc: parsed.pandoc ?? {},
      };
    }
    return {
      engine: DEFAULT_CONFIG_DATA.engine,
      converter: parsed,
      document: {},
      style: {},
      captions: {},
      pandoc: {},
    };
  } catch (error) {
    if (error.code === "ENOENT" && absoluteConfigPath === DEFAULT_CONFIG) {
      return structuredClone(DEFAULT_CONFIG_DATA);
    }
    throw error;
  }
}

async function collectMarkdownFiles(inputPath) {
  const absolute = path.resolve(inputPath);
  const stat = await fs.stat(absolute);
  if (stat.isFile()) {
    if (path.extname(absolute).toLowerCase() !== ".md") {
      throw new Error(`Input is not a Markdown file: ${inputPath}`);
    }
    return [{ file: absolute, root: path.dirname(absolute) }];
  }

  if (!stat.isDirectory()) {
    throw new Error(`Input is neither a file nor a directory: ${inputPath}`);
  }

  const files = await walkMarkdown(absolute);
  return files.map((file) => ({ file, root: absolute }));
}

async function walkMarkdown(directory) {
  const entries = await fs.readdir(directory, { withFileTypes: true });
  const results = [];

  for (const entry of entries) {
    if (entry.name.startsWith(".")) {
      continue;
    }

    const fullPath = path.join(directory, entry.name);
    if (entry.isDirectory()) {
      results.push(...await walkMarkdown(fullPath));
      continue;
    }

    if (entry.isFile() && path.extname(entry.name).toLowerCase() === ".md") {
      results.push(fullPath);
    }
  }

  return results.sort();
}

function extractPandocTitleBlock(markdown) {
  const normalized = markdown.replace(/\r\n/g, "\n");
  const lines = normalized.split("\n");
  if (!lines[0]?.startsWith("%")) {
    return { markdown, metadata: {} };
  }

  const titleLines = [];
  let index = 0;
  while (index < lines.length && lines[index].startsWith("%")) {
    titleLines.push(lines[index].slice(1).trim());
    index += 1;
  }

  if (lines[index] === "") {
    index += 1;
  }

  return {
    markdown: lines.slice(index).join("\n"),
    metadata: {
      title: titleLines[0],
      creator: titleLines[1],
      date: titleLines[2],
    },
  };
}

function applyAcademicCaptions(markdown, captionsConfig) {
  const captions = {
    ...DEFAULT_CONFIG_DATA.captions,
    ...(captionsConfig ?? {}),
  };

  if (!captions.enabled) {
    return markdown;
  }

  const lines = markdown.replace(/\r\n/g, "\n").split("\n");
  const output = [];
  let figureNumber = 1;
  let tableNumber = 1;
  let pendingTableCaption = null;

  for (let index = 0; index < lines.length; index += 1) {
    const line = lines[index];
    const tableCaptionMatch = line.match(/^\s*<!--\s*(?:table-caption|кесте)\s*:\s*(.*?)\s*-->\s*$/i);
    if (tableCaptionMatch) {
      pendingTableCaption = tableCaptionMatch[1].trim();
      continue;
    }

    const imageMatch = line.match(/^(\s*)!\[([^\]]*)\]\((.+)\)\s*$/);
    if (imageMatch) {
      output.push(line);
      if (!nextLineStartsWith(lines, index, captions.figurePrefix)) {
        const title = stripCaptionPrefix(imageMatch[2]);
        output.push("");
        output.push(`${FIGURE_CAPTION_MARKER}${captions.figurePrefix} ${figureNumber}${title ? `. ${title}` : ""}`);
        output.push("");
      }
      figureNumber += 1;
      continue;
    }

    if (isMarkdownTableStart(lines, index)) {
      const table = collectMarkdownTable(lines, index);
      let caption = pendingTableCaption;
      pendingTableCaption = null;

      const afterTableLine = lines[table.endIndex + 1];
      const pandocCaptionMatch = afterTableLine?.match(/^\s*:\s*(.+?)\s*$/);
      if (!caption && pandocCaptionMatch) {
        caption = pandocCaptionMatch[1].trim();
      }

      if (!previousOutputLineStartsWith(output, captions.tablePrefix)) {
        output.push("");
        output.push(`${TABLE_CAPTION_MARKER}${captions.tablePrefix} ${tableNumber}`);
        if (caption) {
          output.push(`${TABLE_CAPTION_MARKER}${stripCaptionPrefix(caption)}`);
        }
        output.push("");
      }

      output.push(...table.lines);
      if (pandocCaptionMatch) {
        index = table.endIndex + 1;
      } else {
        index = table.endIndex;
      }
      tableNumber += 1;
      continue;
    }

    output.push(line);
  }

  return output.join("\n");
}

function nextLineStartsWith(lines, currentIndex, prefix) {
  for (let index = currentIndex + 1; index < lines.length; index += 1) {
    const line = lines[index].trim();
    if (!line) {
      continue;
    }
    return line.startsWith(prefix);
  }
  return false;
}

function previousOutputLineStartsWith(output, prefix) {
  for (let index = output.length - 1; index >= 0; index -= 1) {
    const line = output[index].trim();
    if (!line) {
      continue;
    }
    return line.startsWith(prefix) || line.startsWith(`${TABLE_CAPTION_MARKER}${prefix}`);
  }
  return false;
}

function isMarkdownTableStart(lines, index) {
  return isTableRow(lines[index]) && isTableSeparator(lines[index + 1]);
}

function isTableRow(line) {
  return typeof line === "string" && line.includes("|") && line.trim().length > 0;
}

function isTableSeparator(line) {
  if (typeof line !== "string") {
    return false;
  }
  return /^\s*\|?\s*:?-{3,}:?\s*(\|\s*:?-{3,}:?\s*)+\|?\s*$/.test(line);
}

function collectMarkdownTable(lines, startIndex) {
  let endIndex = startIndex + 1;
  while (endIndex + 1 < lines.length && isTableRow(lines[endIndex + 1])) {
    endIndex += 1;
  }

  return {
    lines: lines.slice(startIndex, endIndex + 1),
    endIndex,
  };
}

function stripCaptionPrefix(text) {
  return String(text ?? "")
    .trim()
    .replace(/^(?:figure|fig\.?|сурет|рисунок|кесте|table)\s*\d+[\.\):\-–—]?\s*/i, "")
    .trim();
}

function normalizePandocMarkdown(markdown, keepPandocTitleBlock) {
  if (keepPandocTitleBlock) {
    return { markdown, metadata: {} };
  }

  const { markdown: body, metadata } = extractPandocTitleBlock(markdown);
  if (!metadata.title) {
    return { markdown, metadata };
  }

  const header = [`# ${metadata.title}`];
  const subtitle = [metadata.creator, metadata.date].filter(Boolean).join(" - ");
  header[0] = `# ${DOC_TITLE_MARKER}${metadata.title}`;
  if (subtitle) {
    header.push("", `_${DOC_SUBTITLE_MARKER}${subtitle}_`);
  }

  return {
    markdown: `${header.join("\n")}\n\n${body.trimStart()}`,
    metadata,
  };
}

function outputPathFor(item, allItems, options) {
  if (options.output) {
    return path.resolve(options.output);
  }

  const sourceParsed = path.parse(item.file);
  if (!options.outDir) {
    return path.join(sourceParsed.dir, `${sourceParsed.name}.docx`);
  }

  const outputRoot = path.resolve(options.outDir);
  const relative = allItems.length === 1
    ? `${sourceParsed.name}.docx`
    : path.relative(item.root, path.join(sourceParsed.dir, `${sourceParsed.name}.docx`));

  return path.join(outputRoot, relative);
}

function applyCliOverrides(converterConfig, options) {
  const config = structuredClone(converterConfig);

  if (options.ignoreImages) {
    config.ignoreImage = true;
  }

  if (options.libreOfficeCompat) {
    config.math = {
      ...(config.math ?? {}),
      libreOfficeCompat: true,
    };
  }

  return config;
}

function createDocumentConfig(source, normalized, baseConfig) {
  return {
    ...baseConfig.document,
    title: normalized.metadata.title ?? baseConfig.document.title ?? path.basename(source),
    creator: normalized.metadata.creator ?? baseConfig.document.creator,
    description: baseConfig.document.description,
  };
}

async function renderWithMarkdownDocx(markdown, source, normalized, baseConfig, options) {
  const converterConfig = applyCliOverrides(baseConfig.converter, options);
  const maxImageWidthPx = getMaxImageWidthPx(converterConfig, options);
  delete converterConfig.image;
  if (!converterConfig.ignoreImage) {
    converterConfig.imageAdapter = createImageAdapter(path.dirname(source), maxImageWidthPx);
  }

  const converter = new MarkdownDocx(markdown, converterConfig);
  const document = await converter.toDocument(createDocumentConfig(source, normalized, baseConfig));
  return Packer.toBuffer(document);
}

async function renderWithPandoc(markdown, source, baseConfig, options) {
  if (options.ignoreImages) {
    markdown = removeMarkdownImages(markdown);
  }

  const pandocConfig = mergeDeep(DEFAULT_CONFIG_DATA.pandoc, baseConfig.pandoc ?? {});
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "md-to-docx-pandoc-"));
  const tempDocx = path.join(tempDir, "output.docx");
  const sourceDir = path.dirname(source);

  const args = [
    "-f",
    pandocConfig.from || DEFAULT_CONFIG_DATA.pandoc.from,
    "-t",
    "docx",
    "-o",
    tempDocx,
  ];

  if (pandocConfig.standalone !== false) {
    args.push("--standalone");
  }

  if (pandocConfig.referenceDoc) {
    args.push(`--reference-doc=${path.resolve(pandocConfig.referenceDoc)}`);
  }

  const resourcePath = [sourceDir, process.cwd()].filter(Boolean).join(path.delimiter);
  args.push(`--resource-path=${resourcePath}`);

  if (Array.isArray(pandocConfig.extraArgs)) {
    args.push(...pandocConfig.extraArgs.map(String));
  }

  args.push("-");

  try {
    await runPandoc(pandocConfig.binary || "pandoc", args, markdown, sourceDir);
    return await fs.readFile(tempDocx);
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
}

function removeMarkdownImages(markdown) {
  return String(markdown ?? "")
    .replace(/!\[[^\]]*]\([^)]+\)/g, "")
    .replace(/<img\b[^>]*>/gi, "");
}

function runPandoc(binary, args, input, cwd) {
  return new Promise((resolve, reject) => {
    const child = spawn(binary, args, {
      cwd,
      stdio: ["pipe", "pipe", "pipe"],
    });

    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
    });

    child.on("error", (error) => {
      if (error.code === "ENOENT") {
        reject(new Error(`Pandoc executable not found: ${binary}. Install pandoc or use --engine markdown-docx.`));
        return;
      }
      reject(error);
    });

    child.on("close", (code) => {
      if (code === 0) {
        resolve();
        return;
      }

      const details = [stderr.trim(), stdout.trim()].filter(Boolean).join("\n");
      reject(new Error(`Pandoc failed with exit code ${code}${details ? `:\n${details}` : "."}`));
    });

    child.stdin.end(input);
  });
}

function createImageAdapter(sourceDir, maxWidthPx) {
  return async function imageAdapter(token) {
    const href = token.href;
    if (!href) {
      return null;
    }

    try {
      const { buffer, resolvedSource } = await loadImageBuffer(href, sourceDir);
      const size = imageSize(buffer);
      const type = normalizeImageType(size.type, resolvedSource);
      if (!type) {
        return null;
      }

      const dimensions = scaleImageDimensions(size.width, size.height, maxWidthPx);

      return {
        type,
        data: buffer,
        width: dimensions.width,
        height: dimensions.height,
      };
    } catch (error) {
      console.error(`[md-to-docx] imageLoadError ${href}: ${error.message}`);
      return null;
    }
  };
}

function scaleImageDimensions(width, height, maxWidthPx) {
  if (!maxWidthPx || width <= maxWidthPx) {
    return { width, height };
  }

  const scale = maxWidthPx / width;
  return {
    width: Math.round(width * scale),
    height: Math.round(height * scale),
  };
}

async function loadImageBuffer(href, sourceDir) {
  if (isHttp(href)) {
    return {
      buffer: await downloadHttpImage(href),
      resolvedSource: href,
    };
  }

  const imagePath = resolveLocalImagePath(href, sourceDir);
  return {
    buffer: await fs.readFile(imagePath),
    resolvedSource: imagePath,
  };
}

function resolveLocalImagePath(href, sourceDir) {
  if (href.startsWith("file://")) {
    return fileURLToPath(href);
  }

  const cleanHref = stripUrlTail(safeDecodeUri(href));
  if (path.isAbsolute(cleanHref)) {
    return cleanHref;
  }

  return path.resolve(sourceDir, cleanHref);
}

function stripUrlTail(href) {
  const queryIndex = href.indexOf("?");
  const hashIndex = href.indexOf("#");
  const indexes = [queryIndex, hashIndex].filter((index) => index >= 0);
  if (indexes.length === 0) {
    return href;
  }
  return href.slice(0, Math.min(...indexes));
}

function safeDecodeUri(value) {
  try {
    return decodeURI(value);
  } catch {
    return value;
  }
}

function isHttp(src) {
  return /^https?:\/\//i.test(src);
}

function downloadHttpImage(url) {
  return new Promise((resolve, reject) => {
    const client = url.startsWith("https://") ? https : http;
    client.get(url, (response) => {
      if (response.statusCode >= 300 && response.statusCode < 400 && response.headers.location) {
        response.resume();
        downloadHttpImage(new URL(response.headers.location, url).toString()).then(resolve, reject);
        return;
      }

      if (response.statusCode < 200 || response.statusCode >= 300) {
        response.resume();
        reject(new Error(`HTTP ${response.statusCode}`));
        return;
      }

      const chunks = [];
      response.on("data", (chunk) => chunks.push(chunk));
      response.on("end", () => resolve(Buffer.concat(chunks)));
    }).on("error", reject);
  });
}

function normalizeImageType(detectedType, source) {
  const type = (detectedType || path.extname(stripUrlTail(source)).slice(1)).toLowerCase();
  if (type === "webp") {
    console.error("[md-to-docx] WebP images are not supported in the standalone Node converter.");
    return null;
  }
  if (!IMAGE_TYPE_WHITELIST.has(type)) {
    console.error(`[md-to-docx] Unsupported image type: ${type || "unknown"}`);
    return null;
  }
  return type === "jpeg" ? "jpg" : type;
}

async function applyDocxHouseStyle(buffer, styleConfig) {
  const style = mergeDeep(DEFAULT_CONFIG_DATA.style, styleConfig ?? {});
  const zip = await JSZip.loadAsync(buffer);

  const documentFile = zip.file("word/document.xml");
  if (documentFile) {
    let documentXml = await documentFile.async("string");
    documentXml = applyPageMargins(documentXml, style.pageMarginsCm);
    documentXml = applyDocumentTitleParagraphs(documentXml, style.title);
    documentXml = applyHeadingLayout(documentXml, style.headings);
    documentXml = applyMathParagraphLayout(documentXml, style.math);
    documentXml = applyParagraphBlankLineLayout(documentXml, style.paragraph);
    documentXml = applyTableSpacingAfter(documentXml, style.tables);
    documentXml = applyImageParagraphProperties(documentXml, style.captions);
    documentXml = applyListItemParagraphProperties(documentXml, style.listItem);
    documentXml = applyCaptionParagraphProperties(documentXml, style.captions);
    zip.file("word/document.xml", documentXml);
  }

  const numberingFile = zip.file("word/numbering.xml");
  if (numberingFile) {
    let numberingXml = await numberingFile.async("string");
    numberingXml = applyNumberingIndents(numberingXml, style.listItem);
    zip.file("word/numbering.xml", numberingXml);
  }

  const stylesFile = zip.file("word/styles.xml");
  if (stylesFile) {
    let stylesXml = await stylesFile.async("string");
    stylesXml = applyTextColor(stylesXml, style.textColor);
    stylesXml = applyParagraphDefaults(stylesXml, style.paragraph);
    stylesXml = applySpaceStyle(stylesXml, style.paragraph);
    stylesXml = applyBodyParagraphStyle(stylesXml, style.paragraph);
    stylesXml = applyTitleStyle(stylesXml, style.title, style.textColor);
    stylesXml = applyHeadingStyles(stylesXml, style.headings, style.textColor);
    stylesXml = applyListItemStyle(stylesXml, style.listItem);
    stylesXml = applyTableParagraphStyles(stylesXml, style.paragraph);
    zip.file("word/styles.xml", stylesXml);
  }

  return zip.generateAsync({ type: "nodebuffer" });
}

function mergeDeep(base, override) {
  const result = structuredClone(base);
  for (const [key, value] of Object.entries(override ?? {})) {
    if (value && typeof value === "object" && !Array.isArray(value) && result[key] && typeof result[key] === "object") {
      result[key] = mergeDeep(result[key], value);
    } else {
      result[key] = value;
    }
  }
  return result;
}

function applyPageMargins(xml, marginsCm) {
  if (!marginsCm) {
    return xml;
  }

  const attrs = {
    "w:top": cmToTwips(marginsCm.top ?? 2),
    "w:right": cmToTwips(marginsCm.right ?? 1),
    "w:bottom": cmToTwips(marginsCm.bottom ?? 2),
    "w:left": cmToTwips(marginsCm.left ?? 3),
    "w:header": 708,
    "w:footer": 708,
    "w:gutter": 0,
  };

  if (/<w:pgMar\b[^>]*\/>/.test(xml)) {
    return xml.replace(/<w:pgMar\b([^>]*)\/>/g, (_match, attrText) => {
      return `<w:pgMar${setXmlAttributes(attrText, attrs)}/>`;
    });
  }

  return xml.replace(/<w:sectPr\b([^>]*)>/, (match) => {
    return `${match}<w:pgMar${formatXmlAttributes(attrs)}/>`;
  });
}

function applyTextColor(stylesXml, textColor) {
  const color = normalizeHexColor(textColor ?? "000000");
  let xml = stylesXml.replace(/<w:color\b[^>]*\/>/g, (match) => {
    return match.replace(/w:val="[^"]*"/, `w:val="${color}"`);
  });

  xml = xml.replace(/(<w:rPrDefault><w:rPr>)([\s\S]*?)(<\/w:rPr><\/w:rPrDefault>)/, (_match, start, inner, end) => {
    return `${start}${setOrAddEmptyElement(inner, "w:color", { "w:val": color })}${end}`;
  });

  return xml;
}

function applyParagraphDefaults(stylesXml, paragraphStyle) {
  const line = lineSpacingToTwips(paragraphStyle?.lineSpacing ?? 1);
  const before = pointsToTwips(paragraphStyle?.spacingBeforePt ?? 0);
  const after = pointsToTwips(paragraphStyle?.spacingAfterPt ?? 0);

  return stylesXml.replace(/(<w:pPrDefault><w:pPr>)([\s\S]*?)(<\/w:pPr><\/w:pPrDefault>)/, (_match, start, inner, end) => {
    const updated = setOrAddEmptyElement(inner, "w:spacing", {
      "w:before": before,
      "w:after": after,
      "w:line": line,
      "w:lineRule": "auto",
    });
    return `${start}${updated}${end}`;
  });
}

function applySpaceStyle(stylesXml, paragraphStyle) {
  const line = lineSpacingToTwips(paragraphStyle?.lineSpacing ?? 1);
  return patchStyleParagraphProperties(stylesXml, "MdSpace", (inner) => {
    return setOrAddEmptyElement(inner, "w:spacing", {
      "w:before": 0,
      "w:after": 0,
      "w:line": line,
      "w:lineRule": "auto",
    });
  }).replace(
    /(<w:style\b(?=[^>]*w:styleId="MdSpace")[\s\S]*?<w:rPr>)([\s\S]*?)(<\/w:rPr>[\s\S]*?<\/w:style>)/,
    (_match, start, inner, end) => {
      let updated = setOrAddEmptyElement(inner, "w:sz", { "w:val": 24 });
      updated = setOrAddEmptyElement(updated, "w:szCs", { "w:val": 24 });
      return `${start}${updated}${end}`;
    },
  );
}

function applyBodyParagraphStyle(stylesXml, paragraphStyle) {
  const line = lineSpacingToTwips(paragraphStyle?.lineSpacing ?? 1);
  const before = pointsToTwips(paragraphStyle?.spacingBeforePt ?? 0);
  const after = pointsToTwips(paragraphStyle?.spacingAfterPt ?? 0);
  const firstLine = cmToTwips(paragraphStyle?.firstLineCm ?? 1.25);

  return patchStyleParagraphProperties(stylesXml, "MdParagraph", (inner) => {
    let updated = setOrAddEmptyElement(inner, "w:spacing", {
      "w:before": before,
      "w:after": after,
      "w:line": line,
      "w:lineRule": "auto",
    });
    updated = setOrAddEmptyElement(updated, "w:ind", { "w:firstLine": firstLine });
    updated = removeAttributesFromEmptyElement(updated, "w:ind", ["w:left", "w:hanging"]);
    return updated;
  });
}

function applyTitleStyle(stylesXml, titleStyle, textColor) {
  const fontSizePt = Number(titleStyle?.fontSizePt ?? 16);
  const line = lineSpacingToTwips(titleStyle?.lineSpacing ?? 1);
  const before = pointsToTwips(titleStyle?.spacingBeforePt ?? 0);
  const after = pointsToTwips(titleStyle?.spacingAfterPt ?? 0);
  const color = normalizeHexColor(textColor ?? "000000");

  stylesXml = patchStyleParagraphProperties(stylesXml, "Title", (inner) => {
    let updated = setOrAddEmptyElement(inner, "w:spacing", {
      "w:before": before,
      "w:after": after,
      "w:line": line,
      "w:lineRule": "auto",
    });
    updated = setOrAddEmptyElement(updated, "w:ind", { "w:firstLine": 0 });
    return updated;
  });

  return patchStyleRunProperties(stylesXml, "Title", (inner) => {
    let updated = removeEmptyElement(inner, "w:i");
    updated = removeEmptyElement(updated, "w:iCs");
    updated = setOrAddEmptyElement(updated, "w:b", { "w:val": "true" });
    updated = setOrAddEmptyElement(updated, "w:bCs", { "w:val": "true" });
    updated = setOrAddEmptyElement(updated, "w:color", { "w:val": color });
    updated = setOrAddEmptyElement(updated, "w:sz", { "w:val": pointsToHalfPoints(fontSizePt) });
    updated = setOrAddEmptyElement(updated, "w:szCs", { "w:val": pointsToHalfPoints(fontSizePt) });
    return updated;
  });
}

function applyHeadingStyles(stylesXml, headingStyle, textColor) {
  const fontSizePt = Number(headingStyle?.fontSizePt ?? 12);
  const line = lineSpacingToTwips(headingStyle?.lineSpacing ?? 1);
  const before = pointsToTwips(headingStyle?.spacingBeforePt ?? 0);
  const after = pointsToTwips(headingStyle?.spacingAfterPt ?? 0);
  const firstLine = cmToTwips(headingStyle?.firstLineCm ?? 1.25);
  const color = normalizeHexColor(textColor ?? "000000");

  const styleIds = [];
  for (let level = 1; level <= 6; level += 1) {
    styleIds.push(`Heading${level}`, `MdHeading${level}`);
  }

  for (const styleId of styleIds) {
    stylesXml = patchStyleParagraphProperties(stylesXml, styleId, (inner) => {
      let updated = setOrAddEmptyElement(inner, "w:spacing", {
        "w:before": before,
        "w:after": after,
        "w:line": line,
        "w:lineRule": "auto",
      });
      updated = setOrAddEmptyElement(updated, "w:ind", { "w:firstLine": firstLine });
      return updated;
    });

    stylesXml = patchStyleRunProperties(stylesXml, styleId, (inner) => {
      let updated = removeEmptyElement(inner, "w:i");
      updated = removeEmptyElement(updated, "w:iCs");
      updated = setOrAddEmptyElement(updated, "w:b", { "w:val": "true" });
      updated = setOrAddEmptyElement(updated, "w:bCs", { "w:val": "true" });
      updated = setOrAddEmptyElement(updated, "w:color", { "w:val": color });
      updated = setOrAddEmptyElement(updated, "w:sz", { "w:val": pointsToHalfPoints(fontSizePt) });
      updated = setOrAddEmptyElement(updated, "w:szCs", { "w:val": pointsToHalfPoints(fontSizePt) });
      return updated;
    });
  }

  return stylesXml;
}

function applyListItemStyle(stylesXml, listItemStyle) {
  const line = lineSpacingToTwips(listItemStyle?.lineSpacing ?? 1);
  const before = pointsToTwips(listItemStyle?.spacingBeforePt ?? 0);
  const after = pointsToTwips(listItemStyle?.spacingAfterPt ?? 0);
  const indent = getListIndentAttributes(listItemStyle, 0);

  return patchStyleParagraphProperties(stylesXml, "MdListItem", (inner) => {
    let updated = setOrAddEmptyElement(inner, "w:spacing", {
      "w:before": before,
      "w:after": after,
      "w:line": line,
      "w:lineRule": "auto",
    });
    updated = setOrAddEmptyElement(updated, "w:ind", indent);
    updated = removeAttributesFromEmptyElement(updated, "w:ind", ["w:firstLine"]);
    return updated;
  });
}

function applyTableParagraphStyles(stylesXml, paragraphStyle) {
  const line = lineSpacingToTwips(paragraphStyle?.lineSpacing ?? 1);
  const before = pointsToTwips(0);
  const after = pointsToTwips(0);

  for (const styleId of ["MdTable", "MdTableHeader", "MdTableCell"]) {
    stylesXml = patchStyleParagraphProperties(stylesXml, styleId, (inner) => {
      let updated = setOrAddEmptyElement(inner, "w:spacing", {
        "w:before": before,
        "w:after": after,
        "w:line": line,
        "w:lineRule": "auto",
      });
      updated = setOrAddEmptyElement(updated, "w:ind", { "w:firstLine": 0 });
      return updated;
    });
  }

  return stylesXml;
}

function applyListItemParagraphProperties(documentXml, listItemStyle) {
  const line = lineSpacingToTwips(listItemStyle?.lineSpacing ?? 1);
  const before = pointsToTwips(listItemStyle?.spacingBeforePt ?? 0);
  const after = pointsToTwips(listItemStyle?.spacingAfterPt ?? 0);

  return documentXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraphXml) => {
    if (!paragraphXml.includes('<w:pStyle w:val="MdListItem"')) {
      return paragraphXml;
    }

    const levelMatch = paragraphXml.match(/<w:ilvl w:val="(\d+)"/);
    const level = levelMatch ? Number(levelMatch[1]) : 0;
    const indent = getListIndentAttributes(listItemStyle, level);

    return patchParagraphProperties(paragraphXml, (inner) => {
      let updated = setOrAddEmptyElement(inner, "w:spacing", {
        "w:before": before,
        "w:after": after,
        "w:line": line,
        "w:lineRule": "auto",
      });
      updated = setOrAddEmptyElement(updated, "w:ind", indent);
      updated = removeAttributesFromEmptyElement(updated, "w:ind", ["w:firstLine"]);
      return updated;
    });
  });
}

function applyImageParagraphProperties(documentXml, captionStyle) {
  const alignment = captionStyle?.imageAlign ?? "left";
  return documentXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraphXml) => {
    if (!paragraphXml.includes("<w:drawing>")) {
      return paragraphXml;
    }

    return patchParagraphProperties(paragraphXml, (inner) => {
      let updated = setOrAddEmptyElement(inner, "w:jc", { "w:val": alignment });
      updated = setOrAddEmptyElement(updated, "w:spacing", {
        "w:before": 0,
        "w:after": 0,
        "w:line": 240,
        "w:lineRule": "auto",
      });
      updated = setOrAddEmptyElement(updated, "w:ind", {
        "w:firstLine": 0,
        "w:left": 0,
        "w:right": 0,
      });
      return updated;
    });
  });
}

function applyMathParagraphLayout(documentXml, mathStyle) {
  const alignment = mathStyle?.align ?? "center";
  const line = lineSpacingToTwips(mathStyle?.lineSpacing ?? 1);
  const before = pointsToTwips(mathStyle?.spacingBeforePt ?? 0);
  const after = pointsToTwips(mathStyle?.spacingAfterPt ?? 0);
  const removeAdjacentBlankLines = mathStyle?.removeAdjacentBlankLines !== false;

  return documentXml.replace(/<w:body>([\s\S]*?)<\/w:body>/, (_match, bodyXml) => {
    let parts = splitBodyParts(bodyXml).map((part) => {
      if (!isMathOnlyParagraph(part)) {
        return part;
      }

      return patchParagraphProperties(part, (inner) => {
        let updated = setOrAddEmptyElement(inner, "w:jc", { "w:val": alignment });
        updated = setOrAddEmptyElement(updated, "w:spacing", {
          "w:before": before,
          "w:after": after,
          "w:line": line,
          "w:lineRule": "auto",
        });
        updated = setOrAddEmptyElement(updated, "w:ind", {
          "w:firstLine": 0,
          "w:left": 0,
          "w:right": 0,
        });
        return updated;
      });
    });

    if (removeAdjacentBlankLines) {
      parts = removeSpaceParagraphsAdjacentToMath(parts);
    }

    return `<w:body>${parts.join("")}</w:body>`;
  });
}

function removeSpaceParagraphsAdjacentToMath(parts) {
  return parts.filter((part, index) => {
    if (!isSpaceParagraph(part)) {
      return true;
    }

    const previous = findPreviousNonSpacePart(parts, index);
    const next = findNextNonSpacePart(parts, index);

    if (isMathOnlyParagraph(next) && !isHeadingParagraph(previous)) {
      return false;
    }

    if (isMathOnlyParagraph(previous) && !isHeadingParagraph(next)) {
      return false;
    }

    return true;
  });
}

function applyParagraphBlankLineLayout(documentXml, paragraphStyle) {
  if (paragraphStyle?.removeBlankLinesBetweenParagraphs === false) {
    return documentXml;
  }

  return documentXml.replace(/<w:body>([\s\S]*?)<\/w:body>/, (_match, bodyXml) => {
    const parts = splitBodyParts(bodyXml).filter((part, index, allParts) => {
      if (!isSpaceParagraph(part)) {
        return true;
      }

      const previous = findPreviousNonSpacePart(allParts, index);
      const next = findNextNonSpacePart(allParts, index);
      return !(isFlowTextParagraph(previous) && isFlowTextParagraph(next));
    });

    return `<w:body>${parts.join("")}</w:body>`;
  });
}

function findPreviousNonSpacePart(parts, index) {
  for (let current = index - 1; current >= 0; current -= 1) {
    if (!isSpaceParagraph(parts[current])) {
      return parts[current];
    }
  }
  return "";
}

function findNextNonSpacePart(parts, index) {
  for (let current = index + 1; current < parts.length; current += 1) {
    if (!isSpaceParagraph(parts[current])) {
      return parts[current];
    }
  }
  return "";
}

function applyTableSpacingAfter(documentXml, tableStyle) {
  const blankAfter = normalizedNonNegativeInteger(tableStyle?.blankLinesAfter ?? 1);

  return documentXml.replace(/<w:body>([\s\S]*?)<\/w:body>/, (_match, bodyXml) => {
    const parts = splitBodyParts(bodyXml);
    const output = [];

    for (let index = 0; index < parts.length; index += 1) {
      const current = parts[index];
      output.push(current);

      if (!isTablePart(current)) {
        continue;
      }

      let nextIndex = index + 1;
      while (isSpaceParagraph(parts[nextIndex])) {
        nextIndex += 1;
      }

      const next = parts[nextIndex];
      if (blankAfter > 0 && next && !isSectionPropertiesPart(next)) {
        output.push(...Array.from({ length: blankAfter }, createSpaceParagraph));
      }
      index = nextIndex - 1;
    }

    return `<w:body>${output.join("")}</w:body>`;
  });
}

function applyHeadingLayout(documentXml, headingStyle) {
  const blankAfter = normalizedNonNegativeInteger(headingStyle?.blankLinesAfter ?? 1);
  const blankBetween = normalizedNonNegativeInteger(headingStyle?.blankLinesBetweenConsecutive ?? 1);
  const pageBreakBeforeTopLevel = headingStyle?.pageBreakBeforeTopLevel === true;
  const pageBreakBeforeNumberedTopLevel = headingStyle?.pageBreakBeforeNumberedTopLevel === true;
  return documentXml.replace(/<w:body>([\s\S]*?)<\/w:body>/, (_match, bodyXml) => {
    const paragraphs = splitBodyParts(bodyXml);
    const output = [];

    for (let index = 0; index < paragraphs.length; index += 1) {
      let current = paragraphs[index];

      if (!isHeadingParagraph(current)) {
        output.push(current);
        continue;
      }

      if (
        (pageBreakBeforeTopLevel && isTopLevelHeading(current))
        || (pageBreakBeforeNumberedTopLevel && isNumberedTopLevelHeading(current))
      ) {
        current = addPageBreakBefore(current);
      }

      output.push(current);

      let nextIndex = index + 1;
      while (isSpaceParagraph(paragraphs[nextIndex])) {
        nextIndex += 1;
      }

      const next = paragraphs[nextIndex];
      const desiredSpaces = isHeadingParagraph(next) ? blankBetween : blankAfter;
      output.push(...Array.from({ length: desiredSpaces }, createSpaceParagraph));
      index = nextIndex - 1;
    }

    return `<w:body>${output.join("")}</w:body>`;
  });
}

function splitBodyParts(bodyXml) {
  return bodyXml.match(/<w:tbl\b[\s\S]*?<\/w:tbl>|<w:p\b[\s\S]*?<\/w:p>|[\s\S]*?(?=<w:tbl\b|<w:p\b|$)/g)
    ?.filter((part) => part.length > 0) ?? [];
}

function normalizedNonNegativeInteger(value) {
  const number = Number(value);
  return Number.isInteger(number) && number >= 0 ? number : 0;
}

function isHeadingParagraph(paragraphXml) {
  return /<w:pStyle w:val="(?:Md)?Heading[1-6]"/.test(paragraphXml ?? "");
}

function isTopLevelHeading(paragraphXml) {
  return /<w:pStyle w:val="(?:Md)?Heading1"/.test(paragraphXml ?? "");
}

function isSpaceParagraph(paragraphXml) {
  return /<w:pStyle w:val="MdSpace"/.test(paragraphXml ?? "");
}

function isTablePart(partXml) {
  return /<w:tbl\b/.test(partXml ?? "");
}

function isSectionPropertiesPart(partXml) {
  return /<w:sectPr\b/.test(partXml ?? "");
}

function isMathOnlyParagraph(paragraphXml) {
  return /<w:p\b/.test(paragraphXml ?? "")
    && paragraphXml.includes("<m:oMath")
    && extractParagraphText(paragraphXml).trim() === "";
}

function isFlowTextParagraph(paragraphXml) {
  return /<w:pStyle w:val="MdParagraph"/.test(paragraphXml ?? "")
    && !isHeadingParagraph(paragraphXml)
    && !isMathOnlyParagraph(paragraphXml)
    && !paragraphXml.includes("<w:drawing>")
    && !paragraphXml.includes(FIGURE_CAPTION_MARKER)
    && !paragraphXml.includes(TABLE_CAPTION_MARKER);
}

function isNumberedTopLevelHeading(paragraphXml) {
  if (!/<w:pStyle w:val="(?:Md)?Heading1"/.test(paragraphXml ?? "")) {
    return false;
  }

  return /^\s*\d+\./.test(extractParagraphText(paragraphXml));
}

function extractParagraphText(paragraphXml) {
  const textParts = [];
  for (const match of paragraphXml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)) {
    textParts.push(decodeXmlText(match[1]));
  }
  return textParts.join("");
}

function decodeXmlText(text) {
  return String(text ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&amp;", "&")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'");
}

function addPageBreakBefore(paragraphXml) {
  return patchParagraphProperties(paragraphXml, (inner) => {
    return setOrAddEmptyElement(inner, "w:pageBreakBefore", {});
  });
}

function createSpaceParagraph() {
  return '<w:p><w:pPr><w:pStyle w:val="MdSpace"/></w:pPr></w:p>';
}

function applyNumberingIndents(numberingXml, listItemStyle) {
  return numberingXml.replace(/<w:lvl\b(?=[^>]*w:ilvl="(\d+)")[\s\S]*?<\/w:lvl>/g, (levelXml, rawLevel) => {
    const level = Number(rawLevel);
    const indent = getListIndentAttributes(listItemStyle, level);
    return levelXml.replace(/<w:ind\b([^>]*)\/>/, (_match, attrText) => {
      return `<w:ind${setXmlAttributes(attrText, indent)}/>`;
    });
  });
}

function getListIndentAttributes(listItemStyle, level) {
  const markerStart = cmToTwips(listItemStyle?.firstLineCm ?? 1.25);
  const gap = cmToTwips(listItemStyle?.numberTextGapCm ?? 0.5);
  const levelIndent = cmToTwips(listItemStyle?.levelIndentCm ?? 0.75);
  return {
    "w:left": markerStart + gap + level * levelIndent,
    "w:hanging": gap,
  };
}

function applyDocumentTitleParagraphs(documentXml, titleStyle) {
  return documentXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraphXml) => {
    if (paragraphXml.includes(DOC_TITLE_MARKER)) {
      return patchDocumentTitleParagraph(paragraphXml, titleStyle);
    }
    if (paragraphXml.includes(DOC_SUBTITLE_MARKER)) {
      return patchDocumentSubtitleParagraph(paragraphXml, titleStyle);
    }
    return paragraphXml;
  });
}

function patchDocumentTitleParagraph(paragraphXml, titleStyle) {
  let xml = paragraphXml.replaceAll(DOC_TITLE_MARKER, "");
  return patchParagraphProperties(xml, (inner) => {
    let updated = removeEmptyElement(inner, "w:pStyle");
    updated = setOrAddEmptyElement(updated, "w:pStyle", { "w:val": "Title" });
    updated = setOrAddEmptyElement(updated, "w:spacing", {
      "w:before": pointsToTwips(titleStyle?.spacingBeforePt ?? 0),
      "w:after": pointsToTwips(titleStyle?.spacingAfterPt ?? 0),
      "w:line": lineSpacingToTwips(titleStyle?.lineSpacing ?? 1),
      "w:lineRule": "auto",
    });
    updated = setOrAddEmptyElement(updated, "w:ind", { "w:firstLine": 0 });
    return updated;
  });
}

function patchDocumentSubtitleParagraph(paragraphXml, titleStyle) {
  let xml = paragraphXml.replaceAll(DOC_SUBTITLE_MARKER, "");
  return patchParagraphProperties(xml, (inner) => {
    let updated = setOrAddEmptyElement(inner, "w:spacing", {
      "w:before": pointsToTwips(0),
      "w:after": pointsToTwips(titleStyle?.spacingAfterPt ?? 0),
      "w:line": lineSpacingToTwips(titleStyle?.lineSpacing ?? 1),
      "w:lineRule": "auto",
    });
    updated = setOrAddEmptyElement(updated, "w:ind", { "w:firstLine": 0 });
    return updated;
  });
}

function applyCaptionParagraphProperties(documentXml, captionStyle) {
  return documentXml.replace(/<w:p\b[\s\S]*?<\/w:p>/g, (paragraphXml) => {
    if (paragraphXml.includes(FIGURE_CAPTION_MARKER)) {
      return patchCaptionParagraph(paragraphXml, FIGURE_CAPTION_MARKER, captionStyle?.figureAlign ?? "center");
    }
    if (paragraphXml.includes(TABLE_CAPTION_MARKER)) {
      return patchCaptionParagraph(paragraphXml, TABLE_CAPTION_MARKER, captionStyle?.tableAlign ?? "left");
    }
    return paragraphXml;
  });
}

function patchCaptionParagraph(paragraphXml, marker, alignment) {
  let xml = paragraphXml.replaceAll(marker, "");
  return patchParagraphProperties(xml, (inner) => {
    let updated = setOrAddEmptyElement(inner, "w:jc", { "w:val": alignment });
    updated = setOrAddEmptyElement(updated, "w:spacing", {
      "w:before": 120,
      "w:after": 120,
      "w:line": 240,
      "w:lineRule": "auto",
    });
    updated = setOrAddEmptyElement(updated, "w:ind", { "w:firstLine": 0 });
    return updated;
  });
}

function patchStyleParagraphProperties(stylesXml, styleId, updater) {
  const styleRegex = new RegExp(`<w:style\\b(?=[^>]*w:styleId="${escapeRegex(styleId)}")[\\s\\S]*?<\\/w:style>`, "g");
  return stylesXml.replace(styleRegex, (styleXml) => patchParagraphProperties(styleXml, updater));
}

function patchStyleRunProperties(stylesXml, styleId, updater) {
  const styleRegex = new RegExp(`<w:style\\b(?=[^>]*w:styleId="${escapeRegex(styleId)}")[\\s\\S]*?<\\/w:style>`, "g");
  return stylesXml.replace(styleRegex, (styleXml) => patchRunProperties(styleXml, updater));
}

function patchRunProperties(xml, updater) {
  if (/<w:rPr>[\s\S]*?<\/w:rPr>/.test(xml)) {
    return xml.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/, (_match, inner) => {
      return `<w:rPr>${updater(inner)}</w:rPr>`;
    });
  }

  if (xml.startsWith("<w:style")) {
    return xml.replace("</w:style>", `<w:rPr>${updater("")}</w:rPr></w:style>`);
  }

  return xml;
}

function patchParagraphProperties(xml, updater) {
  if (/<w:pPr>[\s\S]*?<\/w:pPr>/.test(xml)) {
    return xml.replace(/<w:pPr>([\s\S]*?)<\/w:pPr>/, (_match, inner) => {
      return `<w:pPr>${updater(inner)}</w:pPr>`;
    });
  }

  if (xml.startsWith("<w:style")) {
    if (xml.includes("<w:rPr>")) {
      return xml.replace("<w:rPr>", `<w:pPr>${updater("")}</w:pPr><w:rPr>`);
    }
    return xml.replace("</w:style>", `<w:pPr>${updater("")}</w:pPr></w:style>`);
  }

  const openingTag = xml.match(/^<w:p\b[^>]*>/)?.[0];
  if (!openingTag) {
    return xml;
  }

  return xml.replace(openingTag, `${openingTag}<w:pPr>${updater("")}</w:pPr>`);
}

function setOrAddEmptyElement(xml, tagName, attrs) {
  const tagRegex = new RegExp(`<${escapeRegex(tagName)}\\b([^>]*)\\/>`);
  if (tagRegex.test(xml)) {
    return xml.replace(tagRegex, (_match, attrText) => {
      return `<${tagName}${setXmlAttributes(attrText, attrs)}/>`;
    });
  }

  return `${xml}<${tagName}${formatXmlAttributes(attrs)}/>`;
}

function removeEmptyElement(xml, tagName) {
  return xml.replace(new RegExp(`<${escapeRegex(tagName)}\\b[^>]*\\/>`, "g"), "");
}

function removeAttributesFromEmptyElement(xml, tagName, attrNames) {
  const tagRegex = new RegExp(`<${escapeRegex(tagName)}\\b([^>]*)\\/>`);
  return xml.replace(tagRegex, (_match, attrText) => {
    let updated = attrText ?? "";
    for (const attrName of attrNames) {
      updated = updated.replace(new RegExp(`\\s${escapeRegex(attrName)}="[^"]*"`, "g"), "");
    }
    return `<${tagName}${updated}/>`;
  });
}

function setXmlAttributes(attrText, attrs) {
  let result = attrText ?? "";
  for (const [name, value] of Object.entries(attrs)) {
    const attrRegex = new RegExp(`\\s${escapeRegex(name)}="[^"]*"`);
    if (attrRegex.test(result)) {
      result = result.replace(attrRegex, ` ${name}="${value}"`);
    } else {
      result += ` ${name}="${value}"`;
    }
  }
  return result;
}

function formatXmlAttributes(attrs) {
  return Object.entries(attrs).map(([name, value]) => ` ${name}="${value}"`).join("");
}

function cmToTwips(value) {
  return Math.round(Number(value) / 2.54 * 1440);
}

function pointsToTwips(value) {
  return Math.round(Number(value) * 20);
}

function pointsToHalfPoints(value) {
  return Math.round(Number(value) * 2);
}

function lineSpacingToTwips(value) {
  return Math.round(Number(value) * 240);
}

function normalizeHexColor(value) {
  return String(value ?? "000000").replace(/^#/, "").toUpperCase();
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function getMaxImageWidthPx(converterConfig, options) {
  if (options.maxImageWidth !== null) {
    return options.maxImageWidth;
  }

  const configured = converterConfig.image?.maxWidthPx;
  if (configured === undefined || configured === null) {
    return null;
  }

  const value = Number(configured);
  if (!Number.isInteger(value) || value < 0) {
    throw new Error("converter.image.maxWidthPx must be a non-negative integer.");
  }
  return value;
}

async function convertOne(item, allItems, baseConfig, options) {
  const source = item.file;
  const target = outputPathFor(item, allItems, options);
  const rawMarkdown = await fs.readFile(source, "utf8");
  const normalized = normalizePandocMarkdown(rawMarkdown, options.keepPandocTitleBlock);
  const captionedMarkdown = applyAcademicCaptions(normalized.markdown, baseConfig.captions);
  const engine = normalizeEngineName(baseConfig.engine ?? DEFAULT_CONFIG_DATA.engine);

  if (options.dryRun) {
    console.log(`${source} -> ${target} [engine=${engine}]`);
    return;
  }

  await fs.mkdir(path.dirname(target), { recursive: true });
  let buffer;
  if (engine === "pandoc") {
    buffer = await renderWithPandoc(captionedMarkdown, source, baseConfig, options);
  } else {
    buffer = await renderWithMarkdownDocx(captionedMarkdown, source, normalized, baseConfig, options);
  }
  buffer = await applyDocxHouseStyle(buffer, baseConfig.style);
  await fs.writeFile(target, buffer);
  console.log(`${source} -> ${target} [engine=${engine}]`);
}

async function main() {
  const { inputs, options } = parseArgs(process.argv.slice(2));
  if (options.help) {
    printHelp();
    return;
  }

  if (inputs.length === 0) {
    printHelp();
    process.exitCode = 1;
    return;
  }

  const groups = await Promise.all(inputs.map(collectMarkdownFiles));
  const items = groups.flat();
  if (items.length === 0) {
    throw new Error("No Markdown files found.");
  }

  const loadedConfig = await loadConfig(options.config);
  const baseConfig = mergeDeep(loadedConfig, options.configOverrides);
  for (const item of items) {
    await convertOne(item, items, baseConfig, options);
  }
}

main().catch((error) => {
  console.error(`md-to-docx: ${error.message}`);
  process.exitCode = 1;
});
