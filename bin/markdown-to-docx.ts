#!/usr/bin/env bun
import type { ConvertOptions } from "../lib/markdown-to-docx"
import { resolve, extname, basename } from "node:path"
import { Packer } from "docx"
import JSZip from "jszip"
import { convertMarkdownToDocx } from "../lib/markdown-to-docx"
import pkg from "../package.json" with { type: "json" }

const USAGE = `
Usage: markdown-to-docx <input.md> [output.docx] [options]

Arguments:
  input.md        Path to the input Markdown file
  output.docx     Path for the output Word document (default: same name as input)

Options:
  --font-size <n>    Base font size in pt; all styles scale from this (default: 12)
  --footer <text>    Text to display in the bottom-left footer
  --header <text>    Text to display left-aligned in the header (skipped on the first page)
  --line-numbers     Enable Word's built-in document line numbering
  --no-bookmarks     Disable automatic bookmark generation for headings
  --page-numbers     Add a right-aligned page number to the footer
  --template <path>  Path to a .dotx template file to load styles from

Global Options:
  -h, --help         Show this help message
  -v, --version      Print the version number
`.trim()

function parseArgs(args: string[]): {
  inputPath: string
  outputPath: string
  templatePath: string | null
  options: ConvertOptions
} | null {
  if (args.includes("--version") || args.includes("-v")) {
    process.stdout.write(pkg.version + "\n")
    return null
  }

  if (args.length === 0 || args.includes("--help") || args.includes("-h")) {
    process.stdout.write(USAGE + "\n")
    return null
  }

  const filteredArgs = [...args]

  // Extract --template <path>
  let templatePath: string | null = null
  const templateIdx = filteredArgs.indexOf("--template")
  if (templateIdx !== -1) {
    const templateArg = filteredArgs[templateIdx + 1]
    if (templateArg === undefined || templateArg.startsWith("--")) {
      process.stderr.write(`--template requires a path argument\n${USAGE}\n`)
      return null
    }
    templatePath = resolve(templateArg)
    filteredArgs.splice(templateIdx, 2)
  }

  // Extract --header <text>
  let headerLabel: string | undefined
  const headerIdx = filteredArgs.indexOf("--header")
  if (headerIdx !== -1) {
    const headerArg = filteredArgs[headerIdx + 1]
    if (headerArg === undefined || headerArg.startsWith("--")) {
      process.stderr.write(`--header requires a text argument\n${USAGE}\n`)
      return null
    }
    headerLabel = headerArg
    filteredArgs.splice(headerIdx, 2)
  }

  // Extract --footer <text>
  let footerLabel: string | undefined
  const footerIdx = filteredArgs.indexOf("--footer")
  if (footerIdx !== -1) {
    const footerArg = filteredArgs[footerIdx + 1]
    if (footerArg === undefined || footerArg.startsWith("--")) {
      process.stderr.write(`--footer requires a text argument\n${USAGE}\n`)
      return null
    }
    footerLabel = footerArg
    filteredArgs.splice(footerIdx, 2)
  }

  // Extract --font-size <n>
  let fontSize: number | undefined
  const fontSizeIdx = filteredArgs.indexOf("--font-size")
  if (fontSizeIdx !== -1) {
    const fontSizeArg = filteredArgs[fontSizeIdx + 1]
    if (fontSizeArg === undefined || fontSizeArg.startsWith("--")) {
      process.stderr.write(`--font-size requires a numeric argument\n${USAGE}\n`)
      return null
    }
    const parsed = parseInt(fontSizeArg, 10)
    if (isNaN(parsed) || parsed <= 0 || String(parsed) !== fontSizeArg) {
      process.stderr.write(
        `--font-size must be a positive integer (e.g. --font-size 12)\n${USAGE}\n`,
      )
      return null
    }
    fontSize = parsed
    filteredArgs.splice(fontSizeIdx, 2)
  }

  const flags = filteredArgs.filter((a) => a.startsWith("-"))
  const positional = filteredArgs.filter((a) => !a.startsWith("-"))

  const knownFlags = new Set([
    "--line-numbers",
    "--no-bookmarks",
    "--page-numbers",
    "--help",
    "--version",
    "-h",
    "-v",
  ])
  const unknownFlags = flags.filter((f) => !knownFlags.has(f))
  if (unknownFlags.length > 0) {
    process.stderr.write(`Unknown flag(s): ${unknownFlags.join(", ")}\n${USAGE}\n`)
    return null
  }

  const inputPath = positional[0]
  if (inputPath === undefined) {
    process.stderr.write(`Missing required argument: input.md\n${USAGE}\n`)
    return null
  }

  const resolvedInput = resolve(inputPath)
  const resolvedOutput =
    positional[1] !== undefined
      ? resolve(positional[1])
      : resolve(resolve(inputPath, ".."), `${basename(inputPath, extname(inputPath))}.docx`)

  return {
    inputPath: resolvedInput,
    outputPath: resolvedOutput,
    templatePath,
    options: {
      lineNumbers: flags.includes("--line-numbers"),
      noBookmarks: flags.includes("--no-bookmarks"),
      footerPageNumber: flags.includes("--page-numbers"),
      headerLabel,
      footerLabel,
      fontSize,
    },
  }
}

async function loadTemplateStyles(templatePath: string): Promise<string> {
  const buf = await Bun.file(templatePath).arrayBuffer()
  const zip = await JSZip.loadAsync(buf)
  const stylesFile = zip.file("word/styles.xml")
  if (!stylesFile) throw new Error(`No word/styles.xml found in template: ${templatePath}`)
  return stylesFile.async("string")
}

const _cliArgs = Bun.argv.slice(2)
const parsed = parseArgs(_cliArgs)
if (parsed === null) {
  const isExplicitRequest =
    _cliArgs.includes("--version") ||
    _cliArgs.includes("-v") ||
    _cliArgs.includes("--help") ||
    _cliArgs.includes("-h")
  process.exit(isExplicitRequest ? 0 : 1)
}

const { inputPath, outputPath, templatePath, options } = parsed

if (!(await Bun.file(inputPath).exists())) {
  process.stderr.write(`Error: Cannot read input file: ${inputPath}\n`)
  process.exit(1)
}

if (templatePath !== null) {
  if (!(await Bun.file(templatePath).exists())) {
    process.stderr.write(`Error: Cannot read template file: ${templatePath}\n`)
    process.exit(1)
  }
  try {
    options.externalStylesXml = await loadTemplateStyles(templatePath)
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err)
    process.stderr.write(`Error loading template: ${message}\n`)
    process.exit(1)
  }
}

try {
  const doc = await convertMarkdownToDocx(inputPath, outputPath, options)
  const buffer = await Packer.toBuffer(doc)
  await Bun.write(outputPath, buffer)
  process.stdout.write(`Written to ${outputPath}\n`)
} catch (err) {
  const message = err instanceof Error ? err.message : String(err)
  process.stderr.write(`Error: ${message}\n`)
  process.exit(1)
}
