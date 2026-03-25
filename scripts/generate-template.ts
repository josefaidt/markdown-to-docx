#!/usr/bin/env bun
import type { IPropertiesOptions } from "docx"
import {
  Document,
  Packer,
  Paragraph,
  AlignmentType,
  BorderStyle,
  ShadingType,
  convertInchesToTwip,
  TabStopPosition,
  TabStopType,
  Footer,
  TextRun,
  PageNumber,
} from "docx"
/**
 * Generates template.dotx from the styles defined in lib/md-to-docx.ts.
 *
 * Builds an empty Document with only the style/numbering definitions (no body
 * content), packs it as a .docx buffer, then patches [Content_Types].xml inside
 * the ZIP to flip the content type from .docx → .dotx, producing a valid Word
 * template file.
 *
 * Usage: bun run scripts/generate-template.ts [output.dotx]
 */
import JSZip from "jszip"

const DOCX_MAIN_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
const DOTX_MAIN_CONTENT_TYPE =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"

const LIST_BULLET_INDENT = convertInchesToTwip(0.2)
const LIST_TEXT_INDENT = convertInchesToTwip(0.45)
const LIST_LEVEL_INDENT = convertInchesToTwip(0.25)

function listIndent(level: number) {
  const bulletAt = LIST_BULLET_INDENT + level * LIST_LEVEL_INDENT
  const textAt = LIST_TEXT_INDENT + level * LIST_LEVEL_INDENT
  return { left: textAt, hanging: textAt - bulletAt }
}

const listLevels = [0, 1, 2, 3, 4, 5].map((level) => {
  const { left, hanging } = listIndent(level)
  return { left, hanging }
})

const footer = new Footer({
  children: [
    new Paragraph({
      style: "FooterText",
      tabStops: [{ type: TabStopType.RIGHT, position: TabStopPosition.MAX }],
      children: [
        new TextRun({ text: "" }),
        new TextRun({ text: "\t" }),
        new TextRun({ children: [PageNumber.CURRENT] }),
      ],
    }),
  ],
})

const docOptions: IPropertiesOptions = {
  styles: {
    paragraphStyles: [
      {
        id: "Normal",
        name: "Normal",
        run: { font: "Aptos", size: 28, color: "333333" },
        paragraph: { spacing: { after: 160, line: 276 } },
      },
      {
        id: "Heading1",
        name: "Heading 1",
        basedOn: "Normal",
        next: "Normal",
        run: { size: 70, bold: true },
        paragraph: { spacing: { before: 0, after: 240 } },
      },
      {
        id: "Heading2",
        name: "Heading 2",
        basedOn: "Normal",
        next: "Normal",
        run: { size: 44, bold: true },
        paragraph: { spacing: { before: 360, after: 160 } },
      },
      {
        id: "Heading3",
        name: "Heading 3",
        basedOn: "Normal",
        next: "Normal",
        run: { size: 36, bold: true },
        paragraph: { spacing: { before: 300, after: 120 } },
      },
      {
        id: "Heading4",
        name: "Heading 4",
        basedOn: "Normal",
        next: "Normal",
        run: { size: 30, bold: true },
        paragraph: { spacing: { before: 240, after: 80 } },
      },
      {
        id: "Heading5",
        name: "Heading 5",
        basedOn: "Normal",
        next: "Normal",
        run: { size: 30, bold: true, italics: true },
        paragraph: { spacing: { before: 240, after: 80 } },
      },
      {
        id: "Heading6",
        name: "Heading 6",
        basedOn: "Normal",
        next: "Normal",
        run: { size: 30, underline: {} },
        paragraph: { spacing: { before: 240, after: 80 } },
      },
      {
        id: "ListItem",
        name: "List Item",
        basedOn: "Normal",
        next: "ListItem",
        paragraph: { spacing: { before: 0, after: 80, line: 276 } },
      },
      {
        id: "Blockquote",
        name: "Blockquote",
        basedOn: "Normal",
        next: "Normal",
        run: { italics: true, color: "777777" },
        paragraph: {
          spacing: { before: 240, after: 80, line: 276, lineRule: "auto" as const },
          indent: { left: convertInchesToTwip(0.15) },
          border: {
            left: { style: BorderStyle.SINGLE, size: 24, color: "A6A6A6", space: 6 },
          },
        },
      },
      {
        id: "CodeBlock",
        name: "Code Block",
        basedOn: "Normal",
        next: "Normal",
        run: { font: "Consolas", size: 28, color: "555555" },
        paragraph: { spacing: { before: 0, after: 0, line: 240 } },
      },
      {
        id: "CodeBlockSpacing",
        name: "Code Block Spacing",
        basedOn: "Normal",
        next: "Normal",
        run: { size: 12 },
        paragraph: { spacing: { before: 0, after: 0, line: 240 } },
      },
      {
        id: "ListSpacing",
        name: "List Spacing",
        basedOn: "Normal",
        next: "Normal",
        run: { size: 8 },
        paragraph: { spacing: { before: 0, after: 0, line: 240 } },
      },
      {
        id: "FooterText",
        name: "Footer Text",
        basedOn: "Normal",
        paragraph: { spacing: { before: 0, after: 0 } },
        run: { size: 18, color: "888888" },
      },
    ],
    characterStyles: [
      {
        id: "InlineCode",
        name: "Inline Code",
        run: {
          font: "Consolas",
          size: 26,
          shading: { type: ShadingType.CLEAR, color: "F2F2F2", fill: "F2F2F2" },
        },
      },
      {
        id: "Hyperlink",
        name: "Hyperlink",
        run: { color: "0563C1", underline: {} },
      },
    ],
  },
  numbering: {
    config: [
      {
        reference: "ordered-numbering",
        levels: [
          {
            level: 0,
            format: "decimal",
            text: "%1.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: listLevels[0] } },
          },
          {
            level: 1,
            format: "lowerLetter",
            text: "%2.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: listLevels[1] } },
          },
          {
            level: 2,
            format: "lowerRoman",
            text: "%3.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: listLevels[2] } },
          },
          {
            level: 3,
            format: "decimal",
            text: "%4.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: listLevels[3] } },
          },
          {
            level: 4,
            format: "lowerLetter",
            text: "%5.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: listLevels[4] } },
          },
          {
            level: 5,
            format: "lowerRoman",
            text: "%6.",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: listLevels[5] } },
          },
        ],
      },
      {
        reference: "bullet-numbering",
        levels: [
          {
            level: 0,
            format: "bullet",
            text: "\u2022",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: listLevels[0] }, run: { font: "Aptos" } },
          },
          {
            level: 1,
            format: "bullet",
            text: "\u25e6",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: listLevels[1] }, run: { font: "Aptos" } },
          },
          {
            level: 2,
            format: "bullet",
            text: "\u25aa",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: listLevels[2] }, run: { font: "Aptos" } },
          },
          {
            level: 3,
            format: "bullet",
            text: "\u2022",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: listLevels[3] }, run: { font: "Aptos" } },
          },
          {
            level: 4,
            format: "bullet",
            text: "\u25e6",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: listLevels[4] }, run: { font: "Aptos" } },
          },
          {
            level: 5,
            format: "bullet",
            text: "\u25aa",
            alignment: AlignmentType.LEFT,
            style: { paragraph: { indent: listLevels[5] }, run: { font: "Aptos" } },
          },
        ],
      },
    ],
  },
  sections: [
    {
      properties: {},
      footers: { default: footer },
      children: [],
    },
  ],
}

const doc = new Document(docOptions)
const docxBuffer = await Packer.toBuffer(doc)

// Patch [Content_Types].xml inside the ZIP to convert .docx → .dotx
const zip = await JSZip.loadAsync(docxBuffer)
const contentTypesFile = zip.file("[Content_Types].xml")
if (!contentTypesFile) throw new Error("Missing [Content_Types].xml in generated document")

const contentTypesXml = await contentTypesFile.async("string")
const patchedXml = contentTypesXml.replace(DOCX_MAIN_CONTENT_TYPE, DOTX_MAIN_CONTENT_TYPE)
zip.file("[Content_Types].xml", patchedXml)

const dotxBuffer = await zip.generateAsync({ type: "nodebuffer" })

const outPath = Bun.argv[2] ?? "scripts/template.dotx"
await Bun.write(outPath, dotxBuffer)
console.log(`Written to ${outPath}`)
