import type { Tokens } from "marked"
import {
  AlignmentType,
  Bookmark,
  BorderStyle,
  Document,
  Footer,
  Header,
  ImageRun,
  LineNumberRestartFormat,
  PageNumber,
  Paragraph,
  ShadingType,
  Tab,
  TabStopType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  convertInchesToTwip,
} from "docx"
import { marked } from "marked"
import { buildNumbering, buildStyleOptions } from "./document-styles"
import { parseFrontmatter } from "./frontmatter"
import { loadImage } from "./image"
import { inlineTokensToRuns } from "./inline"
import { listItemsToParagraphs } from "./list"
import { buildTable } from "./table"

function headingSlug(text: string): string {
  return text
    .toLowerCase()
    .replace(/[^\w\s-]/g, "")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
}

const HEADING_STYLES: Record<number, string> = {
  1: "Heading1",
  2: "Heading2",
  3: "Heading3",
  4: "Heading4",
  5: "Heading5",
  6: "Heading6",
}

async function tokensToDocx(
  tokens: Tokens.Generic[],
  markdownPath: string,
): Promise<Array<Paragraph | Table>> {
  const elements: Array<Paragraph | Table> = []
  let firstAppendix = true

  for (const token of tokens) {
    switch (token.type) {
      case "heading": {
        const text = token["text"] as string | undefined
        const inlineTokens = (token["tokens"] as Tokens.Generic[] | undefined) ?? []
        if (firstAppendix && text?.includes("Appendix")) {
          elements.push(new Paragraph({ pageBreakBefore: true }))
          firstAppendix = false
        }
        const slug = text ? headingSlug(text) : undefined
        const runs = inlineTokensToRuns(inlineTokens)
        elements.push(
          new Paragraph({
            children: slug
              ? [new Bookmark({ id: slug, children: runs })]
              : runs,
            style: HEADING_STYLES[(token["depth"] as number | undefined) ?? 1] ?? "Heading1",
          }),
        )
        break
      }

      case "paragraph": {
        const inlineTokens = (token["tokens"] as Tokens.Generic[] | undefined) ?? []
        if (inlineTokens.length === 1 && inlineTokens[0]?.type === "image") {
          const imgToken = inlineTokens[0]
          const url = (imgToken["href"] as string | undefined) ?? ""
          const alt = (imgToken["text"] as string | undefined) ?? ""
          const result = url ? await loadImage(url, markdownPath) : null
          if (result) {
            elements.push(
              new Paragraph({
                children: [
                  new ImageRun({
                    type: result.type,
                    data: result.data,
                    transformation: { width: result.width, height: result.height },
                    altText: { name: alt, description: alt },
                  }),
                ],
                spacing: { before: 320, after: 320 },
              }),
            )
          } else {
            elements.push(
              new Paragraph({
                style: "Normal",
                children: [
                  new TextRun({ text: `[Image: ${alt}]`, italics: true, color: "888888" }),
                ],
              }),
            )
          }
        } else {
          elements.push(
            new Paragraph({ children: inlineTokensToRuns(inlineTokens), style: "Normal" }),
          )
        }
        break
      }

      case "list": {
        elements.push(
          ...listItemsToParagraphs(
            (token["items"] as Tokens.Generic[] | undefined) ?? [],
            (token["ordered"] as boolean | undefined) ?? false,
            0,
          ),
        )
        elements.push(new Paragraph({ text: "", style: "ListSpacing", suppressLineNumbers: true }))
        break
      }

      case "code": {
        const codeLines = ((token["text"] as string | undefined) ?? "")
          .split("\n")
          .map((line) => new Paragraph({ text: line, style: "CodeBlock" }))
        elements.push(
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.NONE, size: 0 },
              bottom: { style: BorderStyle.NONE, size: 0 },
              left: { style: BorderStyle.NONE, size: 0 },
              right: { style: BorderStyle.NONE, size: 0 },
              insideHorizontal: { style: BorderStyle.NONE, size: 0 },
              insideVertical: { style: BorderStyle.NONE, size: 0 },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: codeLines,
                    shading: { type: ShadingType.CLEAR, color: "F2F2F2", fill: "F2F2F2" },
                    margins: {
                      top: convertInchesToTwip(0.1),
                      bottom: convertInchesToTwip(0.1),
                      left: convertInchesToTwip(0.15),
                      right: convertInchesToTwip(0.15),
                    },
                    borders: {
                      top: { style: BorderStyle.NONE, size: 0 },
                      bottom: { style: BorderStyle.NONE, size: 0 },
                      left: { style: BorderStyle.NONE, size: 0 },
                      right: { style: BorderStyle.NONE, size: 0 },
                    },
                  }),
                ],
              }),
            ],
          }),
        )
        elements.push(new Paragraph({ text: "", style: "CodeBlockSpacing", suppressLineNumbers: true }))
        break
      }

      case "blockquote": {
        for (const inner of (token["tokens"] as Tokens.Generic[] | undefined) ?? []) {
          if (inner.type === "paragraph" && inner["text"]) {
            elements.push(new Paragraph({ text: inner["text"] as string, style: "Blockquote" }))
          }
        }
        elements.push(new Paragraph({ text: "", style: "ListSpacing", suppressLineNumbers: true }))
        break
      }

      case "table": {
        elements.push(buildTable(token))
        elements.push(new Paragraph({ text: "", style: "CodeBlockSpacing", suppressLineNumbers: true }))
        break
      }

      case "image": {
        const url = (token["href"] as string | undefined) ?? ""
        const alt = (token["text"] as string | undefined) ?? ""
        const result = url ? await loadImage(url, markdownPath) : null
        if (result) {
          elements.push(
            new Paragraph({
              children: [
                new ImageRun({
                  type: result.type,
                  data: result.data,
                  transformation: { width: result.width, height: result.height },
                  altText: { name: alt, description: alt },
                }),
              ],
              spacing: { before: 320, after: 320 },
            }),
          )
        } else {
          elements.push(
            new Paragraph({
              style: "Normal",
              children: [new TextRun({ text: `[Image: ${alt}]`, italics: true, color: "888888" })],
            }),
          )
        }
        break
      }

      case "hr": {
        elements.push(
          new Paragraph({
            border: {
              bottom: { style: BorderStyle.SINGLE, size: 12, color: "AAAAAA", space: 1 },
            },
            spacing: { before: 240, after: 240 },
          }),
        )
        break
      }

      case "space":
        break

      default:
        if ((token["tokens"] as Tokens.Generic[] | undefined)?.length || token["text"]) {
          elements.push(
            new Paragraph({
              children: inlineTokensToRuns((token["tokens"] as Tokens.Generic[] | undefined) ?? []),
            }),
          )
        }
    }
  }

  return elements
}

export interface ConvertOptions {
  footerLabel?: string
  /** Include a right-aligned page number in the footer */
  footerPageNumber?: boolean
  headerLabel?: string
  /** Enable Word's built-in document line numbering */
  lineNumbers?: boolean
  /** Raw XML string extracted from word/styles.xml in a .dotx template */
  externalStylesXml?: string
  /** Base font size in pt; all readable styles scale from this (default: 12) */
  fontSize?: number
}

export async function convertMarkdownToDocx(
  markdownPath: string,
  _docxPath: string,
  options: ConvertOptions = {},
): Promise<Document> {
  const content = await Bun.file(markdownPath).text()
  const { body, data } = parseFrontmatter(content)
  const tokens = marked.lexer(body) as Tokens.Generic[]
  const elements = await tokensToDocx(tokens, markdownPath)

  // CLI flags take precedence; frontmatter.title is the fallback for headerLabel
  const headerLabel =
    options.headerLabel ?? (typeof data["title"] === "string" ? data["title"] : undefined)

  const headerSection =
    headerLabel !== undefined
      ? {
          headers: {
            default: new Header({
              children: [
                new Paragraph({
                  style: "FooterText",
                  alignment: AlignmentType.LEFT,
                  children: [new TextRun({ text: headerLabel })],
                }),
              ],
            }),
            first: new Header({ children: [] }),
          },
          properties: { titlePage: true },
        }
      : {}

  const footerChildren: TextRun[] = []
  if (options.footerLabel) {
    footerChildren.push(new TextRun({ text: options.footerLabel }))
  }
  if (options.footerPageNumber) {
    // Tab and page number in separate runs so field chars are isolated from the tab element
    footerChildren.push(new TextRun({ children: [new Tab()] }))
    footerChildren.push(new TextRun({ children: [PageNumber.CURRENT] }))
  }
  // A4 default (11906 twips wide) minus 1" left + 1" right margins = 9026 twips text width
  const textWidthTwip = 11906 - 1440 - 1440
  const footer = new Footer({
    children: [
      new Paragraph({
        style: "FooterText",
        tabStops: [{ type: TabStopType.RIGHT, position: textWidthTwip }],
        children: footerChildren,
      }),
    ],
  })

  return new Document({
    ...buildStyleOptions(options.fontSize ?? 12, options.externalStylesXml),
    numbering: buildNumbering(),
    sections: [
      {
        ...headerSection,
        footers: { default: footer, ...(headerLabel !== undefined ? { first: footer } : {}) },
        properties: {
          ...(headerSection as { properties?: object }).properties,
          ...(options.lineNumbers
            ? { lineNumbers: { countBy: 1, restart: LineNumberRestartFormat.CONTINUOUS } }
            : {}),
        },
        children: elements,
      },
    ],
  })
}
