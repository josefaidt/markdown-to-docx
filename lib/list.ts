import type { Tokens } from "marked"
import {
  BorderStyle,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
  convertInchesToTwip,
} from "docx"
import { highlightCode, isSupportedLang } from "./highlight"
import { inlineTokensToRuns } from "./inline"

const LIST_BULLET_INDENT = convertInchesToTwip(0.2)
const LIST_TEXT_INDENT = convertInchesToTwip(0.45)
const LIST_LEVEL_INDENT = convertInchesToTwip(0.35)

export function listIndent(level: number) {
  const bulletAt = LIST_BULLET_INDENT + level * LIST_LEVEL_INDENT
  const textAt = LIST_TEXT_INDENT + level * LIST_LEVEL_INDENT
  return { left: textAt, hanging: textAt - bulletAt }
}

async function codeTokenToTable(token: Tokens.Generic, level: number): Promise<Table> {
  const codeText = (token["text"] as string | undefined) ?? ""
  const codeLang = token["lang"] as string | undefined
  const codeLines = isSupportedLang(codeLang)
    ? (await highlightCode(codeText, codeLang)).map(
        (lineTokens) =>
          new Paragraph({
            style: "CodeBlock",
            children: lineTokens.map(
              (t) =>
                new TextRun({
                  text: t.text,
                  font: "Consolas",
                  color: t.color,
                  bold: t.bold || undefined,
                  italics: t.italic || undefined,
                }),
            ),
          }),
      )
    : codeText.split("\n").map((line) => new Paragraph({ text: line, style: "CodeBlock" }))

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    indent: { size: listIndent(level).left, type: WidthType.DXA },
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
            shading: { type: ShadingType.CLEAR, color: "F6F8FA", fill: "F6F8FA" },
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
  })
}

export async function listItemsToParagraphs(
  items: Tokens.Generic[],
  ordered: boolean,
  level: number,
  orderedRef: string,
): Promise<Array<Paragraph | Table>> {
  const elements: Array<Paragraph | Table> = []

  for (const item of items) {
    if (item["type"] !== "list_item") continue

    const textTokens: Tokens.Generic[] = []
    const nestedLists: Tokens.Generic[] = []
    const codeTokens: Tokens.Generic[] = []

    for (const t of (item["tokens"] as Tokens.Generic[] | undefined) ?? []) {
      if (t.type === "list") {
        nestedLists.push(t)
      } else if (t.type === "text" || t.type === "paragraph") {
        textTokens.push(t)
      } else if (t.type === "code") {
        codeTokens.push(t)
      }
    }

    const inlineTokens = textTokens.flatMap(
      (t) => (t["tokens"] as Tokens.Generic[] | undefined) ?? (t["text"] ? [t] : []),
    )

    elements.push(
      new Paragraph({
        style: "ListItem",
        children: inlineTokensToRuns(inlineTokens),
        numbering: {
          reference: ordered ? orderedRef : "bullet-numbering",
          level,
        },
      }),
    )

    for (const codeToken of codeTokens) {
      elements.push(await codeTokenToTable(codeToken, level))
      // Word has no conditional next-sibling spacing, so we add an explicit spacer
      // to balance the gap after the table with the ListItem `after: 160` before it.
      elements.push(
        new Paragraph({
          children: [new TextRun({ size: 12 })],
          spacing: { before: 0, after: 0, line: 160, lineRule: "exact" as const },
        }),
      )
    }

    for (const nested of nestedLists) {
      elements.push(
        ...(await listItemsToParagraphs(
          (nested["items"] as Tokens.Generic[] | undefined) ?? [],
          (nested["ordered"] as boolean | undefined) ?? false,
          level + 1,
          orderedRef,
        )),
      )
    }
  }

  return elements
}
