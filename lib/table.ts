import type { Tokens } from "marked"
import { AlignmentType, Paragraph, Table, TableCell, TableRow, WidthType } from "docx"
import { inlineTokensToRuns } from "./inline"

interface ParsedTableCell {
  text: string
  tokens: Tokens.Generic[]
  header: boolean
  align: "center" | "left" | "right" | null
}

function cellAlignType(
  align: "center" | "left" | "right" | null,
): (typeof AlignmentType)[keyof typeof AlignmentType] {
  if (align === "center") return AlignmentType.CENTER
  if (align === "right") return AlignmentType.RIGHT
  return AlignmentType.LEFT
}

export function buildTable(token: Tokens.Generic): Table {
  const rows: TableRow[] = []

  const header = token["header"] as ParsedTableCell[] | undefined
  if (header && header.length > 0) {
    rows.push(
      new TableRow({
        children: header.map(
          (cell) =>
            new TableCell({
              children: [
                new Paragraph({
                  children: inlineTokensToRuns(cell.tokens),
                  alignment: cellAlignType(cell.align),
                }),
              ],
            }),
        ),
        tableHeader: true,
      }),
    )
  }

  for (const row of (token["rows"] as ParsedTableCell[][] | undefined) ?? []) {
    rows.push(
      new TableRow({
        children: row.map(
          (cell) =>
            new TableCell({
              children: [
                new Paragraph({
                  children: inlineTokensToRuns(cell.tokens),
                  alignment: cellAlignType(cell.align),
                }),
              ],
            }),
        ),
      }),
    )
  }

  return new Table({
    rows,
    style: "GridTable1Light",
    width: { size: 100, type: WidthType.PERCENTAGE },
  })
}
