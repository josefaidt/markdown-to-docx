import type { Tokens } from "marked"
import {
  AlignmentType,
  ImportedXmlComponent,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  VerticalAlignTable,
} from "docx"
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

  const cellParagraph = (cell: ParsedTableCell, bold = false) =>
    new Paragraph({
      children: inlineTokensToRuns(cell.tokens, { bold }),
      alignment: cellAlignType(cell.align),
      spacing: { after: 0 },
    })

  const tableCell = (cell: ParsedTableCell, bold = false) =>
    new TableCell({
      children: [cellParagraph(cell, bold)],
      verticalAlign: VerticalAlignTable.CENTER,
    })

  const header = token["header"] as ParsedTableCell[] | undefined
  if (header && header.length > 0) {
    rows.push(
      new TableRow({
        children: header.map((cell) => tableCell(cell, true)),
        tableHeader: true,
      }),
    )
  }

  for (const row of (token["rows"] as ParsedTableCell[][] | undefined) ?? []) {
    rows.push(
      new TableRow({
        children: row.map((cell) => tableCell(cell)),
      }),
    )
  }

  const table = new Table({
    rows,
    style: "TableGridLight",
  })
  // The docx library mis-serializes WidthType.PERCENTAGE as e.g. "5000%" and always
  // emits tblBorders even when a named style owns them. Patch the TableProperties root
  // directly to replace the broken width element and strip the redundant borders.
  const tblPr = (table as unknown as { root: { root: unknown[] }[] }).root[0]
  tblPr.root = tblPr.root.filter((el) => {
    const name = (el as { rootKey?: string }).rootKey
    return name !== "w:tblW" && name !== "w:tblBorders"
  })
  const tblW = (
    ImportedXmlComponent.fromXmlString('<w:tblW w:w="5000" w:type="pct"/>') as unknown as {
      root: unknown[]
    }
  ).root[0]
  tblPr.root.push(tblW)
  return table
}
