import type { Tokens } from "marked"
import {
  AlignmentType,
  Paragraph,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  WidthType,
} from "docx"
import { inlineTokensToRuns } from "./inline"

interface ParsedTableCell {
  text: string
  tokens: Tokens.Generic[]
  header: boolean
  align: "center" | "left" | "right" | null
}

/**
 * Base grid-column hint every column gets, plus per-character growth. These are
 * proportional hints for an autofit table, not absolute widths, so the exact
 * scale is immaterial — only the ratios matter. The base ensures even a single
 * narrow column stays a legible width relative to its neighbours.
 */
const COLUMN_BASE_TWIP = 300
const PER_CHAR_TWIP = 100

/**
 * Compute proportional column-width hints for the table grid, weighted by the
 * longest cell text in each column. These are hints only: the table is emitted
 * as a full-width autofit table, so Word distributes the columns to fill the
 * page (whatever its size and margins) starting from these ratios.
 */
function computeColumnWidths(charCounts: number[]): number[] {
  return charCounts.map((c) => COLUMN_BASE_TWIP + c * PER_CHAR_TWIP)
}

export function buildTable(token: Tokens.Generic): Table {
  const header = (token["header"] as ParsedTableCell[] | undefined) ?? []
  const bodyRows = (token["rows"] as ParsedTableCell[][] | undefined) ?? []

  const columnCount = Math.max(header.length, ...bodyRows.map((r) => r.length), 0)

  // Weight each column by the longest cell text it contains (header + body),
  // so the grid tracks content without depending on how the author spaced the
  // markdown pipes.
  const chars: number[] = Array.from({ length: columnCount }, () => 0)
  const measure = (cells: ParsedTableCell[]) => {
    for (let i = 0; i < cells.length; i++) {
      const len = cells[i]?.text.length ?? 0
      if (len > (chars[i] ?? 0)) chars[i] = len
    }
  }
  measure(header)
  for (const row of bodyRows) measure(row)

  const columnWidths = computeColumnWidths(chars)

  const cellParagraph = (cell: ParsedTableCell, bold = false) =>
    new Paragraph({
      children: inlineTokensToRuns(cell.tokens, { bold }),
      ...(cell.align === "center" && { alignment: AlignmentType.CENTER }),
      ...(cell.align === "right" && { alignment: AlignmentType.RIGHT }),
      spacing: { after: 0 },
    })

  const tableCell = (cell: ParsedTableCell, bold = false) =>
    new TableCell({
      children: [cellParagraph(cell, bold)],
    })

  const rows: TableRow[] = []
  if (header.length > 0) {
    rows.push(
      new TableRow({
        children: header.map((cell) => tableCell(cell, true)),
        tableHeader: true,
      }),
    )
  }
  for (const row of bodyRows) {
    rows.push(new TableRow({ children: row.map((cell) => tableCell(cell)) }))
  }

  const table = new Table({
    rows,
    columnWidths,
    layout: TableLayoutType.AUTOFIT,
    width: { size: 100, type: WidthType.PERCENTAGE },
    style: "TableGridLight",
  })

  // The docx library always emits tblBorders even when a named style owns them.
  // Strip the redundant element so TableGridLight's borders are used.
  const tblPr = (table as unknown as { root: [{ root: unknown[] }, ...unknown[]] }).root[0]
  tblPr.root = tblPr.root.filter((el) => (el as { rootKey?: string }).rootKey !== "w:tblBorders")
  return table
}
