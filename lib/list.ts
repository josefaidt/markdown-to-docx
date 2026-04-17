import type { Tokens } from "marked"
import { Paragraph, TextRun, convertInchesToTwip } from "docx"
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

async function codeTokenToParagraphs(token: Tokens.Generic): Promise<Paragraph[]> {
  const codeText = (token["text"] as string | undefined) ?? ""
  const codeLang = token["lang"] as string | undefined
  if (isSupportedLang(codeLang)) {
    const lines = await highlightCode(codeText, codeLang)
    return lines.map(
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
  }
  return codeText.split("\n").map((line) => new Paragraph({ text: line, style: "CodeBlock" }))
}

export async function listItemsToParagraphs(
  items: Tokens.Generic[],
  ordered: boolean,
  level: number,
  orderedRef: string,
): Promise<Paragraph[]> {
  const paragraphs: Paragraph[] = []

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

    paragraphs.push(
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
      paragraphs.push(...(await codeTokenToParagraphs(codeToken)))
    }

    for (const nested of nestedLists) {
      paragraphs.push(
        ...(await listItemsToParagraphs(
          (nested["items"] as Tokens.Generic[] | undefined) ?? [],
          (nested["ordered"] as boolean | undefined) ?? false,
          level + 1,
          orderedRef,
        )),
      )
    }
  }

  return paragraphs
}
