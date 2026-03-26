import type { Tokens } from "marked"
import { Paragraph, convertInchesToTwip } from "docx"
import { inlineTokensToRuns } from "./inline"

const LIST_BULLET_INDENT = convertInchesToTwip(0.2)
const LIST_TEXT_INDENT = convertInchesToTwip(0.45)
const LIST_LEVEL_INDENT = convertInchesToTwip(0.25)

export function listIndent(level: number) {
  const bulletAt = LIST_BULLET_INDENT + level * LIST_LEVEL_INDENT
  const textAt = LIST_TEXT_INDENT + level * LIST_LEVEL_INDENT
  return { left: textAt, hanging: textAt - bulletAt }
}

export function listItemsToParagraphs(
  items: Tokens.Generic[],
  ordered: boolean,
  level: number,
  orderedRef: string,
): Paragraph[] {
  const paragraphs: Paragraph[] = []

  for (const item of items) {
    if (item["type"] !== "list_item") continue

    const textTokens: Tokens.Generic[] = []
    const nestedLists: Tokens.Generic[] = []

    for (const t of (item["tokens"] as Tokens.Generic[] | undefined) ?? []) {
      if (t.type === "list") {
        nestedLists.push(t)
      } else if (t.type === "text" || t.type === "paragraph") {
        textTokens.push(t)
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

    for (const nested of nestedLists) {
      paragraphs.push(
        ...listItemsToParagraphs(
          (nested["items"] as Tokens.Generic[] | undefined) ?? [],
          (nested["ordered"] as boolean | undefined) ?? false,
          level + 1,
          orderedRef,
        ),
      )
    }
  }

  return paragraphs
}
