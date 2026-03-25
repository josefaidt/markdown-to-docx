import type { Tokens } from "marked"
import { ExternalHyperlink, InternalHyperlink, TextRun } from "docx"

export type InlineChild = TextRun | ExternalHyperlink | InternalHyperlink

interface InlineCtx {
  bold?: boolean
  italics?: boolean
}

export function inlineTokensToRuns(tokens: Tokens.Generic[], ctx: InlineCtx = {}): InlineChild[] {
  const runs: InlineChild[] = []
  for (const token of tokens) {
    // Use bracket notation because Tokens.Generic has an index signature
    const text = token["text"] as string | undefined
    const href = token["href"] as string | undefined
    const children = token["tokens"] as Tokens.Generic[] | undefined
    switch (token.type) {
      case "text":
        if (children?.length) {
          runs.push(...inlineTokensToRuns(children, ctx))
        } else {
          runs.push(new TextRun({ text: text ?? "", ...ctx }))
        }
        break
      case "strong":
        runs.push(...inlineTokensToRuns(children ?? [], { ...ctx, bold: true }))
        break
      case "em":
        runs.push(...inlineTokensToRuns(children ?? [], { ...ctx, italics: true }))
        break
      case "codespan":
        runs.push(new TextRun({ text: text ?? "", ...ctx, style: "InlineCode" }))
        break
      case "link": {
        const linkText =
          children?.map((t) => (t["text"] as string | undefined) ?? "").join("") ?? text ?? ""
        const linkHref = href ?? ""
        if (linkHref.startsWith("#")) {
          runs.push(
            new InternalHyperlink({
              anchor: linkHref.slice(1),
              children: [new TextRun({ text: linkText, style: "Hyperlink" })],
            }),
          )
        } else {
          runs.push(
            new ExternalHyperlink({
              link: linkHref,
              children: [new TextRun({ text: linkText, style: "Hyperlink" })],
            }),
          )
        }
        break
      }
      case "br":
        runs.push(new TextRun({ break: 1 }))
        break
      case "escape":
      default:
        if (text) runs.push(new TextRun({ text, ...ctx }))
        break
    }
  }
  return runs
}
