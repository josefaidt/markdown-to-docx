import type { MarkedExtension, Tokens } from "marked"
import { Lexer } from "marked"

export interface FootnoteRefToken {
  type: "footnote_ref"
  raw: string
  label: string
}

export interface FootnoteDefToken {
  type: "footnote_def"
  raw: string
  label: string
  tokens: Tokens.Generic[]
}

export const footnote: MarkedExtension = {
  extensions: [
    {
      name: "footnote_def",
      level: "block",
      start(src) {
        return src.match(/^\[\^\d+\]\s/)?.index ?? -1
      },
      tokenizer(src) {
        const match = /^\[\^(\d+)\]\s(.+)/.exec(src)
        if (!match) return undefined
        const label = match[1]!
        const content = match[2]!
        const tokens = Lexer.lexInline(content) as Tokens.Generic[]
        return {
          type: "footnote_def",
          raw: match[0],
          label,
          tokens,
        } satisfies FootnoteDefToken
      },
      childTokens: ["tokens"],
    },
    {
      name: "footnote_ref",
      level: "inline",
      start(src) {
        return src.match(/\[\^\d+\]/)?.index ?? -1
      },
      tokenizer(src) {
        const match = /^\[\^(\d+)\]/.exec(src)
        if (!match) return undefined
        return {
          type: "footnote_ref",
          raw: match[0],
          label: match[1]!,
        } satisfies FootnoteRefToken
      },
    },
  ],
}
