import type { BundledLanguage, Highlighter } from "shiki"
import { bundledLanguages, createHighlighter } from "shiki"

export interface HighlightToken {
  text: string
  color: string
  bold: boolean
  italic: boolean
}

// FontStyle bitmask values from shiki
const FONT_STYLE_ITALIC = 1
const FONT_STYLE_BOLD = 2

const THEME = "github-light"

// Singleton highlighter — initialized once, lazily
let highlighter: Highlighter | undefined

async function getHighlighter(): Promise<Highlighter> {
  if (!highlighter) {
    highlighter = await createHighlighter({
      themes: [THEME],
      langs: Object.keys(bundledLanguages) as BundledLanguage[],
    })
  }
  return highlighter
}

export function isSupportedLang(lang: string | undefined): lang is BundledLanguage {
  if (!lang) return false
  return lang in bundledLanguages
}

/**
 * Tokenize `code` for the given language using Shiki's github-light theme.
 * Returns one array of tokens per line.
 */
export async function highlightCode(
  code: string,
  lang: BundledLanguage,
): Promise<HighlightToken[][]> {
  const h = await getHighlighter()
  const result = h.codeToTokens(code, { lang, theme: THEME })
  return result.tokens.map((line) =>
    line.map((token) => ({
      text: token.content,
      color: (token.color ?? "#24292E").replace("#", ""),
      bold: Boolean(token.fontStyle && token.fontStyle & FONT_STYLE_BOLD),
      italic: Boolean(token.fontStyle && token.fontStyle & FONT_STYLE_ITALIC),
    })),
  )
}
