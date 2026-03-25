import { describe, expect, test } from "bun:test"
import { ExternalHyperlink, InternalHyperlink, TextRun } from "docx"
import { marked } from "marked"
import { inlineTokensToRuns } from "./inline"

function parseInline(markdown: string) {
  const tokens = marked.lexer(markdown)
  const para = tokens.find((t) => t.type === "paragraph") as any
  return inlineTokensToRuns(para?.tokens ?? [])
}

describe("inlineTokensToRuns", () => {
  test("plain text → single TextRun", () => {
    const runs = parseInline("hello world")
    expect(runs).toHaveLength(1)
    expect(runs[0]).toBeInstanceOf(TextRun)
  })

  test("**bold** → single TextRun", () => {
    expect(parseInline("**bold**")).toHaveLength(1)
    expect(parseInline("**bold**")[0]).toBeInstanceOf(TextRun)
  })

  test("*italic* → single TextRun", () => {
    expect(parseInline("*italic*")).toHaveLength(1)
    expect(parseInline("*italic*")[0]).toBeInstanceOf(TextRun)
  })

  test("`code` → single TextRun", () => {
    expect(parseInline("`code`")).toHaveLength(1)
    expect(parseInline("`code`")[0]).toBeInstanceOf(TextRun)
  })

  test("[link](url) → ExternalHyperlink", () => {
    const runs = parseInline("[click here](https://example.com)")
    expect(runs).toHaveLength(1)
    expect(runs[0]).toBeInstanceOf(ExternalHyperlink)
  })

  test("mixed inline → multiple runs", () => {
    expect(parseInline("plain **bold** and *italic*").length).toBeGreaterThan(1)
  })

  test("**`bold code`** — nested codespan inside strong → single TextRun (not lost)", () => {
    const runs = parseInline("**`bold code`**")
    expect(runs).toHaveLength(1)
    expect(runs[0]).toBeInstanceOf(TextRun)
  })

  test("**bold *italic*** — nested em inside strong → all TextRuns", () => {
    const runs = parseInline("**bold *italic***")
    expect(runs.length).toBeGreaterThan(0)
    expect(runs.every((r) => r instanceof TextRun)).toBe(true)
  })

  test("[**bold**](url) — nested strong in link → ExternalHyperlink", () => {
    const runs = parseInline("[**bold link**](https://example.com)")
    expect(runs).toHaveLength(1)
    expect(runs[0]).toBeInstanceOf(ExternalHyperlink)
  })

  test("[link](#anchor) → InternalHyperlink", () => {
    const runs = parseInline("[go there](#my-section)")
    expect(runs).toHaveLength(1)
    expect(runs[0]).toBeInstanceOf(InternalHyperlink)
  })

  test("[link](#anchor) → ExternalHyperlink is not emitted", () => {
    const runs = parseInline("[go there](#my-section)")
    expect(runs[0]).not.toBeInstanceOf(ExternalHyperlink)
  })

  test("empty token list → empty array", () => {
    expect(inlineTokensToRuns([])).toHaveLength(0)
  })
})
