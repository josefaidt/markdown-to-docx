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

  test("`--flag` codespan — double-dash is not converted to em dash", () => {
    const runs = parseInline("`--some-flag`")
    expect(runs).toHaveLength(1)
    const run = runs[0] as TextRun
    const wt = (run as any).root.find((n: any) => n.rootKey === "w:t")
    expect(wt.root[1]).toBe("--some-flag")
  })

  test("plain text -- double-dash is converted to em dash", () => {
    const runs = parseInline("foo -- bar")
    expect(runs).toHaveLength(1)
    const wt = (runs[0] as TextRun as any).root.find((n: any) => n.rootKey === "w:t")
    expect(wt.root[1]).toBe("foo \u2014 bar")
  })

  test("empty token list → empty array", () => {
    expect(inlineTokensToRuns([])).toHaveLength(0)
  })
})
