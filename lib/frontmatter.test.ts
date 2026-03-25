import { describe, expect, test } from "bun:test"
import { parseFrontmatter } from "./frontmatter"

describe("parseFrontmatter", () => {
  test("returns empty data and full content when no frontmatter", () => {
    const { body, data } = parseFrontmatter("# Hello\n\nParagraph.")
    expect(body).toBe("# Hello\n\nParagraph.")
    expect(data).toEqual({})
  })

  test("strips the YAML block and returns the body", () => {
    const { body } = parseFrontmatter("---\ntitle: My Doc\n---\n# Hello")
    expect(body).toBe("# Hello")
    expect(body).not.toContain("title")
  })

  test("parses scalar fields into data", () => {
    const { data } = parseFrontmatter("---\ntitle: My Doc\nauthor: Alice\n---\n# Hello")
    expect(data["title"]).toBe("My Doc")
    expect(data["author"]).toBe("Alice")
  })

  test("parses numeric fields", () => {
    const { data } = parseFrontmatter("---\nfont-size: 14\n---\n")
    expect(data["font-size"]).toBe(14)
  })

  test("parses boolean fields", () => {
    const { data } = parseFrontmatter("---\ndraft: true\n---\n")
    expect(data["draft"]).toBe(true)
  })

  test("parses list fields", () => {
    const { data } = parseFrontmatter("---\ntags:\n  - foo\n  - bar\n---\n")
    expect(data["tags"]).toEqual(["foo", "bar"])
  })

  test("treats a lone opening --- with no closing delimiter as no frontmatter", () => {
    const input = "---\ntitle: oops\n# No closing delimiter"
    const { body, data } = parseFrontmatter(input)
    expect(body).toBe(input)
    expect(data).toEqual({})
  })

  test("body can contain --- without being treated as frontmatter", () => {
    const input = "---\ntitle: Doc\n---\ntext\n---\nmore"
    const { body } = parseFrontmatter(input)
    expect(body).toBe("text\n---\nmore")
  })

  test("malformed YAML returns empty data without throwing", () => {
    const { body, data } = parseFrontmatter("---\n: invalid: yaml: [\n---\n# Body")
    expect(body).toBe("# Body")
    expect(data).toEqual({})
  })
})
