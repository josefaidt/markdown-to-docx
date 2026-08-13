import { describe, test, expect } from "bun:test"
import { mkdtempSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { Packer } from "docx"
import JSZip from "jszip"
import { convertMarkdownToDocx } from "./markdown-to-docx"
import { parsePageSize } from "./page-size"

/** Minimal PNG whose IHDR declares the given dimensions — enough for probe to size it */
function widePng(width: number, height: number): Buffer {
  const png = Buffer.from([
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
    0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
    0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
    0x44, 0xae, 0x42, 0x60, 0x82,
  ])
  png.writeUInt32BE(width, 16)
  png.writeUInt32BE(height, 20)
  return png
}

async function buildDocx(
  markdown: string,
  options: Parameters<typeof convertMarkdownToDocx>[2] = {},
) {
  const mdPath = join(tmpdir(), `test-${Date.now()}.md`)
  writeFileSync(mdPath, markdown)

  const doc = await convertMarkdownToDocx(mdPath, "", options)
  const buf = await Packer.toBuffer(doc)
  const zip = await JSZip.loadAsync(buf)

  async function xml(name: string) {
    const file = zip.file(name)
    if (!file) throw new Error(`${name} not found in docx zip`)
    return file.async("string")
  }

  return { zip, xml }
}

async function headerFiles(zip: JSZip): Promise<Record<string, string>> {
  const relsXml = await zip.file("word/_rels/document.xml.rels")!.async("string")
  const docXml = await zip.file("word/document.xml")!.async("string")

  const relMap: Record<string, string> = {}
  for (const m of relsXml.matchAll(/Id="(rId\d+)"[^>]+Target="(header\d+\.xml)"/g)) {
    relMap[m[1]!] = `word/${m[2]}`
  }

  const result: Record<string, string> = {}
  for (const m of docXml.matchAll(/w:headerReference w:type="(\w+)" r:id="(rId\d+)"/g)) {
    const filename = relMap[m[2]!]
    if (filename) result[m[1]!] = filename
  }
  return result
}

describe("footer", () => {
  test("has a right-aligned tab stop at text width (9746 twips for Moderate margins)", async () => {
    const { xml } = await buildDocx("# Hello", { footerLabel: "My Label" })
    const footer = await xml("word/footer1.xml")
    expect(footer).toContain('w:val="right"')
    expect(footer).toContain('w:pos="9746"')
  })

  test("renders the footer label text", async () => {
    const { xml } = await buildDocx("# Hello", { footerLabel: "Acme Corp" })
    const footer = await xml("word/footer1.xml")
    expect(footer).toContain("Acme Corp")
  })

  test("page number field appears after the tab, not inline with the label", async () => {
    const { xml } = await buildDocx("# Hello", { footerLabel: "Left Side", footerPageNumber: true })
    const footer = await xml("word/footer1.xml")

    const labelPos = footer.indexOf("Left Side")
    const tabPos = footer.indexOf("<w:tab/>")
    const pageFieldPos = footer.indexOf("<w:instrText")

    expect(labelPos).toBeGreaterThan(-1)
    expect(tabPos).toBeGreaterThan(-1)
    expect(pageFieldPos).toBeGreaterThan(-1)

    expect(labelPos).toBeLessThan(tabPos)
    expect(tabPos).toBeLessThan(pageFieldPos)
  })

  test("page number uses PAGE field instruction", async () => {
    const { xml } = await buildDocx("# Hello", { footerLabel: "Footer", footerPageNumber: true })
    const footer = await xml("word/footer1.xml")
    expect(footer).toContain(">PAGE<")
  })

  test("label and page number are in the same paragraph", async () => {
    const { xml } = await buildDocx("# Hello", { footerLabel: "Footer", footerPageNumber: true })
    const footer = await xml("word/footer1.xml")
    const paragraphs = footer.match(/<w:p[ >]/g) ?? []
    expect(paragraphs).toHaveLength(1)
  })

  test("page number is present even without a footerLabel", async () => {
    const { xml } = await buildDocx("# Hello", { footerPageNumber: true })
    const footer = await xml("word/footer1.xml")
    expect(footer).toContain(">PAGE<")
  })

  test("label text is absent when footerLabel is not provided", async () => {
    const { xml } = await buildDocx("# Hello", { footerPageNumber: true })
    const footer = await xml("word/footer1.xml")
    // Only the tab + PAGE runs should be present; no stray label text run
    const runs = footer.match(/<w:r>/g) ?? []
    expect(runs).toHaveLength(2)
  })

  test("label and page number runs share the same font size via FooterText style", async () => {
    const { xml } = await buildDocx("# Hello", { footerLabel: "Corp" })
    const footer = await xml("word/footer1.xml")
    // Neither run should carry an inline w:sz override — size comes from the style
    expect(footer).not.toContain("<w:sz ")
  })

  test("footer appears on the first page when a header is configured", async () => {
    const { zip } = await buildDocx("# Hello", { headerLabel: "My Header", footerPageNumber: true })
    const docXml = await zip.file("word/document.xml")!.async("string")
    // A first-page footer reference must exist alongside the default one
    expect(docXml).toContain('w:footerReference w:type="first"')
    // Find the first-page footer file and confirm it has the PAGE field
    const relsXml = await zip.file("word/_rels/document.xml.rels")!.async("string")
    const firstFooterId = docXml.match(/w:footerReference w:type="first" r:id="(rId\d+)"/)?.[1]
    expect(firstFooterId).toBeDefined()
    const firstFooterFile = relsXml.match(
      new RegExp(`Id="${firstFooterId}"[^>]+Target="(footer\\d+\\.xml)"`),
    )?.[1]
    expect(firstFooterFile).toBeDefined()
    const firstFooter = await zip.file(`word/${firstFooterFile!}`)!.async("string")
    expect(firstFooter).toContain(">PAGE<")
  })
})

async function bodyXml(
  markdown: string,
  options: Parameters<typeof convertMarkdownToDocx>[2] = {},
) {
  const { zip } = await buildDocx(markdown, options)
  const docXml = await zip.file("word/document.xml")!.async("string")
  return docXml.match(/<w:body>(.*)<\/w:body>/s)?.[1] ?? ""
}

describe("page size", () => {
  async function pageSizeAttrs(options: Parameters<typeof convertMarkdownToDocx>[2] = {}) {
    const { xml } = await buildDocx("# Hello", options)
    const doc = await xml("word/document.xml")
    const pgSz = doc.match(/<w:pgSz [^/]*\/>/)?.[0] ?? ""
    return {
      width: Number(pgSz.match(/w:w="(\d+)"/)?.[1]),
      height: Number(pgSz.match(/w:h="(\d+)"/)?.[1]),
      orientation: pgSz.match(/w:orient="(\w+)"/)?.[1],
    }
  }

  test("defaults to A4", async () => {
    expect(await pageSizeAttrs()).toMatchObject({ width: 11906, height: 16838 })
  })

  test("applies a named size", async () => {
    expect(await pageSizeAttrs({ pageSize: parsePageSize("legal") })).toMatchObject({
      width: 12240,
      height: 20160,
    })
  })

  test("applies a custom size", async () => {
    expect(await pageSizeAttrs({ pageSize: parsePageSize("9x12in") })).toMatchObject({
      width: 12960,
      height: 17280,
    })
  })

  test("accepts a size spec string directly", async () => {
    expect(await pageSizeAttrs({ pageSize: "Legal" })).toMatchObject({
      width: 12240,
      height: 20160,
    })
  })

  test("rejects an invalid size spec string", async () => {
    await expect(buildDocx("# Hello", { pageSize: "nope" })).rejects.toThrow(/Unknown page size/)
  })

  test("a taller-than-wide size is emitted as a portrait page", async () => {
    expect((await pageSizeAttrs({ pageSize: parsePageSize("letter") })).orientation).toBe(
      "portrait",
    )
  })

  test("a wider-than-tall size is emitted as a landscape page", async () => {
    const { width, height, orientation } = await pageSizeAttrs({
      pageSize: parsePageSize("11x8.5"),
    })
    expect(orientation).toBe("landscape")
    // Word expects the printed dimensions, not the pre-rotation ones
    expect(width).toBe(15840)
    expect(height).toBe(12240)
  })

  test("margins are unchanged by the page size", async () => {
    const { xml } = await buildDocx("# Hello", { pageSize: parsePageSize("legal") })
    const doc = await xml("word/document.xml")
    expect(doc).toContain('w:top="1440"')
    expect(doc).toContain('w:bottom="1440"')
    expect(doc).toContain('w:left="1080"')
    expect(doc).toContain('w:right="1080"')
  })

  test("footer tab stop tracks the page width", async () => {
    const { xml } = await buildDocx("# Hello", {
      footerLabel: "Label",
      footerPageNumber: true,
      pageSize: parsePageSize("letter"),
    })
    const footer = await xml("word/footer1.xml")
    // Letter (12240) less 0.75in left + right margins
    expect(footer).toContain(`w:pos="${12240 - 1080 - 1080}"`)
  })

  test("images are capped to the text width of a narrow page", async () => {
    const dir = mkdtempSync(join(tmpdir(), "page-size-"))
    const imgPath = join(dir, "wide.png")
    writeFileSync(imgPath, widePng(1000, 500))
    const mdPath = join(dir, "doc.md")
    writeFileSync(mdPath, "![wide](./wide.png)")

    const doc = await convertMarkdownToDocx(mdPath, "", { pageSize: parsePageSize("4x6in") })
    const buf = await Packer.toBuffer(doc)
    const zip = await JSZip.loadAsync(buf)
    const body = await zip.file("word/document.xml")!.async("string")

    // 4in page less 1.5in of margins = 2.5in of text width = 240px = 2286000 EMU
    const cx = Number(body.match(/<wp:extent cx="(\d+)"/)?.[1])
    expect(cx).toBe(2286000)
  })
})

describe("frontmatter", () => {
  test("YAML frontmatter is stripped from output", async () => {
    const body = await bodyXml("---\ntitle: My Doc\nauthor: Alice\n---\n# Hello")
    expect(body).not.toContain("title: My Doc")
    expect(body).not.toContain("author: Alice")
    expect(body).toContain("Hello")
  })

  test("document without frontmatter renders normally", async () => {
    const body = await bodyXml("# Hello")
    expect(body).toContain("Hello")
  })

  test("HTML comment between frontmatter and first heading does not add an empty paragraph before the heading", async () => {
    const md = "---\ntitle: A New Doc\n---\n\n<!-- internal comment -->\n\n# The First Heading"
    const body = await bodyXml(md)
    expect(body).not.toContain("internal comment")
    // Only one <w:p element should exist — the heading itself (no preceding empty paragraph)
    const paragraphs = body.match(/<w:p[ >]/g) ?? []
    expect(paragraphs).toHaveLength(1)
  })

  test("HTML comment mid-document is silently ignored", async () => {
    const md = "# Heading\n\n<!-- a comment -->\n\nSome text."
    const body = await bodyXml(md)
    expect(body).not.toContain("a comment")
    expect(body).toContain("Heading")
    expect(body).toContain("Some text.")
  })
})

describe("headings", () => {
  for (const [depth, style] of [
    [1, "Heading1"],
    [2, "Heading2"],
    [3, "Heading3"],
    [4, "Heading4"],
    [5, "Heading5"],
    [6, "Heading6"],
  ] as const) {
    test(`h${depth} uses style ${style}`, async () => {
      const hashes = "#".repeat(depth)
      const body = await bodyXml(`${hashes} My Heading`)
      expect(body).toContain(`w:val="${style}"`)
      expect(body).toContain("My Heading")
    })
  }

  test("heading with inline `code` uses InlineCode style", async () => {
    const body = await bodyXml("## Using `foo()` in headings")
    expect(body).toContain('w:val="Heading2"')
    expect(body).toContain('w:val="InlineCode"')
    expect(body).toContain("foo()")
  })

  test("heading containing 'Appendix' inserts a page break before it", async () => {
    const body = await bodyXml("# Appendix A")
    expect(body).toContain("<w:pageBreakBefore/>")
  })

  test("non-appendix heading does not insert a page break", async () => {
    const body = await bodyXml("# Introduction")
    expect(body).not.toContain("<w:pageBreakBefore/>")
  })

  test("headings do not get bookmarks by default", async () => {
    const body = await bodyXml("## My Section")
    expect(body).not.toContain("<w:bookmarkStart")
    expect(body).not.toContain("<w:bookmarkEnd")
  })

  test("heading gets a bookmark with a slugified id when bookmarks enabled", async () => {
    const body = await bodyXml("## My Section", { bookmarks: true })
    expect(body).toContain("<w:bookmarkStart")
    expect(body).toContain('w:name="my-section"')
  })

  test("heading slug strips punctuation and lowercases", async () => {
    const body = await bodyXml("# Hello, World!", { bookmarks: true })
    expect(body).toContain('w:name="hello-world"')
  })

  test("heading slug collapses multiple spaces/hyphens", async () => {
    const body = await bodyXml("## Two  Words", { bookmarks: true })
    expect(body).toContain('w:name="two-words"')
  })

  test("multiple headings get unique bookmark ids", async () => {
    const body = await bodyXml("# One\n\n## Two\n\n### Three", { bookmarks: true })
    const ids = [...body.matchAll(/w:bookmarkStart[^>]+w:id="(\d+)"/g)].map((m) => m[1])
    expect(ids.length).toBe(3)
    expect(new Set(ids).size).toBe(3)
  })
})

describe("inline formatting", () => {
  test("plain text renders as Normal paragraph", async () => {
    const body = await bodyXml("Just some text.")
    expect(body).toContain('w:val="Normal"')
    expect(body).toContain("Just some text.")
  })

  test("**bold** wraps text in <w:b/>", async () => {
    const body = await bodyXml("Hello **world**")
    expect(body).toContain("<w:b/>")
    expect(body).toContain("world")
  })

  test("*italic* wraps text in <w:i/>", async () => {
    const body = await bodyXml("Hello *world*")
    expect(body).toContain("<w:i/>")
    expect(body).toContain("world")
  })

  test("`inline code` uses InlineCode character style", async () => {
    const body = await bodyXml("Use `foo()` here")
    expect(body).toContain('w:val="InlineCode"')
    expect(body).toContain("foo()")
  })

  test("InlineCode style has no explicit color and no explicit size (inherits from parent)", async () => {
    const { zip } = await buildDocx("`foo`")
    const stylesXml = await zip.file("word/styles.xml")!.async("string")
    const inlineCodeChunk =
      stylesXml.split(/<w:style\s/).find((c) => c.includes('"InlineCode"')) ?? ""
    expect(inlineCodeChunk).not.toContain('w:val="555555"')
    expect(inlineCodeChunk).not.toContain("w:sz")
  })

  test("`code#with-hash` renders as a single InlineCode run — not split at #", async () => {
    const body = await bodyXml("Use `package.json#name` here")
    // The full text must appear as a single w:t value, not split into two runs
    expect(body).toContain("package.json#name")
    // Verify there is exactly one InlineCode run (not two)
    const inlineRuns = body.match(/<w:rStyle w:val="InlineCode"\/>/g) ?? []
    expect(inlineRuns).toHaveLength(1)
  })

  test("[link](url) renders as hyperlink with Hyperlink style", async () => {
    const body = await bodyXml("See [the docs](https://example.com)")
    expect(body).toContain("<w:hyperlink")
    expect(body).toContain('w:val="Hyperlink"')
    expect(body).toContain("the docs")
  })

  test("[link](#anchor) renders as internal hyperlink", async () => {
    const body = await bodyXml("See [this section](#my-section)")
    expect(body).toContain('w:anchor="my-section"')
    expect(body).toContain("this section")
  })

  test("[link](#anchor) does not emit an external relationship", async () => {
    const { zip } = await buildDocx("See [this section](#my-section)")
    const relsXml = await zip.file("word/_rels/document.xml.rels")!.async("string")
    expect(relsXml).not.toContain("#my-section")
  })

  test("**`bold code`** renders with both bold and InlineCode style", async () => {
    const body = await bodyXml("**`bold code`**")
    expect(body).toContain("<w:b/>")
    expect(body).toContain('w:val="InlineCode"')
    expect(body).toContain("bold code")
  })

  test("**bold *italic*** renders with both bold and italic runs", async () => {
    const body = await bodyXml("**bold *italic***")
    expect(body).toContain("<w:b/>")
    expect(body).toContain("<w:i/>")
  })

  test("[**bold**](url) renders as hyperlink containing bold text", async () => {
    const body = await bodyXml("[**bold link**](https://example.com)")
    expect(body).toContain("<w:hyperlink")
    expect(body).toContain("bold link")
  })
})

describe("blockquote", () => {
  test("uses Blockquote style", async () => {
    const body = await bodyXml("> Some quoted text")
    expect(body).toContain('w:val="Blockquote"')
    expect(body).toContain("Some quoted text")
  })

  test("**bold** inside blockquote renders <w:b/>", async () => {
    const body = await bodyXml("> **Note** some text")
    expect(body).toContain('w:val="Blockquote"')
    expect(body).toContain("<w:b/>")
    expect(body).toContain("Note")
  })

  test("*italic* inside blockquote renders <w:i/>", async () => {
    const body = await bodyXml("> *emphasis* here")
    expect(body).toContain('w:val="Blockquote"')
    expect(body).toContain("<w:i/>")
  })

  test("`code` inside blockquote renders InlineCode style", async () => {
    const body = await bodyXml("> use `foo()` here")
    expect(body).toContain('w:val="Blockquote"')
    expect(body).toContain('w:val="InlineCode"')
    expect(body).toContain("foo()")
  })
})

describe("lists", () => {
  async function listFormat(markdown: string) {
    const { zip } = await buildDocx(markdown)
    const body =
      (await zip.file("word/document.xml")!.async("string")).match(
        /<w:body>(.*)<\/w:body>/s,
      )?.[1] ?? ""
    const numId = body.match(/w:numId w:val="(\d+)"/)?.[1]
    if (!numId) return null
    const numXml = await zip.file("word/numbering.xml")!.async("string")
    const abstractNumId = numXml.match(
      new RegExp(`<w:num w:numId="${numId}"><w:abstractNumId w:val="(\\d+)"`),
    )?.[1]
    if (!abstractNumId) return null
    const abstractNum =
      numXml.match(
        new RegExp(`<w:abstractNum w:abstractNumId="${abstractNumId}".*?</w:abstractNum>`, "s"),
      )?.[0] ?? ""
    return abstractNum.match(/<w:numFmt w:val="(\w+)"/)?.[1] ?? null
  }

  test("unordered list items use bullet format", async () => {
    const fmt = await listFormat("- alpha\n- beta")
    expect(fmt).toBe("bullet")
  })

  test("unordered list renders item text", async () => {
    const body = await bodyXml("- alpha\n- beta")
    expect(body).toContain("alpha")
    expect(body).toContain("beta")
  })

  test("ordered list items use decimal format", async () => {
    const fmt = await listFormat("1. first\n2. second")
    expect(fmt).toBe("decimal")
  })

  test("ordered list renders item text", async () => {
    const body = await bodyXml("1. first\n2. second")
    expect(body).toContain("first")
    expect(body).toContain("second")
  })

  test("list items use ListItem style", async () => {
    const body = await bodyXml("- item")
    expect(body).toContain('w:val="ListItem"')
  })

  test("ListItem style has contextualSpacing enabled", async () => {
    const { zip } = await buildDocx("- item")
    const stylesXml = await zip.file("word/styles.xml")!.async("string")
    const listItemChunk = stylesXml.split(/<w:style\s/).find((c) => c.includes('"ListItem"')) ?? ""
    expect(listItemChunk).toContain("<w:contextualSpacing/>")
  })

  test("nested list items increment the ilvl", async () => {
    const body = await bodyXml("- parent\n  - child")
    expect(body).toContain('w:val="0"')
    expect(body).toContain('w:val="1"')
  })
})

describe("code blocks", () => {
  test("fenced code uses CodeBlock style", async () => {
    const body = await bodyXml("```js\nconst x = 1\n```")
    expect(body).toContain('w:val="CodeBlock"')
    // with highlighting, tokens are split — verify each token appears
    expect(body).toContain("const")
    expect(body).toContain(">x<")
    expect(body).toContain(">1<")
  })

  test("code block is wrapped in a borderless table with grey shading", async () => {
    const body = await bodyXml("```\nhello\n```")
    expect(body).toContain('w:val="none"')
    expect(body).toContain('w:fill="F6F8FA"')
  })

  test("multi-line code block produces one paragraph per line", async () => {
    const body = await bodyXml("```\nline one\nline two\nline three\n```")
    const matches = body.match(/line one|line two|line three/g) ?? []
    expect(matches).toHaveLength(3)
  })

  test("code block is not followed by a spacer paragraph", async () => {
    const body = await bodyXml("```\nhello\n```")
    expect(body).not.toContain('w:val="CodeBlockSpacing"')
  })

  test("highlighted code block emits colored runs for a known language", async () => {
    const body = await bodyXml("```typescript\nconst x = 1\n```")
    // 'const' is a keyword — shiki github-light renders it in red D73A49
    expect(body).toContain('w:val="D73A49"')
    // numeric literal '1' is rendered in blue 005CC5
    expect(body).toContain('w:val="005CC5"')
  })

  test("code block without a language falls back to plain text (no color runs)", async () => {
    const body = await bodyXml("```\nhello world\n```")
    expect(body).toContain("hello world")
    // plain fallback uses CodeBlock style color via stylesheet, no per-run color
    expect(body).not.toContain("<w:color")
  })
})

describe("tables", () => {
  test("table uses Table Grid Light built-in style", async () => {
    const body = await bodyXml("| A | B |\n|---|---|\n| 1 | 2 |")
    expect(body).toContain('w:val="TableGridLight"')
  })

  test("table header row has tblHeader set", async () => {
    const body = await bodyXml("| A | B |\n|---|---|\n| 1 | 2 |")
    expect(body).toContain("<w:tblHeader/>")
  })

  test("table body cell text is present", async () => {
    const body = await bodyXml("| A | B |\n|---|---|\n| foo | bar |")
    expect(body).toContain("foo")
    expect(body).toContain("bar")
  })

  test("table is full-width and autofit so it inherits the page width", async () => {
    const body = await bodyXml("| A |\n|---|\n| x |")
    // Percentage width + autofit lets Word distribute columns to fill whatever
    // page size and margins are in effect, rather than a hardcoded width.
    expect(body).toContain('w:type="pct"')
    expect(body).toContain('<w:tblLayout w:type="autofit"/>')
  })

  test("table grid columns are non-degenerate and track content", async () => {
    const body = await bodyXml(
      "| # | Condition | Response |\n| :-- | :-- | :-- |\n| 1 | a running deployment | passes through |",
    )
    const grid = body.match(/<w:tblGrid>(.*?)<\/w:tblGrid>/s)?.[1] ?? ""
    const cols = [...grid.matchAll(/<w:gridCol w:w="(\d+)"\/>/g)].map((m) => Number(m[1]))
    expect(cols).toHaveLength(3)
    // No degenerate placeholder widths (the old bug emitted identical 100-twip
    // columns that Word Online could not reconcile with the declared width).
    expect(cols.every((w) => w > 0)).toBe(true)
    expect(new Set(cols).size).toBeGreaterThan(1)
    // The content-heavy columns start wider than the "#" column.
    expect(cols[1]!).toBeGreaterThan(cols[0]!)
    expect(cols[2]!).toBeGreaterThan(cols[0]!)
  })
})

describe("list spacing", () => {
  test("blank line between bullet list and ordered list inserts a spacer paragraph", async () => {
    const md = "- alpha\n- beta\n\n1. first\n2. second"
    const body = await bodyXml(md)
    // Both list items must appear
    expect(body).toContain("alpha")
    expect(body).toContain("first")
    // A spacer paragraph (size 12, exact line height 160) must separate them
    expect(body).toMatch(/<w:sz w:val="12"\/>/)
  })

  test("no blank line between bullet list and ordered list produces no spacer", async () => {
    // When there is no blank line, marked treats it as a single list; no spacer expected
    const md = "- alpha\n- beta\n1. first\n2. second"
    const body = await bodyXml(md)
    expect(body).toContain("alpha")
    expect(body).toContain("first")
    // No tiny spacer run between adjacent items in the same list block
    const spacerCount = (body.match(/<w:sz w:val="12"\/>/g) ?? []).length
    expect(spacerCount).toBe(0)
  })

  test("blank line between list and following paragraph does not insert a spacer", async () => {
    const md = "- alpha\n- beta\n\nSome paragraph text."
    const body = await bodyXml(md)
    expect(body).toContain("alpha")
    expect(body).toContain("Some paragraph text.")
    const spacerCount = (body.match(/<w:sz w:val="12"\/>/g) ?? []).length
    expect(spacerCount).toBe(0)
  })

  test("blank line between list and following heading does not insert a spacer", async () => {
    const md = "- alpha\n- beta\n\n## Next Section"
    const body = await bodyXml(md)
    expect(body).toContain("alpha")
    expect(body).toContain("Next Section")
    const spacerCount = (body.match(/<w:sz w:val="12"\/>/g) ?? []).length
    expect(spacerCount).toBe(0)
  })
})

describe("list inline formatting", () => {
  test("**bold** inside list item renders <w:b/>", async () => {
    const body = await bodyXml("- **bold item**")
    expect(body).toContain("<w:b/>")
    expect(body).toContain("bold item")
  })

  test("`code` inside list item renders InlineCode style", async () => {
    const body = await bodyXml("- use `foo()` here")
    expect(body).toContain('w:val="InlineCode"')
    expect(body).toContain("foo()")
  })

  // Regression tests for https://github.com/josefaidt/markdown-to-docx/issues/63
  // The reported bug: inline code adjacent to other decorations in a list item was skipped.
  // Inline code inside a link is covered separately because the link handler flattens children.
  test("`code` + **bold** + `code` in one list item → both codespans survive", async () => {
    const body = await bodyXml(
      "- this is some `inline code` but only after this **bolded** text do `inline code` blocks actually render correctly",
    )
    const codespans = body.match(/<w:rStyle w:val="InlineCode"\/>/g) ?? []
    expect(codespans).toHaveLength(2)
    expect(body).toContain("<w:b/>")
    expect(body).toContain("this is some")
    expect(body).toContain("bolded")
    expect(body).toContain("blocks actually render correctly")
  })

  test("`code` + *italic* + `code` in one list item → both codespans survive", async () => {
    const body = await bodyXml("- pre `one` mid *emph* post `two` end")
    const codespans = body.match(/<w:rStyle w:val="InlineCode"\/>/g) ?? []
    expect(codespans).toHaveLength(2)
    expect(body).toContain("<w:i/>")
  })

  test("codespan adjacent to **bold** with no separating space still renders", async () => {
    const body = await bodyXml("- lead**bold**`code`tail")
    const codespans = body.match(/<w:rStyle w:val="InlineCode"\/>/g) ?? []
    expect(codespans).toHaveLength(1)
    expect(body).toContain("<w:b/>")
    expect(body).toContain("bold")
    expect(body).toContain("code")
  })

  test("**bold** then `code` then *italic* then `code` — decoration chain around codespans", async () => {
    const body = await bodyXml("- **b** `c1` *i* `c2` end")
    const codespans = body.match(/<w:rStyle w:val="InlineCode"\/>/g) ?? []
    expect(codespans).toHaveLength(2)
    expect(body).toContain("<w:b/>")
    expect(body).toContain("<w:i/>")
  })

  test("codespan inside **nested strong** alongside a sibling codespan — both render", async () => {
    const body = await bodyXml("- **bold `inner`** then `outer`")
    const codespans = body.match(/<w:rStyle w:val="InlineCode"\/>/g) ?? []
    expect(codespans).toHaveLength(2)
    expect(body).toContain("<w:b/>")
  })

  test("ordered list: `code` + **bold** + `code` → both codespans survive", async () => {
    const body = await bodyXml("1. pre `one` mid **bold** post `two` end")
    const codespans = body.match(/<w:rStyle w:val="InlineCode"\/>/g) ?? []
    expect(codespans).toHaveLength(2)
    expect(body).toContain("<w:b/>")
  })

  test("nested list item: `code` + **bold** + `code` → both codespans survive", async () => {
    const body = await bodyXml("- parent\n  - pre `one` mid **bold** post `two` end")
    const codespans = body.match(/<w:rStyle w:val="InlineCode"\/>/g) ?? []
    expect(codespans).toHaveLength(2)
    expect(body).toContain("<w:b/>")
  })

  test("loose list item (blank line between items): codespans + **bold** all render", async () => {
    const body = await bodyXml("- first\n\n- pre `one` mid **bold** post `two` end")
    const codespans = body.match(/<w:rStyle w:val="InlineCode"\/>/g) ?? []
    expect(codespans).toHaveLength(2)
    expect(body).toContain("<w:b/>")
  })
})

describe("nested code blocks in list items", () => {
  const md = [
    "1. item one",
    "",
    "2. item two",
    "",
    "   ```bash",
    "   some-command",
    "   ```",
    "",
    "3. item three",
  ].join("\n")

  test("list item text is present for all items", async () => {
    const body = await bodyXml(md)
    expect(body).toContain("item one")
    expect(body).toContain("item two")
    expect(body).toContain("item three")
  })

  test("fenced code block inside list item renders with CodeBlock style", async () => {
    const body = await bodyXml(md)
    expect(body).toContain('w:val="CodeBlock"')
  })

  test("fenced code block text is not stripped", async () => {
    const body = await bodyXml(md)
    expect(body).toContain("some-command")
  })

  test("fenced code block inside list item has grey background shading", async () => {
    const body = await bodyXml(md)
    expect(body).toContain('w:fill="F6F8FA"')
  })

  test("fenced code block inside list item is indented to match list level", async () => {
    const body = await bodyXml(md)
    // level 0: LIST_TEXT_INDENT = convertInchesToTwip(0.45) = 648 twips
    expect(body).toContain('w:w="648"')
  })

  test("spacer paragraph after nested code block uses exact 160-twip line height with no after spacing", async () => {
    const body = await bodyXml(md)
    expect(body).toContain('w:lineRule="exact"')
    expect(body).toContain('w:line="160"')
  })

  test("plain (no lang) fenced code block inside list item renders CodeBlock style", async () => {
    const plain = "1. item\n\n   ```\n   plain code\n   ```\n\n2. next"
    const body = await bodyXml(plain)
    expect(body).toContain('w:val="CodeBlock"')
    expect(body).toContain("plain code")
  })

  test("highlighted fenced code block inside list item emits colored runs", async () => {
    const highlighted = "1. item\n\n   ```typescript\n   const x = 1\n   ```\n\n2. next"
    const body = await bodyXml(highlighted)
    expect(body).toContain('w:val="D73A49"')
  })

  test("unordered list item with nested code block renders CodeBlock", async () => {
    const bullet = "- item\n\n  ```\n  bullet-code\n  ```\n\n- next"
    const body = await bodyXml(bullet)
    expect(body).toContain('w:val="CodeBlock"')
    expect(body).toContain("bullet-code")
  })
})

describe("numbering level formats", () => {
  function orderedAbstractNum(numXml: string): string {
    const blocks = numXml.match(/<w:abstractNum\b[^>]*>[\s\S]*?<\/w:abstractNum>/g) ?? []
    return blocks.find((b) => b.includes('"decimal"')) ?? ""
  }

  function bulletAbstractNum(numXml: string): string {
    const blocks = numXml.match(/<w:abstractNum\b[^>]*>[\s\S]*?<\/w:abstractNum>/g) ?? []
    // The custom bullet numbering block has all-bullet levels; pick the last one
    // (docx emits a default bullet block first, then ours)
    return [...blocks].reverse().find((b) => !b.includes('"decimal"')) ?? ""
  }

  test("ordered list level 0 uses decimal format", async () => {
    const { zip } = await buildDocx("1. a")
    const numXml = await zip.file("word/numbering.xml")!.async("string")
    const levels = orderedAbstractNum(numXml).match(/<w:numFmt w:val="(\w+)"/g) ?? []
    expect(levels[0]).toBe('<w:numFmt w:val="decimal"')
  })

  test("ordered list level 1 uses lowerLetter format", async () => {
    const { zip } = await buildDocx("1. a")
    const numXml = await zip.file("word/numbering.xml")!.async("string")
    const levels = orderedAbstractNum(numXml).match(/<w:numFmt w:val="(\w+)"/g) ?? []
    expect(levels[1]).toBe('<w:numFmt w:val="lowerLetter"')
  })

  test("ordered list level 2 uses lowerRoman format", async () => {
    const { zip } = await buildDocx("1. a")
    const numXml = await zip.file("word/numbering.xml")!.async("string")
    const levels = orderedAbstractNum(numXml).match(/<w:numFmt w:val="(\w+)"/g) ?? []
    expect(levels[2]).toBe('<w:numFmt w:val="lowerRoman"')
  })

  test("bullet numbering uses bullet numFmt at all defined levels", async () => {
    const { zip } = await buildDocx("- a")
    const numXml = await zip.file("word/numbering.xml")!.async("string")
    const levels = bulletAbstractNum(numXml).match(/<w:numFmt w:val="(\w+)"/g) ?? []
    expect(levels.length).toBeGreaterThan(0)
    expect(levels.every((l) => l === '<w:numFmt w:val="bullet"')).toBe(true)
  })
})

describe("line numbers", () => {
  test("lineNumbers: true adds w:lnNumType to section properties", async () => {
    const { zip } = await buildDocx("# Hello", { lineNumbers: true })
    const docXml = await zip.file("word/document.xml")!.async("string")
    expect(docXml).toContain("<w:lnNumType")
  })

  test("lineNumbers not set produces no w:lnNumType", async () => {
    const { zip } = await buildDocx("# Hello")
    const docXml = await zip.file("word/document.xml")!.async("string")
    expect(docXml).not.toContain("<w:lnNumType")
  })
})

describe("horizontal rule", () => {
  test("--- renders a paragraph with a bottom border", async () => {
    const body = await bodyXml("text\n\n---\n\nmore")
    expect(body).toContain("<w:bottom")
  })
})

describe("images", () => {
  test("unresolvable image renders italic alt-text placeholder", async () => {
    const body = await bodyXml("![alt text](./nonexistent.png)")
    expect(body).toContain("[Image: alt text]")
    expect(body).toContain("<w:i/>")
  })
})

describe("font size scaling", () => {
  async function styleSizes(
    options: Parameters<typeof convertMarkdownToDocx>[2] = {},
  ): Promise<Map<string, number>> {
    const { zip } = await buildDocx("# Hello\n\nParagraph.", options)
    const stylesXml = await zip.file("word/styles.xml")!.async("string")
    const map = new Map<string, number>()
    const chunks = stylesXml.split(/<w:style\s/)
    for (const chunk of chunks) {
      const idMatch = chunk.match(/w:styleId="([^"]+)"/)
      if (!idMatch) continue
      const szMatch = chunk.match(/<w:sz w:val="(\d+)"/)
      if (szMatch) map.set(idMatch[1]!, parseInt(szMatch[1]!, 10))
    }
    return map
  }

  test("default (no fontSize) → Normal = 22 half-points", async () => {
    const sizes = await styleSizes()
    expect(sizes.get("Normal")).toBe(22)
  })

  test("fontSize: 11 explicit === no fontSize implicit", async () => {
    const implicit = await styleSizes()
    const explicit = await styleSizes({ fontSize: 11 })
    for (const [id, size] of implicit) {
      expect(explicit.get(id)).toBe(size)
    }
  })

  test("fontSize: 10 → Normal = 20 half-points", async () => {
    const sizes = await styleSizes({ fontSize: 10 })
    expect(sizes.get("Normal")).toBe(20)
  })

  test("fontSize: 10 → Heading1 = 48 half-points (10+14=24pt)", async () => {
    const sizes = await styleSizes({ fontSize: 10 })
    expect(sizes.get("Heading1")).toBe(48)
  })

  test("fontSize: 10 → Heading2 = 34 half-points (10+7=17pt)", async () => {
    const sizes = await styleSizes({ fontSize: 10 })
    expect(sizes.get("Heading2")).toBe(34)
  })

  test("fontSize: 10 → Heading3 = 26 half-points (10+3=13pt)", async () => {
    const sizes = await styleSizes({ fontSize: 10 })
    expect(sizes.get("Heading3")).toBe(26)
  })

  test("fontSize: 10 → Heading4 = 20 half-points (same as Normal, 10pt)", async () => {
    const sizes = await styleSizes({ fontSize: 10 })
    expect(sizes.get("Heading4")).toBe(20)
  })

  test("fontSize: 10 → CodeBlock = 20 half-points (same as Normal)", async () => {
    const sizes = await styleSizes({ fontSize: 10 })
    expect(sizes.get("CodeBlock")).toBe(20)
  })

  test("FooterText is always 20 half-points (fixed 10pt) regardless of fontSize", async () => {
    const sizesDefault = await styleSizes()
    expect(sizesDefault.get("FooterText")).toBe(20)
    const sizesCustom = await styleSizes({ fontSize: 10 })
    expect(sizesCustom.get("FooterText")).toBe(20)
  })

  test("CodeBlockSpacing style is not present", async () => {
    const sizes = await styleSizes()
    expect(sizes.has("CodeBlockSpacing")).toBe(false)
  })

  test("ListSpacing style is not present", async () => {
    const sizes = await styleSizes()
    expect(sizes.has("ListSpacing")).toBe(false)
  })
})

describe("style invariants", () => {
  let stylesXml: string

  async function getStylesXml() {
    if (!stylesXml) {
      const { zip } = await buildDocx("# H1\n## H2\n### H3\n\nParagraph with `code`.")
      stylesXml = await zip.file("word/styles.xml")!.async("string")
    }
    return stylesXml
  }

  function styleChunk(xml: string, styleId: string) {
    return xml.split(/<w:style\s/).find((c) => c.includes(`"${styleId}"`)) ?? ""
  }

  function halfPoints(xml: string, styleId: string): number | undefined {
    const chunk = styleChunk(xml, styleId)
    const m = chunk.match(/<w:sz w:val="(\d+)"/)
    return m ? parseInt(m[1]!, 10) : undefined
  }

  test("InlineCode has no explicit font size (inherits from parent paragraph)", async () => {
    const xml = await getStylesXml()
    expect(styleChunk(xml, "InlineCode")).not.toContain("w:sz")
  })

  test("InlineCode has no explicit color (inherits prose color)", async () => {
    const xml = await getStylesXml()
    expect(styleChunk(xml, "InlineCode")).not.toContain("<w:color")
  })

  test("Heading1 is larger than Heading2", async () => {
    const xml = await getStylesXml()
    expect(halfPoints(xml, "Heading1")!).toBeGreaterThan(halfPoints(xml, "Heading2")!)
  })

  test("Heading2 is larger than Heading3", async () => {
    const xml = await getStylesXml()
    expect(halfPoints(xml, "Heading2")!).toBeGreaterThan(halfPoints(xml, "Heading3")!)
  })

  test("Heading3 is larger than Normal", async () => {
    const xml = await getStylesXml()
    expect(halfPoints(xml, "Heading3")!).toBeGreaterThan(halfPoints(xml, "Normal")!)
  })

  test("CodeBlock matches Normal font size", async () => {
    const xml = await getStylesXml()
    expect(halfPoints(xml, "CodeBlock")).toBe(halfPoints(xml, "Normal"))
  })

  test("FooterText is always 20 half-points regardless of Normal size", async () => {
    const xml = await getStylesXml()
    expect(halfPoints(xml, "FooterText")).toBe(20)
  })
})

describe("externalStylesXml", () => {
  test("when provided, built-in Aptos font style is not present", async () => {
    const externalStylesXml = [
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`,
      `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">`,
      `  <w:style w:type="paragraph" w:styleId="Normal">`,
      `    <w:name w:val="Normal"/>`,
      `    <w:rPr><w:sz w:val="24"/></w:rPr>`,
      `  </w:style>`,
      `</w:styles>`,
    ].join("")
    const { zip } = await buildDocx("# Hello", { externalStylesXml })
    const stylesXml = await zip.file("word/styles.xml")!.async("string")
    expect(stylesXml).not.toContain("Aptos")
  })
})

describe("header", () => {
  test("is omitted when headerLabel is not provided", async () => {
    const { zip } = await buildDocx("# Hello")
    const relsXml = await zip.file("word/_rels/document.xml.rels")!.async("string")
    expect(relsXml).not.toContain("header")
  })

  test("default header contains the label text", async () => {
    const { zip, xml } = await buildDocx("# Hello", { headerLabel: "My Header" })
    const files = await headerFiles(zip)
    expect(files["default"]).toBeDefined()
    const header = await xml(files["default"]!)
    expect(header).toContain("My Header")
  })

  test("default header is left-aligned", async () => {
    const { zip, xml } = await buildDocx("# Hello", { headerLabel: "Left" })
    const files = await headerFiles(zip)
    const header = await xml(files["default"]!)
    expect(header).toContain('w:val="left"')
  })

  test("first-page header is empty", async () => {
    const { zip, xml } = await buildDocx("# Hello", { headerLabel: "My Header" })
    const files = await headerFiles(zip)
    expect(files["first"]).toBeDefined()
    const firstHeader = await xml(files["first"]!)
    expect(firstHeader).not.toContain("<w:r>")
    expect(firstHeader).not.toContain("<w:t")
  })

  test("titlePage is set when headerLabel is provided", async () => {
    const { zip } = await buildDocx("# Hello", { headerLabel: "My Header" })
    const docXml = await zip.file("word/document.xml")!.async("string")
    expect(docXml).toContain("<w:titlePg/>")
  })

  test("titlePage is not set when headerLabel is omitted", async () => {
    const { zip } = await buildDocx("# Hello")
    const docXml = await zip.file("word/document.xml")!.async("string")
    expect(docXml).not.toContain("<w:titlePg/>")
  })

  test("frontmatter title is used as header when no headerLabel option is passed", async () => {
    const { zip, xml } = await buildDocx("---\ntitle: Frontmatter Title\n---\n# Hello")
    const files = await headerFiles(zip)
    expect(files["default"]).toBeDefined()
    const header = await xml(files["default"]!)
    expect(header).toContain("Frontmatter Title")
  })

  test("explicit headerLabel overrides frontmatter title", async () => {
    const { zip, xml } = await buildDocx("---\ntitle: FM Title\n---\n# Hello", {
      headerLabel: "CLI Header",
    })
    const files = await headerFiles(zip)
    const header = await xml(files["default"]!)
    expect(header).toContain("CLI Header")
    expect(header).not.toContain("FM Title")
  })
})

describe("footnotes", () => {
  const md = [
    "Some text with a citation.[^1] Another sentence.[^2]",
    "",
    "[^1] [Some Link](https://example.com)",
    "",
    "[^2] Plain text footnote",
  ].join("\n")

  test("footnote ref in prose renders as FootnoteRef character style", async () => {
    const body = await bodyXml(md)
    expect(body).toContain('w:val="FootnoteRef"')
  })

  test("footnote ref renders the label number as text", async () => {
    const body = await bodyXml(md)
    expect(body).toContain(">1<")
    expect(body).toContain(">2<")
  })

  test("footnote definition uses FootnoteText paragraph style", async () => {
    const body = await bodyXml(md)
    expect(body).toContain('w:val="FootnoteText"')
  })

  test("footnote definition label renders as FootnoteRef character style", async () => {
    const body = await bodyXml(md)
    const footnoteChunk = body.slice(body.indexOf('w:val="FootnoteText"'))
    expect(footnoteChunk).toContain('w:val="FootnoteRef"')
  })

  test("footnote definition with hyperlink renders as hyperlink", async () => {
    const body = await bodyXml(md)
    expect(body).toContain("<w:hyperlink")
    expect(body).toContain("Some Link")
  })

  test("footnote definition plain text renders the text", async () => {
    const body = await bodyXml(md)
    expect(body).toContain("Plain text footnote")
  })

  test("footnote ref run uses FootnoteRef style", async () => {
    const body = await bodyXml("text[^1]\n\n[^1] note")
    // Both the inline ref and definition label should carry the built-in character style
    const refs = body.match(/w:val="FootnoteRef"/g) ?? []
    expect(refs.length).toBeGreaterThanOrEqual(2)
  })
})
