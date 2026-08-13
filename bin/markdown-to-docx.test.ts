import { describe, expect, test } from "bun:test"
import { mkdtempSync, writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import JSZip from "jszip"
import pkg from "../package.json" with { type: "json" }

const CLI = join(import.meta.dir, "markdown-to-docx.ts")

function run(...args: string[]) {
  return Bun.spawnSync(["bun", CLI, ...args], { stderr: "pipe" })
}

/** Converts a one-line document with the given flags and returns its document.xml */
async function convert(...args: string[]): Promise<string> {
  const dir = mkdtempSync(join(tmpdir(), "cli-"))
  const mdPath = join(dir, "doc.md")
  const docxPath = join(dir, "doc.docx")
  writeFileSync(mdPath, "# Hello")

  const result = run(mdPath, docxPath, ...args)
  expect(result.stderr.toString()).toBe("")
  expect(result.exitCode).toBe(0)

  const zip = await JSZip.loadAsync(await Bun.file(docxPath).arrayBuffer())
  return zip.file("word/document.xml")!.async("string")
}

describe("CLI --version / -v", () => {
  test("--version prints version and exits 0", () => {
    const result = run("--version")
    expect(result.exitCode).toBe(0)
    expect(result.stdout.toString().trim()).toBe(pkg.version)
  })

  test("-v prints version and exits 0", () => {
    const result = run("-v")
    expect(result.exitCode).toBe(0)
    expect(result.stdout.toString().trim()).toBe(pkg.version)
  })
})

describe("CLI --help / -h", () => {
  test("--help prints usage and exits 0", () => {
    const result = run("--help")
    expect(result.exitCode).toBe(0)
    expect(result.stdout.toString()).toContain("Usage:")
  })

  test("-h prints usage and exits 0", () => {
    const result = run("-h")
    expect(result.exitCode).toBe(0)
    expect(result.stdout.toString()).toContain("Usage:")
  })

  test("no args prints usage and exits 1", () => {
    const result = run()
    expect(result.exitCode).toBe(1)
    expect(result.stdout.toString()).toContain("Usage:")
  })

  test("help documents --size", () => {
    const result = run("--help")
    expect(result.stdout.toString()).toContain("--size")
    expect(result.stdout.toString()).toContain("legal")
  })
})

describe("CLI --size", () => {
  test("a named size sets the page dimensions", async () => {
    const doc = await convert("--size", "legal")
    expect(doc).toContain('w:w="12240"')
    expect(doc).toContain('w:h="20160"')
  })

  test("a name is matched case-insensitively", async () => {
    const doc = await convert("--size", "Legal")
    expect(doc).toContain('w:h="20160"')
  })

  test("a custom size in inches sets the page dimensions", async () => {
    const doc = await convert("--size", "9x12")
    expect(doc).toContain('w:w="12960"')
    expect(doc).toContain('w:h="17280"')
  })

  test("a custom size in millimetres sets the page dimensions", async () => {
    const doc = await convert("--size", "210x297mm")
    expect(doc).toContain('w:w="11906"')
    expect(doc).toContain('w:h="16838"')
  })

  test("omitting --size leaves the A4 default", async () => {
    const doc = await convert()
    expect(doc).toContain('w:w="11906"')
    expect(doc).toContain('w:h="16838"')
  })

  test("an unknown size fails with a message listing the known names", () => {
    const result = run("doc.md", "--size", "quarto")
    expect(result.exitCode).toBe(1)
    expect(result.stderr.toString()).toContain("Unknown page size: quarto")
    expect(result.stderr.toString()).toContain("letter")
  })

  test("an out-of-range size fails with a message", () => {
    const result = run("doc.md", "--size", "40x50")
    expect(result.exitCode).toBe(1)
    expect(result.stderr.toString()).toContain("width must be between")
  })

  test("a missing size argument fails", () => {
    const result = run("doc.md", "--size")
    expect(result.exitCode).toBe(1)
    expect(result.stderr.toString()).toContain("--size requires a size argument")
  })

  test("a size followed by another flag fails rather than consuming it", () => {
    const result = run("doc.md", "--size", "--page-numbers")
    expect(result.exitCode).toBe(1)
    expect(result.stderr.toString()).toContain("--size requires a size argument")
  })
})
