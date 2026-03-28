import { describe, expect, test } from "bun:test"
import { join } from "node:path"
import pkg from "../package.json" with { type: "json" }

const CLI = join(import.meta.dir, "markdown-to-docx.ts")

function run(...args: string[]) {
  return Bun.spawnSync(["bun", CLI, ...args], { stderr: "pipe" })
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
})
