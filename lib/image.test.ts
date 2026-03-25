import { describe, expect, test } from "bun:test"
import { writeFileSync } from "node:fs"
import { tmpdir } from "node:os"
import { join } from "node:path"
import { loadImage } from "./image"

// Minimal valid 1×1 RGB PNG (no external deps needed)
const ONE_PX_PNG = Buffer.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53,
  0xde, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0xd7, 0x63, 0xf8, 0xcf, 0xc0, 0x00,
  0x00, 0x00, 0x02, 0x00, 0x01, 0xe2, 0x21, 0xbc, 0x33, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e,
  0x44, 0xae, 0x42, 0x60, 0x82,
])

function writeTmpPng(): string {
  const path = join(tmpdir(), `test-${Date.now()}.png`)
  writeFileSync(path, ONE_PX_PNG)
  return path
}

describe("loadImage (local)", () => {
  test("returns ImageResult for a valid local PNG", async () => {
    const imgPath = writeTmpPng()
    const result = await loadImage(imgPath, imgPath)
    expect(result).not.toBeNull()
    expect(result!.type).toBe("png")
    expect(result!.data).toBeInstanceOf(Buffer)
    expect(result!.width).toBeGreaterThan(0)
    expect(result!.height).toBeGreaterThan(0)
  })

  test("returns null for a nonexistent file", async () => {
    const result = await loadImage("./does-not-exist.png", "/tmp/doc.md")
    expect(result).toBeNull()
  })

  test("resolves relative paths against the markdown file's directory", async () => {
    const imgPath = writeTmpPng()
    // Pass the image filename as a relative href, markdown is in the same dir
    const dir = imgPath.replace(/\/[^/]+$/, "")
    const filename = imgPath.split("/").pop()!
    const result = await loadImage(`./${filename}`, join(dir, "doc.md"))
    expect(result).not.toBeNull()
  })
})
