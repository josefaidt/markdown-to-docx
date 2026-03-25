import { describe, expect, test } from "bun:test"
import { convertInchesToTwip } from "docx"
import { listIndent } from "./list"

// Constants mirrored from list.ts
const BULLET_INDENT = convertInchesToTwip(0.2) // 288
const TEXT_INDENT = convertInchesToTwip(0.45) // 648
const LEVEL_STEP = convertInchesToTwip(0.25) // 360

describe("listIndent", () => {
  test("level 0: left = TEXT_INDENT, hanging = TEXT_INDENT - BULLET_INDENT", () => {
    const { left, hanging } = listIndent(0)
    expect(left).toBe(TEXT_INDENT)
    expect(hanging).toBe(TEXT_INDENT - BULLET_INDENT)
  })

  test("level 1: left advances by one LEVEL_STEP", () => {
    const { left } = listIndent(1)
    expect(left).toBe(TEXT_INDENT + LEVEL_STEP)
  })

  test("level 2: left advances by two LEVEL_STEP", () => {
    const { left } = listIndent(2)
    expect(left).toBe(TEXT_INDENT + 2 * LEVEL_STEP)
  })

  test("hanging width is constant across levels (bullet offset stays fixed)", () => {
    const h0 = listIndent(0).hanging
    const h1 = listIndent(1).hanging
    const h2 = listIndent(2).hanging
    expect(h0).toBe(h1)
    expect(h1).toBe(h2)
  })
})
