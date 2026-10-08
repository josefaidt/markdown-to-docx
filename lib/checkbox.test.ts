import { describe, expect, test } from "bun:test"
import { CHECKBOX_MODES, DEFAULT_CHECKBOX_MODE, parseCheckboxMode } from "./checkbox"

describe("parseCheckboxMode", () => {
  test("accepts every mode", () => {
    for (const mode of CHECKBOX_MODES) expect(parseCheckboxMode(mode)).toBe(mode)
  })

  test("is case-insensitive and ignores surrounding whitespace", () => {
    expect(parseCheckboxMode(" Interactive ")).toBe("interactive")
    expect(parseCheckboxMode("STATIC")).toBe("static")
  })

  test("rejects an unknown mode with a message listing the known modes", () => {
    expect(() => parseCheckboxMode("clickable")).toThrow(
      "Unknown checkbox mode: clickable\nExpected one of: static, interactive",
    )
  })

  test("rejects a non-string value", () => {
    expect(() => parseCheckboxMode(true)).toThrow("Unknown checkbox mode: true")
  })

  test("defaults to static", () => {
    expect(DEFAULT_CHECKBOX_MODE).toBe("static")
  })
})
