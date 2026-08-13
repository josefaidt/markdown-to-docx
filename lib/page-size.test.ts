import { describe, expect, test } from "bun:test"
import { DEFAULT_PAGE_SIZE, PAGE_SIZE_NAMES, parsePageSize, toLandscape } from "./page-size"

describe("parsePageSize named sizes", () => {
  test("resolves letter", () => {
    expect(parsePageSize("letter")).toEqual({ width: 12240, height: 15840 })
  })

  test("resolves legal", () => {
    expect(parsePageSize("legal")).toEqual({ width: 12240, height: 20160 })
  })

  test("resolves a4 to the default size", () => {
    expect(parsePageSize("a4")).toEqual(DEFAULT_PAGE_SIZE)
  })

  test("name lookup is case-insensitive", () => {
    expect(parsePageSize("Legal")).toEqual(parsePageSize("legal"))
    expect(parsePageSize("A4")).toEqual(parsePageSize("a4"))
  })

  test("surrounding whitespace is ignored", () => {
    expect(parsePageSize("  letter  ")).toEqual(parsePageSize("letter"))
  })

  test("every advertised name resolves", () => {
    for (const name of PAGE_SIZE_NAMES) {
      const size = parsePageSize(name)
      expect(size.width).toBeGreaterThan(0)
      expect(size.height).toBeGreaterThan(0)
    }
  })
})

describe("parsePageSize custom sizes", () => {
  test("defaults to inches when no unit is given", () => {
    expect(parsePageSize("9x12")).toEqual({ width: 12960, height: 17280 })
  })

  test("accepts an explicit in unit", () => {
    expect(parsePageSize("9inx12in")).toEqual({ width: 12960, height: 17280 })
  })

  test("accepts inch marks", () => {
    expect(parsePageSize('8.5"x11"')).toEqual(parsePageSize("letter"))
  })

  test("accepts millimetres", () => {
    // A4 is 210x297mm, so the conversion must land on the named size
    expect(parsePageSize("210x297mm")).toEqual(parsePageSize("a4"))
  })

  test("accepts centimetres", () => {
    expect(parsePageSize("21x29.7cm")).toEqual(parsePageSize("210x297mm"))
  })

  test("accepts points", () => {
    expect(parsePageSize("612x792pt")).toEqual(parsePageSize("letter"))
  })

  test("a unit on one side applies to the other", () => {
    expect(parsePageSize("210mmx297")).toEqual(parsePageSize("210x297mm"))
  })

  test("mixed units are honoured per dimension", () => {
    expect(parsePageSize("1inx25.4mm")).toEqual({ width: 1440, height: 1440 })
  })

  test("tolerates whitespace around the separator", () => {
    expect(parsePageSize("9 in x 12 in")).toEqual(parsePageSize("9x12"))
  })

  test("accepts an uppercase X separator", () => {
    expect(parsePageSize("9X12")).toEqual(parsePageSize("9x12"))
  })

  test("accepts the × separator", () => {
    expect(parsePageSize("9×12")).toEqual(parsePageSize("9x12"))
  })

  test("accepts fractional dimensions", () => {
    expect(parsePageSize("8.5x11")).toEqual(parsePageSize("letter"))
  })
})

describe("toLandscape", () => {
  test("swaps a portrait size", () => {
    expect(toLandscape(parsePageSize("a4"))).toEqual({ width: 16838, height: 11906 })
  })

  test("leaves a size that is already wider than tall alone", () => {
    expect(toLandscape(parsePageSize("11x8.5"))).toEqual(parsePageSize("11x8.5"))
  })

  test("is idempotent", () => {
    const once = toLandscape(parsePageSize("legal"))
    expect(toLandscape(once)).toEqual(once)
  })

  test("leaves a square size alone", () => {
    expect(toLandscape(parsePageSize("9x9"))).toEqual(parsePageSize("9x9"))
  })
})

describe("parsePageSize errors", () => {
  test("rejects an unknown name", () => {
    expect(() => parsePageSize("a9")).toThrow(/Unknown page size: a9/)
  })

  test("lists the known names in the error", () => {
    expect(() => parsePageSize("nope")).toThrow(/letter/)
  })

  test("rejects a missing dimension", () => {
    expect(() => parsePageSize("9x")).toThrow(/Unknown page size/)
  })

  test("rejects a non-numeric dimension", () => {
    expect(() => parsePageSize("axb")).toThrow(/Unknown page size/)
  })

  test("rejects an unsupported unit", () => {
    expect(() => parsePageSize("9emx12em")).toThrow(/Unknown page size/)
  })

  test("rejects a dimension above Word's 22in limit", () => {
    expect(() => parsePageSize("23x11")).toThrow(/width must be between/)
  })

  test("rejects a dimension below the minimum", () => {
    expect(() => parsePageSize("9x0.05")).toThrow(/height must be between/)
  })

  test("rejects a zero dimension", () => {
    expect(() => parsePageSize("0x11")).toThrow(/width must be between/)
  })
})
