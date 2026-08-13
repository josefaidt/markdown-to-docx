/** Page dimensions in twips (1/1440 of an inch), the unit OOXML uses for `w:pgSz` */
export interface PageSize {
  width: number
  height: number
}

const TWIPS_PER_INCH = 1440
const TWIPS_PER_UNIT: Record<string, number> = {
  in: TWIPS_PER_INCH,
  '"': TWIPS_PER_INCH,
  mm: TWIPS_PER_INCH / 25.4,
  cm: TWIPS_PER_INCH / 2.54,
  pt: TWIPS_PER_INCH / 72,
}

/** Word rejects pages outside this range, so reject them here with a clearer message */
const MIN_DIMENSION_TWIP = TWIPS_PER_INCH / 10
const MAX_DIMENSION_TWIP = TWIPS_PER_INCH * 22

/** Named paper sizes in twips, portrait orientation */
const NAMED_SIZES: Record<string, PageSize> = {
  a3: { width: 16838, height: 23811 },
  a4: { width: 11906, height: 16838 },
  a5: { width: 8391, height: 11906 },
  b5: { width: 9978, height: 14173 },
  executive: { width: 10440, height: 15120 },
  legal: { width: 12240, height: 20160 },
  letter: { width: 12240, height: 15840 },
  tabloid: { width: 15840, height: 24480 },
}

/** Sorted names, for help text and error messages */
export const PAGE_SIZE_NAMES = Object.keys(NAMED_SIZES).sort()

/** Matches `9x12`, `9in x 12in`, `210x297mm`, `8.5"x11"` — the unit is optional on either side */
const CUSTOM_SIZE_PATTERN =
  /^(\d+(?:\.\d+)?)\s*(in|"|mm|cm|pt)?\s*[x×]\s*(\d+(?:\.\d+)?)\s*(in|"|mm|cm|pt)?$/i

export const DEFAULT_PAGE_SIZE: PageSize = NAMED_SIZES["a4"]!

/**
 * Turn a page on its side, so its longer edge runs horizontally. Idempotent:
 * a size that is already wider than it is tall comes back unchanged.
 */
export function toLandscape(size: PageSize): PageSize {
  return {
    width: Math.max(size.width, size.height),
    height: Math.min(size.width, size.height),
  }
}

function toTwips(value: string, unit: string | undefined, label: string, spec: string): number {
  const perUnit = TWIPS_PER_UNIT[(unit ?? "in").toLowerCase()]!
  const twips = Math.round(Number(value) * perUnit)
  if (twips < MIN_DIMENSION_TWIP || twips > MAX_DIMENSION_TWIP) {
    throw new Error(
      `Invalid page size "${spec}": ${label} must be between 0.1in (2.5mm) and 22in (558mm)`,
    )
  }
  return twips
}

/**
 * Resolve a page size spec to its dimensions in twips. Accepts a named size
 * (case-insensitive, e.g. `Legal`) or a custom `<width>x<height>` spec with an
 * optional unit — `in` (the default), `"`, `mm`, `cm`, or `pt`. A unit given on
 * one side applies to the other when that side omits it, so `210x297mm` works.
 *
 * @throws {Error} If the spec is neither a known name nor a valid custom size.
 */
export function parsePageSize(spec: string): PageSize {
  const trimmed = spec.trim()

  const named = NAMED_SIZES[trimmed.toLowerCase()]
  if (named) return named

  const match = CUSTOM_SIZE_PATTERN.exec(trimmed)
  if (!match) {
    throw new Error(
      `Unknown page size: ${spec}\n` +
        `Expected a named size (${PAGE_SIZE_NAMES.join(", ")}) ` +
        `or a custom size such as 9x12, 9x12in, or 210x297mm`,
    )
  }

  const [, rawWidth, widthUnit, rawHeight, heightUnit] = match
  return {
    width: toTwips(rawWidth!, widthUnit ?? heightUnit, "width", trimmed),
    height: toTwips(rawHeight!, heightUnit ?? widthUnit, "height", trimmed),
  }
}
