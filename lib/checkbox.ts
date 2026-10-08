import { BuilderElement, CheckBox, TextRun } from "docx"

/**
 * How GFM task-list items (`- [ ]` / `- [x]`) render:
 * - `static` — the box is a plain text glyph, so the checked state from the
 *   Markdown is fixed and looks the same in every renderer
 * - `interactive` — the box is a Word checkbox content control that starts in
 *   the Markdown's state and can be toggled in Word
 */
export const CHECKBOX_MODES = ["static", "interactive"] as const

export type CheckboxMode = (typeof CHECKBOX_MODES)[number]

export const DEFAULT_CHECKBOX_MODE: CheckboxMode = "static"

/**
 * The font Microsoft suggests for a checkmark checkbox. Word's default, MS
 * Gothic, draws ☑'s check poking out of the box; both states use the same font
 * so the boxes match.
 */
const CHECKBOX_FONT = "Segoe UI Symbol"

/** Checkbox content control symbols, as hex code points */
const CHECKBOX_SYMBOLS = {
  checked: "2611", // ☑
  unchecked: "2610", // ☐
} as const

/**
 * Resolve a checkbox mode from a CLI argument or frontmatter value
 * (case-insensitive).
 *
 * @throws {Error} If the value is not one of {@link CHECKBOX_MODES}.
 */
export function parseCheckboxMode(value: unknown): CheckboxMode {
  const normalized = typeof value === "string" ? value.trim().toLowerCase() : value
  const mode = CHECKBOX_MODES.find((m) => m === normalized)
  if (!mode) {
    throw new Error(
      `Unknown checkbox mode: ${String(value)}\nExpected one of: ${CHECKBOX_MODES.join(", ")}`,
    )
  }
  return mode
}

/** The box glyph as a text run */
function glyphRun(checked: boolean): TextRun {
  const symbol = checked ? CHECKBOX_SYMBOLS.checked : CHECKBOX_SYMBOLS.unchecked
  return new TextRun({
    text: String.fromCodePoint(parseInt(symbol, 16)),
    font: CHECKBOX_FONT,
  })
}

/** A `w14:` element carrying `w14:val` (and optionally `w14:font`) attributes */
function w14Element(name: string, val: string, font?: string): BuilderElement {
  return new BuilderElement<{ val: string; font?: string }>({
    name: `w14:${name}`,
    attributes: {
      val: { key: "w14:val", value: val },
      ...(font ? { font: { key: "w14:font", value: font } } : {}),
    },
  })
}

/** Content control ids only need to be unique within a document */
let nextContentControlId = 1

/**
 * A Word checkbox content control written the way Word writes its own:
 * - a unique `w:id` before `w14:checkbox`, which docx's `CheckBox` omits but
 *   Word always writes
 * - the box as a text run rather than a `w:sym` element, which LibreOffice
 *   mis-renders: every unchecked `w:sym` box after the first shows as checked
 *
 * It extends `CheckBox` only because paragraphs accept that type as a child;
 * all of its XML is replaced.
 */
class TextCheckBox extends CheckBox {
  constructor(checked: boolean) {
    super()
    this.root.splice(0)
    this.root.push(
      new BuilderElement({
        name: "w:sdtPr",
        children: [
          new BuilderElement<{ val: number }>({
            name: "w:id",
            attributes: { val: { key: "w:val", value: nextContentControlId++ } },
          }),
          new BuilderElement({
            name: "w14:checkbox",
            children: [
              w14Element("checked", checked ? "1" : "0"),
              w14Element("checkedState", CHECKBOX_SYMBOLS.checked, CHECKBOX_FONT),
              w14Element("uncheckedState", CHECKBOX_SYMBOLS.unchecked, CHECKBOX_FONT),
            ],
          }),
        ],
      }),
      new BuilderElement({ name: "w:sdtContent", children: [glyphRun(checked)] }),
    )
  }
}

/** The box that leads a task-list item, as a glyph or a content control per `mode` */
export function checkboxRun(checked: boolean, mode: CheckboxMode): CheckBox | TextRun {
  return mode === "interactive" ? new TextCheckBox(checked) : glyphRun(checked)
}
