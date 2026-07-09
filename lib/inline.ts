import type { ImageResult } from "./image"
import type { Tokens } from "marked"
import { ExternalHyperlink, ImageRun, InternalHyperlink, ShadingType, TextRun } from "docx"

export type InlineChild = TextRun | ExternalHyperlink | InternalHyperlink | ImageRun

interface InlineCtx {
  bold?: boolean
  italics?: boolean
  useTemplate?: boolean
}

const NAMED_ENTITIES: Record<string, string> = {
  nbsp: "\u00a0",
  amp: "&",
  lt: "<",
  gt: ">",
  quot: '"',
  apos: "'",
  copy: "\u00a9",
  reg: "\u00ae",
  trade: "\u2122",
  mdash: "\u2014",
  ndash: "\u2013",
  hellip: "\u2026",
  laquo: "\u00ab",
  raquo: "\u00bb",
  lsquo: "\u2018",
  rsquo: "\u2019",
  ldquo: "\u201c",
  rdquo: "\u201d",
  bull: "\u2022",
}

function decodeHtmlEntities(text: string): string {
  return text.replace(/&(#x?[0-9a-fA-F]+|\w+);/g, (match, entity: string) => {
    if (entity.startsWith("#x") || entity.startsWith("#X")) {
      const code = parseInt(entity.slice(2), 16)
      return isNaN(code) ? match : String.fromCodePoint(code)
    }
    if (entity.startsWith("#")) {
      const code = parseInt(entity.slice(1), 10)
      return isNaN(code) ? match : String.fromCodePoint(code)
    }
    return NAMED_ENTITIES[entity] ?? match
  })
}

function applyTypography(text: string): string {
  return decodeHtmlEntities(text).replace(/--/g, "\u2014")
}

// Word silently truncates bookmark names to 40 characters, so we enforce the
// same limit here so that BookmarkStart names and InternalHyperlink anchors
// always agree.
const WORD_BOOKMARK_MAX_LENGTH = 40

export function headingSlug(text: string): string {
  const slug = text
    .toLowerCase()
    .replace(/[^\w\s-]/g, "")
    .replace(/\s+/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
  return slug.slice(0, WORD_BOOKMARK_MAX_LENGTH).replace(/-+$/g, "")
}

export function inlineTokensToRuns(
  tokens: Tokens.Generic[],
  ctx: InlineCtx = {},
  images: ReadonlyMap<string, ImageResult> = new Map(),
  useTemplate = false,
): InlineChild[] {
  const runs: InlineChild[] = []
  for (const token of tokens) {
    // Use bracket notation because Tokens.Generic has an index signature
    const rawText = token["text"] as string | undefined
    const text = rawText !== undefined ? applyTypography(rawText) : undefined
    const href = token["href"] as string | undefined
    const children = token["tokens"] as Tokens.Generic[] | undefined
    switch (token.type) {
      case "text":
        if (children?.length) {
          runs.push(...inlineTokensToRuns(children, ctx, images, useTemplate))
        } else {
          runs.push(new TextRun({ text: text ?? "", ...ctx }))
        }
        break
      case "strong":
        runs.push(
          ...inlineTokensToRuns(children ?? [], { ...ctx, bold: true }, images, useTemplate),
        )
        break
      case "em":
        runs.push(
          ...inlineTokensToRuns(children ?? [], { ...ctx, italics: true }, images, useTemplate),
        )
        break
      case "codespan":
        runs.push(
          new TextRun({
            text: rawText ?? "",
            ...ctx,
            style: "InlineCode",
            ...(useTemplate
              ? {}
              : { shading: { type: ShadingType.CLEAR, color: "F2F2F2", fill: "F2F2F2" } }),
          }),
        )
        break
      case "link": {
        const linkText =
          children?.map((t) => applyTypography((t["text"] as string | undefined) ?? "")).join("") ??
          text ??
          ""
        const linkHref = href ?? ""
        if (linkHref.startsWith("#")) {
          runs.push(
            new InternalHyperlink({
              anchor: headingSlug(linkHref.slice(1)),
              children: [new TextRun({ text: linkText, style: "Hyperlink" })],
            }),
          )
        } else {
          runs.push(
            new ExternalHyperlink({
              link: linkHref,
              children: [new TextRun({ text: linkText, style: "Hyperlink" })],
            }),
          )
        }
        break
      }
      case "image": {
        const imgHref = href ?? ""
        const alt = (token["text"] as string | undefined) ?? ""
        const result = images.get(imgHref)
        if (result) {
          runs.push(
            new ImageRun({
              type: result.type,
              data: result.data,
              transformation: { width: result.width, height: result.height },
              altText: { name: alt, description: alt },
            }),
          )
        } else {
          runs.push(new TextRun({ text: `[Image: ${alt}]`, italics: true, color: "888888" }))
        }
        break
      }
      case "footnote_ref": {
        const label = token["label"] as string | undefined
        runs.push(new TextRun({ text: label ?? "", style: "FootnoteRef", ...ctx }))
        break
      }
      case "br":
        runs.push(new TextRun({ break: 1 }))
        break
      case "escape":
      default:
        if (text) runs.push(new TextRun({ text, ...ctx }))
        break
    }
  }
  return runs
}
