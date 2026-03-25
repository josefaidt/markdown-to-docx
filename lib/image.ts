import { resolve, dirname } from "node:path"
import probe from "probe-image-size"

const MAX_IMAGE_WIDTH_EMU = 914400 * 6

export type SupportedImageType = "jpg" | "png" | "gif" | "bmp"
const SUPPORTED_TYPES = new Set<string>(["jpg", "jpeg", "png", "gif", "bmp"])

export interface ImageResult {
  data: Buffer
  type: SupportedImageType
  width: number
  height: number
}

function mimeToType(mime: string): SupportedImageType | null {
  if (mime.includes("jpeg") || mime.includes("jpg")) return "jpg"
  if (mime.includes("png")) return "png"
  if (mime.includes("gif")) return "gif"
  if (mime.includes("bmp")) return "bmp"
  return null
}

function extToType(url: string): SupportedImageType | null {
  const ext = url.split("?")[0]?.split(".").pop()?.toLowerCase() ?? ""
  if (!SUPPORTED_TYPES.has(ext)) return null
  return ext === "jpeg" ? "jpg" : (ext as SupportedImageType)
}

function scaleToEmu(px: number, scale: number) {
  return Math.round(((px * 914400) / 96) * scale)
}

async function fetchImageBuffer(url: string): Promise<ImageResult | null> {
  try {
    const res = await fetch(url, { headers: { "User-Agent": "md-to-docx/1.0" } })
    if (!res.ok) return null
    const data = Buffer.from(await res.arrayBuffer())
    const info = probe.sync(data)
    if (!info) return null
    const type = mimeToType(info.mime) ?? extToType(url)
    if (!type) return null
    const scale = Math.min(1, MAX_IMAGE_WIDTH_EMU / ((info.width * 914400) / 96))
    return {
      data,
      type,
      width: scaleToEmu(info.width, scale),
      height: scaleToEmu(info.height, scale),
    }
  } catch {
    return null
  }
}

async function loadLocalImage(filePath: string): Promise<ImageResult | null> {
  try {
    const data = Buffer.from(await Bun.file(filePath).arrayBuffer())
    const info = probe.sync(data)
    if (!info) return null
    const type = mimeToType(info.mime) ?? extToType(filePath)
    if (!type) return null
    const scale = Math.min(1, MAX_IMAGE_WIDTH_EMU / ((info.width * 914400) / 96))
    return {
      data,
      type,
      width: scaleToEmu(info.width, scale),
      height: scaleToEmu(info.height, scale),
    }
  } catch {
    return null
  }
}

export async function loadImage(href: string, markdownPath: string): Promise<ImageResult | null> {
  if (href.startsWith("http://") || href.startsWith("https://")) {
    return fetchImageBuffer(href)
  }
  return loadLocalImage(resolve(dirname(markdownPath), href))
}
