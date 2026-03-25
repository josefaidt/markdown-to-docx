export type Frontmatter = Record<string, unknown>

export interface ParsedContent {
  /** Raw markdown body with the frontmatter block removed */
  body: string
  /** Parsed YAML frontmatter fields, or empty object if none present */
  data: Frontmatter
}

/**
 * Splits a markdown string into its YAML frontmatter and body.
 * Uses Bun.YAML.parse() for spec-compliant YAML 1.2 parsing.
 * Returns an empty `data` object if no frontmatter is present.
 */
export function parseFrontmatter(content: string): ParsedContent {
  if (!content.startsWith("---\n")) {
    return { body: content, data: {} }
  }

  const end = content.indexOf("\n---\n", 4)
  if (end === -1) {
    return { body: content, data: {} }
  }

  const yaml = content.slice(4, end)
  const body = content.slice(end + 5)

  let data: Frontmatter = {}
  try {
    const parsed = Bun.YAML.parse(yaml)
    if (parsed !== null && typeof parsed === "object" && !Array.isArray(parsed)) {
      data = parsed as Frontmatter
    }
  } catch {
    // Malformed YAML — treat as no frontmatter
  }

  return { body, data }
}
