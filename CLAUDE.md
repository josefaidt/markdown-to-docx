# CLAUDE.md

This file provides guidance to coding agents (Claude Code, Cursor, Copilot, etc.) when working with code in this repository.

## Project Structure

This is a standalone Bun package that converts Markdown files to `.docx` Word documents.

- `bin/markdown-to-docx.ts` — CLI entrypoint
- `lib/markdown-to-docx.ts` — core conversion library (exported)
- `scripts/generate-template.ts` — generates `scripts/template.dotx` from the built-in styles

## Package Manager

This project uses **Bun** as the package manager and runtime.

## Common Commands

```bash
bun run build              # Compile a native binary to build/markdown-to-docx
bun run generate-template  # Generate scripts/template.dotx from built-in styles

# Run the CLI directly (no build needed)
bun run bin/markdown-to-docx.ts <input.md> [output.docx] [options]
```

## CLI Options

```
--template <path>  Load styles from a .dotx template file
--header <text>    Centred header text (skipped on first page)
--footer <text>    Bottom-left footer text (page number on the right)
--line-numbers     Add line numbers to code blocks
--help             Show usage
```

## Tooling

### TypeScript

- Configuration in `tsconfig.json` — strict, `moduleResolution: "bundler"`, `verbatimModuleSyntax`
- Raw TypeScript executed directly by Bun — no build/transpile step needed
- `bun-types` provides Bun global types (`Bun.file`, `Bun.write`, `Bun.argv`, etc.)

## TypeScript Development Approach

**This project uses raw TypeScript without build/transpilation steps.**

### Directory Structure

- `lib/` — library code consumed by the CLI and exported to package consumers
- `bin/` — CLI entrypoint (not exported)
- `scripts/` — dev-only scripts (not exported)
- **NEVER use barrel files** — import directly from specific files

### Import Style

- **NEVER include file extensions (.ts, .js) in import statements**
- Correct: `import { foo } from "./lib/bar"`
- Incorrect: `import { foo } from "./lib/bar.ts"`

#### Import Order

Always use explicit `import type` for type-only imports, and place them first. No blank lines between import statements.

1. `import type` statements
2. `node:` built-in imports
3. Third-party package imports (alphabetical by package name)
4. Relative imports

```typescript
import type { ConvertOptions } from "./lib/markdown-to-docx"
import { resolve } from "node:path"
import { Document, Packer } from "docx"
import { marked } from "marked"
import { convertMarkdownToDocx } from "./lib/markdown-to-docx"
```

## Architecture Notes

- ES modules (`"type": "module"` in package.json)
- Bun APIs used throughout: `Bun.file()`, `Bun.write()`, `Bun.argv`
- `node:path` used for path manipulation (Bun has no native path API)
- `docx` library builds the Word document object model
- `jszip` (devDep) used in `bin/` and `scripts/` to read/write `.dotx` ZIP internals
- `.dotx` template generation works by patching `[Content_Types].xml` inside the ZIP to flip the content type from `.docx` → `.dotx`
