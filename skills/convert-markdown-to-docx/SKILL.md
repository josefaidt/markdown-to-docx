---
name: convert-markdown-to-docx
description: Converts Markdown files to Word (.docx) documents using the md-to-docx CLI. Use when the user wants to convert, export, or transform a Markdown file into a Word document or .docx format.
---

# Convert Markdown to DOCX

## Install

```sh
curl -fsSL https://github.com/josefaidt/markdown-to-docx/releases/latest/download/install.sh | sh
```

## Quick start

```bash
markdown-to-docx <input.md> [output.docx]
```

Output defaults to the same name and directory as the input with a `.docx` extension.

## Options

| Flag                | Description                                                    |
| ------------------- | -------------------------------------------------------------- |
| `--template <path>` | Load styles from a `.dotx` template file                       |
| `--header <text>`   | Left-aligned header text (skipped on first page)               |
| `--footer <text>`   | Bottom-left footer text (page number always on the right)      |
| `--font-size <n>`   | Base font size in pt; all styles scale from this (default: 12) |
| `--line-numbers`    | Add line numbers to code blocks                                |

## Examples

**Basic conversion:**

```bash
markdown-to-docx docs/report.md
# → docs/report.docx
```

**With custom output path:**

```bash
markdown-to-docx docs/report.md output/report.docx
```

**With header, footer, and template:**

```bash
markdown-to-docx report.md \
  --template my-styles.dotx \
  --header "Project Report" \
  --footer "Confidential" \
  --font-size 11
```

## Frontmatter support

If the Markdown file has YAML frontmatter, `title` is used as the header label when `--header` is not provided:

```markdown
---
title: My Document
---

# Introduction

...
```

## Normalizing links and images before conversion

`markdown-to-docx` cannot resolve relative paths at conversion time. Before converting, rewrite any relative links and image paths to absolute URLs or absolute file paths.

### Links

1. Scan the markdown body for relative links — e.g. `](./other-doc.md)`, `](../assets/fig.png)`
2. Resolve each to an absolute URL or absolute file path
3. If a target cannot be resolved, degrade gracefully: remove the href and render the link text as bold (`**link text**`), and warn the user

### Images

Apply the same rule to image references: rewrite `![alt](./img.png)` to an absolute path or URL. If the image cannot be resolved, remove it and warn the user.

### Workflow

Write the rewritten content to a temp file (e.g. `/tmp/converting-<original-name>.md`) and pass that to `markdown-to-docx`. **Never modify the original file.** Always regenerate the temp file from the source before each conversion run — do not reuse a previously created temp file.
