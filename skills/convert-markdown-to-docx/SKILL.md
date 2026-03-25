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
