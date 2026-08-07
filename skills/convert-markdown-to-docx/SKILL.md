---
name: convert-markdown-to-docx
description: Convert, export, or transform a Markdown file into a Word (.docx) document with the markdown-to-docx CLI.
---

# Convert Markdown to DOCX

## Quick start

```bash
markdown-to-docx <input.md> [output.docx]
```

Output defaults to the same name and directory as the input with a `.docx` extension. Pass a second argument to choose the output path.

For the full, current flag list run `markdown-to-docx --help`. Flags combine:

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

`markdown-to-docx` cannot resolve relative paths at conversion time — every relative link and image must be rewritten to an absolute URL or absolute file path first, or it breaks in the output.

Rewrite the source into a temp file, then convert that:

1. Copy the source to a fresh temp file — e.g. `/tmp/converting-<original-name>.md`. Regenerate it from the source every run; never reuse a prior temp file, and never modify the original.
2. Rewrite every relative link and image in the temp file to an absolute URL or absolute file path — `](./other-doc.md)`, `](../assets/fig.png)`, `![alt](./img.png)`.
3. For any target that cannot be resolved, degrade gracefully and warn the user: drop a link's href and render its text as bold (`**link text**`); remove an unresolvable image.
4. Pass the temp file to `markdown-to-docx`.

Done when no relative path remains in the temp file.
