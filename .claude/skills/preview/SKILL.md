---
name: preview
description: Generate a preview .docx from a representative test fixture and open it. Use when the user wants to visually inspect rendering changes.
---

# Preview

Generate a `.docx` from the test fixture at `scripts/preview.md` and open it so the user can visually inspect the output.

```bash
bun run bin/markdown-to-docx.ts .claude/skills/preview/preview.md /tmp/preview.docx && open /tmp/preview.docx
```
