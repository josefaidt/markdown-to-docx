---
name: preview
description: Generate a preview .docx from a representative test fixture and open it. Use when the user wants to visually inspect rendering changes.
allowed-tools:
  - Bash(osascript *)
  - Bash(bun run bin/markdown-to-docx.ts *)
  - Bash(open *)
---

# Preview

Generate a `.docx` from the test fixture at `scripts/preview.md` and open it so the user can visually inspect the output.

```bash
osascript -e '
  tell application "Microsoft Word"
    if it is running then
      if (count of documents) > 0 then
        set docPath to POSIX file "/tmp/preview.docx"
        repeat with d in documents
          if (get full name of d) is (docPath as string) then
            close d saving no
          end if
        end repeat
      end if
    end if
  end tell
' && \
bun run bin/markdown-to-docx.ts .claude/skills/preview/assets/preview.md /tmp/preview.docx && \
open /tmp/preview.docx
```
