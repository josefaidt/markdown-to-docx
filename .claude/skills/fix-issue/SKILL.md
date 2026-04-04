---
name: fix-issue
description: Fetch a GitHub issue, create a worktree for it, and start working
disable-model-invocation: true
permissionMode: acceptEdits
---

Fix GitHub issue #$ARGUMENTS:

1. Run `gh issue view $ARGUMENTS` to read the issue
2. Use the issue number and title to name a worktree (slugified)
3. Enter the worktree and implement the fix following our coding standards
4. Write tests and create a commit referencing the issue
