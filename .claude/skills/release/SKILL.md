---
name: release
description: Orchestrates a release of markdown-to-docx. Use when the user wants to cut a release, bump the version, or publish a new version.
---

# Release

Releases are driven by the `release` workflow (`dispatch`) on GitHub Actions, which handles everything: version bump in `package.json`, README stamping, git commit + tag, GitHub release creation, and binary builds. The preferred path is to trigger that workflow so nothing needs to run locally.

## Preferred: trigger via GitHub Actions

Ask the user which bump type to use if not provided: `patch`, `minor`, `major`, `prepatch`, `preminor`, `premajor`, `prerelease`.

```bash
gh workflow run release.yml --field bump=<bump>
```

Then follow the run:

```bash
gh run watch --exit-status
```

## Alternative: release locally

Use this when GitHub Actions is unavailable or the user explicitly wants to run locally.

### 1. Verify the working tree is clean

```bash
git status --short
```

If there are uncommitted changes, stop and ask the user to commit or stash them first.

### 2. Confirm CI passes

```bash
bun run lint && bun run fmt:check && bun test
```

### 3. Bump version, commit, and tag

`npm version` handles the `package.json` update, commit, and tag in one step:

```bash
git config user.name "your name"
git config user.email "your email"
TAG=$(bun x --bun npm version <bump> --message "release: %s")
```

`<bump>` is one of: `patch`, `minor`, `major`, `prepatch`, `preminor`, `premajor`, `prerelease`.

### 4. Stamp the README and amend

```bash
USAGE=$(bun run bin/markdown-to-docx.ts --help 2>&1 || true)
perl -i -0pe "s|<!-- BEGIN USAGE -->.*<!-- END USAGE -->|<!-- BEGIN USAGE -->\n\`\`\`\n${USAGE}\n\`\`\`\n<!-- END USAGE -->|s" README.md
sed -i "" "s|\"ref\": \".*\"|\"ref\": \"${TAG}\"|g" README.md
sed -i "" "s|<!-- LATEST_TAG: .* -->|<!-- LATEST_TAG: ${TAG} -->|" README.md
git add README.md
git commit --amend --no-edit
git tag -f "$TAG"
```

### 5. Push main and the tag

```bash
git push origin HEAD:main
git push origin "$TAG"
```

### 6. Cut the GitHub release as a draft

```bash
gh release create "$TAG" \
  --title "$TAG" \
  --generate-notes \
  --draft \
  scripts/install.sh
```

### 7. Build and upload binaries, then publish

Build each platform target and upload to the draft release:

```bash
for TARGET in bun-linux-x64 bun-linux-arm64 bun-darwin-x64 bun-darwin-arm64; do
  ARTIFACT="markdown-to-docx-${TARGET#bun-}"
  bun build --compile --target=$TARGET --outfile=$ARTIFACT bin/markdown-to-docx.ts
  gh release upload "$TAG" "$ARTIFACT"
done
# Windows
bun build --compile --target=bun-windows-x64 --outfile=markdown-to-docx-windows-x64.exe bin/markdown-to-docx.ts
gh release upload "$TAG" markdown-to-docx-windows-x64.exe
```

Once all binaries are uploaded, promote the draft to published:

```bash
gh release edit "$TAG" --draft=false
```

## Notes

- **Draft → publish**: the release is created as a draft and only promoted once all platform binaries are successfully uploaded. If any build fails, the release stays as a draft and is never visible to users.

- **sed on macOS**: the `-i ""` flag is required (BSD sed). The workflow uses Linux, so it uses `-i` without the empty string.
- **Binaries**: never built or committed locally. The Actions matrix builds all five targets (`linux-x64`, `linux-arm64`, `darwin-x64`, `darwin-arm64`, `windows-x64`) and uploads them to the release.
- **Version format**: always `X.Y.Z` as input; the `v` prefix is added by the scripts.
