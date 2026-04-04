---
name: authoring-typescript
description: Best practices for authoring TypeScript in this monorepo. Apply when writing or reviewing TypeScript files.
user-invocable: false
---

# Authoring TypeScript

## Configuration

- Base configuration in `tsconfig.base.json` — strict, `moduleResolution: "bundler"`, `verbatimModuleSyntax`
- TypeScript 6: `"types": []` is set explicitly — add type packages per-package as needed (e.g. `"types": ["bun"]`)
- Raw TypeScript executed directly by Bun — no build/transpile step needed for development

## Import Style

Always use explicit `import type` for type-only imports. No blank lines between import statements.

Import order:

1. `import type` statements
2. `node:` built-in imports
3. Third-party package imports (alphabetical)
4. Relative imports (workspace-internal)

```typescript
import type { Foo } from "./types"
import { join } from "node:path"
import { something } from "some-package"
import { bar } from "./bar"
```

## Immutability

Use `as const` for any value that should be treated as readonly — literals, arrays, and objects that serve as fixed data or lookup tables. This narrows types to their literal values and prevents accidental mutation.

```typescript
// prefer
const DIRECTIONS = ["north", "south", "east", "west"] as const
const CONFIG = { retries: 3, timeout: 5000 } as const

// avoid
const DIRECTIONS = ["north", "south", "east", "west"]
const CONFIG = { retries: 3, timeout: 5000 }
```

Derive types from `as const` values rather than duplicating them:

```typescript
const STATUS = { active: "active", inactive: "inactive" } as const
type Status = (typeof STATUS)[keyof typeof STATUS]
```

## Directory Conventions

- The package entrypoint is `<package-name>.ts` at the package root (e.g. `template.ts` for `@workspace/template`)
- **NEVER use barrel files** — import directly from specific files
- **NEVER include file extensions** in import statements

## Type Usage

- **NEVER** use `any`. if anything use `unknown` if the type is truly unknown.
