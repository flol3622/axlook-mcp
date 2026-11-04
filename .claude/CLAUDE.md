# Working on axlook-mcp

- **Runtime.** Bun for development (`bun test`, `bun run typecheck`, `bun run src/server.ts`). The shipped bundle `dist/axlook-mcp.js` must stay Node-compatible: no `Bun.*` or `from "bun"` in `src/`; use Node built-ins.
- **Layout.** `helper/OutlookAX.swift` holds generic Accessibility primitives only. All Outlook-specific selectors and flows live in `src/outlook/*.ts`.
- **Shipping.** `dist/` is committed because the Claude Code plugin runs it as is. After changing `src/` or `helper/`, run `bun run build` and commit `dist/` together with the source. `bin/` is a local dev build and must not be committed: a plugin's top-level `bin/` is added to PATH.
- **Checks before committing:** `bun test`, `bun run typecheck`, `claude plugin validate . --strict`.
- **Live testing** (`scripts/e2e.ts`) drives the real Outlook. It needs the `E2E_*` variables described in its header, and touches only the test folder, drafts and event it names.
- **Privacy.** Never commit real mailbox data: no addresses, subjects or names in tests, fixtures, docs or defaults.
