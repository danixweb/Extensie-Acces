# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Repo layout

This repo has two mostly-independent parts:

- **`access-explorer/`** — a VS Code extension (TypeScript) that lets VS Code open and edit Microsoft Access `.accdb` files directly, via COM automation. This is the only buildable/testable code in the repo.
- **`skills/`** — Claude Code skills (Romanian-language SOPs) for working with the VBA code *inside* Access projects opened through this extension (adding error handling, analyzing routines, fixing logged runtime errors). These are documentation/instructions, not application code.

Root also holds `prompt Acces.txt` (the original Romanian feature spec the extension was built from — useful for "why does X work this way" questions) and a real `.accdb` test/dev database (gitignored, along with `.laccdb` lock files and the prompt file itself — see `.gitignore`).

## Commands (run from `access-explorer/`)

```powershell
npm install          # devDependencies only — zero runtime dependencies
npm run check        # tsc --noEmit
npm run build        # esbuild bundle (dev) -> dist/extension.js
npm run build:prod    # production bundle (minified)
npm run watch         # esbuild --watch
npm run package       # vsce package -> access-explorer-<version>.vsix
```

There is no test suite (no `npm test`). Verification is manual: press **F5** in `access-explorer/` to launch an Extension Development Host, or see the manual test checklist in `access-explorer/README.md` (round-trip module edit, broken-code compile warning, invalid SQL save, exclusive-lock detection, orphaned-process cleanup, etc.).

To install a built `.vsix` locally: `npx vsce package` then `code --install-extension access-explorer-<version>.vsix --force`, followed by "Developer: Reload Window" — the installed extension does **not** auto-update on `git pull`; it always needs a manual rebuild + reinstall.

## Architecture: access-explorer

```
VS Code (extension, src/*.ts)
   │  JSON-lines over stdin/stdout + temp files for code payloads
   ▼
powershell.exe -STA (ps/access-bridge.ps1) — one long-lived process per open database
   │  COM: New-Object -ComObject Access.Application
   ▼
MSACCESS.EXE (hidden instance, opened SHARED)
```

COM automation is isolated in a separate PowerShell process per database so a hung/blocked Access call never freezes the extension's UI. Requests/responses are one JSON object per line; module/query/macro bodies always travel through temp files (`workDir`), never inline in JSON, so framing never depends on payload content. `AccessBridge` (`src/bridge.ts`) owns this process and serializes all requests (Access COM is not reentrant); a per-op timeout triggers a hard-kill of the whole ladder (kill process → taskkill by PID → remove stale `.laccdb`) since a hang almost always means a hidden modal dialog.

**Object model** (`src/model.ts`): every Access object (table/query/form/report/macro/module) is exposed as a virtual document under the custom URI scheme `accessdb:/<dbKey>/<Category>/<name><ext>`, implemented by `AccessFsProvider` (`src/fsProvider.ts`), a `vscode.FileSystemProvider`. `dbKey` is a truncated sha1 of the lowercased db path. `DatabaseRegistry` tracks all currently-open databases and fires change events consumed by the tree view.

**Read/write pipeline** (`src/fsProvider.ts`): `writeFile`/`saveObject` always does conflict-check (re-export and diff against the baseline captured at load time; refuses to overwrite if the object changed in Access since) → backup (`BackupManager`, abort the write if backup fails) → COM write → re-read the canonical form (DAO reformats SQL; Access may normalize other exports) → optional async compile check (`acCmdCompileAndSaveAllModules`, which compiles/saves the *whole* VBA project, not just the touched module). Forms/Reports and standard/class Modules carry an opaque header/design blob (`src/vbaHeader.ts`) that is never parsed — only the trailing user-code body is shown/edited; the header is cached at load and spliced back verbatim on save.

**AI mirror** (`src/aiMirror.ts`): because `accessdb:` documents are virtual (no real file on disk), external tools like Claude Code's file/Read/Edit tools can't see them. `AiMirrorManager` mirrors every object to real files under `<db-folder>/.accdb-ai/<Category>/<name><ext>`, keeps them in sync both ways (Access change → mirror file rewritten; external edit to the mirror file → pushed back into Access through the *same* `saveObject` pipeline as Ctrl+S, then offers a revert), and tears the mirror down when the database closes (it's a derived artifact, never a durable store). This is the reason a Claude Code session can read/edit `.bas`/`.form.txt` files that look like normal files but are actually backed by a live Access connection.

**Other notable pieces:**
- `src/aiContext.ts` + `src/vbaSymbols.ts` — "Select Code for AI" (`Ctrl+Alt+A`): groups Sub/Function/Property procedures the way the VBA IDE's dropdowns do (`cmdSave_Click` → bucket `cmdSave`, event `Click`; unmatched → `(General)`) and sets the editor *selection* to the chosen scope, since selected text (not the virtual URI) is what actually reaches chat/AI tooling.
- `src/linkedCredentials.ts` / `src/vbaUnlock.ts` — handle the two cases where Access would otherwise pop a native modal dialog invisible to COM (ODBC-linked table login; password-protected VBA project) by prompting once and persisting to VS Code's encrypted Secret Storage per database.
- `src/errors.ts` — every COM/bridge failure maps to one of a fixed set of `BridgeErrorCode`s, localized centrally; add new failure modes here rather than surfacing raw COM errors.
- `src/redirectEditorProvider.ts` — claims `.accdb` as VS Code's default editor (so double-click/`code file.accdb` never shows the raw binary editor) by briefly opening a dummy webview and immediately handing off to the real `openDatabase` flow.

**Windows-only, Access 2013 quirks worth knowing**: `SaveAsText` on this environment exports class modules as empty files (no antet/header) — the bridge falls back to reading via the VBE (`viaVbe`) in that case, and such modules must be written back the same way. AccessVBOM ("Trust access to the VBA project object model") being disabled surfaces as `VBE_TRUST_REQUIRED` and requires a manual Trust Center change by the user — never attempt to flip this via the registry.

## Architecture: skills/ (VBA conventions for Access projects)

These are Claude Code skills (Romanian) that operate on VBA source *exported as text* from an Access project (`.bas`/`.form.txt`/`.report.txt`), typically accessed through the AI mirror described above. They encode conventions that are real and dominant across the target application's existing codebase — reuse them, don't invent alternatives:

- **`vba-include-tratare-erori`** — adds the standard `TRATARE_ERORI` error-handling block (label-based, `Erl`-driven line capture, logging via a `ScrieEroare` utility function) to one named Sub/Function. If `ScrieEroare` doesn't already exist in the target project, it must stop and propose creating it + its log table (`references/creare-scrieroare-si-tabela.md`) rather than inventing a different name — this is a common point where a run can appear to "skip" the treatment if the creation proposal wasn't separately confirmed.
- **`vba-analiza`** — read-only analysis (bugs, performance, unclosed resources, SQL injection, duplication, dead code) of a named routine/module; only modifies code on explicit confirmation.
- **`vba-corectare-erori`** — the only one of the three with a *live* connection: reads untreated rows from the linked SQL Server table `dbo_Erori program` via `access-explorer/ps/get-untreated-errors.ps1`, locates the responsible routine, proposes a fix, and (only on explicit per-row confirmation) writes it back via `access-bridge.ps1` and marks the row treated via `mark-error-treated.ps1`.

Shared conventions across all three:
- AI-authored/modified code is wrapped in marker comments so changes are visible in an editor with no VBA diff view: `'===Start Generat AI===` / `'===Final Generat AI===` for general edits/additions; `vba-corectare-erori` uses a distinct, timestamped variant (`'===Start Corectat AI <timestamp>===`) to mark a fix that originated from a logged error.
- None of these skills has a live DB connection except `vba-corectare-erori` — the other two only ever *propose* SQL/code for the user to run manually in Access.
- Line numbers in treated routines are multiples of 10 (for `Erl`); when adding treatment, prefer minimal-diff renumbering (fill gaps) over a full renumber, and never break existing `GoTo <n>`/`Resume <n>`/`Select Case Erl` references.

`access-explorer/ps/` also has a few standalone dev/test PowerShell helpers used only by `vba-corectare-erori` and manual testing (`get-untreated-errors.ps1`, `mark-error-treated.ps1`, `add-test-error.ps1`, `db-link-helper.ps1`) — these are not part of the shipped extension. `access-explorer/ps/db-credentials.local.json` (gitignored) holds the SQL Server credentials they need; copy it from the `.example` file and never commit real credentials.
