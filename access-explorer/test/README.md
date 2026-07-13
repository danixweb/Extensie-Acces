# Suita de regresie pentru bridge (headless)

Testează `ps/access-bridge.ps1` și scripturile standalone din `ps/` pe o bază `.accdb`
reală, fără VS Code — exact protocolul folosit de extensie (`src/bridge.ts`).

## Rulare

```powershell
cd access-explorer
npm run test:bridge                       # Fazele 0–4 (build, citiri, round-trip scrieri, robustețe)
npm run test:bridge -- --with-server      # + Faza 5 (tabele legate SQL Server; cere ps/db-credentials.local.json)
npm run test:bridge -- --only T3.4,T4     # doar testele cu prefixul dat
npm run test:bridge -- --db "C:\alta\baza.accdb" --skip-build
```

Ieșire: TAP pe stdout + `results.json`, `bridge-stderr.log`, `run.log` și payload-urile
în `%TEMP%\access-explorer-tests\run-<timestamp>\` (calea e afișată la start și la final).

## Precondiții

- Microsoft Access instalat (implicit se testează pe `..\64B sursa SQL SERVADENT STOC.accdb`).
- Niciun MSACCESS.EXE pornit — preflight-ul (T0.3) oprește suita altfel.
- Pentru `--with-server`: `ps/db-credentials.local.json` valid (vezi `.example`).

## Siguranță

- Toate scrierile sunt blocuri-comentariu marcate `'===Start Generat AI===` și sunt
  **restaurate în `finally`** — dacă un test moare între salvare și revert, restul de
  cod e ușor de găsit după marker.
- Suita nu creează și nu șterge obiecte în baza de date. `compact` (T4.1) rescrie
  fișierul in-place (comportament non-destructiv, cu reopen).
- T5.4 **inserează un rând real** în tabela legată `dbo_Erori program` de pe SQL Server
  și îl lasă acolo marcat `tratata=True` (raportat în findings).
- T5.5 scrie temporar credențiale greșite în connect-string-uri și le restaurează
  întotdeauna la final; dacă restaurarea eșuează, apare un finding CRITIC.
- Orice blocaj (dialog modal invizibil în Access) e tratat de kill-ladder:
  kill powershell → `taskkill` pe MSACCESS → ștergere `.laccdb`.

Folderul e exclus din pachetul `.vsix` prin `.vscodeignore` (`test/**`).
