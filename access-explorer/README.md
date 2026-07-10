# Access Explorer

Extensie VS Code pentru lucrul cu baze de date **Microsoft Access (.accdb)** prin **automatizare COM** (fără parsare binară a fișierului). Necesită Windows cu Microsoft Access instalat.

## Funcționalități

- **Deschidere** `.accdb` din comandă (`Access: Open Access Database`) sau click-dreapta pe fișier în Explorer.
- **TreeView** în sidebar cu toate componentele: Tabele, Interogări, Formulare, Rapoarte, Macro-uri, Module VBA (standard + clasă).
- **Editare cu scriere înapoi** în baza de date pentru:
  - **Module VBA** (`.bas` / `.cls`) — cu syntax highlighting VBA inclus;
  - **Macro-uri** (`.mac`) — în serializarea text Access;
  - **SQL-ul interogărilor** (`.sql`) — cu validare de sintaxă la salvare (DAO respinge SQL invalid, documentul rămâne dirty).
- **Read-only** pentru definițiile tabelelor, formularelor și rapoartelor (lacăt în editor).
- **Ctrl+S** salvează direct în `.accdb`; după salvarea unui modul rulează o **verificare de compilare VBA** și avertizează dacă proiectul nu compilează.
- **Backup automat** înainte de fiecare scriere în `<folder DB>\.accdb-backups\<nume>.<timestamp>.accdb`, cu retenție configurabilă. Dacă backup-ul eșuează, scrierea este anulată.
- **Detectare conflicte**: dacă obiectul a fost modificat în Access după deschidere, salvarea e refuzată până rulezi Refresh.
- **Refresh** recitește structura bazei de date și reîncarcă editoarele deschise.
- **UI bilingv**: engleză implicit, română cu `code --locale=ro` (sau Display Language: Română).

## Arhitectură

```
VS Code (extensie TypeScript)
   │  JSON-lines pe stdin/stdout + fișiere temp pentru cod
   ▼
powershell.exe -STA  (un proces per bază de date, long-lived)
   │  COM: New-Object -ComObject Access.Application
   ▼
MSACCESS.EXE (instanță ascunsă, deschidere SHARED)
```

- Comunicarea COM este **izolată într-un proces separat** — UI-ul extensiei nu se blochează niciodată.
- Modulele VBA se citesc/scriu cu `Application.SaveAsText / LoadFromText` — **nu** necesită setarea „Trust access to the VBA project object model".
- Antetul tehnic al exportului (`VERSION/BEGIN/Attribute VB_*`) este ascuns la editare și restaurat identic la scriere.
- La închidere (sau crash) extensia închide Access (`Quit acQuitSaveNone`), eliberează obiectele COM și, ca ultimă soluție, omoară procesul MSACCESS.EXE după PID și șterge `.laccdb` rămas blocat — fără instanțe orfane.

## Dependențe

**Zero dependențe runtime și zero module native.** Puntea COM este PowerShell (livrat cu Windows), deci nu e nevoie de `winax`, node-gyp sau Visual Studio Build Tools.

Dependențe doar de dezvoltare (instalate cu `npm install`): `typescript`, `esbuild`, `@types/vscode`, `@types/node`, `@vscode/vsce`, `@vscode/l10n-dev`.

## Build și instalare locală

```powershell
cd access-explorer
npm install          # doar devDependencies
npm run check        # type-check (tsc --noEmit)
npm run build        # bundle dev cu esbuild -> dist/extension.js
```

Rulare în dezvoltare: deschide folderul `access-explorer` în VS Code și apasă **F5** (Extension Development Host).

Împachetare și instalare `.vsix`:

```powershell
npx vsce package                                       # produce access-explorer-0.1.0.vsix
code --install-extension .\access-explorer-0.1.0.vsix
```

## Testare manuală recomandată

1. Creează `TestDb.accdb` cu: 1 tabel, 1 interogare, 1 formular, 1 raport, 1 macro, 1 modul standard, 1 modul de clasă (pune diacritice ăâîșț într-un comentariu VBA).
2. Punte standalone (fără VS Code): `powershell -NoProfile -STA -File ps\access-bridge.ps1`, apoi tastează:
   `{"id":1,"op":"open","args":{"path":"C:\\...\\TestDb.accdb"}}`, `{"id":2,"op":"list"}`, `{"id":3,"op":"quit"}` — stdout trebuie să conțină doar JSON; MSACCESS.EXE apare și dispare; `.laccdb` e șters.
3. Click-dreapta pe `.accdb` → arborele arată cele 6 categorii (fără `MSys*`/`~*`).
4. Round-trip modul: deschide `Module1.bas` (fără linii `Attribute` vizibile), editează, Ctrl+S, verifică în Access.
5. `Sub Broken(` → salvat + warning de compilare.
6. SQL invalid (`SELEC *`) → salvarea eșuează, documentul rămâne dirty.
7. Deschide DB-ul exclusiv în Access (`msaccess /excl`) → extensia raportează DB blocat; deschis normal → coexistă.
8. Omoară `powershell.exe` din Task Manager → extensia anunță pierderea conexiunii și nu rămâne MSACCESS.EXE orfan.

## Limitări și riscuri

- **Doar Windows**, cu Microsoft Access instalat (arhitectura pe biți a Access-ului nu contează — COM este out-of-process).
- **AutoExec / formulare de startup**: deschiderea bazei de date prin COM execută macro-ul AutoExec și setările de startup ale bazei, ca la orice deschidere. Nu există flag de automatizare care să le sară. Pentru baze cu AutoExec agresiv, testează întâi pe o copie.
- **Trust Center**: bazele din locații neîncredere pot fi blocate sau pot afișa dialoguri; extensia detectează blocajul prin timeout și resetează conexiunea, dar soluția corectă este adăugarea folderului în *Trusted Locations*.
- **Macro-urile** se editează în serializarea text a Access (formatul `SaveAsText`), nu într-un designer; formatul este sensibil la versiunea de Access.
- **Scriere în mod shared = last-writer-wins** dacă doi utilizatori editează același obiect simultan; atenuat prin backup automat + detectarea conflictelor la salvare.
- Verificarea de compilare folosește `acCmdCompileAndSaveAllModules`, care **salvează toate modulele**, nu doar cel editat (dezactivabilă din `accessExplorer.compileAfterSave`).
- **Baze protejate cu parolă / criptate** nu sunt suportate în această versiune.
- **Proiect VBA protejat cu parolă** (Tools > VBAProject Properties > Protection): la deschidere,
  extensia detectează blocajul și oferă comanda *Unlock VBA Project* — parola se copiază în
  clipboard, iar tu o lipești în dialogul nativ Access care apare (nu există API COM public pentru
  a o introduce programatic). Parola poate fi reținută în VS Code Secret Storage per bază de date.
- **OneDrive/Dropbox**: fișierele `.accdb` sincronizate în cloud (cum e cazul acestui workspace, aflat pe OneDrive) riscă conflicte de sincronizare și blocaje pe `.laccdb`. Recomandat: pune baza de date pe un folder local nesincronizat sau pune sincronizarea pe pauză cât lucrezi.
- Deși există backup automat, pentru operații ample fă și un **backup manual** înainte (comanda *Open Backups Folder* îți arată copiile existente).

## Setări

| Setare | Implicit | Descriere |
|---|---|---|
| `accessExplorer.backupBeforeWrite` | `true` | Backup înainte de fiecare scriere |
| `accessExplorer.backupCount` | `10` | Câte backup-uri se păstrează per DB |
| `accessExplorer.backupDebounceSeconds` | `30` | Interval minim între backup-uri (0 = la fiecare salvare) |
| `accessExplorer.compileAfterSave` | `true` | Verificare de compilare VBA după salvarea unui modul |
| `accessExplorer.operationTimeoutSeconds` | `20` | Timeout pentru operațiile COM |
| `accessExplorer.vbaUnlockTimeoutSeconds` | `120` | Cât timp se așteaptă introducerea parolei la deblocarea unui proiect VBA protejat |
