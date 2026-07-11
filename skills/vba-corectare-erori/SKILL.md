---
name: vba-corectare-erori
description: Citeste erorile netratate din tabela legata SQL Server dbo_Erori program, localizeaza rutina VBA responsabila si propune/aplica (la confirmare) o corectie in Access, cand utilizatorul cere explicit analiza/corectarea erorilor din log (ex. "/vba-corectare-erori 5").
argument-hint: <numar-erori-de-procesat>
---

## Ce face acest skill

Spre deosebire de `vba-include-tratare-erori` și `vba-analiza` (care presupun "nu există conexiune live" și doar propun cod/SQL de aplicat manual), acest skill **are conexiune live reală**: citește direct din tabela SQL Server `dbo_Erori program` (legată prin ODBC în `.accdb`) și scrie direct în Access, prin scripturile din `access-explorer/ps/`.

Procesează un număr limitat de erori **netratate** (`tratata=False`), în ordine (`id`), una câte una: pentru fiecare, localizează modulul/rutina VBA responsabilă, propune o corecție concretă, **așteaptă mereu confirmarea explicită a utilizatorului** înainte de a scrie ceva în Access, apoi — la confirmare — aplică corecția (marcată cu timestamp) și marchează rândul ca tratat în SQL Server.

## Prerechizite (verifică înainte de a rula orice)

- `access-explorer/ps/db-credentials.local.json` există și are `uid`/`pwd` valide (copiat din `.example` și completat de utilizator) — dacă lipsește, scripturile de mai jos eșuează cu un mesaj clar; nu inventa credențiale, cere-le utilizatorului.
- Calea către `.accdb`-ul țintă (`-DbPath`) — dacă nu e evidentă din context (un singur `.accdb` în proiect), întreabă utilizatorul.
- Argumentul `<numar-erori-de-procesat>` — dacă lipsește la invocare, întreabă utilizatorul câte erori să proceseze în această rulare (nu presupune un număr implicit).

**Vizibilitate automată** — nu configura nimic special. Toate scripturile din `access-explorer/ps/` citesc singure `ps/settings.local.json` (`visibleOperations`); dacă utilizatorul a activat-o, Access și tabelul/modulul pe care se lucrează efectiv apar pe ecran în timp real, la fiecare pas de mai jos, fără nicio acțiune suplimentară din partea skill-ului. Dacă utilizatorul cere explicit să vadă o singură rulare fără să schimbe setarea persistentă, adaugă `-Visible` la comanda PowerShell respectivă.

## Pași de aplicat

### 1. Ia lista erorilor netratate

Rulează:
```
powershell -NoProfile -File access-explorer\ps\get-untreated-errors.ps1 -DbPath "<DbPath>" -Limit <N>
```
Returnează JSON cu până la N rânduri (`id, ora, numar, mesaj, modul, rutina, utilizator, context`), ordonate după `id`.

### 2. Pentru fiecare rând, în ordine

**a. Verifică dacă rândul e identificabil ca eroare de cod.**

Când tabela e populată corect de `TRATARE_ERORI`/`ScrieEroare` (convenția din `skills/vba-include-tratare-erori/references/creare-scrieroare-si-tabela.md`), `modul`/`rutina` conțin deja direct numele exacte (`ScrieEroare` le extrage din formatul `[Modul].[Rutina]` la momentul apelului și le scrie curat în coloane — verificat empiric pe un caz real: `modul="CodBare"`, `rutina="Barcode_128"`).

- Dacă `modul` și `rutina` sunt **ambele nevide** → identificarea e directă, treci la pasul b.
- Dacă sunt **goale**, încearcă fallback: caută în `context` tiparul `Netratata <Modul> rutina <Rutina>` (aceeași convenție). Dacă nici acolo nu găsești nimic — rândul e foarte probabil o **notificare de business, nu o eroare de cod** (exact ce s-a văzut pe date reale: rânduri cu `numar=0`, `context="E"`, mesaje gen "COMENZI CORECTATE..."). **Sari peste acest rând explicit** (spune utilizatorului de ce), nu marca `tratata` și nu inventa un modul/rutină.

**b. Fetch sursa modulului/rutinei via `access-bridge.ps1`** (protocol JSON pe stdin/stdout, documentat în `access-explorer/README.md` — deschide procesul, trimite `{"op":"open","args":{"path":"<DbPath>"}}`, apoi operația de citire, apoi `{"op":"quit"}`):
   - încearcă întâi ca modul standard/clasă: `{"op":"getModule","args":{"name":"<modul>","file":"<tmp>"}}`;
   - dacă modulul nu există cu acel nume, încearcă formular (`Form_<modul>` sau numele direct, convenție din `vba-include-tratare-erori`) cu `getFormDef`, apoi raport cu `getReportDef`.
   - Fișierul `<tmp>` (scos de operație) conține sursa exportată — citește-l cu `Read`.

**c. Localizează rutina (`Sub`/`Function`) numită de `rutina` în sursă.** Nu atinge alte rutine din același modul.

**d. Diagnostichează cauza** folosind `numar` (cod eroare), `mesaj`, `context`:
   - consultă `skills/vba-include-tratare-erori/references/coduri-eroare-comune.md` (tabel cauză → cod → tratare recomandată) ca ghid principal;
   - verifică dacă rutina are deja un `Select Case errNumar` cu un `Case <numar>` specific pentru acest cod — dacă nu are (a căzut pe `Case Else`), cea mai sigură corecție e să adaugi un `Case <numar>` nou, înaintea lui `Case Else`, urmând exact stilul altor `Case`-uri deja din aceeași rutină (vezi exemplul real din `CodBare.Barcode_128`: `Case 94` tratează o cauză similară);
   - dacă eroarea vine dintr-o cauză de fond identificabilă în corpul rutinei (ex. `DLookup` fără `Nz`, împărțire cu numitor posibil 0, conversie neprotejată) — vezi și checklist-ul din `skills/vba-analiza/references/verificari.md` — propune și o corecție la sursă (nu doar un `Case` de logare), dacă e clar și local rutinei.

**e. Prezintă propunerea concretă** (diff-ul exact: ce linii se adaugă/schimbă) utilizatorului și **cere confirmare explicită** — aplic / nu aplic / sari. Nu presupune un răspuns.

**f. Dacă utilizatorul confirmă:**
   1. Generează timestamp-ul curent (`Get-Date -Format 'yyyy-MM-dd HH:mm:ss'`).
   2. În fișierul temporar cu sursa, inserează corecția delimitată:
      ```vb
      '===Start Corectat AI <timestamp>===
      ... cod corectat/adaugat ...
      '===Final Corectat AI <timestamp>===
      ```
      (Convenție distinctă de `===Start/Final Generat AI===` folosită de `/vba` — aceasta marchează specific o intervenție a acestui skill, pornind de la o eroare din log.)
   3. Scrie înapoi în Access cu operația corespunzătoare tipului de obiect (`saveModule` / `saveFormDef` / `saveReportDef`), trimițând același `<tmp>` editat.
   4. Rulează:
      ```
      powershell -NoProfile -File access-explorer\ps\mark-error-treated.ps1 -DbPath "<DbPath>" -Id <id>
      ```

**g. Dacă utilizatorul refuză sau sare:** nu scrie nimic, nu marchezi `tratata`, treci la următorul rând din listă.

### 3. Rezumat final

După ce ai parcurs toate rândurile din listă (sau utilizatorul oprește procesul): raportează câte au fost corectate, câte sărite/refuzate (și de ce — cod neidentificabil vs. refuz explicit), câte rămân netratate în total.

## Ce NU face acest skill

- Nu marchează niciodată `tratata=True` fără să fi aplicat efectiv o corecție confirmată (excepție: rândurile sărite ca neidentificabile rămân `tratata=False`, nu se ascund).
- Nu inventează un modul/rutină când nu poate identifica niciunul — sare peste rând și spune explicit de ce.
- Nu atinge alte rutine din același modul în afara celei indicate de `rutina`.
- Nu proceseaza mai multe erori decât numărul cerut explicit la invocare.

## Utilitar de testare (nu face parte din fluxul normal)

`access-explorer/ps/add-test-error.ps1` inserează un rând simulat în `dbo_Erori program` (cu `modul`/`rutina`/`numar` populate direct, ca un `ScrieEroare` real) — util doar pentru validarea end-to-end a acestui skill pe un caz controlat, nu se rulează ca parte a corectării propriu-zise. Cere mereu confirmarea utilizatorului înainte de inserare (scrie într-o bază de producție, vizibilă altor utilizatori).
