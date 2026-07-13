---
name: vba-progres
description: Afiseaza sau actualizeaza jurnalul de progres per baza de date Access (.accdb-progress/) - ce module/formulare/rapoarte/rutine au fost deja lucrate, cu ce scop, si ce mai exista netratat in fisier - pentru reorientare rapida la redeschiderea unei baze sau inceperea unei sesiuni noi. Se declanseaza la cererea explicita a utilizatorului (ex. "/vba-progres", "ce am lucrat pana acum in baza asta", "arata progresul", "noteaza scopul acestui formular").
argument-hint: [DbPath]
---

## Ce face acest skill

Menține și afișează jurnalul de progres per bază de date (`.accdb-progress/`, lângă `.accdb`) — ce module/formulare/rapoarte/interogări/tabele au fost deja atinse, cu ce scop, ce s-a făcut concret și ce rămâne netratat din inventarul complet al bazei. Schema exactă a jurnalului și a antetului inserat în cod sunt documentate separat, ca sursă unică:

- `references/schema-jurnal.md` — structura `.accdb-progress/` (INDEX.md + fișiere per obiect), câmpurile "Scop"/"Scop general", regulile de regenerare.
- `references/antet-cod.md` — antetul `AI-Track` inserat în codul VBA al rutinelor modificate (identificator + amprentă de conținut), și algoritmul de detectare a modificărilor externe.

Celelalte trei skill-uri (`vba-analiza`, `vba-include-tratare-erori`, `vba-corectare-erori`) urmează aceleași două documente ca parte din pașii lor proprii, după orice modificare confirmată — nu invocă acest skill, dar scriu în aceleași fișiere, cu aceeași schemă. Acest skill e punctul de intrare pentru:

- **Status** (mod implicit) — vezi rapid, la începutul unei sesiuni sau la redeschiderea bazei, imaginea de ansamblu + ce s-a lucrat + ce nu.
- **Notă manuală** — înregistrează scopul unui obiect/rutină dictat explicit de utilizator, sau o lucrare făcută manual în Access (nu prin celelalte skill-uri).

## Prerechizite

- Calea către `.accdb`-ul țintă — dacă nu e evidentă din context (un singur `.accdb` în proiect), întreabă utilizatorul.

## Pași de aplicat

### Mod status (implicit — fără subcomandă, sau "arată"/"status"/"progres")

1. Obține inventarul complet și proaspăt al bazei prin `access-bridge.ps1`, de sine stătător (fără VS Code): pornește procesul, trimite `{"id":1,"op":"open","args":{"path":"<DbPath>"}}`, apoi `{"id":2,"op":"list"}`, citește răspunsul (`tables`/`queries`/`forms`/`reports`/`macros`/`modules`), apoi `{"id":3,"op":"quit"}`.
2. Citește `.accdb-progress/INDEX.md` și fișierele per-obiect existente, dacă există folderul.
3. Regenerează `INDEX.md` conform `references/schema-jurnal.md` (inventarul proaspăt de la pasul 1, cruce-referențiat cu fișierele per-obiect existente — obiectele fără fișier propriu apar doar ca nume în lista "netratate", nu li se creează fișier gol).
4. Afișează utilizatorului, pe scurt:
   - Imaginea de ansamblu (secțiunea liberă din `INDEX.md`, dacă există).
   - Un rezumat pe categorie: câte obiecte atinse din total, cu scopul lor pe scurt.
   - Lista obiectelor complet netratate, grupată pe categorie.
   - Orice mențiune de "propuneri neconfirmate" rămase în fișierele per-obiect (lucru analizat dar neaplicat încă).
5. **Nu recalculează automat hash-urile din `antet-cod.md` pentru toate rutinele atinse la fiecare rulare de status** (ar necesita un `getModule`/`getFormDef`/`getReportDef` per obiect atins — cost inutil doar ca să afișezi un rezumat). Dacă utilizatorul cere explicit o verificare de integritate ("verifică dacă s-a mai umblat la ce am lucrat"), abia atunci fă fetch pe fiecare obiect atins și aplică algoritmul de comparare din `antet-cod.md`, raportând orice discrepanță găsită.

### Mod notă manuală

Când utilizatorul dictează explicit scopul unui obiect/rutină, sau descrie o lucrare făcută manual în Access (nu prin `vba-analiza`/`vba-include-tratare-erori`/`vba-corectare-erori`):

1. Identifică obiectul (și rutina, dacă se aplică) exact ca în celelalte skill-uri.
2. **Fetch sursa curentă** a rutinei/obiectului prin `access-bridge.ps1` (`getModule`/`getFormDef`/`getReportDef`) — chiar și pentru o notă pur descriptivă, ca hash-ul din antet (dacă se scrie unul) să corespundă codului real, nu unei presupuneri. Nu inventa niciodată un hash fără să fi citit efectiv codul curent.
3. Scrie/actualizează fișierul `.accdb-progress/<Category>/<obiect>.md` corespunzător conform `references/schema-jurnal.md` (Scop/Scop general dictat de utilizator, marcat explicit ca atare — nu ca dedus de Claude).
4. Dacă utilizatorul confirmă că vrea și antetul `AI-Track` inserat (de exemplu, lucrarea manuală a modificat efectiv codul rutinei), aplică `references/antet-cod.md` și scrie înapoi în Access. Dacă nota e doar despre scop/intenție, fără nicio schimbare de cod, **nu** insera niciun antet — antetul e strict legat de o modificare de cod reală.

## Ce NU face acest skill

- Nu modifică logica de business a rutinelor — doar jurnalul extern și, când aplicabil, antetul de tracking.
- Nu inserează antetul `AI-Track` fără să fi citit efectiv codul curent al rutinei (nu inventează hash-uri).
- Nu șterge fișiere per-obiect existente din `.accdb-progress/` fără cerere explicită a utilizatorului.
- Nu se conectează la SQL Server și nu are nicio legătură cu comutarea de rețea (`accessExplorer.syncErrorLog`/`markErrorsTreated`) — acelea sunt exclusiv apanajul `vba-corectare-erori`.
- Nu regenerează `INDEX.md` recalculând hash-uri de integritate pentru toate obiectele la fiecare status — doar la cerere explicită (vezi pasul 5 de mai sus).
