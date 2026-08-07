---
name: vba-analiza
description: Analizeaza, optimizeaza, cauta bug-uri sau alte probleme (resurse neinchise, tranzactii incomplete, securitate SQL, compatibilitate 32/64-bit, cod duplicat, cod mort) intr-o rutina/modul VBA dintr-un proiect Access exportat ca text, la cererea explicita a utilizatorului (ex. "/vba-analiza analizeaza NumeRutina", "/vba-analiza bug-uri NumeModul").
argument-hint: <operatie> <nume-rutina-sau-modul>
---

## Ce face acest skill

Analizează o rutină (sau un modul întreg) VBA dintr-un proiect Access exportat ca text (`.bas`/`.form.txt`/`.report.txt`) și raportează constatări, fără să modifice codul decât la cerere explicită. Operații suportate, ca subcomenzi ale lui `/vba-analiza`:

**Prezentare**

- `analizeaza` — explică ce face rutina: scop, parametri, valoare returnată, dependențe (alte rutine/formulare/tabele apelate), riscuri generale observate.

**Defecte**

- `bug-uri` — erori de logică.
- `resurse` — obiecte/recordset-uri deschise fără eliberare, stare globală nerestaurată.
- `tranzactii` — operații corelate fără tranzacție, `dbFailOnError` lipsă, eșecuri tăcute la scriere.
- `compatibilitate` — cod care nu rulează pe Access 64-bit sau pe altă versiune de Office (`PtrSafe`, `LongPtr`, referințe legate timpuriu).
- `securitate` — risc de SQL injection prin interogări construite prin concatenare, plus `Eval`/`Shell` cu input de utilizator.

**Calitate**

- `optimizeaza` — probleme de performanță specifice VBA/Access.
- `duplicat` — cod repetat între rutine, candidat la extragere într-o funcție comună.
- `cod-mort` — variabile/parametri nefolosiți, cod inaccesibil.
- `lizibilitate` — `Option Explicit`, valori „magice", rutine prea lungi sau prea imbricate, `GoTo` pentru flux normal. **Nu** include denumirea variabilelor/rutinelor.
- `formulare` — tipare fragile specifice formularelor și rapoartelor (`Screen.ActiveForm`, `DoCmd.Close` fără argumente, dependență de ordinea evenimentelor).

**Verificări țintite (subseturi ale lui `resurse`, rulate singure când asta interesează)**

- `obiecte-neinchise` — doar închiderea obiectelor, pe toate căile de ieșire.
- `utilizare-currentdb` — doar folosirea `CurrentDb` și a variabilei de bază de date.

**Altele**

- `documenteaza` — propune un comentariu sumar deasupra rutinei.
- `erori` — verifică dacă rutina are deja tratare de erori; dacă nu are, recomandă skill-ul `vba-include-tratare-erori` în loc să o implementeze aici.

**Fără operație specificată** → rulează toate verificările de raportare de mai sus, cu două excepții: `documenteaza` (produce cod, nu constatări) și cele două verificări țintite `obiecte-neinchise` / `utilizare-currentdb` (ar dubla constatările lui `resurse`). Produce un raport unic, grupat pe severitate, nu pe operație.

Pentru checklist-urile detaliate ale fiecărei operații, vezi `references/verificari.md` — încarcă-l doar când ai nevoie de tiparele concrete de căutat, nu de fiecare dată. Titlurile de secțiune din acel fișier corespund 1:1 cu numele operațiilor de mai sus.

## Context necesar înainte de aplicare

1. Identifică exact fișierul și rutina cerută — delimitarea unei rutine e `Sub|Function|Property Get|Let|Set ... End Sub|End Function|End Property`.
2. Dacă utilizatorul nu specifică o rutină/modul prin nume în comandă, determină ținta din contextul editorului, în această ordine:
   - Dacă există o selecție curentă în editor (tipic rezultatul comenzii "Select Code for AI" din extensie) — analizează exact acea selecție, nu presupune "tot fișierul" în locul ei.
   - Altfel, caută `.accdb-ai/.cursor-context.json` lângă baza de date curentă — extensia îl scrie automat de fiecare dată când cursorul se mută pe o altă rutină în editor (inclusiv la alegerea unei rutini din combourile de tip "obiect"/"procedură" din bara de breadcrumbs a editorului, care doar mută cursorul, fără să creeze o selecție reală). Conține `{category, module, routine, kindWord, startLine, endLine}` (linii 1-based) — dacă există și pare proaspăt (corespunde documentului/modulului activ), analizează exact acea rutină, citindu-i codul din fișierul-oglindă `.accdb-ai/<category>/<module>...` la liniile indicate.
   - Altfel, dacă există un document deschis în editor în context — analizează tot documentul (modulul întreg).
   - Doar dacă niciuna din cele de mai sus nu e disponibilă, cere clarificare în loc să presupui o țintă.
3. Dacă rutina are părți deja tratate anterior (ex. are deja tratare de erori de la skill-ul `vba-include-tratare-erori`), nu le rescrie — ia-le ca atare în analiză.
4. Unele verificări au nevoie de context din afara rutinei. Dacă nu îl ai, spune-o în constatare în loc să presupui:
   - `Option Explicit`, `Option Compare` și declarațiile `Declare`/API se văd doar în antetul modulului, nu în rutina izolată.
   - Dacă un tabel e local sau legat (și către ce backend) schimbă verdictul pentru `dbSeeChanges`, `Seek` și tipul implicit de recordset.
   - `duplicat` are sens doar la nivel de modul sau mai larg — pe o rutină izolată, spune că verificarea nu se aplică.

## Reguli generale (valabile pentru toate operațiile)

- **Implicit doar raportează.** Nu editează fișierul decât dacă utilizatorul cere explicit aplicarea unei propuneri anume.
- Orice text chiar inserat în fișier (ex. la `documenteaza`, sau la aplicarea unui fix cerut explicit) e delimitat cu:
  ```vb
  '===Start Generat AI===
  ... cod ...
  '===Final Generat AI===
  ```
- După orice editare aplicată (nu doar propusă), recitește fișierul și confirmă că textul chiar apare acolo înainte de a raporta succesul — un apel de editare eșuat sau neexecutat nu se raportează niciodată ca aplicat.
- Nu se declanșează proactiv — doar la cerere explicită, pe o rutină/modul numit(ă), la fel ca `vba-include-tratare-erori`.
- Nu inventează convenții noi (nume de funcții, stil de mesaje, denumiri) — reutilizează ce există deja în fișier/proiect.
- Pentru `erori`: nu reimplementează tratarea de erori aici — recomandă rularea skill-ului `vba-include-tratare-erori`, care deține acea convenție.
- Nu verifică convenții de denumire (notație ungurească) — nu face parte din acest skill, nici sub operația `lizibilitate`.
- **O problemă = o constatare.** Unele tipare apar în checklist-urile mai multor operații (ex. `dbFailOnError` lipsă apare la `bug-uri` și la `tranzactii`; `dbSeeChanges` la `bug-uri` și `tranzactii`). Într-un raport combinat, raporteaz-o o singură dată, sub operația cea mai specifică.

## Format de raport

Pentru fiecare constatare (cu excepția codului propus la `documenteaza`):

```
Constatare: <ce problemă/observație>
Severitate: blocant | risc | imbunatatire
Locatie: <linie/fragment de cod>
De ce: <motivul pentru care e o problemă sau un risc>
Propunere: <ce ar trebui schimbat, concret>
```

Severitatea:

- `blocant` — codul dă eroare de rulare sau de compilare, sau pierde/corupe date (ex. eroare 3622 lipsind `dbSeeChanges`, `Declare` fără `PtrSafe` pe 64-bit, tranzacție fără `Rollback`).
- `risc` — funcționează în cazul de bază, dar cade la o valoare `Null`, la un input neașteptat, într-un context multi-utilizator, sau expune o vulnerabilitate.
- `imbunatatire` — performanță, lizibilitate, duplicare, cod mort.

La un raport combinat, ordonează constatările descrescător după severitate, nu în ordinea operațiilor.

Dacă o operație nu găsește nimic relevant, spune explicit asta ("nu am găsit probleme de tip X") — nu inventa constatări ca să umpli raportul.

## Ce NU face acest skill

- Nu modifică rutine nenumite explicit de utilizator.
- Nu aplică modificări comportamentale (fix-uri de bug, refactorizări, optimizări) fără confirmare explicită — deciziile astea schimbă comportamentul programului.
- Nu verifică convenții de denumire.
- Nu creează/rulează nimic direct în baza de date (nu există conexiune live din acest mediu) — propune doar cod/SQL, la fel ca `vba-include-tratare-erori`.
- Nu modifică schema bazei de date. Lipsa unui index sau un tip de coloană nepotrivit se raportează ca recomandare, nu se aplică.