---
name: vba-analiza
description: Analizeaza, optimizeaza, cauta bug-uri sau alte probleme (resurse neinchise, securitate SQL, cod duplicat, cod mort) intr-o rutina/modul VBA dintr-un proiect Access exportat ca text, la cererea explicita a utilizatorului (ex. "/vba-analiza analizeaza NumeRutina", "/vba-analiza bug-uri NumeModul").
argument-hint: <operatie> <nume-rutina-sau-modul>
---

## Ce face acest skill

Analizează o rutină (sau un modul întreg) VBA dintr-un proiect Access exportat ca text (`.bas`/`.form.txt`/`.report.txt`) și raportează constatări, fără să modifice codul decât la cerere explicită. Operații suportate, ca subcomenzi ale lui `/vba-analiza`:

- `analizeaza` — explică ce face rutina: scop, parametri, valoare returnată, dependențe (alte rutine/formulare/tabele apelate), riscuri generale observate.
- `optimizeaza` — probleme de performanță specifice VBA/Access.
- `bug-uri` — erori de logică.
- `resurse` — obiecte/recordset-uri deschise fără eliberare.
- `securitate` — risc de SQL injection prin interogări construite prin concatenare.
- `duplicat` — cod repetat între rutine, candidat la extragere într-o funcție comună.
- `cod-mort` — variabile/parametri nefolosiți, cod inaccesibil.
- `documenteaza` — propune un comentariu sumar deasupra rutinei.
- `erori` — verifică dacă rutina are deja tratare de erori; dacă nu are, recomandă skill-ul `vba-include-tratare-erori` în loc să o implementeze aici.
- fără operație specificată → rulează toate verificările de raportare de mai sus (nu `documenteaza`) pe rutina/modulul indicat și produce un raport unic.

Pentru checklist-urile detaliate ale operațiilor de analiză (`optimizeaza`, `bug-uri`, `resurse`, `securitate`, `duplicat`, `cod-mort`), vezi `references/verificari.md` — încarcă-l doar când ai nevoie de tiparele concrete de căutat, nu de fiecare dată.

## Context necesar înainte de aplicare

1. Identifică exact fișierul și rutina cerută — delimitarea unei rutine e `Sub|Function|Property Get|Let|Set ... End Sub|End Function|End Property`.
2. Dacă utilizatorul nu specifică o rutină/modul prin nume în comandă, determină ținta din contextul editorului, în această ordine:
   - Dacă există o selecție curentă în editor (tipic rezultatul comenzii "Select Code for AI" din extensie) — analizează exact acea selecție, nu presupune "tot fișierul" în locul ei.
   - Altfel, dacă există un document deschis în editor în context — analizează tot documentul (modulul întreg).
   - Doar dacă niciuna din cele de mai sus nu e disponibilă, cere clarificare în loc să presupui o țintă.
3. Dacă rutina are părți deja tratate anterior (ex. are deja tratare de erori de la skill-ul `vba-include-tratare-erori`), nu le rescrie — ia-le ca atare în analiză.

## Reguli generale (valabile pentru toate operațiile)

- **Implicit doar raportează.** Nu editează fișierul decât dacă utilizatorul cere explicit aplicarea unei propuneri anume.
- Orice text chiar inserat în fișier (ex. la `documenteaza`, sau la aplicarea unui fix cerut explicit) e delimitat cu:
  ```vb
  '===Start Generat AI===
  ... cod ...
  '===Final Generat AI===
  ```
- Nu se declanșează proactiv — doar la cerere explicită, pe o rutină/modul numit(ă), la fel ca `vba-include-tratare-erori`.
- Nu inventează convenții noi (nume de funcții, stil de mesaje, denumiri) — reutilizează ce există deja în fișier/proiect.
- Pentru `erori`: nu reimplementează tratarea de erori aici — recomandă rularea skill-ului `vba-include-tratare-erori`, care deține acea convenție.
- Nu verifică convenții de denumire (notație ungurească) — nu face parte din acest skill.

## Format de raport

Pentru fiecare constatare (cu excepția codului propus la `documenteaza`):

```
Constatare: <ce problemă/observație>
Locatie: <linie/fragment de cod>
De ce: <motivul pentru care e o problemă sau un risc>
Propunere: <ce ar trebui schimbat, concret>
```

Dacă o operație nu găsește nimic relevant, spune explicit asta ("nu am găsit probleme de tip X") — nu inventa constatări ca să umpli raportul.

## Ce NU face acest skill

- Nu modifică rutine nenumite explicit de utilizator.
- Nu aplică modificări comportamentale (fix-uri de bug, refactorizări, optimizări) fără confirmare explicită — deciziile astea schimbă comportamentul programului.
- Nu verifică convenții de denumire.
- Nu creează/rulează nimic direct în baza de date (nu există conexiune live din acest mediu) — propune doar cod/SQL, la fel ca `vba-include-tratare-erori`.
