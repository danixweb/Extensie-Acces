# Antet în cod: identificator universal + detectare modificări externe

Document canonic — referențiat de `vba-progres`, `vba-analiza`, `vba-include-tratare-erori` și `vba-corectare-erori`. Se aplică **doar** la momentul unei modificări de cod efectiv confirmate/aplicate (același declanșator ca marcajele `'===Start/Final Generat AI===` / `'===Start/Final Corectat AI <timestamp>===` deja existente) — niciodată la o simplă analiză read-only, fără scriere.

## De ce există

Jurnalul extern (`.accdb-progress/`) e local și gitignored — nu călătorește cu fișierul `.accdb` dacă acesta ajunge pe altă stație (copiat, mutat pe un share, deschis de alt coleg). Ca identitatea rutinei și semnalul "codul s-a schimbat de la ultima atingere Claude" să supraviețuiască oriunde e deschis fișierul, informația minimă necesară trăiește direct **în codul VBA**, sub formă de comentariu.

## Format

Prima linie a corpului rutinei (imediat după linia `Sub`/`Function`/`Property Get|Let|Set ...`, înaintea oricărui cod și înaintea oricărui bloc `===Start...Generat/Corectat AI===` existent):

```vb
Sub Barcode_128()
    '===AI-Track id=3f9a2b41 hash=8f3e2a1c touched=2026-07-13===
    10  On Error GoTo TRATARE_ERORI
    ...
End Sub
```

Un comentariu simplu, neetichetat (fără număr de linie) — nu interferează cu numerotarea multiplilor de 10 folosită pentru `Erl` (vezi `skills/vba-include-tratare-erori/SKILL.md`), pentru că nu e o linie executabilă numerotată.

## Calculul câmpurilor

Ambele valori sunt SHA1 trunchiat (primele 10 caractere hex), calculat printr-un one-liner PowerShell asupra textului relevant, transportat prin fișier temporar (același tipar deja folosit de `access-bridge.ps1` pentru corpul rutinelor — niciodată inline):

```powershell
$sha1 = [System.Security.Cryptography.SHA1]::Create()
$bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
$hex = -join ($sha1.ComputeHash($bytes) | ForEach-Object { $_.ToString('x2') })
$hex.Substring(0, 10)
```

- **`id`** — SHA1 asupra `"<Category>|<NumeObiect>|<NumeRutina>"` (ex. `"Modules|CodBare|Barcode_128"`). **Deterministic, nu aleator**: aceeași rutină produce mereu același `id`, oricând, pe orice mașină, fără niciun registru central și fără generare/persistare de UUID. Nu necesită să fie citit dintr-o sursă anterioară — se recalculează identic de fiecare dată din nume.
  - Limitare cunoscută: o redenumire a obiectului sau a rutinei schimbă `id`-ul — rutina redenumită e tratată ca "nouă" din perspectiva urmăririi (fișierul de jurnal vechi rămâne orfan sub numele vechi; dacă utilizatorul confirmă că e o redenumire, mută manual conținutul relevant în fișierul nou).
- **`hash`** — SHA1 asupra corpului rutinei, **de la linia imediat următoare antetului, până la `End Sub`/`End Function`/`End Property` inclusiv** (antetul însuși nu intră niciodată în calcul — altfel ar fi circular). Recalculat și rescris de fiecare dată când se scrie efectiv o modificare confirmată în acea rutină.
- **`touched`** — data curentă (`yyyy-MM-dd`), aceeași sursă ca timestamp-ul din marcajele `Generat/Corectat AI` deja existente.

## Când se scrie/actualizează

- **Se inserează prima dată** când o rutină e modificată prin unul dintre skill-urile care scriu cod (`vba-include-tratare-erori`, `vba-corectare-erori`, sau `vba-analiza` la aplicarea explicită a unei propuneri) — nu la o analiză fără scriere.
- **Se actualizează** (recalculează `hash`, păstrează `id`, actualizează `touched`) de fiecare dată când aceeași rutină primește o nouă modificare confirmată.
- Nu se atinge la o rutină pe care skill-ul curent n-a modificat-o, chiar dacă e în același modul.

## Verificare la fiecare atingere ulterioară a unei rutine care are deja antet

Înainte de a te baza pe notele din `.accdb-progress/` pentru o rutină care are deja `'===AI-Track...===`:

1. Fetch sursa curentă a rutinei (deja necesar pentru orice analiză/corecție).
2. Recalculează `hash`-ul peste corpul curent (aceeași rețetă, aceleași limite: de după linia de antet până la `End Sub`/`End Function`/`End Property`).
3. Compară cu `hash`-ul din antetul găsit în cod:
   - **Coincide** → nimic nu s-a schimbat de la ultima atingere Claude; notele din jurnal (`Scop`, `Ce s-a făcut`) rămân de încredere ca punct de plecare.
   - **Nu coincide** → codul a fost modificat de atunci — manual în Access, sau pe altă stație, de altcineva sau de o sesiune Claude fără acces la același `.accdb-progress/` local. **Semnalează explicit utilizatorului** ("codul rutinei `<Nume>` pare modificat de la ultima lucrare Claude din `<touched>` — notele din jurnal ar putea fi depășite") și tratează nota veche cu precauție: re-analizează relevant codul curent în loc să presupui orbește că starea descrisă în notă mai e valabilă.

## Ce NU face acest mecanism

- Nu inserează niciodată antetul fără o modificare de cod confirmată în acel moment (o analiză pur read-only nu scrie nimic în sursă).
- Nu inventează un `hash` fără să fi citit efectiv codul curent al rutinei.
- Nu renumerotează și nu atinge alte rutine din același modul în afara celei modificate.
- Nu înlocuiește sau interferează cu marcajele `===Start/Final Generat AI===`/`===Start/Final Corectat AI===` — antetul e o linie separată, despre identitatea/integritatea rutinei ca întreg, nu despre delimitarea codului inserat.
