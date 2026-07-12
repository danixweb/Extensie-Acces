# Coduri de eroare VBA/Access frecvente — cauză și tratare recomandată

Tabel de referință pentru Pasul 5 din SKILL.md. Nu insera un `Case` decât dacă găsești în rutina analizată o operație concretă care poate produce acel cod — nu adăuga `Case`-uri speculative fără corespondent în cod.

| Cod | Nume / descriere | Cauză tipică în cod | Tratare recomandată |
|---|---|---|---|
| 94 | Invalid use of Null | `DLookup`/`DFirst`/`DMax`/`DSum` fără `Nz(...)`, sau un control gol folosit direct într-o expresie | Ideal: repară sursa cu `Nz(DLookup(...), 0)` sau `Nz(..., "")`. Ca plasă de siguranță suplimentară, `Case 94`: dacă eroarea vine de la o linie/anumit context cunoscut, `Resume Next` cu valoare implicită; altfel `MsgBox` explicit + `Resume Next`. |
| 11 | Division by zero | `x / y`, `Int(x / y)`, `x Mod y` cu numitor ce poate fi 0 (ex. un `DateDiff` sau o cantitate care poate fi 0) | `Case 11`: fie verifică numitorul înainte (recomandat, elimină eroarea la sursă), fie în handler `Resume Next` cu rezultat 0/implicit și un mesaj scurt către utilizator. |
| 13 | Type mismatch | `CDbl`/`CInt`/`CLng`/`CDate`/`CStr` aplicat pe un control/variabilă ce poate fi Null, text gol sau format neașteptat | `Case 13`: `MsgBox` care indică ce valoare era invalidă (dacă se poate identifica din context) + `Resume Next` sau `Exit Sub`/`Function`, în funcție de dacă restul rutinei poate continua fără acea valoare. |
| 91 | Object variable or With block variable not set | Folosirea unui obiect (`Recordset`, `Control`, `Form`) înainte de `Set`, sau după ce a fost `Set ... = Nothing` | `Case 91`: de regulă eroare de logică reală, nu doar "de mediu" — recomandă `Exit Sub`/`Function` după log, nu `Resume Next` orb (ar continua cu obiectul tot nesetat). |
| 2450 | Can't find the form referenced in a macro or Visual Basic code expression | `Forms![NumeFormular]` către un formular care nu e deschis în acel moment | `Case 2450`: verifică ideal cu `CurrentProject.AllForms("NumeFormular").IsLoaded` înainte; ca tratare de avarie, mesaj explicit ("Deschideți mai întâi formularul X") + `Exit Sub`/`Function`. |
| 2465 | Application-defined or object-defined error (câmp/control inexistent) | `!NumeControl` sau `!NumeCamp` care nu (mai) există pe formularul/tabela țintă | `Case 2465`: mesaj cu numele controlului/câmpului suspectat (din contextul rutinei) + `Exit Sub`/`Function`; de regulă nu se poate face `Resume Next` sigur. |
| 3021 | No current record | Navigare recordset (`.MoveNext`/`.MovePrevious`/citire câmp) fără verificare `EOF`/`BOF`, sau recordset gol | `Case 3021`: dacă apariția e așteptată la finalul unei bucle, `Resume Next` (vezi modelul din `references/exemplu-complet.md` #5, unde e chiar sub-ramificat pe linie); altfel mesaj + `Exit Sub`/`Function`. |
| 3167 | Record is deleted | Operație pe o înregistrare ștearsă între citire și scriere (concurență) | `Case 3167`: mesaj clar de tip "înregistrarea a fost ștearsă între timp, reluați operația" + `Exit Sub`/`Function` (nu `Resume Next` — datele sursă nu mai există). |
| 3061 | Too few parameters. Expected N | Query parametrizat cu nume de parametru greșit/lipsă (des la interogări salvate apelate din cod) | De regulă bug de configurare, nu excepție de runtime normală — `Case 3061`: log + mesaj explicit către dezvoltator/admin, nu recuperare automată. |
| 3078 | The Microsoft Access database engine cannot find the input table or query | Nume de tabelă/interogare greșit sau redenumit, folosit într-un `DLookup`/`OpenRecordset`/SQL direct din cod | `Case 3078`: log + mesaj explicit, `Exit Sub`/`Function` — nu recuperare automată, e o eroare de configurare/schemă. |
| 2001 / 2501 | Acțiune anulată de utilizator (ex. a închis un dialog `SaveAs`/print) | `DoCmd.OutputTo`/`DoCmd.TransferSpreadsheet`/print preview anulat de user | `Case 2001` (sau 2501, verifică care apare efectiv): tratare silențioasă, fără mesaj de eroare — `Resume Next` sau `Exit Sub` curat, utilizatorul a anulat intenționat. |

## Reguli generale la aplicarea acestui tabel

- Adaugă `Case`-urile relevante **înaintea** lui `Case Else`, niciodată după (VBA execută prima ramură care se potrivește).
- Nu introduce mai mult de un `Case` pentru același cod de eroare.
- Dacă mai multe coduri din tabel au aceeași tratare exactă, poți grupa: `Case 3021, 3167`.
- Codul exact returnat de Access pentru "anulare de utilizator" variază între versiuni/acțiuni (2001 vs 2501) — verifică ce apare de fapt în alte rutine din același proiect înainte să presupui unul anume.
- Aceste propuneri se prezintă utilizatorului înainte de aplicare (vezi Pasul 5 din SKILL.md) — nu se inserează tacit.
- La aplicare, comentează atât semnificația codului (coloana "Nume / descriere") cât și motivul tratării alese (coloana "Tratare recomandată" adaptată la rutina concretă) — vezi formatul exact cerut în Pasul 5 din SKILL.md.
