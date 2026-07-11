# Checklist-uri per operație — `/vba-analiza`

Tipare concrete VBA/Access de căutat, per operație. Referință încărcată doar la nevoie de `SKILL.md`.

## optimizeaza

- `DLookup`/`DFirst`/`DMax`/`DSum`/`DCount` apelate în interiorul unei bucle → înlocuiește cu un recordset deschis o singură dată sau cu o interogare agregată.
- `Screen.Updating` / `DoCmd.Echo False` lipsă în jurul unei secvențe de operații UI grele (multe `.Requery`, multe scrieri pe formular) → adaugă și restaurează la final.
- `.Requery` apelat de mai multe ori inutil pe același control/recordset în aceeași rutină.
- `SELECT *` într-o interogare folosită doar pentru câteva coloane → limitează la coloanele efectiv necesare.
- Aceeași expresie/valoare recalculată de mai multe ori în interiorul unei bucle, deși nu depinde de indexul buclei → calculeaz-o o singură dată înainte de buclă.
- `.FindFirst` pe un recordset mare unde există un index potrivit → `Seek` (necesită recordset de tip table, `dbOpenTable`).
- Concatenare de string cu `&` în interiorul unei bucle cu multe iterații → adună într-un array și `Join` la final, dacă volumul justifică.

## bug-uri

- Comparație/operație pe o valoare posibil `Null` fără `Nz(...)` în jur (ex. `DLookup` folosit direct într-un calcul sau `If`).
- Off-by-one în `For i = 1 To ...`/`For i = 0 To ...` raportat la limitele reale ale colecției/array-ului folosit.
- Condiție care e mereu adevărată sau mereu falsă (ex. comparație pe un tip incompatibil, sau o variabilă atribuită mereu aceeași valoare înainte de test).
- Variabilă refolosită între iterațiile unei bucle fără resetare explicită la începutul iterației, când logica presupune o valoare "curată".
- Conversie/tip de date nepotrivit pentru operația efectuată (ex. `Integer` unde valoarea poate depăși ~32767, `String` comparat direct cu un `Date`).
- Parametru opțional (`Optional`) folosit fără verificare `IsMissing(...)` sau fără valoare implicită.
- `Resume Next` folosit generic într-un handler, care ascunde o eroare reală în loc să o trateze specific.

## resurse

- `Recordset`/`Database`/`QueryDef` deschis (`.OpenRecordset`, `CurrentDb.OpenRecordset`, etc.) fără `.Close` corespunzător, inclusiv pe ramura de eroare a rutinei.
- Obiect `Set` fără eliberare (`Set obj = Nothing`) la finalul rutinei sau în handler-ul de eroare, când obiectul nu mai e necesar după ce rutina se termină.
- Referințe la `Forms!<Nume>`/`Reports!<Nume>` fără o verificare prealabilă că formularul/raportul respectiv e deschis (risc de eroare 2450/2465, vezi și `vba-include-tratare-erori`).

## securitate

- Interogări SQL construite prin concatenare directă a unui input de la utilizator (control de formular, parametru) în șirul SQL, de tipul `"SELECT ... WHERE camp = '" & control & "'"` → risc de SQL injection și de eroare la apostrof în input.
- Propunere: parametrizare prin `QueryDef.Parameters`, sau — dacă interogarea rămâne dinamică prin concatenare — escaparea corectă a apostrofului (`Replace(valoare, "'", "''")`) și validarea tipului de date înainte de concatenare.

## duplicat

- Blocuri de cod (peste câteva linii) aproape identice, repetate în mai multe rutine din același fișier/modul, care diferă doar prin nume de control/variabilă sau o constantă.
- Propunere concretă: semnătura unei funcții/proceduri comune care înlocuiește duplicarea, cu parametrii care variază între apeluri.

## cod-mort

- `Dim` declarat dar niciodată folosit în rutină.
- Parametru al rutinei niciodată folosit în corpul ei.
- Cod plasat imediat după un `GoTo <eticheta>`/`Exit Sub`/`Exit Function`/`End` necondiționat, care nu mai poate fi atins din fluxul normal (exceptând eticheta țintă a unui `GoTo` din interiorul aceleiași rutini).

## documenteaza

Format propus pentru comentariul de sumar, inserat direct deasupra liniei `Sub`/`Function`, doar la cerere explicită de aplicare:

```vb
'===Start Generat AI===
' <o propoziție: ce face rutina>
' Parametri: <nume> - <rol>, ...   (omis dacă nu are parametri)
' Returneaza: <ce returnează>       (doar pentru Function)
'===Final Generat AI===
```

Nu inventa detalii despre parametri/retur care nu reies din cod — dacă rolul unui parametru nu e clar din utilizare, formulează generic sau întreabă.

## erori

Detectarea tratării existente: caută eticheta `TRATARE_ERORI:` și `On Error GoTo TRATARE_ERORI` (convenția din `vba-include-tratare-erori`), sau orice alt `On Error GoTo <eticheta>` din rutină.

- Dacă rutina **nu are** nicio formă de `On Error` → raportează lipsa și recomandă explicit: "rulează skill-ul `vba-include-tratare-erori` pe această rutină" — nu construi tu blocul de tratare aici.
- Dacă rutina **are deja** tratare de erori → poți verifica dacă e completă (există `Case Else`, se apelează `ScrieEroare` cu semnătura corectă folosită în restul proiectului, există `Exit Sub`/`Exit Function` înainte de eticheta handler-ului) și raportează eventuale lipsuri, tot ca recomandare către `vba-include-tratare-erori` pentru completare — nu aplica tu modificări pe blocul de tratare a erorilor.
