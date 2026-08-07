# Checklist-uri per operație — `/vba-analiza`

Tipare concrete VBA/Access de căutat, per operație. Referință încărcată doar la nevoie de `SKILL.md`.

## optimizeaza

- `DLookup`/`DFirst`/`DMax`/`DSum`/`DCount` apelate în interiorul unei bucle → înlocuiește cu un recordset deschis o singură dată sau cu o interogare agregată.
- Funcții domeniu (`DLookup` etc.) folosite în sursa unui control calculat de pe formular/raport sau într-o coloană de interogare care se evaluează pe fiecare rând → mută logica într-un `JOIN` sau într-o subinterogare.
- `Screen.Updating` / `DoCmd.Echo False` lipsă în jurul unei secvențe de operații UI grele (multe `.Requery`, multe scrieri pe formular) → adaugă și restaurează la final.
- `.Requery` apelat de mai multe ori inutil pe același control/recordset în aceeași rutină.
- `Me.Repaint` / `Me.Recalc` / `DoEvents` apelate în interiorul unei bucle fără o justificare clară (progres vizibil, cedare de control).
- `SELECT *` într-o interogare folosită doar pentru câteva coloane → limitează la coloanele efectiv necesare.
- Aceeași expresie/valoare recalculată de mai multe ori în interiorul unei bucle, deși nu depinde de indexul buclei → calculeaz-o o singură dată înainte de buclă.
- `.OpenRecordset` / `CurrentDb` / `.CreateQueryDef` apelate **în interiorul** unei bucle, deși sursa nu depinde de iterație → deschide o singură dată înainte de buclă.
- Buclă de `.Edit` … `.Update` pe recordset pentru a modifica multe înregistrări după un criteriu uniform → o singură interogare `UPDATE`/`DELETE`/`INSERT` executată cu `db.Execute …, dbFailOnError`.
- `dbOpenDynaset` (sau recordset implicit) folosit pentru citire pură → `dbOpenSnapshot`, mai rapid și fără blocări; pentru o singură trecere înainte, `dbOpenSnapshot + dbForwardOnly`.
- `OpenRecordset("<NumeTabel>")` pe tot tabelul, urmat de filtrare în cod (`.FindFirst`, `If` în buclă) → deschide direct pe un `SELECT ... WHERE ...`, ca filtrarea să fie făcută de motor.
- Recordset deschis exclusiv pentru a adăuga înregistrări, fără `dbAppendOnly` → încarcă inutil setul existent.
- `.MoveLast` executat doar pentru a obține `.RecordCount` pe un recordset mare → `DCount` sau o interogare agregată `COUNT(*)`.
- `DoCmd.RunSQL` pentru interogări de acțiune → `db.Execute …, dbFailOnError` (mai rapid, fără dialoguri de confirmare, raportează erorile).
- Câmpuri folosite în `WHERE`, `JOIN` sau `ORDER BY` care nu au index în tabelul sursă → semnalează, ca recomandare de structură (nu modifica schema).
- `.FindFirst` pe un recordset mare unde există un index potrivit → `Seek` (necesită recordset de tip table, `dbOpenTable`).
- Concatenare de string cu `&` în interiorul unei bucle cu multe iterații → adună într-un array și `Join` la final, dacă volumul justifică.
- `Variant` folosit pentru contoare/variabile de buclă unde un tip explicit (`Long`) e suficient.
- Legare târzie (`Dim x As Object` + `CreateObject`) în cod apelat des, unde legarea timpurie e posibilă — semnalează ca tradeoff (viteză vs. portabilitate între versiuni de Office), nu ca defect automat.

## bug-uri

- Comparație/operație pe o valoare posibil `Null` fără `Nz(...)` în jur (ex. `DLookup` folosit direct într-un calcul sau `If`).
- Comparație directă cu `Null` (`If x = Null`, `If x <> Null`) → rezultatul e mereu `Null`, deci ramura nu se execută niciodată; folosește `IsNull(x)`.
- `Nz(valoare)` fără al doilea argument într-un context numeric → returnează string vid; folosește `Nz(valoare, 0)`.
- `And`/`Or` nu fac scurtcircuit în VBA: ambii membri se evaluează întotdeauna. Tipar periculos: `If Not rs.EOF And rs!Camp = ...` → mută condiția de gardă într-un `If` imbricat.
- `Option Explicit` lipsă la începutul modulului → greșelile de tipar în numele variabilelor devin variabile noi, `Empty`.
- `Dim a, b As Long` → doar `b` primește tipul, `a` rămâne `Variant`; declară tipul pentru fiecare variabilă.
- Off-by-one în `For i = 1 To ...`/`For i = 0 To ...` raportat la limitele reale ale colecției/array-ului folosit.
- Condiție care e mereu adevărată sau mereu falsă (ex. comparație pe un tip incompatibil, sau o variabilă atribuită mereu aceeași valoare înainte de test).
- Buclă `Do Until rs.EOF` / `Do While Not rs.EOF` fără `.MoveNext` pe toate ramurile interne (ex. lipsește pe ramura unui `If`) → buclă infinită.
- `.EOF`/`.BOF` neverificat înainte de citirea câmpurilor dintr-un recordset proaspăt deschis.
- `.FindFirst`/`.FindNext`/`.Seek` fără verificarea ulterioară a `.NoMatch`.
- `OpenRecordset` fără argumentul `Type` explicit, într-o rutină care folosește `.Seek`, `.FindFirst`/`.FindNext` sau `.RecordCount`, ori care scrie în recordset. DAO alege singur tipul (table-type pentru tabel local, dynaset pentru interogare sau tabel legat), deci comportamentul depinde de sursă, nu de cod: `.Seek` cere `dbOpenTable`, iar `.FindFirst` nu funcționează pe table-type — ambele dau eroarea 3251. Aceeași rutină merge pe o sursă și cade pe alta după ce un tabel local e legat sau înlocuit cu un `SELECT`.
- `.RecordCount` interpretat ca număr total pe un recordset care nu e table-type → dynaset/snapshot returnează câte înregistrări au fost accesate până în acel moment; e corect doar după `.MoveLast`.
- `.Edit`/`.AddNew` pe un recordset deschis pe o sursă neactualizabilă (`dbOpenSnapshot`, `dbReadOnly`, interogare cu agregare, `DISTINCT` sau `UNION`) → eroare 3027.
- `OpenRecordset` actualizabil pe tabel legat SQL Server cu coloană `IDENTITY`/`timestamp` fără `dbSeeChanges` → eroare 3622 (vezi și `tranzactii`).
- `.AddNew`/`.Edit` fără `.Update` corespunzător pe toate ramurile de ieșire din rutină (inclusiv la eroare).
- Dată concatenată într-un șir SQL fără delimitatori `#` și fără format US → interpretare greșită zi/lună; folosește `Format(data, "\#mm\/dd\/yyyy\#")`.
- Număr zecimal concatenat într-un șir SQL pe locale românească → virgula zecimală ajunge în SQL și rupe sintaxa; folosește `Str(valoare)` sau `Replace(CStr(valoare), ",", ".")`.
- Variabilă refolosită între iterațiile unei bucle fără resetare explicită la începutul iterației, când logica presupune o valoare "curată".
- Conversie/tip de date nepotrivit pentru operația efectuată (ex. `Integer` unde valoarea poate depăși ~32767, `String` comparat direct cu un `Date`).
- Valori monetare ținute în `Single`/`Double` → erori de rotunjire; `Currency` sau `Decimal`.
- `/` (împărțire cu rezultat flotant) folosit unde se aștepta `\` (împărțire întreagă) sau invers.
- Comparație de string-uri a cărei corectitudine depinde de `Option Compare` (Database/Text = fără diferențiere de majuscule, Binary = cu diferențiere) fără ca asta să fie intenționat explicit.
- Colecție/`Controls` modificată în timp ce e parcursă cu `For Each` → elemente sărite sau eroare.
- Parametru opțional (`Optional`) folosit fără verificare `IsMissing(...)` sau fără valoare implicită.
- `Me.Dirty = False` lipsă înainte de `Requery`/reinterogare, când codul se bazează pe valoarea deja salvată a înregistrării curente.
- Rutină recursivă fără o condiție de oprire evidentă din cod.
- `Resume Next` folosit generic într-un handler, care ascunde o eroare reală în loc să o trateze specific.
- `db.Execute` fără `dbFailOnError` → eșecul interogării trece tăcut, iar codul continuă ca și cum ar fi reușit.

## resurse

- `Recordset`/`Database`/`QueryDef` deschis (`.OpenRecordset`, `CurrentDb.OpenRecordset`, etc.) fără `.Close` corespunzător, inclusiv pe ramura de eroare a rutinei.
- Obiect `Set` fără eliberare (`Set obj = Nothing`) la finalul rutinei sau în handler-ul de eroare, când obiectul nu mai e necesar după ce rutina se termină.
- Eliberare făcută în altă ordine decât inversă creării (copil înainte de părinte: `Recordset` → `QueryDef` → `Database`; `Range` → `Worksheet` → `Workbook` → `Application`).
- Stare globală comutată și nerestaurată pe toate ramurile de ieșire, inclusiv la eroare: `DoCmd.SetWarnings False`, `DoCmd.Hourglass True`, `Application.Echo False`, `DoCmd.Echo False`.
- `LockEdit` lăsat implicit (`dbPessimistic`) pe un recordset actualizabil → blocarea se menține între `.Edit` și `.Update`; într-o rutină lungă sau într-un context multi-utilizator preferă `dbOptimistic`, iar pentru citire pură `dbOpenSnapshot`/`dbReadOnly`.
- `BeginTrans` fără `CommitTrans` sau `Rollback` pe toate căile de ieșire → tranzacție rămasă deschisă, blocaje pe backend.
- Automatizare Office (`Excel.Application`, `Word.Application`, `Outlook.Application`) fără `.Quit` și fără eliberarea obiectelor copil → procese orfane rămase în memorie.
- Fișier deschis cu `Open ... For Input/Output/Append` fără `Close`, sau `TextStream`/`FileSystemObject` fără `.Close`.
- Obiecte ADODB (`Connection`, `Recordset`, `Command`) deschise fără `.Close` + `Set ... = Nothing`.
- Referințe la `Forms!<Nume>`/`Reports!<Nume>` fără o verificare prealabilă că formularul/raportul respectiv e deschis (risc de eroare 2450/2465, vezi și `vba-include-tratare-erori`).

## securitate

- Interogări SQL construite prin concatenare directă a unui input de la utilizator (control de formular, parametru) în șirul SQL, de tipul `"SELECT ... WHERE camp = '" & control & "'"` → risc de SQL injection și de eroare la apostrof în input.
- Propunere: parametrizare prin `QueryDef.Parameters`, sau — dacă interogarea rămâne dinamică prin concatenare — escaparea corectă a apostrofului (`Replace(valoare, "'", "''")`) și validarea tipului de date înainte de concatenare.
- Același tipar de concatenare în argumentele `DoCmd.OpenForm`/`OpenReport` (`WhereCondition`), în `Filter`/`RecordSource` setate din cod și în argumentele criteriu ale funcțiilor domeniu (`DLookup(..., criteriu)`).
- `Eval(...)` sau `Application.Run` apelate cu un șir care conține input de la utilizator → execuție arbitrară de expresii.
- `Shell(...)` cu cale sau parametri concatenați din input de utilizator, fără validare și fără ghilimele în jurul căii.
- `LIKE` cu wildcard-uri provenite din input neescapate (`*`, `%`, `?`, `[`) → interogare mult mai largă decât se intenționa.
- Parole, șiruri de conexiune sau chei API hard-codate în cod, sau scrise în clar într-un tabel de configurare.

## duplicat

- Blocuri de cod (peste câteva linii) aproape identice, repetate în mai multe rutine din același fișier/modul, care diferă doar prin nume de control/variabilă sau o constantă.
- Propunere concretă: semnătura unei funcții/proceduri comune care înlocuiește duplicarea, cu parametrii care variază între apeluri.

## cod-mort

- `Dim` sau `Public` declarat dar niciodată folosit în rutină.
- `Const` declarat și nefolosit.
- Parametru al rutinei niciodată folosit în corpul ei.
- Cod plasat imediat după un `GoTo <eticheta>`/`Exit Sub`/`Exit Function`/`End` necondiționat, care nu mai poate fi atins din fluxul normal (exceptând eticheta țintă a unui `GoTo` din interiorul aceleiași rutini).
- Etichetă de salt declarată dar niciodată țintită de vreun `GoTo`/`Resume` din rutină.
- Rutină `Private` din modul care nu e apelată nicăieri în același modul (raportează ca posibil cod mort — poate fi apelată prin `Application.Run` sau dintr-o proprietate de formular, deci semnalează, nu șterge).
- `Case` duplicat într-un `Select Case`, sau `Case` plasat după `Case Else`.
- Blocuri mari de cod comentat lăsate în rutină (peste câteva linii consecutive).

## obiecte-neinchise

Subset țintit al operației `resurse`, rulat singur când interesează doar închiderea obiectelor.

- Obiecte deschise dar neînchise după procesare, sau la ieșirea din rutină cauzată de o eroare.
- Verifică fiecare cale de ieșire, nu doar sfârșitul rutinei: `Exit Sub`/`Exit Function` intermediare, `GoTo` către handler, ramuri de `If` care returnează devreme.
- Închiderea trebuie să fie precedată de verificare (`If Not rs Is Nothing Then`), altfel handler-ul de eroare aruncă el însuși eroare când obiectul nu a apucat să fie deschis.

## utilizare-currentdb

- Utilizează `CurrentDb` fără să declare o variabilă pe care să o închidă la finalizarea operațiilor sau la ieșirea accidentală din rutină datorată erorilor.
- `CurrentDb` apelat de mai multe ori în aceeași rutină → fiecare apel reconstruiește obiectul și colecțiile lui; declară o singură dată `Dim db As DAO.Database` / `Set db = CurrentDb`.
- `CurrentDb` vs. `DBEngine(0)(0)`: al doilea returnează aceeași instanță (nu o reîmprospătează) — semnalează folosirea lui acolo unde codul se așteaptă la o stare actualizată a schemei.
- Declarații DAO fără prefix de bibliotecă (`Dim rs As Recordset` în loc de `DAO.Recordset`) → ambiguitate dacă proiectul are și referință ADO.

## compatibilitate

- Declarații API (`Declare Function ... Lib ...`) fără `PtrSafe` → nu compilează pe Access pe 64 de biți.
- Handle-uri, pointeri sau adrese ținute în `Long` în loc de `LongPtr` în declarațiile API și în variabilele care le primesc.
- Lipsa blocurilor de compilare condiționată `#If VBA7 Then` / `#If Win64 Then` acolo unde codul trebuie să ruleze pe ambele arhitecturi.
- Legare timpurie către o versiune specifică de bibliotecă Office, într-un fișier distribuit pe stații cu versiuni diferite → risc de referință MISSING; semnalează ca tradeoff față de legarea târzie.
- Funcții/obiecte depreciate sau specifice unei versiuni vechi de Access folosite fără alternativă.

## lizibilitate

- `Option Explicit` lipsă din antetul modulului.
- Numere sau șiruri „magice" repetate în cod (coduri de stare, nume de tabele, praguri) → `Const` cu nume descriptiv.
- Rutină foarte lungă (peste ~100 de linii) sau cu peste 3–4 niveluri de imbricare → propune extragerea unor sub-rutine, cu semnături concrete.
- Rutină cu foarte mulți parametri, dintre care mai mulți opționali cu semnificație corelată.
- `GoTo` folosit pentru control de flux obișnuit, altul decât tratarea erorilor.
- Mai multe instrucțiuni pe aceeași linie separate prin `:`, în afara unor cazuri scurte și evidente.

Denumirea variabilelor și a rutinelor (notație ungurească, prefixe, stil de nume) **nu** face parte din acest skill — nu o raporta nici sub această operație.

## formulare

- `Screen.ActiveForm`/`Screen.ActiveControl` folosite acolo unde există o referință directă și sigură (`Me`, `Me!Control`, `Forms!<Nume>`).
- Referire la controale proprii prin `Forms!<Nume>!<Control>` din interiorul aceluiași formular, în loc de `Me!<Control>`.
- `DoCmd.Close` fără specificarea tipului și numelui obiectului → închide obiectul activ, care nu e garantat cel intenționat.
- `DoCmd.OpenForm` fără `acDialog` acolo unde codul care urmează presupune că formularul s-a închis deja (execuția continuă imediat).
- Cod care depinde de ordinea evenimentelor (`Current`, `BeforeUpdate`, `AfterUpdate`, `BeforeInsert`) fără `Me.Dirty` / `Cancel = True` explicit.
- Verificări de validare puse doar în `AfterUpdate` (prea târziu pentru a anula) în loc de `BeforeUpdate` cu `Cancel`.

## tranzactii

- Operații corelate pe mai multe tabele (master/detail, transfer între conturi, ștergeri în cascadă manuale) executate fără tranzacție → date parțiale dacă apare o eroare la mijloc.
- `BeginTrans` fără `Rollback` în handler-ul de eroare.
- `db.Execute` fără `dbFailOnError` (vezi și `bug-uri`) și fără verificarea `db.RecordsAffected` acolo unde numărul de rânduri afectate contează.
- `OpenRecordset` pe tabele legate la SQL Server cu coloane `IDENTITY`/`timestamp` fără `dbSeeChanges` → eroare 3622.

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

Nu modifica și nu șterge comentariile existente scrise de om; dacă există deja un bloc `'===Start Generat AI===`, înlocuiește-l pe acela, nu adăuga unul nou.

## erori

Detectarea tratării existente: caută eticheta `TRATARE_ERORI:` și `On Error GoTo TRATARE_ERORI` (convenția din `vba-include-tratare-erori`), sau orice alt `On Error GoTo <eticheta>` din rutină.

- Dacă rutina **nu are** nicio formă de `On Error` → raportează lipsa și recomandă explicit: "rulează skill-ul `vba-include-tratare-erori` pe această rutină" — nu construi tu blocul de tratare aici.
- Dacă rutina **are deja** tratare de erori → poți verifica dacă e completă (există `Case Else`, se apelează `ScrieEroare` cu semnătura corectă folosită în restul proiectului, există `Exit Sub`/`Exit Function` înainte de eticheta handler-ului) și raportează eventuale lipsuri obiecte neinchise sau alte probleme de genul, tot ca recomandare către `vba-include-tratare-erori` pentru completare — nu aplica tu modificări pe blocul de tratare a erorilor.
- Semnalează și: `On Error Resume Next` lăsat activ până la finalul rutinei fără un `On Error GoTo 0` care să-l anuleze după secvența pentru care era intenționat.
- Handler care nu eliberează resursele deschise înainte de a ieși (vezi `obiecte-neinchise`).
- `Err.Clear` sau `Resume` care reia execuția fără ca a cauza erorii să fi fost tratată → risc de buclă infinită între handler și linia problematică.