---
name: vba-include-tratare-erori
description: Adauga tratamentul standard de erori (stil VBA/Access, TRATARE_ERORI + ScrieEroare) intr-o subrutina Sub/Function, cand utilizatorul cere explicit "adauga tratare erori", "pune tratare de erori" sau echivalent pe o rutina VBA dintr-un proiect Access exportat ca text (.bas/.form.txt/.report.txt).
---

## Ce face acest skill

Adaugă, într-o rutină VBA (Sub sau Function) punctuală, blocul standard de tratare a erorilor folosit în acest tip de aplicație Access — cel bazat pe eticheta `TRATARE_ERORI`, capturarea liniei prin `Erl` și logarea prin funcția utilitară `ScrieEroare`. Pe lângă schelet, analizează codul rutinei și **propune** (nu inserează tacit) `Case`-uri suplimentare de coduri de eroare specifice, plauzibile pentru operațiile din acea rutină.

Acest tipar NU e o invenție a acestui skill — e convenția reală, dominantă în acest tip de cod (sute de rutine o folosesc identic). Regula de bază: **reutilizează, nu reinventa**.

## Context necesar înainte de a aplica

1. Caută în proiectul curent funcția de logare (de regulă `Public Function ScrieEroare(...)`, într-un modul utilitar gen `modMain.bas`).
2. **Dacă există** — reține semnătura exactă găsită (parametrii pot diferi ușor de la un proiect la altul) și folosește-o exact așa cum e definită, nu cum e în exemplul din `references/exemplu-complet.md`.
3. **Dacă NU există nimic echivalent în acest proiect** — nu te opri, nu inventa un alt nume pe loc: urmează procedura din `references/creare-scrieroare-si-tabela.md`, care acoperă crearea funcției `ScrieEroare` și a tabelei de log necesare. Pe scurt:
   - propune codul funcției `ScrieEroare` (adaptat, cerând confirmarea utilizatorului privind modulul unde se adaugă și dacă include și părțile opționale — screenshot, clipboard);
   - propune schema tabelei de log (nume, coloane) — implicit aceeași structură confirmată în acest tip de aplicație (`id`, `ora`, `numar`, `mesaj`, `modul`, `rutina`, `utilizator`, `vazut`, `tratata`, `context`);
   - **nu poți crea tabela direct** (nu ai o conexiune live la baza de date Access/SQL Server din acest mediu) — generează în schimb interogarea SQL de tip `CREATE TABLE` (sintaxă Access SQL) și explică exact utilizatorului cum să o ruleze din Access (Query nou → SQL View → lipește → Run, ca query de definire a datelor);
   - abia după ce funcția și tabela există (create de utilizator sau deja existente), continuă cu pașii de mai jos folosind noua funcție.
4. Dacă ai nevoie de exemplul complet, verificat, de tipar (schelet + `Select Case` cu coduri specifice + template generator), citește `references/exemplu-complet.md`. Dacă ai nevoie de lista codurilor de eroare comune și tratarea recomandată, citește `references/coduri-eroare-comune.md`. Nu le încărca dacă nu ai nevoie de detalii suplimentare — pașii de mai jos sunt suficienți pentru cazul obișnuit.

## Când se aplică

Doar când utilizatorul cere explicit adăugarea tratării de erori pe o rutină (sau rutine) anume. Nu se aplică proactiv, "din oficiu", pe rutine pe care utilizatorul nu le-a menționat, și nu se rescrie tot fișierul.

Rutina/rutinele "anume" cerute nu trebuie neapărat numite explicit în text — dacă utilizatorul nu dă un nume, identifică ținta din contextul editorului, în această ordine:
- Dacă există o selecție curentă în editor (tipic rezultatul comenzii "Select Code for AI") — aplică tratarea doar rutinei/rutinelor din acea selecție.
- Altfel, caută `.accdb-ai/.cursor-context.json` lângă baza de date curentă — extensia îl scrie automat de fiecare dată când cursorul se mută pe o altă rutină în editor (inclusiv la alegerea unei rutini din combourile de tip "obiect"/"procedură" din bara de breadcrumbs, care doar mută cursorul, fără să creeze o selecție reală). Conține `{category, module, routine, kindWord, startLine, endLine}` — dacă există și pare proaspăt, aplică tratarea doar acelei rutini.
- Altfel, dacă există un document deschis în editor în context — aplică tratarea tuturor rutinelor din acel modul care nu o au deja.
- Doar dacă niciuna din cele de mai sus nu e disponibilă, cere clarificare în loc să presupui o țintă.

## Pași de aplicat

### 1. Identifică rutina și contextul

- Numele exact al modulului (ex. `Form_1 1 calcul pontaj subform` pentru un formular, sau numele fișierului `.bas` pentru un modul standard) — folosește exact formatul găsit în alte rutine tratate din același fișier, dacă există.
- Numele rutinei și tipul ei (`Sub` → `Exit Sub`; `Function` → `Exit Function`).
- Dacă rutina are deja parțial tratare de erori (un fragment de `On Error`/label), nu duplica — completează ce lipsește.

### 2. Numerotează liniile executabile (pentru `Erl`) — cu grijă la referințe existente

Fiecare instrucțiune executabilă din corpul rutinei trebuie să aibă un număr de linie în față (convenție: multipli de 10), pentru ca `lngLinia = Erl` din handler să poată identifica exact linia care a picat. `Dim`-urile NU se numerotează.

**Regulă critică — nu strica referințele existente la numere de linie.** În acest cod, numărul unei linii poate fi țintit direct din alte locuri:
- `GoTo <numar>` — goto numeric direct către o linie (distinct de `GoTo TRATARE_ERORI`, care e goto la etichetă, nu la număr);
- `Select Case lngLinia` sau `Select Case Erl` cu `Case <numar1>, <numar2>, ...` — folosit ca sub-ramificare *în interiorul* unui `Case <errNumar>` din handler, pentru recuperare diferită în funcție de linia exactă unde a picat eroarea (vezi exemplul real din `references/exemplu-complet.md`).

Nu confunda aceste `Case <numar>` (linii) cu `Case <numar>` dintr-un `Select Case errNumar` obișnuit (unde `<numar>` e un cod de eroare, nu o linie) — pe cele din urmă nu le atingi la renumerotare.

Algoritm de aplicat:
1. Listează toate numerele de linie deja existente în rutină, în ordine.
2. Găsește toate referințele la numere de linie ca valoare: orice `GoTo <n>`, `Resume <n>`, și orice `Case <n>` dintr-un `Select Case` care testează `Erl` sau o variabilă alimentată din `Erl` (de regulă `lngLinia`).
3. **Strategie implicită (minimizează diff-ul):** completează doar liniile care nu au încă număr, folosind valori libere intermediare — de exemplu, dacă există deja 10 și 20, o linie nouă între ele primește 15; o linie nouă după ultimul număr existent primește următorul multiplu de 10 liber. Așa liniile deja numerotate nu-și schimbă numărul, iar toate `GoTo`/`Resume`/`Select Case` pe linie existente rămân valide automat, fără nicio rescriere.
4. **Doar dacă e nevoie de renumerotare completă** (numerotare haotică, numere duplicate, sau utilizatorul cere explicit renumerotare/curățare): construiește o mapare `număr vechi → număr nou`, apoi rescrie renumerotarea ȘI toate referințele găsite la pasul 2 conform acestei mapări, astfel încât fiecare `GoTo`, `Resume` și `Case` din `Select Case lngLinia`/`Erl` să indice în continuare exact aceeași linie logică de dinainte.
5. După orice renumerotare, verifică că nu a rămas nicio referință orfană (un `GoTo`/`Case` către un număr care nu mai există în mapare). Dacă găsești una, oprește-te și semnalează utilizatorului — nu continua silențios cu o referință coruptă.

### 3. Inserează `On Error GoTo TRATARE_ERORI`

Ca prima linie numerotată executabilă a rutinei (după `Dim`-uri).

### 4. Inserează blocul standard la finalul corpului rutinei (înainte de `End Sub`/`End Function`)

```vb
'========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
N   Exit Sub                          ' sau Exit Function, dupa tip
    
TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
N   lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
N   Select Case errNumar
    Case 0
    ' aici se pot adauga Case-uri suplimentare pentru coduri specifice - vezi Pasul 5
N   Case Else
N       ScrieEroare "Eroare in [<NumeModul>].[<NumeRutina>] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata <NumeModul> rutina <NumeRutina>"
N       RaspunsMesaj = MsgBox("[Eroare in <NumeModul>].[<NumeRutina>] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
N       If RaspunsMesaj = vbYes Then
N           Resume Next
        Else
N           GoTo TRATARE_ERORI_iesire
        End If
N   End Select
'========================== terminat tratare erori
```

- Înlocuiește `<NumeModul>`/`<NumeRutina>` cu numele reale, exact în formatul folosit deja în alte rutine tratate din același fișier.
- `DBEngine.Rollback` rămâne comentat implicit (așa e în marea majoritate a codului existent). Îl activezi (necomentat) doar dacă rutina face operații de scriere pe mai mulți pași și utilizatorul cere explicit rollback la eroare.
- Păstrează textul mesajelor identic (`"DORITI SA CONTINUATI EXECUTIA"`, formatul `"Eroare in [...]."`) — e convenția din tot restul aplicației, nu-l parafraza.
- Nu atinge alte rutine din același fișier în afara celei/celor cerute.
- **După ce aplici editarea, recitește fișierul și confirmă că textul scheletului chiar apare acolo, înainte să anunți utilizatorul că s-a făcut.** Un apel de editare care nu s-a executat de fapt (sau a eșuat silențios) nu trebuie niciodată raportat ca succes.

### 5. Analizează codul rutinei și propune `Case`-uri suplimentare de eroare (înaintea lui `Case Else`)

Nu te opri la `Case Else` generic. Citește corpul rutinei linie cu linie și identifică operațiile predispuse la coduri de eroare VBA/DAO specifice — consultă `references/coduri-eroare-comune.md` pentru tabelul complet cauză → cod → tratare recomandată. Pe scurt, verifică:

- `DLookup`/`DFirst`/`DMax`/`DSum` fără `Nz(...)` în jur → risc **94** (Invalid use of Null);
- împărțiri (`/`, `Int(x/y)`, `Mod`) cu numitor ce poate fi 0 → **11** (Division by zero);
- conversii explicite (`CDbl`, `CInt`, `CLng`, `CDate`) pe input dintr-un control/variabilă → **13** (Type mismatch);
- `With Forms![AltFormular]` / `!Control` către un formular ce s-ar putea să nu fie deschis → **2450**/**2465**;
- navigare recordset (`.MoveNext`/`.MovePrevious`/`.Edit` fără verificare EOF/BOF) → **3021** (No current record) sau **3167**;
- obiecte folosite fără `Set` prealabil → **91** (Object variable not set);
- interogări cu parametri lipsă/greșiți → **3061**/**3078**.

Pentru fiecare cod plauzibil identificat în ACEASTĂ rutină (nu generic — trebuie să existe o linie de cod în rutină care chiar poate produce acel eroare), formulează o propunere concretă: un `Case <numar>` poziționat înaintea lui `Case Else`, cu strategia de recuperare potrivită contextului (continuă cu `Resume Next`, valoare implicită prin `Nz`, mesaj specific, sau `Exit Sub`/`Function` curat) — nu doar logare generică, exact ca modelul real din `references/coduri-eroare-comune.md` / `CodBare.bas`.

**Prezintă aceste propuneri explicit utilizatorului** (listă scurtă: cod eroare → linia/operația din rutină care îl poate provoca → tratare propusă) și aplică-le doar pe cele confirmate — deciziile de recuperare la erori specifice schimbă comportamentul programului, spre deosebire de scheletul din Pasul 4, care e pur mecanic și se aplică direct.

**După ce aplici oricare `Case` confirmat, recitește fișierul (nu te baza pe ce crezi că ai scris) și confirmă explicit că acel `Case` chiar apare acolo, cu exact conținutul intenționat, înainte de a raporta succesul utilizatorului.** Un apel de editare care nu s-a executat, sau care a eșuat (ex. textul-ancoră nu s-a mai potrivit din cauza reformatării SQL/DAO după salvarea anterioară), poate să nu producă nicio eroare vizibilă pentru utilizator — dar tot nu trebuie raportat ca aplicat dacă recitirea nu confirmă modificarea.

**Fiecare `Case <numar>` adăugat trebuie însoțit de comentarii explicative, în trei părți:**

1. Pe linia lui `Case <numar>`: ce înseamnă codul de eroare (numele standard VBA/DAO, ex. "Division by zero") și de ce a fost inclus pentru această rutină anume (ce linie/operație din cod îl poate provoca).
2. În corpul `Case`-ului (deasupra sau lângă instrucțiunea de recuperare): **de ce s-a ales exact acea tratare** — de ce `Resume Next` și nu `Exit Sub`/`Function` (sau invers), de ce o valoare implicită anume, de ce un mesaj către utilizator în loc de recuperare silențioasă.
3. **Rezultatul scontat — în ce stare rămân rutina și datele după această tratare.** Nu e suficient un comentariu de tipul "ieșim din rutină" — dacă `Exit Sub`/`Function` (sau `GoTo TRATARE_ERORI_iesire`) survine **după** ce rutina a executat deja alte operații (scrieri în tabelă, `.Edit`/`.Update` pe recordset, alte apeluri care au produs efecte), comentariul trebuie să spună explicit ce s-a executat deja și ce a rămas netratat/incomplet — ex. "ATENȚIE: la acest punct linia din COMENZI a fost deja actualizată, dar STOC nu mai apucă să se recalculeze — de corectat manual" — nu doar semnala că "a apărut o problemă". Dacă `Resume Next` presupune continuarea cu o valoare implicită, precizează concret ce rezultat va produce restul rutinei cu acea valoare (ex. "totalul va ieși 0 pentru acest rând, nu blochează restul comenzii").

Nu e suficient ca utilizatorul să afle doar că a existat o eroare și că rutina a ieșit — trebuie să știe și ce a apucat să facă rutina până acolo și ce a rămas neexecutat, ca să poată decide dacă e nevoie de o corecție manuală a datelor.

Exemplu:

```vb
N   Case 94 ' Invalid use of Null - DLookup de la linia 40 poate intoarce Null daca nu gaseste inregistrarea
N       ' tratam ca stoc 0: lipsa inregistrarii inseamna produs neintrodus inca, nu eroare blocanta
N       ' rezultat: rutina continua cu stoc 0, restul calculului comenzii nu este afectat
N       sngStoc = 0
N       Resume Next
N   Case 11 ' Division by zero - impartirea la Cantitate de la linia 60 poate fi 0
N       ' Cantitate 0 e o stare valida (comanda goala) - continuam fara pretul unitar, nu oprim rutina
N       ' rezultat: pretul unitar ramane 0 pentru acest rand, restul liniilor comenzii se calculeaza normal
N       Resume Next
N   Case 3167 ' Record is deleted - inregistrarea din COMENZI a fost stearsa intre citire (linia 20) si scrierea de la linia 90
N       ' rezultat: liniile 20-80 s-au executat deja (COMANDA a fost deja marcata "in lucru" in linia 50);
N       ' scrierea finala de la linia 90 NU se mai executa - starea "in lucru" ramane setata si trebuie corectata manual
N       MsgBox "Comanda a fost stearsa de alt utilizator in timpul procesarii. Verificati manual starea comenzii!"
N       Exit Sub
```

Nu adăuga aceste comentarii pe `Case 0` sau `Case Else` din scheletul de bază (Pasul 4) — doar pe `Case`-urile de cod specific propuse aici.

## Ce NU face acest skill

- Nu copiază convenția `ErrHandler:` (fără `Erl`, fără `ScrieEroare`) găsită în module vizibil terțe/importate (biblioteci externe, controale calendar, etc.) — aceea e o convenție diferită, nu a aplicației.
- Nu inserează niciodată `Case`-uri de eroare speculative fără legătură clară cu codul din rutina respectivă.
- Nu renumerotează/atinge rutine pe care utilizatorul nu le-a cerut explicit.
- Nu creează tabela de log direct în baza de date (nu există conexiune live la Access/SQL Server din acest mediu) — generează doar interogarea SQL `CREATE TABLE` pe care utilizatorul o rulează manual în Access.
