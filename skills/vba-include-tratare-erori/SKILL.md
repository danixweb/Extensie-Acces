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

## Ce NU face acest skill

- Nu copiază convenția `ErrHandler:` (fără `Erl`, fără `ScrieEroare`) găsită în module vizibil terțe/importate (biblioteci externe, controale calendar, etc.) — aceea e o convenție diferită, nu a aplicației.
- Nu inserează niciodată `Case`-uri de eroare speculative fără legătură clară cu codul din rutina respectivă.
- Nu renumerotează/atinge rutine pe care utilizatorul nu le-a cerut explicit.
- Nu creează tabela de log direct în baza de date (nu există conexiune live la Access/SQL Server din acest mediu) — generează doar interogarea SQL `CREATE TABLE` pe care utilizatorul o rulează manual în Access.
