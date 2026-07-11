# Ce faci când `ScrieEroare` (sau echivalent) NU există în proiectul curent

Nu opri lucrul și nu inventa alt nume de funcție pe loc. Propune utilizatorului crearea funcției și a tabelei de log, folosind ca model structura reală, verificată, din `modMain.bas:46-83` a proiectului de referință. Cere confirmare înainte de a scrie codul funcției într-un modul (alegerea modulului țintă e a utilizatorului), și explică pas cu pas cum se creează tabela (tu nu poți crea tabela direct — nu ai acces live la motorul Access/SQL Server din acest mediu).

## Pasul A — propune schema tabelei de log

Schema implicită recomandată (identică cu cea găsită deja funcțională în acest tip de aplicație, `Tables/dbo_Erori program.table.txt`):

| Coloană | Tip Access | Note |
|---|---|---|
| `id` | AutoNumber (Long) | cheie primară |
| `ora` | Date/Time | `Now()` la fiecare inserare |
| `numar` | Long | codul de eroare (`Err.Number`) |
| `mesaj` | Text(255) | mesajul complet formatat |
| `modul` | Text(100) | numele formularului/modulului |
| `rutina` | Text(100) | numele rutinei |
| `utilizator` | Text(50) | userul curent |
| `vazut` | Yes/No | flag, implicit 0 — pentru trierea ulterioară a erorilor |
| `tratata` | Yes/No | flag trimis de apelant (`ScrieEroare`, parametrul `tratata`) |
| `context` | Memo (Long Text) | context/descriere suplimentară |

Întreabă utilizatorul dacă vrea alt nume de tabelă decât `Erori program` (evită `dbo_` ca prefix dacă tabela e nativă Access — acel prefix apare de regulă doar pe tabele legate dintr-un backend SQL Server) și dacă vrea coloane suplimentare.

## Pasul B — generează SQL-ul de creare (Access SQL), NU-l executa tu

Tu nu ai o conexiune activă la baza de date din acest mediu, deci nu poți crea tabela direct. Generează exact acest text și explică utilizatorului cum să-l ruleze:

```sql
CREATE TABLE [Erori program] (
    id AUTOINCREMENT PRIMARY KEY,
    ora DATETIME,
    numar LONG,
    mesaj TEXT(255),
    modul TEXT(100),
    rutina TEXT(100),
    utilizator TEXT(50),
    vazut YESNO,
    tratata YESNO,
    context LONGTEXT
);
```

Instrucțiuni de rulare în Access (de inclus mereu în răspunsul către utilizator):
1. În Access: fila **Create** → **Query Design** → închide dialogul de adăugare tabele fără să alegi nimic.
2. Comută la vedere SQL: click-dreapta pe query în bara de titlu / View → **SQL View**.
3. Șterge conținutul, lipește instrucțiunea `CREATE TABLE` de mai sus.
4. Rulează cu **Run** (!) — Access va afișa un avertisment că e o "data definition query" fără rezultate; confirmă. Tabela apare imediat în panoul de navigare.
5. Dacă tabela trebuie să existe pe un backend SQL Server (nu local), utilizatorul trebuie să ruleze echivalentul T-SQL acolo (nu în Access) — semnalează asta explicit dacă proiectul pare să folosească tabele legate `dbo_*` (indiciu de backend SQL Server extern).

## Pasul C — propune funcția `ScrieEroare`

Variantă **minimă** (recomandată ca implicit dacă nu se confirmă altfel — nu presupune existența unor helper-e precum `CapturaEcranGDIPlus`/`ClipBoard_SetData`/`PERSOANA_ACTIVA` dacă nu sunt găsite deja în proiect):

```vb
Public Function ScrieEroare(modul As String, MESAJ As String, Optional _
    NumarEroare As Long = 0, Optional tratata As Boolean = False, Optional _
    context As String = "-")
    On Error Resume Next

    Dim formular As String, ruttina As String, pos1, pos2, LINIA

    pos1 = InStr(1, modul, "[")
    pos2 = InStr(pos1, modul, "]")
    formular = Mid(modul, pos1 + 1, pos2 - pos1 - 1)
    pos1 = InStr(pos2, modul, "[")
    pos2 = InStr(pos1, modul, "]")
    ruttina = Mid(modul, pos1 + 1, pos2 - pos1 - 1)
    LINIA = Right(modul, Len(modul) - pos2 - 7)

    CurrentDb.Execute _
        "INSERT INTO [Erori program] ( ora, modul, mesaj, rutina, utilizator, vazut, numar, tratata, context ) values (Now(),'" _
        & formular & "' ,'" & Replace(MESAJ, "'", "''") & "', '" & ruttina & "', '" & Environ("username") & _
            "',0," & NumarEroare & "," & tratata & ",'" & Replace(context, "'", "''") & "');"

    Debug.Print Now() & " " & MESAJ & "' '" & modul & " NUMAR=" & NumarEroare
End Function
```

Note importante de menționat utilizatorului când propui asta:
- Numele tabelei din `INSERT INTO` trebuie să corespundă exact cu ce s-a creat la Pasul B (dacă utilizatorul a ales alt nume, actualizează aici).
- Adaugă escapare de apostrof (`Replace(..., "'", "''")`) pe câmpurile text libere (`MESAJ`, `context`) — spre deosebire de exemplul original din `references/exemplu-complet.md`, care nu are această protecție; e o îmbunătățire minimă de robustețe, nu o abatere de la convenție.
- `Environ("username")` e un fallback simplu pentru identificarea utilizatorului dacă proiectul nu are deja o variabilă globală de tip `PERSOANA_ACTIVA`. Dacă o astfel de variabilă globală există deja în proiect (verifică înainte), folosește-o pe aceea, nu `Environ("username")`, pentru consistență cu restul aplicației.
- Părțile opționale din exemplul original (captură de ecran via `CapturaEcranGDIPlus`, copiere în clipboard via `ClipBoard_SetData`) NU se adaugă implicit — doar dacă utilizatorul cere explicit acest comportament ȘI helper-ele respective există deja (sau sunt de asemenea create la cerere) în proiect.
- Întreabă în ce modul se adaugă funcția (un modul utilitar existent, dacă există unul similar cu `modMain.bas`, altfel un modul standard nou, ex. `modErori.bas`).

## Pasul D — abia după ce A-C sunt confirmate/create

Continuă cu pașii normali din `SKILL.md` (numerotare linii, inserare bloc `TRATARE_ERORI`, propunere `Case`-uri specifice), folosind numele real al funcției și al tabelei nou create.
