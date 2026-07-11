# Exemple reale, verificate, ale tiparului `TRATARE_ERORI`

Toate citatele de mai jos au fost citite direct din codul exportat sub `.accdb-ai/` al aplicației Access care a originat acest skill. Servesc ca referință de format exact — nu le parafraza, copiază structura.

## 1. Șablonul canonic (comentat, generator) — `modMain.bas:115-139`

```vb
Public Function ScrieEroareSablon()
'''   On Error GoTo TRATARE_ERORI
'''
'''    {PROCEDURE_BODY}
''''========================== incep tratare erori
'''TRATARE_ERORI_iesire:
'''     'DBEngine.Rollback
'''    exit {PROCEDURE_TYPE}
'''
'''TRATARE_ERORI:
'''Dim lngLinia As Long, errNumar As Long, errDescriere As String
'''lngLinia = Erl: errNumar = err.Number: errDescriere = err.description
'''          Dim RaspunsMesaj As String
'''       Select Case errNumar
'''          Case 0
'''       Case Else
'''           ScrieEroare "Eroare in [{MODULE_NAME}].[{PROCEDURE_NAME}] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata {MODULE_NAME} rutina {PROCEDURE_NAME}"
'''           RaspunsMesaj = MsgBox("[Eroare in {MODULE_NAME}].[{PROCEDURE_NAME}] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
'''            If RaspunsMesaj = vbYes Then
'''                Resume Next
'''                Else
'''                GoTo TRATARE_ERORI_iesire
'''            End If
'''       End Select
```

Placeholderele `{MODULE_NAME}` / `{PROCEDURE_NAME}` / `{PROCEDURE_TYPE}` / `{PROCEDURE_BODY}` confirmă că acesta e tiparul intenționat, "stampilat" identic în fiecare rutină a aplicației.

## 2. Instanță completă, funcțională — `modMain.bas:85-114` (`Function EROARE`)

```vb
Function EROARE()
      Dim a As Integer
      Dim B As Integer
10       On Error GoTo TRATARE_ERORI
a = 23
20    B = a * 1000000
      '========================== incep tratare erori
TRATARE_ERORI_iesire:
           'DBEngine.Rollback
30        Exit Function

TRATARE_ERORI:
      Dim lngLinia As Long, errNumar As Long, errDescriere As String
40    lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
                Dim RaspunsMesaj As String
50           Select Case errNumar
                Case 0
60           Case Else
70               ScrieEroare "Eroare in [modMain].[EROARE] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata modMain rutina EROARE"
80               RaspunsMesaj = MsgBox("[Eroare in modMain].[EROARE] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
90                If RaspunsMesaj = vbYes Then
100                   Resume Next
110                   Else
120                   GoTo TRATARE_ERORI_iesire
130               End If
140          End Select
      '========================== terminat tratare erori

End Function
```

## 3. Funcția de logare — `modMain.bas:46-83`

```vb
Public Function ScrieEroare(modul As String, MESAJ As String, Optional _
NumarEroare As Long = 0, Optional tratata As Boolean = False, Optional _
context As String = "-")
    On Error Resume Next

    Dim formular As String
    Dim ruttina As String
    Dim pos1, pos2
    Dim LINIA

    pos1 = InStr(1, modul, "[")
    pos2 = InStr(pos1, modul, "]")
    formular = Mid(modul, pos1 + 1, pos2 - pos1 - 1)
    pos1 = InStr(pos2, modul, "[")
    pos2 = InStr(pos1, modul, "]")
    ruttina = Mid(modul, pos1 + 1, pos2 - pos1 - 1)
    LINIA = Right(modul, Len(modul) - pos2 - 7)

    CurrentDb.Execute _
            "INSERT INTO [dbo_Erori program] ( ora, modul, mesaj, rutina, utilizator, vazut, numar, tratata, context ) values (Now(),'" _
            & formular & "' ,'" & MESAJ & "', '" & ruttina & "', '" & PERSOANA_ACTIVA & _
                "',0," & NumarEroare & "," & tratata & ",'" & LINIA & "');"

    Dim ERORICAPTURATE As String
    ERORICAPTURATE = DLookup("valoare", "setari", "parametru='ERORICAPTURATE'")
    Dim PERSS As String
    If IsEmpty(PERSOANA_ACTIVA) Or PERSOANA_ACTIVA = "" Then PERSS = ""

    Call CapturaEcranGDIPlus("png", ERORICAPTURATE & "\" & Format(Now, "yyyy-mm-dd hh nn") & " " & Environ("userdomain") & "_" & Environ("username") & "-" & PERSS & " " & modul & ".png")
    ClipBoard_SetData "EROARE- " & PERSOANA_ACTIVA & " " & Now() & " " & MESAJ & "' '" & modul
    Debug.Print Now() & " " & MESAJ & "' '" & modul & " NUMAR=" & NumarEroare
End Function
```

Notă: parametrii și logica exactă (tabelă, capturi de ecran, clipboard) pot diferi de la un proiect la altul — verifică mereu definiția reală din proiectul curent înainte de a o folosi, nu presupune că e identică cu acest exemplu.

## 4. Rutină cu tratare de erori aplicată (referința inițială care a generat acest skill) — `Forms/1 1 calcul pontaj subform.form.txt:17-65` (`executate_Click`)

```vb
Private Sub executate_Click()

    Dim ART1
    Dim inte As String
10  On Error GoTo TRATARE_ERORI

20  ART1 = DLookup("NRART", "COMENZI", "NRCOMANDA='" & Me.COMANDA & "'")
30  If InStr(4, ART1, "-104-") Then MsgBox "hiflow"
40  If InStr(4, ART1, "-164-") Then MsgBox "hiflow"
50  If InStr(4, ART1, "-184-") Then MsgBox "hiflow"
60  If InStr(4, ART1, "-155-") Then MsgBox "hiflow"
If InStr(4, ART1, "-147-") Then MsgBox "hiflow"
70  With Forms![4 intrare produse realizate]
80      !start = Format(Me.INCEPUT, "hh:mm")
90      !STOP1 = Format(Me.TERMINAT, "hh:mm")
100     inte = aDURATA(!start, !STOP1)

110     !Text31 = Format(CStr(Int(inte / 60)), "0#") & ":" & Format(CStr(inte Mod 60), "0#")

120     !Text34 = val(Left(!Text31, 2)) & "," & Format(Int(1.666 * val(Right(!Text31, 2))), "0#")
130     !txtSuma = CDbl(!txtSuma) + CDbl(Nz(!Text34, 0))

140 End With
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
150 Exit Sub

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
160 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
170 Select Case errNumar
    Case 0
180 Case Else
190     ScrieEroare "Eroare in [Form_1 1 calcul pontaj subform].[executate_Click] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata Form_1 1 calcul pontaj subform rutina executate_Click"
200     RaspunsMesaj = MsgBox("[Eroare in Form_1 1 calcul pontaj subform].[executate_Click] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
210     If RaspunsMesaj = vbYes Then
220         Resume Next
230     Else
240         GoTo TRATARE_ERORI_iesire
250     End If
260 End Select
    '========================== terminat tratare erori
End Sub
```

(Notă: linia `If InStr(4, ART1, "-147-")...` nu are număr — exemplu real de numerotare incompletă, exact genul de gol pe care Pasul 2 din SKILL.md trebuie să-l completeze, cu o valoare liberă între 60 și 70, ex. `65`, nu prin renumerotarea a tot ce urmează.)

## 5. `GoTo` numeric + `Select Case lngLinia` imbricat (recuperare pe linie exactă) — `Forms/10 clisee subform planificare.form.txt:1186-1222`

Exemplu real (cod comentat în fișierul sursă, dar structural valid și reprezentativ) care arată de ce renumerotarea trebuie să actualizeze și `GoTo`-urile numerice, și `Case`-urile dintr-un `Select Case` pe linie:

```vb
Case 75
    Resume Next
Case 53
    Resume Next

Case 94
    If lngLinia = 2350 Or lngLinia = 2360 Then
        Resume Next
    Else
        ''AppActivate "CorelDraw"
        ''GoTo 2180
        ''Exit Sub
    End If

Case 3021    'No current record.
    Select Case lngLinia
    Case 1350, 2360, 3010
        Resume Next
    Case Else
        AppActivate "CorelDraw"
        GoTo 2380
        Exit Function
    End Select
Case 462    'The remote server machine does not exist or is unavailable
    MsgBox "Aplicatia Corel s-a inchis neasteptat. Reporniti formularul si reeditati cliseul."
    Exit Function
Case Else
    ScrieEroare "Eroare in [Form_10 clisee subform PLANIFICARE].[CliseuCorel_9] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata Form_10 clisee subform rutina Command9_Click"
    RaspunsMesaj = MsgBox("[Eroare in Form_10 clisee subform PLANIFICARE].[CliseuCorel_9] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
    If RaspunsMesaj = vbYes Then
        Resume Next
    Else
        GoTo TRATARE_ERORI_iesire
    End If
End Select
```

Observă: `Case 94` și `Case 3021` sunt coduri de eroare (nu se renumerotează niciodată). În schimb, `1350`, `2360`, `3010`, `2180`, `2380` sunt **numere de linie** din corpul rutinei — dacă renumerotezi rutina, acestea trebuie actualizate ca să indice în continuare aceleași linii logice.

## 6. `Case`-uri specifice de eroare înaintea lui `Case Else` — `Modules/CodBare.bas:99-114` (`Barcode_128`)

```vb
Select Case errNumar
Case 0
Case 94
    If lngLinia = 80 Then
        Exit Function
    Else
        Resume Next
    End If
Case 2427
    MsgBox "Codul de bare nu poate fi generat pentru ca raportul nu are inregistrari.ANUNTATI PROGRAMATORUL!!!"
    Exit Function
Case Else
    ScrieEroare "Eroare in [CodBare].[Barcode_128] linia " & lngLinia, ...
```

Acesta e modelul de urmat la Pasul 5 din SKILL.md: `Case`-uri specifice, cu tratare adaptată (nu doar logare + prompt), poziționate înaintea lui `Case Else`.
