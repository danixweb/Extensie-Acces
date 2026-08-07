Option Compare Database
Option Explicit
Public ScoateBonuri As String
Public PERSOANA_ACTIVA As String
Public marca_ACTIVA As Integer
Public SARJA_ACTIVA As String
Public blnGermana As Boolean
Public strCodcomanda As String
Public GasitServerulActiv As Boolean

'''#If Win64 Then
Public Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Public Declare PtrSafe Function GetUserName Lib "advapi32.dll" Alias _
       "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As LongPtr, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
Private Declare PtrSafe Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As LongPtr
Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
        (ByVal hWnd As LongPtr, _
         ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, _
         ByVal nIndex As Long, _
         ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" _
        (ByVal hWnd As Long, _
         ByVal crKey As Long, _
         ByVal bAlpha As Byte, _
         ByVal dwFlags As Long) As Long
'''
'''#Else
'''
'''Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'''
'''    Public Declare Function GetUserName Lib "advapi32.dll" Alias _
 '''                                        "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
'''    Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'''    Private Declare Function ShellExecuteForExplore Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, lpParameters As Any, lpDirectory As Any, ByVal nShowCmd As Long) As Long
''''    Private Declare  Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
 '''                                           (ByVal hwnd As Long, _
 '''                                            ByVal nIndex As Long) As Long
''''    Private Declare  Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
 '''                                           (ByVal hwnd As Long, _
 '''                                            ByVal nIndex As Long, _
 '''                                            ByVal dwNewLong As Long) As Long
''''    Private Declare  Function SetLayeredWindowAttributes Lib "user32" _
 '''                                                        (ByVal hwnd As Long, _
 '''                                                         ByVal crKey As Long, _
 '''                                                         ByVal bAlpha As Byte, _
 '''                                                         ByVal dwFlags As Long) As Long
'''
'''#End If




'Private Const essSW_HIDE = 0
'Private Const essSW_MAXIMIZE = 3
'Private Const essSW_MINIMIZE = 6
'Private Const essSW_SHOWMAXIMIZED = 3
'Private Const essSW_SHOWMINIMIZED = 2
Private Const essSW_SHOWNORMAL = 1
'Private Const essSW_SHOWNOACTIVATE = 4
'Private Const essSW_SHOWNA = 8
'Private Const essSW_SHOWMINNOACTIVE = 7
'Private Const essSW_SHOWDEFAULT = 10
'Private Const essSW_RESTORE = 9
'Private Const essSW_SHOW = 5
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&
Private Const SE_ERR_ACCESSDENIED = 5            '  access denied
Private Const SE_ERR_ASSOCINCOMPLETE = 27
Private Const SE_ERR_DDEBUSY = 30
Private Const SE_ERR_DDEFAIL = 29
Private Const SE_ERR_DDETIMEOUT = 28
Private Const SE_ERR_DLLNOTFOUND = 32
Private Const SE_ERR_FNF = 2                     '  file not found
Private Const SE_ERR_NOASSOC = 31
Private Const SE_ERR_PNF = 3                     '  path not found
Private Const SE_ERR_OOM = 8                     '  out of memory
Private Const SE_ERR_SHARE = 26
Private Declare PtrSafe Sub sapiSleep Lib "kernel32" _
        Alias "Sleep" _
        (ByVal dwMilliseconds As Long)


'Private Const LWA_COLORKEY As Long = &H1
'Private Const LWA_ALPHA As Long = &H2
'Private Const GWL_EXSTYLE As Long = -20
'Private Const WS_EX_LAYERED As Long = &H80000



Private lngRowNumber As Long
Private colPrimaryKeys As VBA.Collection
Public strModificat As String
Private Declare PtrSafe Function SetThreadExecutionState Lib "kernel32" (ByVal esFlags As Long) As Long


Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
        ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
        ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long

Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
        ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
        ByVal samDesired As Long, ByRef phkResult As Long) As Long

Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Const HKEY_CURRENT_USER As Long = &H80000001
Const KEY_SET_VALUE As Long = &H2
Const REG_DWORD As Long = 4
Const HKEY_LOCAL_MACHINE As Long = &H80000002

'''#If VBA7 Then
'''    Private Declare PtrSafe Function BitBlt Lib "gdi32" ( _
 '''        ByVal hDestDC As LongPtr, ByVal X As Long, ByVal Y As Long, _
 '''        ByVal nWidth As Long, ByVal nHeight As Long, _
 '''        ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, _
 '''        ByVal dwRop As Long) As Long
'''
'''    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
'''    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long
'''
'''    Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
'''    Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
'''    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
'''    Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
'''    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
'''    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'''    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" ( _
 '''        ByRef picdesc As uPicDesc, _
 '''        ByRef RefIID As GUID, _
 '''        ByVal fPictureOwnsHandle As Long, _
 '''        ByRef IPic As IPicture) As Long
'''#Else
'''    Private Declare Function BitBlt Lib "gdi32" ( _
 '''        ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
 '''        ByVal nWidth As Long, ByVal nHeight As Long, _
 '''        ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
 '''        ByVal dwRop As Long) As Long
'''
'''    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
'''    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
'''
'''    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
'''    Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'''    Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'''    Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
'''    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'''    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'''         Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'''#End If

Private Type Guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4_0 As Byte
    Data4_1 As Byte
    Data4_2 As Byte
    Data4_3 As Byte
    Data4_4 As Byte
    Data4_5 As Byte
    Data4_6 As Byte
    Data4_7 As Byte

End Type

Private Type uPicDesc
    size As Long
type As Long
    hPic As LongPtr
    hPal As LongPtr
End Type


#If VBA7 Then
    Private Declare PtrSafe Function GdiplusStartup Lib "GDIPlus" ( _
            token As LongPtr, inputbuf As Any, Optional ByVal outputbuf As LongPtr = 0) As Long

    Private Declare PtrSafe Sub GdiplusShutdown Lib "GDIPlus" ( _
            ByVal token As LongPtr)

    Private Declare PtrSafe Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" ( _
            ByVal hbm As LongPtr, ByVal hPal As LongPtr, Bitmap As LongPtr) As Long

    Private Declare PtrSafe Function GdipSaveImageToFile Lib "GDIPlus" ( _
            ByVal Image As LongPtr, ByVal FileName As LongPtr, _
            clsidEncoder As Guid, ByVal encoderParams As LongPtr) As Long

    Private Declare PtrSafe Function GdipDisposeImage Lib "GDIPlus" ( _
            ByVal Image As LongPtr) As Long

    Private Declare PtrSafe Function CLSIDFromString Lib "ole32" ( _
            ByVal lpsz As LongPtr, pclsid As Guid) As Long
    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" ( _
            ByRef picdesc As uPicDesc, _
            ByRef RefIID As Guid, _
            ByVal fPictureOwnsHandle As Long, _
            ByRef IPic As IPicture) As Long

#Else
    Private Declare Function GdiplusStartup Lib "GDIPlus" ( _
                                            token As Long, inputbuf As Any, Optional ByVal outputbuf As Long = 0) As Long

    Private Declare Sub GdiplusShutdown Lib "GDIPlus" ( _
                                        ByVal token As Long)

    Private Declare Function GdipCreateBitmapFromHBITMAP Lib "GDIPlus" ( _
                                                         ByVal hbm As Long, ByVal hPal As Long, Bitmap As Long) As Long

    Private Declare Function GdipSaveImageToFile Lib "GDIPlus" ( _
                                                 ByVal Image As Long, ByVal FileName As Long, _
                                                 clsidEncoder As Guid, ByVal encoderParams As Long) As Long

    Private Declare Function GdipDisposeImage Lib "GDIPlus" ( _
                                              ByVal Image As Long) As Long

    Private Declare Function CLSIDFromString Lib "ole32" ( _
                                             ByVal lpsz As Long, pclsid As Guid) As Long

    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" ( _
                                                      ByRef picdesc As uPicDesc, _
                                                      ByRef RefIID As Guid, _
                                                      ByVal fPictureOwnsHandle As Long, _
                                                      ByRef IPic As IPicture) As Long

#End If

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As LongPtr
    SuppressExternalCodecs As LongPtr
End Type

'''Private Type GUID
'''    Data1 As Long
'''    Data2 As Integer
'''    Data3 As Integer
'''    Data4_0 As Byte
'''    Data4_1 As Byte
'''    Data4_2 As Byte
'''    Data4_3 As Byte
'''    Data4_4 As Byte
'''    Data4_5 As Byte
'''    Data4_6 As Byte
'''    Data4_7 As Byte
'''End Type


#If VBA7 Then
    Private Declare PtrSafe Function BitBlt Lib "gdi32" ( _
            ByVal hDestDC As LongPtr, ByVal X As Long, ByVal Y As Long, _
            ByVal nWidth As Long, ByVal nHeight As Long, _
            ByVal hSrcDC As LongPtr, ByVal xSrc As Long, ByVal ySrc As Long, _
            ByVal dwRop As Long) As Long

    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hDC As LongPtr) As Long

    Private Declare PtrSafe Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As LongPtr) As LongPtr
    Private Declare PtrSafe Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As LongPtr, ByVal nWidth As Long, ByVal nHeight As Long) As LongPtr
    Private Declare PtrSafe Function SelectObject Lib "gdi32" (ByVal hDC As LongPtr, ByVal hObject As LongPtr) As LongPtr
    Private Declare PtrSafe Function DeleteDC Lib "gdi32" (ByVal hDC As LongPtr) As Long
    Private Declare PtrSafe Function DeleteObject Lib "gdi32" (ByVal hObject As LongPtr) As Long
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Function BitBlt Lib "gdi32" ( _
                                    ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
                                    ByVal nWidth As Long, ByVal nHeight As Long, _
                                    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
                                    ByVal dwRop As Long) As Long

    Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long

    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
    Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
    Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
'''#If Win64 Then
Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias _
        "SHGetPathFromIDListA" (ByVal pidl As Long, _
                                ByVal pszPath As String) As Long

Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias _
        "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) _
        As Long

'''
'''#Else
'''Private Declare PtrSafe Function SHGetPathFromIDList Lib "shell32.dll" Alias _
 '''            "SHGetPathFromIDListA" (ByVal pidl As Long, _
 '''            ByVal pszPath As String) As Long
'''
'''Private Declare PtrSafe Function SHBrowseForFolder Lib "shell32.dll" Alias _
 '''            "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) _
 '''            As Long
'''
'''
'''#End If


Private Const BIF_RETURNONLYFSDIRS = &H1

 Function testPaste()
 Debug.Print "A doua zi de Paste " & Format(DateAdd("d", PasteOrtodox(2026), 1), "dd.mm.yy")
 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL(DateAdd("d", PasteOrtodox(2026), 1)) & " where explicatia='A doua zi de Paste'"
 Debug.Print "A doua zi de Rusalii " & Format(DateAdd("d", PasteOrtodox(2026), 50), " ddd dd.mm.yy")
  CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL(DateAdd("d", PasteOrtodox(2026), 50)) & " where explicatia='A doua zi de Rusalii'"
 End Function
Public Function PersoanaActiva() As Boolean
10  If PERSOANA_ACTIVA = "" Or PERSOANA_ACTIVA = "nelogat" Then
20      MesajMod "Trebuie sa va logati inainte de a face operari.", "", "OK"
30      PersoanaActiva = False
40  Else
50      PersoanaActiva = True
60  End If
End Function


Function PasteOrtodox(an As Integer) As Date
    '===Start Generat AI===
    ' Aceasta functie calculeaza data Pastelui Ortodox pentru anul specificat.
    ' Utilizeaza algoritmul Meeus-Jones-Butcher pentru calculul datei iuliene a Pastelui,
    ' la care adauga corectia de +13 zile (valabila pentru secolele XX si XXI) pentru transpunerea in calendarul gregorian.
    ' De asemenea, functia actualizeaza automat in tabela [zile libere] datele pentru sarbatorile dependente:
    ' - A doua zi de Paste (+1 zi)
    ' - Vinerea Pastelor (-2 zile)
    ' - A doua zi de Rusalii (+50 zile)
    ' Precum si restul sarbatorilor cu data fixa din anul respectiv.
    '===Final Generat AI===
    Dim a As Integer, B As Integer, C As Integer
    Dim D As Integer, E As Integer
    Dim zi As Integer, Luna As Integer

10  a = an Mod 4
20  B = an Mod 7
30  C = an Mod 19
40  D = (19 * C + 15) Mod 30
50  E = (2 * a + 4 * B - D + 34) Mod 7

60  Luna = Int((D + E + 114) / 31)
70  zi = ((D + E + 114) Mod 31) + 1

    ' Data Pastelui �n calendar iulian ? +13 zile (sec. XX�XXI)
80  PasteOrtodox = DateSerial(an, Luna, zi) + 13
90  CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL(DateAdd("d", PasteOrtodox, 1)) & " where explicatia='A doua zi de Paste'"
100 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL(DateAdd("d", PasteOrtodox, -2)) & " where explicatia='Vinerea Pastelor'"
110 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL(DateAdd("d", PasteOrtodox, 50)) & " where explicatia='A doua zi de Rusalii'"


120 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL("01.01." & an) & " where explicatia='Prima zi de an nou'"
130 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL("02.01." & an) & " where explicatia='A doua zi de an nou'"
140 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL("06.01." & an) & " where explicatia='Boboteaza'"
150 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL("07.01." & an) & " where explicatia='Sfantul Ioan'"
160 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL("24.01." & an) & " where explicatia='Ziua Unirii'"
170 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL("01.05." & an) & " where explicatia='1 MAI'"
180 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL("01.06." & an) & " where explicatia='Ziua copilului'"
190 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL("15.08." & an) & " where explicatia='Adormirea Maici Domnului'"
200 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL("30.11." & an) & " where explicatia='Sf. Andrei'"
210 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL("01.12." & an) & " where explicatia='ZIUA NATIONALA'"
220 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL("25.12." & an) & " where explicatia='Prima zi de Craciun'"
230 CurrentDb.Execute "update [zile libere] set data=" & DataPentruSQL("26.12." & an) & " where explicatia='Prima zi de Craciun'"

End Function
Public Sub StartComplex()
'''DoCmd.Echo False
'''Application.Echo False
'''DoCmd.SetWarnings False
'Application.SetOption "AutoCompact", False
End Sub

Public Sub StopComplex()
'''DoCmd.Echo True
'''Application.Echo True
'''DoCmd.SetWarnings True
'Application.SetOption "AutoCompact", True
End Sub
Public Function VerificProdusInSituatieComenzi() As String
    Dim idComanda As Long, setiparesc As String, ProdusNomenclator As String
    Dim ComandaGresita, nrArticolComanda
    Dim SirEroare As String
    Dim MESAJ As String
    Dim N As Long
    '    Dim IesiLa3 As Integer
    '    Dim comandaInainte, comandaDupa
10  On Error GoTo TRATARE_ERORI    '

20  If LCase(DLookup("valoare", "setari", "parametru=' CorectezSitComenzi'")) <> "da" Then Exit Function

30  MESAJ = "COMENZI CORECTATE"
    'GoTo 40
40  ComandaGresita = Nz(DLookup("COMANDA", "qry_Situatie comenzi produs gresit"), "")
50  If ComandaGresita = "" Then
60      ComandaGresita = Nz(DLookup("COMANDA", "qry_Situatie comenzi produs DUPLICAT gresit"), "")
70  End If

80  If ComandaGresita <> "" Then
90      N = N + 1
100     nrArticolComanda = DLookup("NRART", "comenzi", "nr_comanda_interna=" & ComandaGresita & "")
110     ProdusNomenclator = DLookup("PRODUS", "nomenclator articole", "NRART='" & nrArticolComanda & "'")
120     MESAJ = MESAJ & vbCrLf & N & " = " & ComandaGresita
130     CurrentDb.Execute "update [dbo_Situatie comenzi] set produs='" & ProdusNomenclator & "' where comanda=" & ComandaGresita & " and exportate=1"
140     CurrentDb.Execute "update [dbo_Situatie comenzi] set produs='" & "S" & Right(ProdusNomenclator, Len(ProdusNomenclator) - 1) & "' where comanda=" & ComandaGresita & " and exportate=0"
150 End If

160 ComandaGresita = Nz(DLookup("COMANDA", "qry_Situatie comenzi produs gresit"), "")
170 If ComandaGresita = "" Then
180     ComandaGresita = Nz(DLookup("COMANDA", "qry_Situatie comenzi produs DUPLICAT gresit"), "")
190 End If

200 If N > 18 Then

210     ScrieEroare "CORECTIE", "Verifica situatie " & ComandaGresita, , , "Situatie comenzi"
220     GoTo 250
230 End If

240 If ComandaGresita <> "" And N < 20 Then GoTo 40
    'GoTo 250
250 If MESAJ <> "COMENZI CORECTATE" Then
260     VerificProdusInSituatieComenzi = MESAJ

270     ScrieEroare "CORECTIE", MESAJ, , , "Situatie comenzi"
280 Else
290     VerificProdusInSituatieComenzi = "Totul este in regula"
300 End If
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
310 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
320 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
330 Select Case errNumar
    Case 0
340 Case Else
350     ScrieEroare "Eroare in [functii].[VerificProdusInSituatieComenzi] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
360     RaspunsMesaj = MsgBox("Eroare in [functii].[VerificProdusInSituatieComenzi] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "Eroare in [functii].[VerificProdusInSituatieComenzi] linia " & lngLinia & " cu numarul " & Errnumar & vbCrLf & Errdescriere"
370     If RaspunsMesaj = vbYes Then
380         Resume Next
390     Else
400         GoTo TRATARE_ERORI_iesire
410     End If
420 End Select
    '========================== terminat tratare erori

End Function


Public Sub TimpProdus(Articol As String)


40  CurrentDb.Execute "drop table productivitati"

50  CurrentDb.QueryDefs("qryAnalizaProductivitate").sql = "SELECT TOP 100 CDate(Format([INCEPUT],'dd\.mm\.yyyy')) AS ZIUA, [23 PONTARI].PERSOANA, [23 PONTARI].Comanda, [NOMENCLATOR ARTICOLE].NRART," & _
     " [NOMENCLATOR ARTICOLE].VolumTimp, [NOMENCLATOR ARTICOLE].precizietimp, [NOMENCLATOR ARTICOLE].datastabiliretimp, [23 PONTARI].INCEPUT, IIf((CDate(#12/30/1899 10:0:0#)>CDate(Format([23 PONTARI].INCEPUT,'hh:nn:ss')) And CDate(#12/30/1899 10:20:0#)<CDate(Format([23 PONTARI].terminat,'hh:nn:ss'))),20,0) AS pauza10, IIf((CDate(#12/30/1899 13:0:0#)>CDate(Format([23 PONTARI].[INCEPUT],'hh:nn:ss')) And CDate(#12/30/1899 13:10:0#)<CDate(Format([23 PONTARI].[terminat],'hh:nn:ss'))),10,0) AS pauza13, DateDiff('n',[23 PONTARI].[INCEPUT],[23 PONTARI].[TERMINAT]) AS durata, Val((DateDiff('n',[23 PONTARI].[INCEPUT],[23 PONTARI].[TERMINAT])-IIf((CDate(#12/30/1899 10:0:0#)>CDate(Format([23 PONTARI].[INCEPUT],'hh:nn:ss')) And CDate(#12/30/1899 10:20:0#)<CDate(Format([23 PONTARI].[terminat],'hh:nn:ss'))),20,0)-IIf((CDate(#12/30/1899 13:0:0#)>CDate(Format([23 PONTARI].[INCEPUT],'hh:nn:ss')) And CDate(#12/30/1899 13:10:0#)<CDate(Format([23 PONTARI].[terminat],'hh:nn:ss'))),10,0))) AS efectiv, " & _
                                                        " [23 PONTARI].TERMINAT, [23 PONTARI].FAZA, [23 PONTARI].BUCATI, CDbl(Format(([bucati]*[volumtimp]*100)/(60*Val((DateDiff('n',[23 PONTARI].[INCEPUT],[23 PONTARI].[TERMINAT])-IIf((CDate(#12/30/1899 10:0:0#)>CDate(Format([23 PONTARI].[INCEPUT],'hh:nn:ss')) And CDate(#12/30/1899 10:20:0#)<CDate(Format([23 PONTARI].[terminat],'hh:nn:ss'))),20,0)-IIf((CDate(#12/30/1899 13:0:0#)>CDate(Format([23 PONTARI].[INCEPUT],'hh:nn:ss')) And CDate(#12/30/1899 13:10:0#)<CDate(Format([23 PONTARI].[terminat],'hh:nn:ss'))),10,0)))),'0.00')) AS [%] INTO productivitati" & _
                                                        " FROM ([23 PONTARI] LEFT JOIN COMENZI ON [23 PONTARI].Comanda = COMENZI.NRCOMANDA) LEFT JOIN [NOMENCLATOR ARTICOLE] ON COMENZI.NRART = [NOMENCLATOR ARTICOLE].NRART" & _
                                                        " WHERE ( [NOMENCLATOR ARTICOLE].NRART='" & Articol & "' AND (([NOMENCLATOR ARTICOLE].VolumTimp) Is Not Null) AND (([23 PONTARI].FAZA) Like 'p*') and ([23 PONTARI].BUCATI>0))" & _
                                                        " ORDER BY [23 PONTARI].INCEPUT DESC;"
60  CurrentDb.QueryDefs("qryAnalizaProductivitateGrafic").sql = Replace(CurrentDb.QueryDefs("qryAnalizaProductivitate").sql, "into productivitati", "")
70    CurrentDb.QueryDefs("qryAnalizaProductivitate").Execute
80  CurrentDb.Execute "delete * from productivitati where [%]=0 or efectiv=0"
DoCmd.OpenForm "analizaproductivitate"


90  Forms("analizaproductivitate").Form.RecordSource = "SELECT productivitati.ZIUA, productivitati.PERSOANA, productivitati.Comanda, productivitati.nrart, productivitati.VolumTimp, productivitati.precizietimp, productivitati.datastabiliretimp, productivitati.FAZA, Avg(productivitati.[%]) AS [AvgOf%], Sum(productivitati.efectiv) AS SumOfefectiv, Sum(productivitati.BUCATI) AS SumOfBUCATI, CInt((100/60)*Sum([productivitati].[BUCATI])*[productivitati].[VolumTimp]/Sum([productivitati].[efectiv])) AS productivitate FROM productivitati GROUP BY productivitati.ZIUA, productivitati.PERSOANA, productivitati.Comanda, productivitati.nrart, productivitati.VolumTimp, productivitati.precizietimp, productivitati.datastabiliretimp, productivitati.FAZA;"
100 Forms("analizaproductivitate").Form.Requery

    ' CDbl(Format(([bucati]*[volumtimp]*100)/(60*Val((DateDiff('n',[23 PONTARI].[INCEPUT],[23 PONTARI].[TERMINAT])-IIf((CDate(#12/30/1899 10:0:0#)>CDate(Format([23 PONTARI].[INCEPUT],'hh:nn:ss')) And CDate(#12/30/1899 10:20:0#)<CDate(Format([23 PONTARI].[terminat],'hh:nn:ss'))),20,0)-IIf((CDate(#12/30/1899 13:0:0#)>CDate(Format([23 PONTARI].[INCEPUT],'hh:nn:ss')) And CDate(#12/30/1899 13:10:0#)<CDate(Format([23 PONTARI].[terminat],'hh:nn:ss'))),10,0)))),'0.00'))
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
120 Exit Sub

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
130 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
140 Select Case errNumar

    Case 0
    
       
    
150 Case 3464
160     MsgBox "Nu s-au putut selecta inregistrari.Verificati corectitudinea pontarilor pe articolul curent."
170     GoTo TRATARE_ERORI_iesire
180 Case 6    'LINIA 60 Overflow

190             Resume Next
200             Case 11
210             Resume Next
220             Case 2580
230             Resume Next
240             Case 3376
250             MsgBox "Nu s-au putut selecta inregistrari"
260              GoTo TRATARE_ERORI_iesire
270          Case Else
280              ScrieEroare "Eroare in [Form_AnalizaProductivitate].[TimpProdus] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata Form_AnalizaProductivitate rutina Articol_Click"
290              RaspunsMesaj = MsgBox("[Eroare in Form_AnalizaProductivitate].[TimpProdus] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
                 'Executasiraspunde="INFORMATIE "[Eroare in Form_AnalizaProductivitate].[TimpProdus] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
300            If RaspunsMesaj = vbYes Then
310       Resume Next
320               Else
330       GoTo TRATARE_ERORI_iesire
340      End If
350          End Select
      '========================== terminat tratare erori
End Sub

Public Function AreDreptDeOperare(operatie As String) As Boolean
    Dim Opereaza As String
    Dim pos, a
    If PERSOANA_ACTIVA = "" Or PERSOANA_ACTIVA = "nelogat" Then
        MsgBox "Logati-va inainte de a face operari." & vbCrLf & vbCrLf & "Apasati tasta ESC pentru a putea continua."
        AreDreptDeOperare = False
        Exit Function
    End If
    AreDreptDeOperare = False
    Opereaza = DLookup("VALOARE", "SETARI", "PARAMETRU='" & operatie & "'")
    pos = InStr(1, PERSOANA_ACTIVA, " ")
    a = Mid(PERSOANA_ACTIVA, pos + 1, Len(PERSOANA_ACTIVA) - pos)
    a = Left(a, 2)
    a = Left(PERSOANA_ACTIVA, 2) & " " & a
    If InStr(1, Opereaza, UCase(a)) Then
        AreDreptDeOperare = True
        Else
                MsgBox "Nu aveti dreptul sa modificati comenzi." & vbCrLf & vbCrLf & "Apasati tasta ESC pentru a putea continua."

    End If
End Function

Function testmesaj()
    Dim ret
    TransmiteMesajModFormular = True
    ret = MesajMod("test6", , "da", "nu", , , 1000)
End Function

Public Function AnalizaAI(CuFisier As Boolean) As String
'tipuri de produse in transa curenta
    Dim SIR As String
    Dim N As Long
    Dim hFile As Integer
    Dim numeFisier As String
    Dim rec As DAO.Recordset
    Dim recPers As DAO.Recordset
    Dim NRART As String
    Dim C As Integer
    Dim P As Integer
    Dim DeExecutat, RamasDeExecutatPeTransa, cantpetransa, SS
    Dim VARRETURN As Variant
    Dim SirEroare As String
10  On Error GoTo TRATARE_ERORI
20  DoCmd.Hourglass True
30  VARRETURN = SysCmd(acSysCmdSetStatus, "Creez fisier")
    If CuFisier Then
40      If Dir("C:\Tempv\", vbDirectory) = "" Then
50          MkDir "C:\Tempv\"
60      End If
70      Close #1
80      numeFisier = "C:\Tempv\Planificare Comenzi AI.txt"
90      If Dir(numeFisier) <> "" Then Kill numeFisier

100     hFile = FreeFile
    End If
110 SIR = DLookup("TEXTACTIUNE", "DE FACUT", "IDACTIUNE=51")

120 If CuFisier Then Open numeFisier For Output As #hFile
    '''120 Print #hFile, SIR & vbCrLf
130 VARRETURN = SysCmd(acSysCmdSetStatus, "Colectez prioritati ambalatori")
140 If CuFisier Then Print #hFile, "DATA CURENTA:" & Format(Date, "yyyy-mm-dd")
    AnalizaAI = AnalizaAI & "DATA CURENTA:" & Format(Date, "yyyy-mm-dd")
150 If CuFisier Then Print #hFile, ""
    AnalizaAI = AnalizaAI & ""
160 If CuFisier Then Print #hFile, "A) LUCRATORI DISPONIBILI SI ORDINEA LA EXECUTAREA PRODUSELOR"
    AnalizaAI = AnalizaAI & "A) LUCRATORI DISPONIBILI SI ORDINEA LA EXECUTAREA PRODUSELOR"
170 If CuFisier Then Print #hFile, ""
    AnalizaAI = AnalizaAI & ""
180 If CuFisier Then Close #hFile

190 If CuFisier Then Open numeFisier For Append As #hFile

200 SIR = "TRANSFORM Sum(Int(([PRODUSE REALIZATE].[NUMAR_ORE])*1000/DateDiff('d',IIf([PERSONAL].[Dataangajarii]>#1/1/" & Right(year(Date), 2) & "#,[PERSONAL].[Dataangajarii],#1/1/" & Right(year(Date), 2) & "#),Date()))) AS Produs" & _
          " SELECT [PRODUSE REALIZATE].PERSOANA AS AMBALATOR" & _
          " FROM ([PRODUSE REALIZATE] LEFT JOIN PERSONAL ON [PRODUSE REALIZATE].PERSOANA = PERSONAL.[Numesiprenume]) LEFT JOIN [7 NORME] ON [PRODUSE REALIZATE].TIP_PRODUS = [7 NORME].PRODUS" & _
          " WHERE ((([PRODUSE REALIZATE].TIP_PRODUS) Not In ('CO','BO','RE','M')) AND (([PRODUSE REALIZATE].PERSOANA) Not Like '*psk') AND (([PRODUSE REALIZATE].DATA)>CDate('01.01." & Right(year(Date), 2) & "')))" & _
          " GROUP BY [PRODUSE REALIZATE].PERSOANA" & _
          " PIVOT [PRODUSE REALIZATE].TIP_PRODUS;"
    '190   Debug.Print SIR
210 Set rec = CurrentDb.OpenRecordset(SIR, dbOpenDynaset, dbSeeChanges)
220 If Not rec.EOF Then
230     rec.MoveLast
240     rec.MoveFirst
        Dim antet As String, continut As String, Activ As Boolean
250     For C = 0 To rec.Fields.Count - 1
260         If antet = "" Then
270             antet = rec.Fields(C).name
280         Else
290             antet = antet & "," & rec.Fields(C).name
300         End If
310     Next
320     If CuFisier Then Print #hFile, antet
        AnalizaAI = AnalizaAI & antet
330     For P = 1 To rec.RecordCount
340         continut = ""
350         If Nz(DLookup("PRODUCTIV", "PERSONAL", "NUMESIPRENUME='" & rec.Fields(0).Value & "'"), 0) = True Then
360             For C = 0 To rec.Fields.Count - 1
370                 If continut = "" Then
380                     continut = Nz(rec.Fields(C).Value, 0)
390                 Else
400                     continut = continut & "," & Nz(rec.Fields(C).Value, 0)
410                 End If
420             Next C
430             If CuFisier Then Print #hFile, continut
                AnalizaAI = AnalizaAI & continut
440         End If
450         rec.MoveNext
460     Next
470 End If
480 If CuFisier Then Close #hFile
490 rec.Close
500 Set rec = Nothing








    'COMENZI PE CARE SE LUCREAZA ACUM
510 VARRETURN = SysCmd(acSysCmdSetStatus, "Colectez comenzi in lucru acum")
520 If CuFisier Then hFile = FreeFile
530 If CuFisier Then Open numeFisier For Append As #hFile
540 SIR = "SELECT [23 PONTARI].Comanda, [23 PONTARI].INCEPUT, [23 PONTARI].PERSOANA AS AMBALATOR,[23 PONTARI].faza as produs" & _
          " FROM [23 PONTARI]" & _
          " WHERE ((([23 PONTARI].TERMINAT) Is Null) AND ((Format([23 PONTARI].[INCEPUT],'dd\.mm\.yyyy'))=Format(Now(),'dd\.mm\.yyyy')))" & _
          " ORDER BY [23 PONTARI].Comanda, [23 PONTARI].INCEPUT;"
    'Debug.Print SIR
550 Set rec = CurrentDb.OpenRecordset(SIR, dbOpenDynaset, dbSeeChanges)
560 If Not rec.EOF Then
570     rec.MoveLast
580     rec.MoveFirst
590     If CuFisier Then Print #hFile, ""
        AnalizaAI = AnalizaAI & ""
600     If CuFisier Then Print #hFile, "B)COMENZI IN LUCRU ACUM " & Format(rec.Fields("INCEPUT"), "yyyy-mm-dd") & ":"
        AnalizaAI = AnalizaAI & "B)COMENZI IN LUCRU ACUM " & Format(rec.Fields("INCEPUT"), "yyyy-mm-dd") & ":"
610     If CuFisier Then Print #hFile, ""
        AnalizaAI = AnalizaAI & ""
620     antet = ""
630     For C = 0 To rec.Fields.Count - 1
640         If antet = "" Then
650             antet = rec.Fields(C).name
660         Else
670             antet = antet & "," & rec.Fields(C).name
680         End If
690     Next C
700     If CuFisier Then Print #hFile, antet
        AnalizaAI = AnalizaAI & antet


710     For P = 1 To rec.RecordCount
720         If CuFisier Then Print #hFile, rec.Fields(0) & "," & Format(rec.Fields(1), "yyyy-mm-dd hh:mm:ss") & "," & rec.Fields(2) & "," & rec.Fields(3)
            AnalizaAI = AnalizaAI & rec.Fields(0) & "," & Format(rec.Fields(1), "yyyy-mm-dd hh:mm:ss") & "," & rec.Fields(2) & "," & rec.Fields(3)
730         rec.MoveNext
740     Next
750     If CuFisier Then Print #hFile, ""
        AnalizaAI = AnalizaAI & ""
760 End If
770 If CuFisier Then Close #hFile
780 rec.Close
790 Set rec = Nothing






    'populare tabel manevra cu comenzi
800 VARRETURN = SysCmd(acSysCmdSetStatus, "Colectez comenzi ramase de executat")
810 CurrentDb.TableDefs.Delete "AI_Comenzi"

820 SIR = "SELECT COMENZI.NRCOMANDA AS [Numar comanda], Sarje.sarja AS Sarja, [dbo_Situatie comenzi].Comanda AS [Nr intern], [dbo_Situatie comenzi].Produs, COMENZI.CANT AS [Cantitate comandata], 0 AS executate, 0 AS [Ramase de executat], 0 AS [Timp necesar pentru un produs (in secunde)], 0 AS [Timp care mai trebuie lucrat pe comanda(in secunde)], [dbo_Situatie comenzi].id AS [Id Situatie], Sarje.prioritate AS Prioritate, Sarje.cant AS [Cantitate pe transa], format(Sarje.data_livrare,'yyyy-mm-dd') AS [Data de livrare], Sarje.cliseu, Sum(INTRARI.cant) AS predate INTO AI_Comenzi" & _
          " FROM (((COMENZI LEFT JOIN Sarje ON COMENZI.NRCOMANDA = Sarje.Comanda) LEFT JOIN [Nomenclator articole] ON COMENZI.NRART = [Nomenclator articole].NrArt) RIGHT JOIN [dbo_Situatie comenzi] ON COMENZI.NR_COMANDA_INTERNA = [dbo_Situatie comenzi].Comanda) LEFT JOIN INTRARI ON (Sarje.sarja = INTRARI.sarja) AND (Sarje.Comanda = INTRARI.nume)" & _
          " GROUP BY COMENZI.NRCOMANDA, Sarje.sarja, [dbo_Situatie comenzi].Comanda, [dbo_Situatie comenzi].Produs, COMENZI.CANT, 0, [dbo_Situatie comenzi].id, Sarje.prioritate, Sarje.cant, Sarje.data_livrare, Sarje.cliseu, 0, 0, 0, [dbo_Situatie comenzi].Comanda, [COMENZI].[CANT]-IIf([COMENZI].[PRODUSE],[COMENZI].[PRODUSE],0), COMENZI.STADIU" & _
          " HAVING (((Sarje.prioritate)>0) AND ((Sarje.cliseu)<>'9999') AND (([COMENZI].[CANT]-IIf([COMENZI].[PRODUSE],[COMENZI].[PRODUSE],0))>0) AND ((COMENZI.STADIU)='IN LUCRU')) OR (((Sarje.prioritate)>0) AND ((Sarje.cliseu) Is Null) AND (([COMENZI].[CANT]-IIf([COMENZI].[PRODUSE],[COMENZI].[PRODUSE],0))>0) AND ((COMENZI.STADIU)='IN LUCRU'))" & _
          " ORDER BY Sarje.data_livrare, [dbo_Situatie comenzi].Comanda;"

    '    Debug.Print SIR

830 CurrentDb.Execute SIR

840 CurrentDb.Execute "update [AI_Comenzi] set predate=0 where predate is null"
    '''850 CurrentDb.Execute "delete * from [AI_Comenzi] where predate>=[cantitate pe transa]"

860 SIR = "select * from [AI_Comenzi]"
870 Set rec = CurrentDb.OpenRecordset(SIR, dbOpenDynaset, dbSeeChanges)
880 If Not rec.EOF Then
890     rec.MoveLast
900     rec.MoveFirst

910     For N = 1 To rec.RecordCount

            ''If rec.Fields("[numar comanda]") = "CCH-463-15A" Then Stop
920         rec.Edit
            '''            If rec.Fields("[numar comanda]") = "KAN-1869-5A" Then Stop
930         Select Case UCase(Left(rec.Fields("produs"), 1))
            Case "P"
940             Select Case UCase(rec.Fields("produs"))
                Case "PPT", "PCT"
                    Dim FACUTE
950                 FACUTE = Nz(DLookup("produse", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "A'"), 0)
960                 If FACUTE > 0 Then    ' daca sunt tubulete de ambalat
970                     RamasDeExecutatPeTransa = rec.Fields("[cantitate comandata]") - Nz(DLookup("produse", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "'"), 0)
980                     If RamasDeExecutatPeTransa > 0 Then
990                         cantpetransa = Nz(DSum("cant", "sarje", "prioritate=0 and sarja='" & rec.Fields("sarja") & "' and comanda='" & rec.Fields("[numar comanda]") & "'"), 0)
1000                        RamasDeExecutatPeTransa = RamasDeExecutatPeTransa - cantpetransa
1010                        For SS = 1 To rec.Fields("prioritate")
1020                            cantpetransa = DLookup("cant", "sarje", "prioritate=" & SS & " and sarja='" & rec.Fields("sarja") & "' and comanda='" & rec.Fields("[numar comanda]") & "'")
1030                            RamasDeExecutatPeTransa = RamasDeExecutatPeTransa - cantpetransa
1040                            If RamasDeExecutatPeTransa < 0 Then
1050                                RamasDeExecutatPeTransa = Abs(RamasDeExecutatPeTransa)
1060                                rec.Fields("[ramase de executat]") = RamasDeExecutatPeTransa
                                    rec.Fields("[executatE]") = rec.Fields("[cantitate comandata]") - RamasDeExecutatPeTransa
1070                                RamasDeExecutatPeTransa = 0
1080                            Else
1090                                rec.Fields("[ramase de executat]") = cantpetransa
                                    rec.Fields("[executatE]") = rec.Fields("[cantitate comandata]") - cantpetransa
1100                            End If
1110                        Next
1120                    Else
1130                        rec.Fields("[ramase de executat]") = -1
                            rec.Fields("[executatE]") = rec.Fields("[cantitate comandata]")
1140                    End If
1150                    If rec.Fields("[ramase de executat]") > rec.Fields("[cantitate pe transa]") Then rec.Fields("[ramase de executat]") = rec.Fields("[cantitate pe transa]")
1160
1170                Else
1180                    rec.Fields("[ramase de executat]") = -1
                        rec.Fields("[executatE]") = rec.Fields("[cantitate comandata]")
1190                End If
1200            Case Else

1210                RamasDeExecutatPeTransa = rec.Fields("[cantitate comandata]") - Nz(DLookup("produse", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "'"), 0)
1220                If RamasDeExecutatPeTransa > Nz(DLookup("cant", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "'"), 0) Then
1230                    RamasDeExecutatPeTransa = 0    'DLookup("cant", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "'")
1240                End If

1250                If RamasDeExecutatPeTransa > 0 Then
1260                    cantpetransa = Nz(DSum("cant", "sarje", "prioritate=0 and sarja='" & rec.Fields("sarja") & "' and comanda='" & rec.Fields("[numar comanda]") & "'"), 0)
1270                    RamasDeExecutatPeTransa = RamasDeExecutatPeTransa - cantpetransa
1280                    For SS = 1 To rec.Fields("prioritate")
1290                        cantpetransa = DLookup("cant", "sarje", "prioritate=" & SS & " and sarja='" & rec.Fields("sarja") & "' and comanda='" & rec.Fields("[numar comanda]") & "'")
1300                        RamasDeExecutatPeTransa = RamasDeExecutatPeTransa - cantpetransa
1310                        If RamasDeExecutatPeTransa < 0 Then
1320                            RamasDeExecutatPeTransa = Abs(RamasDeExecutatPeTransa)
1330                            rec.Fields("[ramase de executat]") = RamasDeExecutatPeTransa
                                rec.Fields("[executatE]") = rec.Fields("[cantitate comandata]") - RamasDeExecutatPeTransa
1340                            RamasDeExecutatPeTransa = 0
1350                        Else
1360                            rec.Fields("[ramase de executat]") = cantpetransa
                                rec.Fields("[executatE]") = rec.Fields("[cantitate comandata]") - cantpetransa
1370                        End If
1380                    Next
1390                Else
1400                    rec.Fields("[ramase de executat]") = -1
                        rec.Fields("[executatE]") = rec.Fields("[cantitate comandata]")
1410                End If
1420                If rec.Fields("[ramase de executat]") > rec.Fields("[cantitate pe transa]") Then rec.Fields("[ramase de executat]") = rec.Fields("[cantitate pe transa]")
1430            End Select
                '''
1440            NRART = DLookup("nrart", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "'")
                ''1150            rec.Fields("[timp necesar pentru un produs (in secunde)]") = DLookup("volumtimp", "[nomenclator articole]", "nrart='" & NRART & "'")
                ''1160            rec.Fields("[timp care mai trebuie lucrat pe comanda(in secunde)]") = rec.Fields("[ramase de executat]") * rec.Fields("[timp necesar pentru un produs (in secunde)]")
1450            rec.Fields("cliseu") = ""

1460        Case "S"
                '''If rec.Fields("[numar comanda]") = "ARK-108-13" Then Stop
1470            RamasDeExecutatPeTransa = Nz(DLookup("semiproduse", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "'"), 0)

1480            If RamasDeExecutatPeTransa >= Nz(DLookup("cant", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "'"), 0) Then
1490                RamasDeExecutatPeTransa = 0    'DLookup("cant", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "'")
                    rec.Fields("[executatE]") = Nz(DLookup("cant", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "'"), 0)
1500            End If
1510            If RamasDeExecutatPeTransa > 0 Then    ' sunt facute
1520                cantpetransa = Nz(DSum("cant", "sarje", "prioritate=0 and sarja='" & rec.Fields("sarja") & "' and comanda='" & rec.Fields("[numar comanda]") & "'"), 0)
1530                RamasDeExecutatPeTransa = RamasDeExecutatPeTransa - cantpetransa
1540                For SS = 1 To rec.Fields("prioritate")
1550                    cantpetransa = DLookup("cant", "sarje", "prioritate=" & SS & " and sarja='" & rec.Fields("sarja") & "' and comanda='" & rec.Fields("[numar comanda]") & "'")
1560                    RamasDeExecutatPeTransa = RamasDeExecutatPeTransa - cantpetransa
1570                    If RamasDeExecutatPeTransa < 0 Then
1580                        RamasDeExecutatPeTransa = Abs(RamasDeExecutatPeTransa)
1590                        rec.Fields("[ramase de executat]") = RamasDeExecutatPeTransa
1600                        RamasDeExecutatPeTransa = 0
1610                    Else
1620                        rec.Fields("[ramase de executat]") = cantpetransa
                            rec.Fields("[executatE]") = DLookup("cant", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "'") - cantpetransa

1630                    End If
1640                Next
1650            Else
1660                rec.Fields("[ramase de executat]") = 0    ' rec.Fields("[cantitate pe transa]")
                    rec.Fields("[executatE]") = DLookup("cant", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "'")
1670            End If
                ''1680            If rec.Fields("[ramase de executat]") > rec.Fields("[cantitate pe transa]") Then rec.Fields("[ramase de executat]") = rec.Fields("[cantitate pe transa]")

1690            NRART = rec.Fields("produs") & DLookup("nrart", "comenzi", "nrcomanda='" & rec.Fields("[numar comanda]") & "'")
1700        End Select


1710        rec.Fields("[timp necesar pentru un produs (in secunde)]") = DLookup("volumtimp", "[nomenclator articole]", "nrart='" & NRART & "'")
1720        rec.Fields("[timp care mai trebuie lucrat pe comanda(in secunde)]") = rec.Fields("[ramase de executat]") * rec.Fields("[timp necesar pentru un produs (in secunde)]")
1730        If InStr(1, UCase(rec.Fields("cliseu")), "M") Then
1740            rec.Fields("cliseu") = "Mare"
1750        Else
1760            rec.Fields("cliseu") = "Mic"
1770        End If
1780
1790        rec.Update
1800        rec.MoveNext
1810    Next
        '''1820    CurrentDb.Execute "delete * from [AI_Comenzi] where [ramase de executat]<=0"
        CurrentDb.Execute "update [AI_Comenzi] set [ramase de executat]=0 where [ramase de executat]<0"
1830    CurrentDb.Execute "update [AI_Comenzi] set cliseu=null where produs like 'p*'"
'''CurrentDb.Execute "delete * from [AI_Comenzi] where EXECUTATE>=[cantitate pe transa]"
1840 End If
1850 rec.Close
1860 Set rec = Nothing
1870 SIR = "select * from [AI_Comenzi]"
1880 Set rec = CurrentDb.OpenRecordset(SIR, dbOpenDynaset, dbSeeChanges)
1890 If Not rec.EOF Then
1900    rec.MoveLast
1910    rec.MoveFirst
1920    If CuFisier Then Open numeFisier For Append As #hFile

1930    If CuFisier Then Print #hFile, "C) COMENZI PENTRU PLANIFICAT:"
        AnalizaAI = AnalizaAI & "C) COMENZI PENTRU PLANIFICAT:"
1940    If CuFisier Then Print #hFile, ""
        AnalizaAI = AnalizaAI & ""
1950    antet = ""
1960    For C = 0 To rec.Fields.Count - 1
1970        If antet = "" Then
1980            antet = rec.Fields(C).name
1990        Else
2000            antet = antet & "," & rec.Fields(C).name
2010        End If
2020    Next C
2030    If CuFisier Then Print #hFile, antet
        AnalizaAI = AnalizaAI & antet
2040    For N = 1 To rec.RecordCount
2050        continut = ""
2060        For C = 0 To rec.Fields.Count - 1
2070            If continut = "" Then
2080                continut = Nz(rec.Fields(C).Value, 0)
2090            Else
2100                continut = continut & "," & Nz(rec.Fields(C).Value, 0)
2110            End If
2120        Next C
2130        If CuFisier Then Print #hFile, continut
            AnalizaAI = AnalizaAI & continut
2140        rec.MoveNext
2150    Next
2160    If CuFisier Then Close #hFile
2170 End If
2180 rec.Close
2190 Set rec = Nothing
    '720 MsgBox "Finalizat colectarea"
2200 VARRETURN = SysCmd(acSysCmdSetStatus, "Salvez fisier")
2210 DoCmd.Hourglass False
    If CuFisier Then
        Dim strCaleUSB
2220    strCaleUSB = CaleUSB
2230    If strCaleUSB <> "nu" Then

2240        ChDrive strCaleUSB & ":\"
2250        If CopiazaFisier(numeFisier, strCaleUSB & ":\Planificare Comenzi AI.txt") Then
2260            MsgBox "Fisierul a fost transferat", vbSystemModal
2270            ShellEx strCaleUSB & ":\Planificare Comenzi AI.txt"
2280        Else
2290            MsgBox "Fisierul nu a putut fi transferat)", vbSystemModal
2300            ShellEx numeFisier
2310        End If
2320    Else
2330        MsgBox "Sticul nu este recunoscut sau nu este introdus. Puteti alege calea pentru salvare.", vbSystemModal
2340        strCaleUSB = BrowseFolder("Alegeti calea pentru salvarea fisierului.")
2350        If CopiazaFisier(numeFisier, strCaleUSB & "\Planificare Comenzi AI.txt") Then
2360            MsgBox "Fisierul a fost transferat", vbSystemModal
2370            ShellEx strCaleUSB & "\Planificare Comenzi AI.txt"
2380        Else
2390            MsgBox "Fisierul nu a putut fi transferat. Va fi deschis fisierul temporar si il veti putea salva pe stic manual." & vbCrLf & numeFisier, vbSystemModal
2400            ShellEx numeFisier
2410        End If

2420        ShellEx numeFisier
2430    End If
    End If
2440 Application.SysCmd (acSysCmdClearStatus)




    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
2450 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
2460 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
2470 Select Case errNumar
    Case 0


2480 Case 3021
2490    Resume Next
2500 Case 3211
2510    MsgBox "Inchideti tabelul"
2520 Case Else
2530    ScrieEroare "Eroare in [functii].[AnalizaAI] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
2540    RaspunsMesaj = MsgBox("[Eroare in functii].[AnalizaAI] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[AnalizaAI] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
2550    If RaspunsMesaj = vbYes Then
2560        Resume Next
2570    Else
2580        GoTo TRATARE_ERORI_iesire
2590    End If
2600 End Select
    '========================== terminat tratare erori
End Function

Public Function BrowseFolder(szDialogTitle As String) As String
    Dim X As Long, BI As BROWSEINFO, dwIList As Long
    Dim szPath As String, wPos As Integer

    With BI
        .hOwner = hWndAccessApp
        .lpszTitle = szDialogTitle

        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    dwIList = SHBrowseForFolder(BI)
    szPath = Space$(512)
    X = SHGetPathFromIDList(ByVal dwIList, ByVal szPath)

    If X Then
        wPos = InStr(szPath, Chr(0))
        BrowseFolder = Left$(szPath, wPos - 1)
    Else
        BrowseFolder = ""
    End If
End Function
Sub CapturaEcranGDIPlus(Optional Format As String = "jpg", Optional ByVal Cale As String = "")
    Dim screenDC As LongPtr, memDC As LongPtr
    Dim hBitmap As LongPtr, hOldBitmap As LongPtr
    Dim Width As Long, Height As Long

    Width = GetSystemMetrics(0)
    Height = GetSystemMetrics(1)

    screenDC = GetDC(0)
    memDC = CreateCompatibleDC(screenDC)
    hBitmap = CreateCompatibleBitmap(screenDC, Width, Height)
    hOldBitmap = SelectObject(memDC, hBitmap)

    BitBlt memDC, 0, 0, Width, Height, screenDC, 0, 0, &HCC0020
    SelectObject memDC, hOldBitmap

    '''    Dim Cale As String
    '''    Cale = CurrentProject.path & "\captura." & LCase(format)

    If Cale = "" Then
        Cale = CurrentProject.path & "\Captura.bmp"
    End If

    If SaveHBitmapAsImage(hBitmap, Cale, Format) Then
        ' MsgBox "Captura salvata: " & Cale
    Else
        ' MsgBox "Eroare la salvare imagine eroare."
    End If

    DeleteObject hBitmap
    DeleteDC memDC
    ReleaseDC 0, screenDC
End Sub

Public Function SaveHBitmapAsImage(hBitmap As LongPtr, filePath As String, Format As String) As Boolean
    On Error Resume Next
    Dim gdipToken As LongPtr
    Dim gdipInput As GdiplusStartupInput
    Dim hImage As LongPtr
    Dim clsidEncoder As Guid

    gdipInput.GdiplusVersion = 1
    GdiplusStartup gdipToken, gdipInput

    ' Format GUID (JPG sau PNG)
    Dim sCLSID As String
    Select Case LCase(Format)
    Case "jpg", "jpeg"
        sCLSID = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
    Case "png"
        sCLSID = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
    Case Else
        MsgBox "Format necunoscut: " & Format: Exit Function
    End Select

    CLSIDFromString StrPtr(sCLSID), clsidEncoder
    GdipCreateBitmapFromHBITMAP hBitmap, 0, hImage

    If hImage <> 0 Then
        GdipSaveImageToFile hImage, StrPtr(filePath), clsidEncoder, 0
        GdipDisposeImage hImage
        SaveHBitmapAsImage = True
    End If

    GdiplusShutdown gdipToken
End Function
Sub CapturaEcran(Optional ByVal Cale As String = "")
    Dim screenDC As LongPtr, memDC As LongPtr
    Dim hBitmap As LongPtr, hOldBitmap As LongPtr
    Dim Width As Long, Height As Long

    Width = GetSystemMetrics(0)  ' SM_CXSCREEN
    Height = GetSystemMetrics(1)    ' SM_CYSCREEN

    screenDC = GetDC(0)
    memDC = CreateCompatibleDC(screenDC)
    hBitmap = CreateCompatibleBitmap(screenDC, Width, Height)
    hOldBitmap = SelectObject(memDC, hBitmap)

    BitBlt memDC, 0, 0, Width, Height, screenDC, 0, 0, &HCC0020
    SelectObject memDC, hOldBitmap

    ' Pregatim IPicture
    Dim IID_IDispatch As Guid
    With IID_IDispatch
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4_0 = &H8B
        .Data4_1 = &HBB
        .Data4_2 = &H0
        .Data4_3 = &HAA
        .Data4_4 = &H0
        .Data4_5 = &H30
        .Data4_6 = &HC
        .Data4_7 = &HAB
    End With

    Dim picdesc As uPicDesc
    picdesc.size = Len(picdesc)
    picdesc.type = 1    ' Bitmap
    picdesc.hPic = hBitmap
    picdesc.hPal = 0

    Dim pic As IPicture
    OleCreatePictureIndirect picdesc, IID_IDispatch, True, pic

    If Cale = "" Then
        Cale = CurrentProject.path & "\Captura.bmp"
    End If

    ' Salveaza imaginea
    SavePicture pic, Cale
    '''MsgBox "Captura de ecran salvata �n: " & Cale, vbInformation

    ' Cura?a memoria
    DeleteObject hBitmap
    DeleteDC memDC
    ReleaseDC 0, screenDC
End Sub

Function testcess()
''
''Dim ERORICAPTURATE As String
''ERORICAPTURATE = DLookup("valoare", "setari", "parametru='ERORICAPTURATE'")
'Call CapturaEcran(ERORICAPTURATE & "\" & Replace(Replace(Now, ".", "-"), ":", "-") & " EROARE .bmp")
''Call CapturaEcranGDIPlus("png", ERORICAPTURATE & "\" & format(Now, "yyyy-mm-dd hh nn") & " EROARE.png")
    ScrieEroare "modul1", ""
    'Call CapturaEcranGDIPlus("png", ERORICAPTURATE & "\" & format(Now, "yyyy-mm-dd hh nn") & formular & ruttina & " .png")

End Function


Public Function RuleazaScriptPowerShell(scriptPath As String) As Boolean

    Dim powershellCmd As String
    Dim retval As Double

    '    ' Calea completa catre scriptul .ps1
    '    scriptPath = "C:\cale\catre\script.ps1"
    '
    ' Comanda pentru rulare PowerShell cu bypass la execu?ie restric?ionata
    powershellCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """"

    ' Ruleaza comanda
    retval = shell(powershellCmd, vbNormalFocus)
    RuleazaScriptPowerShell = 1
    ' retVal con?ine PID-ul procesului lansat (po?i verifica daca > 0)
    If retval = 0 Then
        MsgBox "Nu s-a putut rula scriptul PowerShell.", vbCritical
        RuleazaScriptPowerShell = 0
    End If
End Function

Function RuleazaPowerShellCuOutput(scriptPath As String) As String
    Dim wsh As Object
    Dim exec As Object
    Dim output As String
    Dim line As String

    On Error GoTo EROARE

    ' Creeaza obiectul WScript.Shell
    Set wsh = CreateObject("WScript.Shell")

    ' Ruleaza PowerShell cu scriptul specificat
    ' -NoProfile pentru rulare curata, -ExecutionPolicy Bypass pentru a ocoli restric?iile
    scriptPath = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """"


    Set exec = wsh.exec(scriptPath)

    ' Cite?te linie cu linie ie?irea standard p�na c�nd procesul se termina
    output = ""
    Do While exec.Status = 0
        Do While Not exec.StdOut.AtEndOfStream
            line = exec.StdOut.ReadLine
            output = output & line & vbCrLf
        Loop
        DoEvents    ' permite interac?iunea cu UI-ul �n timpul a?teptarii
    Loop

    ' Daca au mai ramas linii dupa terminarea procesului
    Do While Not exec.StdOut.AtEndOfStream
        line = exec.StdOut.ReadLine
        output = output & line & vbCrLf
    Loop

    RuleazaPowerShellCuOutput = output
    Exit Function

EROARE:
    RuleazaPowerShellCuOutput = "Eroare la rularea PowerShell: " & err.Description
End Function
Public Function testRPSC()
    Dim SIR As String
    Dim shell As Object
    Dim caleDesktop As String
    Dim Cale As String

    Dim AP As Excel.Application
    Dim wor As workbook
10  On Error GoTo TRATARE_ERORI
    Dim SirEroare As String

20  Set shell = CreateObject("WScript.Shell")
    'caleDesktop = shell.SpecialFolder("Desktop")

30  caleDesktop = Environ("USERPROFILE") & "\Desktop"
40  SIR = DLookup("textActiune", "de facut", "idActiune=50")

50  Cale = caleDesktop & "\INFO-" & Environ("COMPUTERNAME") & ".ps1"
60  Kill caleDesktop & "\INFO-" & Environ("COMPUTERNAME") & ".XLS"

70  Open Cale For Output As #1
80  Print #1, SIR
90  Close #1
    Dim MESAJ As String
100 MESAJ = RuleazaPowerShellCuOutput(Cale)
110 If InStr(1, MESAJ, "Script executat.") Then
120     Kill caleDesktop & "\INFO-" & Environ("COMPUTERNAME") & ".ps1"

130     Set AP = New Excel.Application
140     AP.visible = False

150     AP.Workbooks.OpenText _
                FileName:=caleDesktop & "\INFO-" & Environ("COMPUTERNAME") & ".txt", _
                Origin:=65001, _
                StartRow:=1, _
                DataType:=1, _
                TextQualifier:=1, _
                ConsecutiveDelimiter:=False, _
                Tab:=False, _
                Semicolon:=True, _
                Comma:=False, _
                Space:=False, _
                Other:=False, _
                FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1))

160     AP.ActiveWorkbook.SaveAs FileName:=caleDesktop & "\INFO-" & Environ("COMPUTERNAME") & ".XLS", FileFormat:=56
170     AP.ActiveWorkbook.Close False
180     Kill caleDesktop & "\INFO-" & Environ("COMPUTERNAME") & ".txt"

190     AP.Quit


200     SIR = "Delete * from calculatoare where numeStatie like '" & Left(Environ("COMPUTERNAME"), 4) & "*'"
210     CurrentDb.Execute SIR, dbSeeChanges
220     DoCmd.TransferSpreadsheet acImport, , "Calculatoare", caleDesktop & "\INFO-" & Environ("COMPUTERNAME") & ".XLS", True

230     SIR = "Delete * from calculatoare where caracteristici='Memorie dedicata: , Partajata:'"
240     CurrentDb.Execute SIR, dbSeeChanges
250     DoCmd.OpenForm "calculatoare", , , "numestatie='" & Environ("COMPUTERNAME") & "'"
260 Else
270     MsgBox "Nu s-a putut rula scriptul"

280 End If
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
290 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
300 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
310 Select Case errNumar
    Case 0
320 Case 53
330     Resume Next
340 Case Else
350     ScrieEroare "Eroare in [functii].[testRPSC] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
360     RaspunsMesaj = MsgBox("[Eroare in functii].[testRPSC] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[testRPSC] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
370     If RaspunsMesaj = vbYes Then
380         Resume Next
390     Else
400         GoTo TRATARE_ERORI_iesire
410     End If
420 End Select
    '========================== terminat tratare erori
End Function

Public Function RuleazaPowerShellComanda(COMANDA As String) As String
    Dim wsh As Object
    Dim exec As Object
    Dim output As String
    Dim line As String
    Dim cmdLine As String

    On Error GoTo EROARE

    Set wsh = CreateObject("WScript.Shell")

    ' Construim linia completa de comanda PowerShell
    ' -NoProfile pentru pornire rapida fara profiluri
    ' -ExecutionPolicy Bypass sa nu blocheze rularea
    ' -Command sa rulam comanda data
    '    cmdLine = "powershell.exe -NoProfile -ExecutionPolicy Bypass -Command " & Chr(34) & COMANDA & Chr(34)
    cmdLine = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & Chr(34) & COMANDA & Chr(34)

    Set exec = wsh.exec(cmdLine, vbNormalFocus)

    output = ""
    Do While exec.Status = 0
        Do While Not exec.StdOut.AtEndOfStream
            line = exec.StdOut.ReadLine
            output = output & line & vbCrLf
        Loop
        DoEvents
    Loop
    Do While Not exec.StdOut.AtEndOfStream
        line = exec.StdOut.ReadLine
        output = output & line & vbCrLf
    Loop

    RuleazaPowerShellComanda = output
    Exit Function

EROARE:
    RuleazaPowerShellComanda = "Eroare la rularea PowerShell: " & err.Description
End Function
Sub TestRuleazaCmd()
    Dim rezultat As String
    ' Comanda PowerShell simpla, de ex: lista directoare curent
    rezultat = RuleazaPowerShellComanda("Get-ChildItem")
    MsgBox rezultat
End Sub

Public Function RulareBat(SIR As String) As String
    Dim MESAJ
    Dim WshShell, oExec

    Dim SirEroare As String
    Dim hFile As Integer
    Dim numeFisier As String
10  On Error GoTo TRATARE_ERORI
20  If Dir("C:\Tempv\", vbDirectory) = "" Then
30      MkDir "C:\Tempv\"
40  End If
50  numeFisier = "C:\Tempv\Script.bat"
60  hFile = FreeFile
70  Open numeFisier For Output As #hFile
80  Print #hFile, SIR
90  Close #hFile
100 If Dir("C:\Tempv\Script.bat") <> "" Then
110     Set WshShell = CreateObject("WScript.Shell")
120     comand = "cmd /c """ & "C:\Tempv\Script.bat" & """"
130     Set oExec = WshShell.exec(comand)
140     Do While oExec.Status = 0
150     Loop
160     RulareBat = oExec.StdOut.Readall
170     RulareBat = Replace(RulareBat, vbCrLf, "")
180 Else
190     RulareBat = "Nu s-a putut crea fisierul de test."
200 End If

    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
210 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
220 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
230 Select Case errNumar
    Case 0
240 Case Else
250     ScrieEroare "Eroare in [Form_ETICHETA DIFERITA].[Command83_Click] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
260     RaspunsMesaj = MsgBox("[Eroare in Form_ETICHETA DIFERITA].[Command83_Click] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in Form_ETICHETA DIFERITA].[Command83_Click] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
270     If RaspunsMesaj = vbYes Then
280         Resume Next
290     Else
300         GoTo TRATARE_ERORI_iesire
310     End If
320 End Select
    '========================== terminat tratare erori
End Function
Public Function NumeProiect()
    NumeProiect = Application.vbe.ActiveVBProject.name
End Function
Public Sub DisableHibernate()
    Dim hKey As Long
    Dim lpData As Long
    Dim lResult As Long

    ' Deschide cheia din registry unde sunt salvate set?rile de hibernare
    lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Power", 0, KEY_SET_VALUE, hKey)

    If lResult = 0 Then
        ' Seteaz? HibernateEnabled la 0 (dezactiveaz? hibernarea)
        lpData = 0
        lResult = RegSetValueEx(hKey, "HibernateEnabled", 0, REG_DWORD, lpData, 4)

        ' �nchide cheia Registry
        RegCloseKey hKey
    End If
End Sub


Public Sub SetHibernateTime()
    Dim hKey As Long
    Dim lpData As Long
    Dim lResult As Long

    ' Deschide cheia din registry unde sunt salvate set?rile de hibernare
    lResult = RegOpenKeyEx(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Power\User\PowerSchemes", 0, KEY_SET_VALUE, hKey)

    If lResult = 0 Then
        ' Seteaz? timpul de hibernare la 600 secunde (10 minute)
        lpData = 600
        lResult = RegSetValueEx(hKey, "HibernateTimeout", 0, REG_DWORD, lpData, 4)

        ' �nchide cheia Registry
        RegCloseKey hKey
    End If
End Sub


Public Sub SetStandbyTime(TIMP As Long)
    Dim hKey As Long
    Dim lpData As Long
    Dim lResult As Long

    ' Deschide cheia din registry
    lResult = RegOpenKeyEx(HKEY_CURRENT_USER, "Control Panel\PowerCfg\PowerPolicies\0", 0, KEY_SET_VALUE, hKey)

    If lResult = 0 Then
        ' Seteaz? timpul de standby la 10 minute (600 secunde)=0 pentru niciodata
        lpData = TIMP
        lResult = RegSetValueEx(hKey, "Policies", 0, REG_DWORD, lpData, 4)

        ' �nchide cheia Registry
        RegCloseKey hKey
    End If
End Sub

Public Sub OpresteSleep()
    SetThreadExecutionState &H80000002
End Sub
Public Sub PermiteSleep()
    SetThreadExecutionState &H80000000
End Sub

Sub Test()
    ArataLazi "cp-192-037"

End Sub

Public Function ArataLazi(Articol As String)
    Dim oRs As DAO.Recordset
    Dim sResult As String
    Dim N As Integer
    Dim precedent As Integer
    Dim final As Integer
    Dim LADAcURENTA As String
    Dim REZERVA As String
10  On Error GoTo TRATARE_ERORI
20  LADAcURENTA = Nz(DLookup("NIVEL", "NOMENCLATOR ARTICOLE", "NRART='" & Articol & "'"), "")

30  REZERVA = Nz(DLookup("detaliiexecutie", "dbo_detalii articol", "articol='" & Articol & "' and detaliiexecutie like 'Rez*'"), "")

40  LADAcURENTA = val(LADAcURENTA)
50  Set oRs = CurrentDb.OpenRecordset("Select lada from [dbo_Lazi Articole] where articol='" & Articol & "' order by lada")
60  ArataLazi = LADAcURENTA
70  If Not oRs.EOF Then

80      oRs.MoveLast
90      oRs.MoveFirst
100     For N = 1 To oRs.RecordCount
110         If N = 1 Then
120             ArataLazi = ArataLazi & "," & oRs!lada
130             precedent = oRs!lada + 1
140         Else

150             If oRs!lada = precedent Then    'daca curent este succesiv
160                 If N = oRs.RecordCount Then    ' daca a ajuns la final
170                     final = oRs!lada
180                     ArataLazi = Replace(ArataLazi, ";", "")
190                     ArataLazi = ArataLazi & "-" & oRs!lada & ";"
200                 Else    ' daca nu a ajuns la final
210                     final = oRs!lada
220                     If Right(ArataLazi, 1) = ";" And oRs!lada <> precedent Then    ' daca a fost incheiat interval anterior si nu este consecutiv
230                         Replace ArataLazi, ";", ","
240                         ArataLazi = ArataLazi & "," & final
250                     End If    ' daca a fost incheiat interval anterior si nu este consecutiv
260                 End If    ' daca a ajuns la final

270             Else    'daca curent nu este succesiv
280                 ArataLazi = Replace(ArataLazi, ";", "")
290                 If final <> 0 Then    ' daca nu s-a scris incheiere interval
300                     ArataLazi = ArataLazi & "-" & final & "," & oRs!lada & ";"
310                     final = 0
320                 Else
330                     ArataLazi = ArataLazi & "," & oRs!lada & ";"
340                 End If
350             End If    'daca curent este succesiv
360             precedent = oRs!lada + 1
370         End If
380         oRs.MoveNext
390     Next
400 Else
410     ArataLazi = Replace(ArataLazi, ",", "")
420 End If
430 oRs.Close

440 Set oRs = Nothing
450 ArataLazi = ArataLazi & " " & REZERVA
460 Exit Function

    '========================== incep tratare erori
TRATARE_ERORI_iesire:

470 Set oRs = Nothing

480 ArataLazi = ArataLazi & ",Eroare"
490 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
500 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
510 Select Case errNumar
    Case 0
520 Case Else

530     ScrieEroare "Eroare in [functii].[ArataLazi] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False
540     RaspunsMesaj = MsgBox("[Eroare in functii].[ArataLazi] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[ArataLazi] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
550     If RaspunsMesaj = vbYes Then
560         Resume Next
570     Else
580         GoTo TRATARE_ERORI_iesire
590     End If
600 End Select
    '========================== terminat tratare erori

End Function
Function PinLaQuickAccess(Cale As String)

    Dim director As String
    Dim restcale As String
    Dim pos

    If Right(Cale, 1) = "\" Then Cale = Left(Cale, Len(Cale) - 1)
    pos = 1
    Do
        pos = InStr(pos + 1, Cale, "\")
        If pos <> 0 Then
            director = Right(Cale, Len(Cale) - pos)
        End If
    Loop Until pos = 0

    Dim objShell As Object, oFoldItem As Object, item As Object
    Dim oFold As Object, objVerbs As Variant
    Set objShell = CreateObject("Shell.Application")
    restcale = Replace(Cale, "\" & director, "")
    Set oFold = objShell.Namespace(restcale & "\")
    If Not oFold Is Nothing Then
        Set oFoldItem = oFold.parsename(director)
        Set objVerbs = oFoldItem.verbs
        For Each item In objVerbs

            'DIRECTOARE
            '''&Open
            '''Pin to Quick access
            '''QMS Adauga in Colector
            '''QMS Procesare
            '''
            '''Bazaar Chec&kout/Branch...
            '''Bazaar &Init...
            '''Bazaar &Explorer
            '''
            '''&Add to archive...
            '''Add &to "ppp.rar"
            '''Compress and email...
            '''Compress to "ppp.rar" and email
            '''Restore Previous & Versions
            '''
            '''&Pin to Start
            '''
            '''Cu& T
            '''&Copy
            '''Create &shortcut
            '''&Delete
            '''Rena& Me
            '''P& roperties
            'FISIERE
            '''&Open
            '''&Edit
            '''Importa in AutoCorect
            '''&New
            '''&Print
            '''Open wit&h...
            '''
            '''
            '''&Add to archive...
            '''Add &to "licente.rar"
            '''Compress and email...
            '''Compress to "licente.rar" and email
            '''Restore Previous & Versions
            '''
            '''Cu& T
            '''&Copy
            '''Create &shortcut
            '''&Delete
            '''Rena& Me
            '''P& roperties
            ''  Debug.Print item.Name
            If item.name = "Pin to Quick access" Then
                item.doit
                Exit For
            End If
        Next
    End If
End Function






Function InlocuiesteLinie(strModuleName, strText As String, strTextFinal As String) _
         As Boolean
    Dim mdl As MODULE, lngNumLines As Long
    Dim lngSLine As Long, lngSCol As Long
    Dim lngELine As Long, lngECol As Long
    Dim strTemp As String

10  On Error GoTo TRATARE_ERORI

20  Set mdl = Modules(strModuleName)
30  If mdl.Find(strText, lngSLine, lngSCol, lngELine, lngECol) Then
40      lngNumLines = Abs(lngELine - lngSLine) + 1
50      strTemp = LTrim$(mdl.Lines(lngSLine, lngNumLines))
60      strTemp = RTrim$(strTemp)
70      If strTemp = strText Then
            ' mdl.DeleteLines lngSLine, lngNumLines
80          mdl.ReplaceLine lngSLine, strTextFinal
90      Else
100         MsgBox "Line contains text in addition to '" _
                   & strText & "'."
110     End If
120 Else
        '''130     MsgBox "Text '" & strText & "' not found."
140 End If
150 InlocuiesteLinie = True
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
160 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
170 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
180 Select Case errNumar
    Case 0
190 Case Else
200     ScrieEroare "Eroare in [functii].[InlocuiesteLinie] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata Form_99 referinte rutina InlocuiesteLinie"
210     RaspunsMesaj = MsgBox("[functii].[InlocuiesteLinie] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[functii].[InlocuiesteLinie] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
220     If RaspunsMesaj = vbYes Then
230         Resume Next
240     Else
250         InlocuiesteLinie = False
260         GoTo TRATARE_ERORI_iesire
270     End If
280 End Select
    '========================== terminat tratare erori

End Function






Public Function AplicaInlocuire(SIR As String, sector As String, modul As String) As Boolean
    Dim extens
    Select Case modul
    Case "modul"
        DoCmd.OpenModule SIR
        extens = ""
    Case "form"
        DoCmd.OpenForm SIR, acDesign
        extens = "Form_"
    End Select
    Application.vbe.ActiveVBProject.VBComponents(extens & SIR).CodeModule.codePane.Show
    Select Case sector
    Case ""
        InlocuiesteLinie extens & SIR, "c:\windows\sysWOW64\tsclib.dll", "c:\windows\system32\tsclib.dll"
    Case "dep1"

        InlocuiesteLinie extens & SIR, "c:\windows\system32\tsclib.dll", "c:\windows\sysWOW64\tsclib.dll"
    Case "dep2"
        InlocuiesteLinie extens & SIR, "c:\windows\sysWOW64\tsclib.dll", "c:\windows\system32\tsclib.dll"
    Case Else
        InlocuiesteLinie extens & SIR, "c:\windows\sysWOW64\tsclib.dll", "c:\windows\system32\tsclib.dll"
    End Select
    Application.vbe.ActiveVBProject.VBComponents(extens & SIR).CodeModule.codePane.Window.Close
End Function

Public Function AdapteazaLaDomeniu()

    Dim SIR As String

    If UCase(Environ("userdomain")) <> "SER" Then
        AplicaInlocuire "imprimare etichete", "", "modul"
        AplicaInlocuire "001 Scule cautare", "", "form"
        AplicaInlocuire "dbo_Bibliorafturi comenzi", "", "form"
        AplicaInlocuire "ETICHETA DIFERITA", "", "form"
        AplicaInlocuire "ETICHETA DIFERITA mare", "", "form"
        AplicaInlocuire "Mijloace", "", "form"

    Else

        AplicaInlocuire "imprimare etichete", Environ("username"), "modul"
        AplicaInlocuire "001 Scule cautare", Environ("username"), "form"
        AplicaInlocuire "dbo_Bibliorafturi comenzi", Environ("username"), "form"
        AplicaInlocuire "ETICHETA DIFERITA", Environ("username"), "form"
        AplicaInlocuire "ETICHETA DIFERITA mare", Environ("username"), "form"
        AplicaInlocuire "Mijloace", Environ("username"), "form"
    End If



End Function
Public Function numa() As String
    numa = Environ("userdomain") & "\" & Environ("username")

End Function

Public Function userdomain() As String
    On Error Resume Next
    userdomain = ""
    userdomain = Environ("userdomain")

End Function

Public Function campuriFormular()
    Dim rec As DAO.Recordset
    Dim N
    Dim exista As Boolean
    Dim CTRL1 As CONTROL
    Dim ControlCreat As TextBox
    Dim ORDINECOLOANA As Integer
    Dim SIR As String
10    On Error GoTo TRATARE_ERORI
20    ORDINECOLOANA = 0
30    DoCmd.Close acForm, "qrySituatie comenzi"
40    DoCmd.OpenForm "qrySituatie comenzi", acDesign

50    For Each CTRL1 In Forms("qrySituatie comenzi").Controls '
60      Select Case CTRL1.name
        Case "Data", "Produs", "Data_Label", "Produs_Label"
70      Case Else
80          Application.DeleteControl "qrySituatie comenzi", CTRL1.name
90      End Select
100   Next
110   Forms("qrySituatie comenzi").Controls("Data").ColumnWidth = 800
120   Forms("qrySituatie comenzi").Controls("Produs").ColumnWidth = 600

130   SIR = "SELECT left(Numesiprenume,instr(Numesiprenume,' ')) & mid(Numesiprenume,instr(Numesiprenume,' '),2) as nume" & _
          " FROM Personal" & _
          " WHERE (((personal.Productiv)=True) AND ((personal.Functia) In ('ambalator','femeie de serviciu'))) ORDER BY Numesiprenume;"
140   Set rec = CurrentDb.OpenRecordset(SIR, dbOpenDynaset, dbSeeChanges)
150   rec.MoveLast
160   rec.MoveFirst
170   ORDINECOLOANA = 100
180   For N = 1 To rec.RecordCount
190     ORDINECOLOANA = ORDINECOLOANA + 1
200     For Each CTRL1 In Forms("qrySituatie comenzi").Controls
210         If rec.Fields("nume") = CTRL1.name Then
220             Application.DeleteControl "qrySituatie comenzi", CTRL1.name
230         End If
240     Next
250     Set ControlCreat = CreateControl("qrySituatie comenzi", acTextBox, acDetail, "", rec.Fields("nume"), 100, 100, 50, 100)
260     ControlCreat.name = rec.Fields("nume")
270     ControlCreat.visible = True
280     ControlCreat.ColumnOrder = ORDINECOLOANA
290     ControlCreat.ColumnWidth = 1100
300     rec.MoveNext
310   Next

320   DoCmd.Save acForm, "qrySituatie comenzi"
330   Forms("qrySituatie comenzi").RowHeight = 190
340   Forms("qrySituatie comenzi").DatasheetFontHeight = 8
350   DoCmd.Close acForm, "qrySituatie comenzi", acSaveYes
360   DoCmd.OpenForm "PRIORITIZARE COMENZI"
    '========================== incep tratare erori
TRATARE_ERORI_iesire:

370   Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
380   lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
390   Select Case errNumar
    Case 0
400   Case Else
410     ScrieEroare "Eroare in [functii].[campuriFormular] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netrarata functii rutina campuriFormular"
420     RaspunsMesaj = MsgBox("[Eroare in functii].[campuriFormular] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[campuriFormular] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
430     If RaspunsMesaj = vbYes Then
440         Resume Next
450     Else
460         GoTo TRATARE_ERORI_iesire
470     End If
480   End Select
    '========================== terminat tratare erori

End Function





Public Function cauta_text(textCautat As String) As String
'''    Dim db As DAO.Database
'''    Dim cnt As CONTAINER
'''    Dim doc As Document
'''    Dim mdl As Object
'''    Dim frm As Form
'''    Dim rpt As Report
'''    Dim ctl As CONTROL
'''    Set db = CurrentDb
'''    Dim linie As Long
''''''    For Each mdl In db.Containers("Modules").Documents
''''''
''''''        linie = mdl.Find(textcautat, 1, 1, mdl.CountOfLines, mdl.CountOfLines)
''''''        If linie > 0 Then
''''''            MsgBox "Gasit in modul " & mdl.Name & "   Linia " & linie
''''''        End If
''''''    Next
'''
'''    For Each frm In db.Containers("forms").Documents
'''        For Each ctl In frm.Controls
'''            If ctl.HasModule Then
'''                linie = ctl.CodeModule.Find(textcautat, 1, 1, ctl.CodeModule.CountOfLines, ctl.CodeModule.CountOfLines)
'''                If linie > 0 Then
'''                    MsgBox "Gasit in Formularul " & frm.Name & " Controlul " & ctl.Name & "   Linia " & linie
'''                End If
'''            End If
'''        Next ctl
'''    Next frm
'''
'''    For Each rpt In db.Containers("forms").Documents
'''        For Each ctl In rpt.Controls
'''            If ctl.HasModule Then
'''                linie = ctl.CodeModule.Find(textcautat, 1, 1, ctl.CodeModule.CountOfLines, ctl.CodeModule.CountOfLines)
'''                If linie > 0 Then
'''                    MsgBox "Gasit in Raportul " & rpt.Name & " Controlul " & ctl.Name & "   Linia " & linie
'''                End If
'''            End If
'''        Next ctl
'''    Next rpt
'''
'''    Set db = Nothing
'''    Set cnt = Nothing
'''    Set doc = Nothing
'''    Set mdl = Nothing
'''    Set frm = Nothing
'''    Set rpt = Nothing
'''    Set ctl = Nothing
'''



10  On Error GoTo TRATARE_ERORI


    Dim CurrentDb As DAO.Database
    Dim oAcc As Access.Application
    Dim dbCurrent As DAO.Database
    '    Dim RS As DAO.Recordset
    Dim doc As Document
    Dim bolIsLoaded As Boolean, strDocName As String
    Dim lngCount As Long, lngR As Long, intProcOrder As Integer
    Dim lngCountDecl As Long, inti As Integer
    Dim lngI As Long, lngCountProcLines As Long
    Dim strProcName As String, strProcLines As String
    Dim strModuleType() As String, intMT As Integer
    Dim objTemp As Object, intModuleType As Integer
    Dim strDBName As String, strDBPath As String
    Dim strDrive As String, strDir As String, strFName As String, strExt As String
    Dim VARRETURN As Variant
    Dim lngOldProcType As Long, strOldProcName As String, strFinalProcName As String


30  DoCmd.Hourglass True

40  ReDim strModuleType(0 To 2)
50  strModuleType(0) = "Modules"
60  strModuleType(1) = "Forms"
70  strModuleType(2) = "Reports"


    ''    Call SplitPath(strPath, strDrive, strDir, strFName, strExt)
    ''    strDBPath = strDrive & strDir
    ''    strDBName = strFName & "." & strExt

    'open recordset to tblModules in THIS database
    '    Set dbCurrent = CurrentDb()
    ''    Set RS = dbCurrent.OpenRecordset("tblModules")


    'open OTHER database
    'Set oAcc = New Access.Application
    '    Set CurrentDb = CurrentDb ' oAcc.DBEngine.OpenDatabase(strPath, False, False, ";PWD=" & strPassword)
    'oAcc.OpenCurrentDatabase strPath

80  intProcOrder = 0

    'Move through the modules
90  For intMT = 0 To 2 Step 1
100     For Each doc In CurrentDb.Containers(strModuleType(intMT)).Documents
110         strDocName = doc.name
120         Select Case strModuleType(intMT)
            Case "Modules"
130             intModuleType = acModule
140         Case "Forms"
150             intModuleType = acForm
160         Case "Reports"
170             intModuleType = acReport
180         End Select
190         bolIsLoaded = IsObjectOpen(intModuleType, strDocName)
            'you can't open the module if it's already open.
200         If bolIsLoaded = False Then
210             Select Case intModuleType
                Case acModule
220                 oAcc.DoCmd.OpenModule doc.name
230                 Set objTemp = oAcc.Modules(doc.name)
240             Case acForm
250                 oAcc.DoCmd.OpenForm strDocName, acDesign, , , acFormReadOnly, acWindowNormal
260                 Set objTemp = oAcc.Forms(strDocName).MODULE
270             Case acReport
280                 oAcc.DoCmd.OpenReport strDocName, acViewDesign
290                 Set objTemp = oAcc.Reports(doc.name).MODULE
300             End Select
310         Else
320             Select Case intModuleType
                Case acModule
330                 Set objTemp = Modules(doc.name)
340             Case acForm
350                 Set objTemp = Forms(doc.name).MODULE
360             Case acReport
370                 Set objTemp = Reports(doc.name).MODULE
380             End Select
390         End If
            'get number of lines in this module
400         lngCount = objTemp.CountOfLines
            'Declarations
            'find out how many lines in Declaration section
410         lngCountDecl = objTemp.CountOfDeclarationLines
            'get code lines for Declaration section
420         strProcLines = objTemp.Lines(1, lngCountDecl)
430         Do Until (Asc(Left(strProcLines, 1)) <> 13 And Asc(Left(strProcLines, 1)) <> 10 _
                      And Asc(Left(strProcLines, 1)) <> 32)
440             strProcLines = Mid(strProcLines, 2)
450         Loop
460         Do Until (Asc(Right(strProcLines, 1)) <> 13 And Asc(Right(strProcLines, 1)) <> 10 _
                      And Asc(Right(strProcLines, 1)) <> 32)
470             strProcLines = Mid(strProcLines, 1, Len(strProcLines) - 1)
480         Loop
            ''            With RS
            ''                .AddNew
            ''                !DbName = strDBName
            ''                !PathName = strDBPath
            ''                !ModName = strDocName
            ''                !ModType = strModuleType(intMT)
            ''                !ProcName = "Declaratii"
            ''                !ProcLines = strProcLines
            ''                !ProcLinesCount = lngCountDecl
            ''                !ProcOrder = intProcOrder
            ''                !SelectedForHTML = False
            ''                .Update
            ''            End With
490         strOldProcName = "Declaratii"
            'Check to see if there is anything else after the declarations section
            'Are there more lines in module than just the lines of Declaration?
500         If lngCount > lngCountDecl Then
                'start at first line after Declaration section
510             inti = lngCountDecl + 1
520             intProcOrder = intProcOrder + 1
                '***** Get Name of Proc of this line *********
                'inti specifies the number of a line in the module.
                'When return from getting the ProcOfLine,
                'lngR will specify the type of procedure:
                '   vbext_pk_Get    lngR=3   A Property Get proc
                '   vbext_pk_Let    lngR=1   A Property Let proc
                '   vbext_pk_Proc   lngR=0   A Sub or Function proc
                '   vbext_pk_Set    lngR=2   A Property Set proc
530             strProcName = objTemp.ProcOfLine(inti, lngR)
                'save proc name so will know when reach line with new proc
540             strOldProcName = strProcName
                'save type of proc for next compare (to distinguish same-name Property stmts)
550             lngOldProcType = lngR
                'If proc was a property stmt, add type to procname that will save in tblModules
560             Select Case lngOldProcType
                Case vbext_pk_Proc
570                 strFinalProcName = strProcName
580             Case vbext_pk_Get
590                 strFinalProcName = strProcName & " [Property Get]"
600             Case vbext_pk_Let
610                 strFinalProcName = strProcName & " [Property Let]"
620             Case vbext_pk_Set
630                 strFinalProcName = strProcName & " [Property Set]"
640             End Select
                '******  update progress display in status bar  *****************
650             VARRETURN = SysCmd(acSysCmdSetStatus, "Processing " _
                                                      & strModuleType(intMT) & " " & doc.name & ".... Procedure " & strFinalProcName)

                'get the number of lines for this proc
660             lngCountProcLines = objTemp.ProcCountLines(strProcName, lngR)

                'get the code lines for this proc
670             strProcLines = objTemp.Lines(inti, lngCountProcLines)
                'strip CRLF's and SPACES from left side of codelines
680             Do Until (Asc(Left(strProcLines, 1)) <> 13 And Asc(Left(strProcLines, 1)) <> 10 _
                          And Asc(Left(strProcLines, 1)) <> 32)
690                 strProcLines = Mid(strProcLines, 2)
700             Loop
                'strip CRLF's and SPACES from right side of codelines
710             Do Until (Asc(Right(strProcLines, 1)) <> 13 And Asc(Right(strProcLines, 1)) <> 10 _
                          And Asc(Right(strProcLines, 1)) <> 32)
720                 strProcLines = Mid(strProcLines, 1, Len(strProcLines) - 1)
730             Loop
                'for html coding (you probably want to delete the following)
                'if have a Proc (not a Property stmt),
                'add "Sub" or "Function" to start of proc name
740             If lngOldProcType = vbext_pk_Proc Then
750                 If Left(strProcLines, 15) = "Public Function" _
                       Or Left(strProcLines, 16) = "Private Function" Then
760                     strFinalProcName = "Function " & strFinalProcName
770                 Else
780                     strFinalProcName = "Sub " & strFinalProcName
790                 End If
800             End If
                ''                With RS
                ''                    .AddNew
                ''                    !DbName = strDBName
                ''                    !PathName = strDBPath
                ''                    !ModName = strDocName
                ''                    !ModType = strModuleType(intMT)
                ''                    !ProcName = strFinalProcName
                ''                    !ProcLines = strProcLines
                ''                    !ProcLinesCount = lngCountProcLines
                ''                    !ProcOrder = intProcOrder
                ''                    !SelectedForHTML = False
                ''                    .Update
                ''                End With
                'Go through the rest of the module, enumerating the procedures
810             For lngI = inti To lngCount

                    '***** Get Name of Proc of this line *********
                    'inti specifies the number of a line in the module.
                    'When return from getting the ProcOfLine,
                    'lngR will specify the type of procedure:
                    '   vbext_pk_Get    lngR=3   A Property Get proc
                    '   vbext_pk_Let    lngR=1   A Property Let proc
                    '   vbext_pk_Proc   lngR=0   A Sub or Function proc
                    '   vbext_pk_Set    lngR=2   A Property Set proc
820                 strProcName = objTemp.ProcOfLine(lngI, lngR)
                    'see if ProcName for this line has changed
                    'or ProcName is the same but have different ProcType
830                 If (strProcName <> strOldProcName) _
                       Or ((strProcName = strOldProcName) And (lngR <> lngOldProcType)) Then

840                     intProcOrder = intProcOrder + 1

                        'save proc name so will know when reach line with new proc
850                     strOldProcName = strProcName
                        'save type of proc for next compare (to distinguish same-name Property stmts)
860                     lngOldProcType = lngR
                        'If proc was a property stmt, add type to procname that will save in tblModules
870                     Select Case lngOldProcType
                        Case vbext_pk_Proc
880                         strFinalProcName = strProcName
890                     Case vbext_pk_Get
900                         strFinalProcName = strProcName & " [Property Get]"
910                     Case vbext_pk_Let
920                         strFinalProcName = strProcName & " [Property Let]"
930                     Case vbext_pk_Set
940                         strFinalProcName = strProcName & " [Property Set]"
950                     End Select
                        '******  update progress display in status bar  *****************
960                     VARRETURN = SysCmd(acSysCmdSetStatus, "Processing " _
                                                              & strModuleType(intMT) & " " & doc.name & ".... Procedure " & strFinalProcName)

                        'get the number of lines for this proc
970                     lngCountProcLines = objTemp.ProcCountLines(strProcName, lngR)
                        'get the code lines for this proc
980                     strProcLines = objTemp.Lines(lngI, lngCountProcLines)
                        'strip CRLF's and SPACES from left side of codelines
990                     Do Until (Asc(Left(strProcLines, 1)) <> 13 And Asc(Left(strProcLines, 1)) <> 10 _
                                  And Asc(Left(strProcLines, 1)) <> 32)
1000                        strProcLines = Mid(strProcLines, 2)
1010                    Loop
                        'strip CRLF's and SPACES from right side of codelines
1020                    Do Until (Asc(Right(strProcLines, 1)) <> 13 And Asc(Right(strProcLines, 1)) <> 10 _
                                  And Asc(Right(strProcLines, 1)) <> 32)
1030                        strProcLines = Mid(strProcLines, 1, Len(strProcLines) - 1)
1040                    Loop
                        'for html coding (you probably want to delete the following)
                        'if have a Proc (not a Property stmt),
                        'add "Sub" or "Function" to start of proc name
1050                    If lngOldProcType = vbext_pk_Proc Then
1060                        If Left(strProcLines, 15) = "Public Function" _
                               Or Left(strProcLines, 16) = "Private Function" Then
1070                            strFinalProcName = "Function " & strFinalProcName
1080                        Else
1090                            strFinalProcName = "Sub " & strFinalProcName
1100                        End If
1110                    End If
                        ''                        With RS
                        ''                            .AddNew
                        ''                            !DbName = strDBName
                        ''                            !PathName = strDBPath
                        ''                            !ModName = strDocName
                        ''                            !ModType = strModuleType(intMT)
                        ''                            !ProcName = strFinalProcName
                        ''                            !ProcLines = strProcLines
                        ''                            !ProcLinesCount = lngCountProcLines
                        ''                            !ProcOrder = intProcOrder
                        ''                            !SelectedForHTML = False
                        ''                            .Update
                        ''                        End With
1120                End If
                    'look at next line in code module
1130            Next lngI
                'finished getting all info from this module
                '/* End of If lngCount > lngCountDecl Then
1140        End If
            'reinit vars to get a new module
1150        intProcOrder = 0
1160        lngCountProcLines = 0
1170        strProcLines = " "
No_Module:
            'close object that contained last code module
1180        Set objTemp = Nothing
1190        oAcc.DoCmd.Close intModuleType, doc.name, acSaveNo
            'go get another code module of the same module type
1200    Next    'doc
        'have gotten all code modules for this type
        'so start getting another type of module
1210 Next    'intMT

    '******  update progress display in status bar  *****************
1220 VARRETURN = SysCmd(acSysCmdSetStatus, "Processing Complete")
    ''    EnumerateModules = True



    '========================== incep tratare erori
TRATARE_ERORI_iesire:
1230 oAcc.CloseCurrentDatabase
1240 CurrentDb.Close: Set CurrentDb = Nothing
    ''    RS.Close: Set RS = Nothing
1250 dbCurrent.Close: Set dbCurrent = Nothing
1260 DoCmd.Hourglass False
1270 DoCmd.SetWarnings True

1280 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
1290 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
1300 Select Case errNumar
    Case 0
1310 Case Else
1320    ScrieEroare "Eroare in [functii].[cauta_text] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netrarata functii rutina cauta_text"
1330    RaspunsMesaj = MsgBox("[Eroare in functii].[cauta_text] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[cauta_text] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
1340    If RaspunsMesaj = vbYes Then
1350        Resume Next
1360    Else
1370        GoTo TRATARE_ERORI_iesire
1380    End If
1390 End Select
    '========================== terminat tratare erori

End Function

Function InchideToate()
    mciSendString "close all", 0, 0, 0
End Function
Function RegleazaVolum(pista As String, nivel As Integer)
    mciSendString "setaudio " & pista & " volume to " & nivel, 0&, 0, 0    ' functioneaza numai pt mp3
End Function
Function RegleazaViteza(pista As String, VITEZA As Integer)
    mciSendString "set " & pista & " speed " & VITEZA, 0&, 0, 0
End Function

Function DeschideSunet(fisier As String, pista As String)
    fisier = GetShortFileName(App.path & "\SUNET\" & fisier)
    mciSendString "open " & fisier & " type MPEGVideo alias " & pista & " wait", 0, 0, 0
End Function
Function StartSunet(pista As String, continuu As Boolean, Optional dela As Integer = 0)
    Dim delacap As String
    delacap = " from " & dela
    Select Case continuu
    Case True
        mciSendString "play " & pista & delacap & " REPEAT", 0&, 0, 0
    Case False
        mciSendString "play " & pista & delacap, 0&, 0, 0
    End Select
End Function
Function StopSunet(pista As String)
    If MediaStatus(pista) = "PLAYING" Or MediaStatus(pista) = "PAUSED" Then
        mciSendString "stop " & pista, 0&, 0, 0
    End If
End Function
Function redare(fisier As String, Optional nivel As Integer = 1000, Optional VITEZA As Integer = 1000, Optional continuu As Boolean = False)
    fisier = GetShortFileName(App.path & "\SUNET\" & fisier)
    inchideSunet "SunetCurent"
    DeschideSunet fisier, "SunetCurent"
    RegleazaVolum "SunetCurent", nivel
    RegleazaViteza "SunetCurent", VITEZA
    StartSunet "SunetCurent", continuu
End Function


Public Function IncrementeazaVolum(pista As String, VALOARE As Integer)
    Dim VolumSunet
    VolumSunet = VolumSunet + VALOARE
    If MediaStatus(pista) = "PLAYING" Or MediaStatus(pista) = "PAUSED" Or MediaStatus(pista) = "STOPPED" Then
        mciSendString "setaudio " & pista & " volume to " & VolumSunet, 0&, 0, 0    ' functioneaza numai pt mp3
    End If
End Function

Public Function inchideSunet(pista As String)
    If MediaStatus(pista) = "PLAYING" Or MediaStatus(pista) = "PAUSED" Or MediaStatus(pista) = "STOPPED" Then
        mciSendString "close " & pista, 0, 0, 0
    End If
End Function


'To find the length of the media
Function MediaLength(pista As String) As Long
    Dim sLength As String * 255, lNullChar As Long
    mciSendString "status " & pista & " length", sLength, 255, 0
    'To find the null character position
    lNullChar = InStr(sLength, Chr$(0))
    'Return the length of the media
    MediaLength = CLng(val(Left$(sLength, lNullChar - 1)))
End Function

'To find the current position of the media
Function CurrentPos(pista As String) As Long
    Dim lNullChar As Long, sCurPos As String * 255
    mciSendString "status " & pista & " position", sCurPos, 255, 0
    lNullChar = InStr(sCurPos, Chr$(0))
    CurrentPos = CLng(val(Left$(sCurPos, lNullChar - 1)))
End Function

'To find the status of the video
Function MediaStatus(pista As String) As String
    Dim lNullChar As Integer, sStatus As String * 255
    mciSendString "status " & pista & " mode", sStatus, 255, 0
    lNullChar = InStr(sStatus, Chr$(0))
    MediaStatus = UCase$(Left$(sStatus, lNullChar - 1))
End Function

'To find valid state for stop or close
Function IsValidState(pista As String) As Boolean
    If MediaStatus(pista) = "PLAYING" Or MediaStatus(pista) = "PAUSED" Or MediaStatus(pista) = "STOPPED" Then
        IsValidState = True
    Else
        IsValidState = False
    End If
End Function


Function Backward(pista As String)
    Dim MStatus As String, lPos As Long
    MStatus = MediaStatus(pista)
    'To move 5 percent of media length towards back
    lPos = CInt(CurrentPos(pista) - (MediaLength(pista) * 0.05))
    If IsValidState(pista) Then
        If lPos > 0 Then
            mciSendString "seek " & pista & " to " & lPos, "", 0, 0
            If MStatus = "PLAYING" Then
                Play pista
            End If
        End If
    End If
End Function


Function DecreaseSpeed(pista As String)
    Dim speed
    If speed - 50 >= 100 Then
        speed = speed - 50
        mciSendString "set " & pista & " speed " & speed, "", 0, 0
    End If
End Function

Function ExitP(pista As String)
    If MediaStatus(pista) = "PLAYING" Or MediaStatus(pista) = "PAUSED" Then
        inchideSunet (pista)
    End If
    '''    Unload Me
End Function

Function FastBackward(pista As String)
    If IsValidState(pista) Then
        mciSendString "seek " & pista & " to start", "", 0, 0
        mciSendString "play " & pista & " from 1 to 2", "", 0, 0
    End If
End Function

Function FastForward(pista As String)
    If IsValidState(pista) Then
        mciSendString "seek " & pista & " to end", "", 0, 0
    End If
End Function

Function Forward(pista As String)
    Dim MStatus As String, lPos As Long
    MStatus = MediaStatus(pista)

    'To move 5 percent of media lenght toward forward
    lPos = CLng(CurrentPos(pista) + (MediaLength(pista) * 0.05))
    If IsValidState(pista) Then
        If lPos < MediaLength(pista) Then
            mciSendString "seek " & pista & " to " & lPos, "", 0, 0
            If MStatus = "PLAYING" Then
                Play pista
            End If
        End If
    End If
End Function

Function IncreaseSpeed(pista As String)
    Dim speed
    If speed + 50 <= 2000 Then
        speed = speed + 50
        mciSendString "set " & pista & " speed " & speed, "", 0, 0
    End If
End Function

Function OpenP(fisier As String, pista As String)
    Dim Extension As String, LastFile As String
    Dim speed
    LastFile = fisier
    If MediaStatus(pista) = "STOPPED" Or MediaStatus(pista) = "PLAYING" Or MediaStatus(pista) = "PAUSED" Then
        inchideSunet pista
    End If

    If fisier <> "" Then
        mciSendString "open " + ShortPath(fisier) + " alias " & pista, "", 0, 0
        Extension = UCase$(Right$(ShortPath(fisier), 3))
        If Extension = "DAT" Or Extension = "MPG" Or Extension = "MOV" Or Extension = "AVI" Or Extension = "MP2" Then
            mciSendString "put " & pista & " window at 0 0 800 480", "", 0, 0
        End If
        speed = 1000
    Else
        fisier = LastFile    '
    End If
End Function

Function PAUZA(pista As String)
    If MediaStatus(pista) = "PLAYING" Then
        mciSendString "pause " & pista, "", 0, 0
    End If
End Function

Function Play(pista As String)
    If MediaStatus(pista) = "STOPPED" Then
        mciSendString "play " & pista, "", 0, 0
    ElseIf MediaStatus(pista) = "PAUSED" Then
        mciSendString "resume " & pista, "", 0, 0
    End If
End Function


Function ShortPath(sPath As String) As String

    Dim ShortPathName As String
    If Right$(sPath, 1) <> "\" Then
        sPath = sPath & "\"
    End If
    ShortPathName = Space$(255)
    GetShortPathName sPath, ShortPathName, Len(ShortPathName)
    ShortPath = Left$(ShortPathName, Len(Trim$(ShortPathName)) - 2)
End Function




Public Function AdaugaTranse()
    Dim DATATRANSAMARE
    Dim TRANSAMARE
    TRANSAMARE = DMax("TRANSA", "14 1 Date exporturi")
    DATATRANSAMARE = DMax("data_export", "14 1 Date exporturi")
    Dim DataCurenta As Date
    Dim dataViitoare As Date
    Dim numarZile As Integer
    Dim contor As Integer
    DataCurenta = Date
    numarZile = 1
    contor = 0
    Do While contor < 8
        dataViitoare = DataCurenta + numarZile
        If Weekday(dataViitoare) = vbFriday Then
            contor = contor + 1
            If dataViitoare > DATATRANSAMARE Then
                CurrentDb.Execute "INSERT INTO [14 1 Date exporturi] ( Data_Export, Transa ) SELECT '" & dataViitoare & "' AS Expr1, " & TRANSAMARE + 1 & " AS Expr2;"
                TRANSAMARE = TRANSAMARE + 1
            End If
        End If
        numarZile = numarZile + 1
    Loop
End Function
Public Sub SCRIEMODIFICARE(IDINREG As Long, Tabel As String, COMANDA As String, Camp As String, initial As String, modificat As String)
10  On Error GoTo TRATARE_ERORI
    Dim SirEroare As String
20  CurrentDb.Execute "INSERT INTO [CAMPURI MODIFICATE] ( idInreg, tabel, comanda, camp, initial, modificat, ora, persoana )" & _
                      " VALUES (" & IDINREG & ",'" & Tabel & "','" & COMANDA & "','" & Camp & "','" & initial & "','" & modificat & "', Now(),'" & PERSOANA_ACTIVA & "')", dbSeeChanges
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
30  Exit Sub

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
40  lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
50  Select Case errNumar
    Case 0
60  Case Else
70      ScrieEroare "Eroare in [functii].[SCRIEMODIFICARE] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
80      RaspunsMesaj = MsgBox("[Eroare in functii].[SCRIEMODIFICARE] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[SCRIEMODIFICARE] linia " & lngLinia & " cu numarul " & Errnumar & vbCrLf & Errdescriere"
90      If RaspunsMesaj = vbYes Then
100         Resume Next
110     Else
120         GoTo TRATARE_ERORI_iesire
130     End If
140 End Select
    '========================== terminat tratare erori

End Sub


Function testCALCULNECESARPERAMASE()
    CALCULNECESARPERAMASE "HDI-688-46", "SSBT"
End Function


Public Function CALCULNECESARPERAMASE(COMANDA1 As String, produs1 As String) As String
    Dim cantitatecomandata As Long
    Dim cantitateExecutata As Long
    Dim cantitateRamasa As Long
    Dim VolumTimp As Long
    'Dim volumTimpPeBucata As Long
    Dim TimpRamas As Long
    Dim TimpREGATIRE As Long
    Dim PregatiriNecesare As Long
    Dim TIMPTOTAL As Long
    Dim Articol As String
    Dim ore, minutee, ZILEE

    cantitatecomandata = DLookup("cant", "comenzi", "nrcomanda='" & COMANDA1 & "'")
    Articol = DLookup("nrart", "comenzi", "nrcomanda='" & COMANDA1 & "'")
    Select Case UCase(Left(produs1, 1))
    Case "S"

        '''        VolumTimp = Nz(DLookup("NORMA", "7 NORME", "PRODUS='" & PRODUS1 & "'"), 0)
        '''        If VolumTimp = 0 Then VolumTimp = 5
        '''        VolumTimp = 28800 / (VolumTimp * 3600)
        VolumTimp = 5
    Case Else
        VolumTimp = Nz(DLookup("volumtimp", "Nomenclator articole", "nrart='" & Articol & "'"), 0)
    End Select
    TimpREGATIRE = Nz(DLookup("pregatire_secunde", "TIMP DE PREGATIRE MASINA", "ARTICOL='" & produs1 & "'"), 0)
    cantitateExecutata = Nz(DSum("cant", "intrari", "((produs) = '" & produs1 & "') AND ((nume)='" & COMANDA1 & "')"), 0)
    cantitateRamasa = cantitatecomandata - cantitateExecutata
    If cantitateRamasa > 0 Then
        If VolumTimp > 0 Then
            PregatiriNecesare = CLng(((cantitateRamasa * VolumTimp)) / ((8 * 3600) + TimpREGATIRE))
        Else
            PregatiriNecesare = 1
        End If
        '    If PregatiriNecesare = 0 Then PregatiriNecesare = 1
        'TimpRamas = cantitateRamasa * VolumTimp + PregatiriNecesare * TimpREGATIRE
        TimpRamas = cantitateRamasa * VolumTimp + TimpREGATIRE
        ZILEE = Int(TimpRamas / 28800)
        ore = Int(TimpRamas / 3600)
        ore = ore - ZILEE * 8
        TimpRamas = TimpRamas - ZILEE * 28800 - ore * 3600
        minutee = Format(CStr(Int((TimpRamas) / 60)), "0#")
        CALCULNECESARPERAMASE = IIf(ZILEE > 0, ZILEE & "Z:", "") & IIf(ore > 0, ore & "h:", "") & minutee & "m"
    Else
        CALCULNECESARPERAMASE = "0"
    End If
End Function

Public Function CALCULNECESARPECOMANDA(COMANDA1 As String, produs1 As String) As String
    Dim cantitatecomandata As Long
    Dim cantitateExecutata As Long
    Dim cantitateRamasa As Long
    Dim VolumTimp As Long
    'Dim volumTimpPeBucata As Long
    Dim TimpRamas As Long
    Dim TimpREGATIRE As Long
    Dim PregatiriNecesare As Long
    Dim TIMPTOTAL As Long
    Dim Articol As String
    Dim ore, minutee, ZILEE

    cantitatecomandata = DLookup("cant", "comenzi", "nrcomanda='" & COMANDA1 & "'")
    Articol = DLookup("nrart", "comenzi", "nrcomanda='" & COMANDA1 & "'")
    Select Case UCase(Left(produs1, 1))
    Case "S"

        '''        VolumTimp = Nz(DLookup("NORMA", "7 NORME", "PRODUS='" & PRODUS1 & "'"), 0)
        '''        If VolumTimp = 0 Then VolumTimp = 5
        '''        VolumTimp = 28800 / (VolumTimp * 3600)
        VolumTimp = 5
    Case Else
        VolumTimp = Nz(DLookup("volumtimp", "Nomenclator articole", "nrart='" & Articol & "'"), 0)
    End Select
    TimpREGATIRE = Nz(DLookup("pregatire_secunde", "TIMP DE PREGATIRE MASINA", "ARTICOL='" & produs1 & "'"), 0)
'    cantitateExecutata = Nz(DSum("cant", "intrari", "((produs) = '" & produs1 & "') AND ((nume)='" & COMANDA1 & "')"), 0)
    cantitateRamasa = cantitatecomandata ' - cantitateExecutata
    If cantitateRamasa > 0 Then
        If VolumTimp > 0 Then
            PregatiriNecesare = CLng(((cantitateRamasa * VolumTimp)) / ((8 * 3600) + TimpREGATIRE))
        Else
            PregatiriNecesare = 1
        End If
        '    If PregatiriNecesare = 0 Then PregatiriNecesare = 1
        'TimpRamas = cantitateRamasa * VolumTimp + PregatiriNecesare * TimpREGATIRE
        TimpRamas = cantitateRamasa * VolumTimp + TimpREGATIRE
        ZILEE = Int(TimpRamas / 28800)
        ore = Int(TimpRamas / 3600)
        ore = ore - ZILEE * 8
        TimpRamas = TimpRamas - ZILEE * 28800 - ore * 3600
        minutee = Format(CStr(Int((TimpRamas) / 60)), "0#")
        CALCULNECESARPECOMANDA = IIf(ZILEE > 0, ZILEE & "Z:", "") & IIf(ore > 0, ore & "h:", "") & minutee & "m"
    Else
        CALCULNECESARPECOMANDA = "0"
    End If
End Function

Public Function ZZHHMM(DURATA As Long) As String
          Dim ore, minutee, ZILEE
10    On Error GoTo TRATARE_ERORI
      Dim SirEroare As String
20        ZILEE = Int(DURATA / 28800)
30        ore = Int(DURATA / 3600)
40        ore = ore - ZILEE * 8
50        DURATA = DURATA - ZILEE * 28800 - ore * 3600
60        minutee = Format(CStr(Int((DURATA) / 60)), "0#")
70        ZZHHMM = IIf(ZILEE > 0, ZILEE & "z:", "") & IIf(ore > 0, ore & "h:", "") & IIf(minutee > 0, minutee & "m", "")
80        If Right(ZZHHMM, 1) = ":" Then ZZHHMM = Left(ZZHHMM, Len(ZZHHMM) - 1)
      '========================== incep tratare erori
TRATARE_ERORI_iesire:
           'DBEngine.Rollback
90        Exit Function

TRATARE_ERORI:
      Dim lngLinia As Long, errNumar As Long, errDescriere As String
100   lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
                Dim RaspunsMesaj As String
110          Select Case errNumar
                Case 0
120          Case Else
130              ScrieEroare "Eroare in [functii].[ZZHHMM] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
140              RaspunsMesaj = MsgBox("[Eroare in functii].[ZZHHMM] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
                 'Executasiraspunde="INFORMATIE "[Eroare in functii].[ZZHHMM] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
150               If RaspunsMesaj = vbYes Then
160                   Resume Next
170                   Else
180                   GoTo TRATARE_ERORI_iesire
190               End If
200          End Select
      '========================== terminat tratare erori
End Function


Public Function TimpLucrat(COMANDA As String, DATA As Date, faza As String, PERSOANA As String, SCRIE As Boolean) As Long

    Dim rec As DAO.Recordset
    Dim idComanda As Long
    Dim idSituatie As Long
    Dim marca As Long
    Dim N As Integer
    Dim SIR As String
10  On Error GoTo TRATARE_ERORI
    Dim SirEroare As String
    'GoTo 20
20  idComanda = DLookup("NR_COMANDA_INTERNA", "COMENZI", "nrcomanda='" & COMANDA & "'")
    'If lngLinia = 140 then' CORECTEAZA- Case 2427   < 10.10.2025 07:03:38 > stocuri
30  idSituatie = DLookup("ID", "DBO_SITUATIE COMENZI", "COMANDA=" & idComanda & " AND PRODUS='" & faza & "'")
40  marca = Nz(DLookup("ID", "PERSONAL", "NUMESIPRENUME LIKE '" & PERSOANA & "*'"), 0)

50  SIR = "SELECT inceput,terminat FROM [23 PONTARI] WHERE PERSOANA LIKE '" & PERSOANA & "*' AND COMANDA='" & COMANDA & "' AND Format([INCEPUT],'yyyy-mm-dd')='" & Format(DATA, "yyyy-mm-dd") & "' AND FAZA='" & faza & "'"

    ''Debug.Print sir
60  Set rec = CurrentDb.OpenRecordset(SIR, dbOpenDynaset, dbSeeChanges)
70  If Not rec.EOF Then
80      rec.MoveLast
90      rec.MoveFirst
100     For N = 1 To rec.RecordCount
110         If Not IsNull(rec.Fields("INCEPUT")) And Not IsNull(rec.Fields("TERMINAT")) Then
120             TimpLucrat = TimpLucrat + 60 * DURATA(rec.Fields("INCEPUT"), rec.Fields("TERMINAT"))
130         End If
140         rec.MoveNext
150     Next
160     rec.Close
170     Set rec = Nothing
180     If SCRIE Then
190         If TimpLucrat <> 0 Then CurrentDb.Execute "UPDATE [dbo_Situatie comenzi Persoane] SET ORE=" & TimpLucrat & " WHERE DATA=" & DataPentruSQL(Date) & " AND IDSITUATIE=" & idSituatie & " AND IDAMBALATOR=" & marca
200     Else
            ''200         MsgBox TimpLucrat

210     End If
220 End If
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
230 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
240 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
250 Select Case errNumar
    Case 0

260 Case 94

VerificProdusInSituatieComenzi

280     GoTo 20

290 Case Else
300     ScrieEroare "Eroare in [functii].[TimpLucrat] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
310     RaspunsMesaj = MsgBox("[Eroare in functii].[TimpLucrat] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[TimpLucrat] linia " & lngLinia & " cu numarul " & Errnumar & vbCrLf & Errdescriere"
320     If RaspunsMesaj = vbYes Then
330         Resume Next
340     Else
350         GoTo TRATARE_ERORI_iesire
360     End If
370 End Select
    '========================== terminat tratare erori

End Function

Public Function fnDurataComanda(DATA As Date, idAmbalator As Integer, idSituatie As Long) As Long
    Dim timpOcupat As Long
    Dim TimpComanda As String
    Dim COMANDA, produs
10  On Error GoTo TRATARE_ERORI
    Dim SirEroare As String
fnDurataComanda = 0
20  COMANDA = DLookup("comanda", "dbo_Situatie comenzi", "id=" & idSituatie)
30  produs = DLookup("produs", "dbo_Situatie comenzi", "id=" & idSituatie)
40  COMANDA = DLookup("nrcomanda", "comenzi", "nr_comanda_interna=" & COMANDA)
50  timpOcupat = Nz(DSum("ore", "dbo_Situatie comenzi Persoane", "data=" & DataPentruSQL(DATA) & " and idambalator=" & idAmbalator), 0)
60  TimpComanda = Nz(DLookup("ore", "planificare calculata", "nrcomanda='" & COMANDA & "' and (produs='HF')"), "")
70  If TimpComanda = "" Then
80      TimpComanda = Nz(DLookup("ore", "planificare calculata", "nrcomanda='" & COMANDA & "' and (produs='" & produs & "')"), "0")
90  End If
100 fnDurataComanda = SecundeDinTimp(TimpComanda)
110 If fnDurataComanda > 28800 - timpOcupat Then
120     fnDurataComanda = 28800 - timpOcupat
130 End If
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
140 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
150 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
160 Select Case errNumar
    Case 0
170 Case Else
180     ScrieEroare "Eroare in [functii].[fnDurataComanda] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
190     RaspunsMesaj = MsgBox("[Eroare in functii].[fnDurataComanda] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[fnDurataComanda] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
200     If RaspunsMesaj = vbYes Then
210         Resume Next
220     Else
230         GoTo TRATARE_ERORI_iesire
240     End If
250 End Select
    '========================== terminat tratare erori
End Function
Public Function SecundeDinTimp(TIMP As String) As Long
    Dim SIR As String
    Dim pos1, pos2
10  On Error GoTo TRATARE_ERORI
    Dim SirEroare As String
    SecundeDinTimp = 0
20  SIR = TIMP
30  pos1 = InStr(1, SIR, "z:")
40  If pos1 >= 1 Then
50      SecundeDinTimp = Left(SIR, pos1 - 1) * 28800

60      SIR = Right(SIR, Len(SIR) - pos1 - 1)
70  End If
80  pos2 = InStr(1, SIR, "h:")
90  If pos2 >= 1 Then
100     SecundeDinTimp = SecundeDinTimp + Left(SIR, pos2 - 1) * 3600
110     SIR = Right(SIR, Len(SIR) - pos2 - 1)
120     SecundeDinTimp = SecundeDinTimp + Replace(SIR, "m", "") * 60
130 Else
140     SecundeDinTimp = SecundeDinTimp + Replace(SIR, "m", "") * 60
150 End If

160 If SecundeDinTimp > 28800 Then SecundeDinTimp = 28800
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
170 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
180 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
190 Select Case errNumar
    Case 0
200 Case Else
210     ScrieEroare "Eroare in [FUNCTII].[SecundeDinTimp] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
220     RaspunsMesaj = MsgBox("[FUNCTII].[SecundeDinTimp] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[FUNCTII].[SecundeDinTimp] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
230     If RaspunsMesaj = vbYes Then
240         Resume Next
250     Else
260         GoTo TRATARE_ERORI_iesire
270     End If
280 End Select
    '========================== terminat tratare erori

End Function

Public Function ActualizareRamaseComanda(comandaa As String, produss As String)
    Dim SIR As String
    Dim M As Integer
10  On Error GoTo TRATARE_ERORI
    Dim SirEroare As String
    Dim rec As DAO.Recordset
    Dim recAZI As DAO.Recordset
    Dim CateSarje, CatePersoaneAzi, CatePersoaneMaine, CateFacute, CateRamase, VolumTimp, TimpREGATIRE, TimpRamas, ZILEE, ore
    Dim minutee, CALCULNECESARPERAMASE
    Dim totalAzi
    Dim MaiRaman


20  Select Case UCase$(Left(produss, 1))

    Case "S"

30      Set rec = CurrentDb.OpenRecordset("SELECT * FROM [planificare calculata] where nrcomanda='" & comandaa & "' and (produs='" & produss & "' OR PRODUS='HF') order by prioritate", dbOpenDynaset, dbSeeChanges)
40      If Not rec.EOF Then
50          rec.MoveLast
60          rec.MoveFirst

70          totalAzi = 0
            ''                 If rec.Fields("NRCOMANDA") = "PCL-074-2A" Then Stop



80          CatePersoaneAzi = Nz(DCount("id", "dbo_Situatie comenzi persoane", "idsituatie=" & rec.Fields("id") & " and data=date()"), 0)
90          CatePersoaneMaine = Nz(DCount("id", "dbo_Situatie comenzi persoane", "idsituatie=" & rec.Fields("id") & " and data>date()"), 0)

100         CateFacute = Nz(DSum("cant", "intrari", "nume='" & rec.Fields("NRCOMANDA") & "' and produs like 's*'"), 0)
110         CateRamase = rec.Fields("cant") - CateFacute
120         MaiRaman = CateRamase
130         VolumTimp = Nz(DLookup("volumtimp", "Nomenclator articole", "nrart='" & Replace(rec.Fields("PRODUS"), "HF", "ST") & rec.Fields("NRART") & "'"), 0)
140         If VolumTimp = 0 Then
150             MsgBox rec.Fields("NRART") & " nu are Timp de realizare semiprodus stabilit. Scrieti-l si reluati operatia"
160             Edretete Nz(rec.Fields("NRART"), "*")
170             VolumTimp = 6
180         End If
190         TimpREGATIRE = Nz(DLookup("pregatire_secunde", "TIMP DE PREGATIRE MASINA", "ARTICOL='" & rec.Fields("PRODUS") & "'"), 0)
200         If CateRamase > 0 Then
210             TimpRamas = CateRamase * VolumTimp + TimpREGATIRE
220             TimpRamas = TimpRamas - CatePersoaneMaine * 8 * 60 * 60
230             MaiRaman = MaiRaman - Int((CatePersoaneMaine * (8 * 60 * 60 - TimpREGATIRE)) / VolumTimp)
240             If TimpRamas < 0 Then TimpRamas = 0
                ' ATENTIUNE- calculeaza la CatePersoaneAzi  cate pot face de la intrare pana la 15:30 sau daca se termina inainte de 15:30
250             If CatePersoaneAzi > 0 Then
260                 SIR = "SELECT [23 PONTARI].INCEPUT" & _
                          " FROM [23 PONTARI]" & _
                          " WHERE ((([23 PONTARI].[COMANDA])='" & rec.Fields("NRCOMANDA") & "') AND ((Format([INCEPUT],'dd\.mm\.yy'))=Format(Date(),'dd\.mm\.yy')) AND (([23 PONTARI].TERMINAT) Is Null));"

270                 Set recAZI = CurrentDb.OpenRecordset(SIR, dbOpenDynaset, dbSeeChanges)
280                 If Not recAZI.EOF Then
290                     recAZI.MoveLast
300                     recAZI.MoveFirst
310                     For M = 1 To recAZI.RecordCount
320                         totalAzi = totalAzi + DURATA(Format(recAZI.Fields("inceput"), "HH:mm"), "15:30") * 60

330                         MaiRaman = MaiRaman - Int(DURATA(Format(recAZI.Fields("inceput"), "HH:mm"), "15:30") * 60 / VolumTimp)
340                         If MaiRaman < 0 Then MaiRaman = 0
350                         recAZI.MoveNext
360                     Next
370                     TimpRamas = TimpRamas - totalAzi

380                     If TimpRamas < 0 Then TimpRamas = 0
390                     ZILEE = Int(TimpRamas / 28800)
400                     ore = Int(TimpRamas / 3600)
410                     ore = ore - ZILEE * 8
420                     TimpRamas = TimpRamas - ZILEE * 28800 - ore * 3600
430                     minutee = Format(CStr(Int((TimpRamas) / 60)), "0#")
440                     CALCULNECESARPERAMASE = IIf(ZILEE > 0, ZILEE & "Z:", "") & IIf(ore > 0, ore & "h:", "") & minutee & "m"
450                 Else
460                     If TimpRamas < 0 Then TimpRamas = 0
470                     ZILEE = Int(TimpRamas / 28800)
480                     ore = Int(TimpRamas / 3600)
490                     ore = ore - ZILEE * 8
500                     TimpRamas = TimpRamas - ZILEE * 28800 - ore * 3600
510                     minutee = Format(CStr(Int((TimpRamas) / 60)), "0#")
520                     CALCULNECESARPERAMASE = IIf(ZILEE > 0, ZILEE & "Z:", "") & IIf(ore > 0, ore & "h:", "") & minutee & "m"
530                 End If

540             Else
550                 If TimpRamas < 0 Then TimpRamas = 0
560                 ZILEE = Int(TimpRamas / 28800)
570                 ore = Int(TimpRamas / 3600)
580                 ore = ore - ZILEE * 8
590                 TimpRamas = TimpRamas - ZILEE * 28800 - ore * 3600
600                 minutee = Format(CStr(Int((TimpRamas) / 60)), "0#")
610                 CALCULNECESARPERAMASE = IIf(ZILEE > 0, ZILEE & "Z:", "") & IIf(ore > 0, ore & "h:", "") & minutee & "m"
620             End If

630             rec.Edit
640             rec.Fields("plan") = CatePersoaneAzi + CatePersoaneMaine
650             rec.Fields("ramas") = IIf(MaiRaman > 0, MaiRaman, 0)
660             rec.Fields("ore") = CALCULNECESARPERAMASE
670             rec.Fields("AZI") = CatePersoaneAzi
680             rec.Fields("MAINE") = CatePersoaneMaine
690             rec.Update
700         End If

710     Else
720         MsgBox "Comanda " & comandaa & " pt semiprodus, a trecut de faza de planificare. Nu se mai actualizeaza timpul ramas pentru aceasta."
730     End If

780 Case "P"

790     Set rec = CurrentDb.OpenRecordset("SELECT * FROM [planificare calculata] where nrcomanda='" & comandaa & "' and (produs='" & produss & "' OR PRODUS='hf')order by prioritate", dbOpenDynaset, dbSeeChanges)
800     If Not rec.EOF Then
810         rec.MoveLast
820         rec.MoveFirst

830         totalAzi = 0
840         CatePersoaneAzi = Nz(DCount("id", "dbo_Situatie comenzi persoane", "idsituatie=" & rec.Fields("id") & " and data=date()"), 0)
850         CatePersoaneMaine = Nz(DCount("id", "dbo_Situatie comenzi persoane", "idsituatie=" & rec.Fields("id") & " and data>date()"), 0)
860         CateFacute = Nz(DSum("cant", "intrari", "nume='" & rec.Fields("NRCOMANDA") & "' and produs like 'p*'"), 0)
870         CateRamase = rec.Fields("cant") - CateFacute
880         MaiRaman = CateRamase
890         VolumTimp = Nz(DLookup("volumtimp", "Nomenclator articole", "nrart='" & rec.Fields("NRART") & "'"), 0)
900         If VolumTimp = 0 Then
910             MsgBox rec.Fields("NRART") & " nu are Timp de realizare produs stabilit. Scrieti-l si reluati operatia"
920             Edretete Nz(rec.Fields("NRART"), "*")
930             Exit Function
940         End If

950         TimpREGATIRE = Nz(DLookup("pregatire_secunde", "TIMP DE PREGATIRE MASINA", "ARTICOL='" & rec.Fields("PRODUS") & "'"), 0)
960         If CateRamase > 0 Then
970             TimpRamas = CateRamase * VolumTimp + TimpREGATIRE
980             TimpRamas = TimpRamas - CatePersoaneMaine * 8 * 60 * 60
990             MaiRaman = MaiRaman - Int((CatePersoaneMaine * (8 * 60 * 60 - TimpREGATIRE)) / VolumTimp)
1000            If TimpRamas < 0 Then TimpRamas = 0
1010            If CatePersoaneAzi > 0 Then
1020                SIR = "SELECT [23 PONTARI].INCEPUT" & _
                          " FROM [23 PONTARI]" & _
                          " WHERE ((([23 PONTARI].[COMANDA])='" & rec.Fields("NRCOMANDA") & "') AND ((Format([INCEPUT],'dd\.mm\.yy'))=Format(Date(),'dd\.mm\.yy')) AND (([23 PONTARI].TERMINAT) Is Null));"

1030                Set recAZI = CurrentDb.OpenRecordset(SIR, dbOpenDynaset, dbSeeChanges)
1040                If Not recAZI.EOF Then
1050                    recAZI.MoveLast
1060                    recAZI.MoveFirst
1070                    For M = 1 To recAZI.RecordCount
1080                        totalAzi = totalAzi + DURATA(Format(recAZI.Fields("inceput"), "HH:mm"), "15:30") * 60
1090                        MaiRaman = MaiRaman - Int(DURATA(Format(recAZI.Fields("inceput"), "HH:mm"), "15:30") * 60 / VolumTimp)
1100                        If MaiRaman < 0 Then MaiRaman = 0
1110                        recAZI.MoveNext
1120                    Next
1130                    TimpRamas = TimpRamas - totalAzi

1140                    If TimpRamas < 0 Then TimpRamas = 0
1150                    ZILEE = Int(TimpRamas / 28800)
1160                    ore = Int(TimpRamas / 3600)
1170                    ore = ore - ZILEE * 8
1180                    TimpRamas = TimpRamas - ZILEE * 28800 - ore * 3600
1190                    minutee = Format(CStr(Int((TimpRamas) / 60)), "0#")
1200                    CALCULNECESARPERAMASE = IIf(ZILEE > 0, ZILEE & "Z:", "") & IIf(ore > 0, ore & "h:", "") & minutee & "m"
1210                Else
1220                    If TimpRamas < 0 Then TimpRamas = 0
1230                    ZILEE = Int(TimpRamas / 28800)
1240                    ore = Int(TimpRamas / 3600)
1250                    ore = ore - ZILEE * 8
1260                    TimpRamas = TimpRamas - ZILEE * 28800 - ore * 3600
1270                    minutee = Format(CStr(Int((TimpRamas) / 60)), "0#")
1280                    CALCULNECESARPERAMASE = IIf(ZILEE > 0, ZILEE & "Z:", "") & IIf(ore > 0, ore & "h:", "") & minutee & "m"
1290                End If

1300            Else


1310                If TimpRamas < 0 Then TimpRamas = 0
1320                ZILEE = Int(TimpRamas / 28800)
1330                ore = Int(TimpRamas / 3600)
1340                ore = ore - ZILEE * 8
1350                TimpRamas = TimpRamas - ZILEE * 28800 - ore * 3600
1360                minutee = Format(CStr(Int((TimpRamas) / 60)), "0#")
1370                CALCULNECESARPERAMASE = IIf(ZILEE > 0, ZILEE & "Z:", "") & IIf(ore > 0, ore & "h:", "") & minutee & "m"

1380            End If

1390            rec.Edit
1400            rec.Fields("plan") = CatePersoaneAzi + CatePersoaneMaine
1410            rec.Fields("ramas") = IIf(MaiRaman > 0, MaiRaman, 0)
1420            rec.Fields("ore") = CALCULNECESARPERAMASE
1430            rec.Fields("AZI") = CatePersoaneAzi
1440            rec.Fields("MAINE") = CatePersoaneMaine
1450            rec.Update
1460        End If
1470    Else
1480        MsgBox "Comanda " & comandaa & " pt produs, a trecut de faza de planificare. Nu se mai actualizeaza timpul ramas pentru aceasta."
1490    End If


1540 End Select
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
1550 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
1560 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
1570 Select Case errNumar
    Case 0
1580 Case 3265
1590    Resume Next
1600 Case Else
1610    ScrieEroare "Eroare in [Form_Prioritizare comenzi].[ActualizareRamaseComanda] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
1620    RaspunsMesaj = MsgBox("[Eroare in Form_Prioritizare comenzi].[ActualizareRamaseComanda] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in Form_Prioritizare comenzi].[ActualizareRamaseComanda] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
1630    If RaspunsMesaj = vbYes Then
1640        Resume Next
1650    Else
1660        GoTo TRATARE_ERORI_iesire
1670    End If
1680 End Select
    '========================== terminat tratare erori
End Function






Public Function CALCULRamasPtPlanificareAmbalator(COMANDA1 As String, produs1 As String) As Long
    Dim cantitatecomandata As Long
    Dim cantitateExecutata As Long
    Dim cantitateRamasa As Long
    Dim VolumTimp As Long
    'Dim volumTimpPeBucata As Long
    Dim TimpRamas As Long
    Dim TimpREGATIRE As Long
    Dim PregatiriNecesare As Long
    Dim TIMPTOTAL As Long
    Dim Articol As String
    Dim ore, minutee, ZILEE

10  On Error GoTo TRATARE_ERORI
    Dim SirEroare As String
20  cantitatecomandata = DLookup("cant", "comenzi", "nrcomanda='" & COMANDA1 & "'")
30  Articol = DLookup("nrart", "comenzi", "nrcomanda='" & COMANDA1 & "'")
40  VolumTimp = Nz(DLookup("volumtimp", "Nomenclator articole", "nrart='" & Articol & "'"), 0)

50  TimpREGATIRE = Nz(DLookup("pregatire_secunde", "TIMP DE PREGATIRE MASINA", "ARTICOL='" & produs1 & "'"), 0)
60  cantitateExecutata = Nz(DSum("cant", "intrari", "((produs) = '" & produs1 & "') AND ((nume)='" & COMANDA1 & "')"), 0)
70  cantitateRamasa = cantitatecomandata - cantitateExecutata
80  If cantitateRamasa > 0 Then
90      If VolumTimp > 0 Then
100         PregatiriNecesare = CLng(((cantitateRamasa * VolumTimp)) / ((8 * 3600) + TimpREGATIRE))
110     Else
120         PregatiriNecesare = 1
130     End If
        '    If PregatiriNecesare = 0 Then PregatiriNecesare = 1
        'TimpRamas = cantitateRamasa * VolumTimp + PregatiriNecesare * TimpREGATIRE
140     TimpRamas = cantitateRamasa * VolumTimp + TimpREGATIRE
        '''        ZILEE = Int(TimpRamas / 28800)
        '''        ore = Int(TimpRamas / 3600)
        '''        ore = ore - ZILEE * 8
        '''        TimpRamas = TimpRamas - ZILEE * 28800 - ore * 3600
        '''        MINUTEE = format(CStr(Int((TimpRamas) / 60)), "0#")
        '''        CALCULNECESARPERAMASE = IIf(ZILEE > 0, ZILEE & "Z:", "") & IIf(ore > 0, ore & "h:", "") & MINUTEE & "m"

150     CALCULRamasPtPlanificareAmbalator = TimpRamas
160     If CALCULRamasPtPlanificareAmbalator > 28800 Then CALCULRamasPtPlanificareAmbalator = 28800
170 Else
180     CALCULRamasPtPlanificareAmbalator = 0
190 End If
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
200 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
210 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
220 Select Case errNumar
    Case 0
230 Case Else
240     ScrieEroare "Eroare in [functii].[CALCULRamasPtPlanificareAmbalator] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
250     RaspunsMesaj = MsgBox("[Eroare in functii].[CALCULRamasPtPlanificareAmbalator] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[CALCULRamasPtPlanificareAmbalator] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
260     If RaspunsMesaj = vbYes Then
270         Resume Next
280     Else
290         GoTo TRATARE_ERORI_iesire
300     End If
310 End Select
    '========================== terminat tratare erori
End Function






Public Function InRetea() As Boolean
10  On Error GoTo erori
    Dim SIR
    SIR = DLookup("valoare", "configurari", "proprietate='CautaRetea'")
20  If Dir(SIR, vbDirectory) <> "" Then
30      InRetea = True
40      blnINRETEA = True
50  Else
60      InRetea = False
70      blnINRETEA = False
80  End If
Iesire:
90  Exit Function
erori:
100 Select Case err.Number
    Case 0
110 Case Else
120     InRetea = False
130     blnINRETEA = False
        GoTo Iesire:
140 End Select
End Function

Public Function ResetRowNumber() As Boolean
    Set colPrimaryKeys = New VBA.Collection
    lngRowNumber = 0
    ResetRowNumber = True
End Function
Public Function RowNumber(UniqueKeyVariant As Variant) As Long

    Dim lngTemp As Long
    On Error Resume Next
    lngTemp = colPrimaryKeys(CStr(UniqueKeyVariant))
    If err.Number Then
        lngRowNumber = lngRowNumber + 1
        colPrimaryKeys.Add lngRowNumber, CStr(UniqueKeyVariant)
        lngTemp = lngRowNumber
    End If
    RowNumber = lngTemp
End Function
Public Function CaleUSB() As String
''    Dim FSO As Object
    Dim Drv As Object
10  On Error GoTo TRATARE_ERORI

20  CaleUSB = "nu"
    '''    Set FSO = CreateObject("Scripting.FileSystemObject") ' Declarat in STARTT
30  For Each Drv In fso.Drives
40      With Drv
            ''        Debug.Print .DriveLetter
50          If .IsReady And .DriveType = 1 Then
                '                MsgBox "Removable " & .DriveLetter
60              CaleUSB = .DriveLetter
70              Exit Function
80          End If
90      End With
100 Next Drv
    CaleUSB = "nu"
110 MsgBox "Nu exista stick USB conectat", vbSystemModal
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
120 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
130 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
140 Select Case errNumar
    Case 0
    Case 91
        GoTo TRATARE_ERORI_iesire
150 Case Else
160     ScrieEroare "Eroare in [functii].[CaleUSB] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netrarata functii rutina CaleUSB"
170     RaspunsMesaj = MsgBox("[Eroare in functii].[CaleUSB] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[CaleUSB] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
180     If RaspunsMesaj = vbYes Then
190         Resume Next
200     Else
210         GoTo TRATARE_ERORI_iesire
220     End If
230 End Select
    '========================== terminat tratare erori
End Function




Public Function fataverso(rptReport As Report)

    Dim prtLoop As Printer
    Dim defaultP As Printer
10  On Error GoTo TRATARE_ERORI
20  Set defaultP = Application.Printer
30  For Each prtLoop In Application.Printers
40      With prtLoop
50          Application.Printer = prtLoop
60          If rptReport.Printer.Orientation = acPRORLandscape Then
70              prtLoop.duplex = acPRDPVertical
80              Application.Printer.duplex = acPRDPVertical
90          Else
100             prtLoop.duplex = acPRDPHorizontal
110             Application.Printer.duplex = acPRDPHorizontal
120         End If
130     End With
140 Next

150 Application.Printer = defaultP
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
160 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
170 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
180 Select Case errNumar
    Case 0
190 Case Else
200     ScrieEroare "Eroare in [functii].[fataverso] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata functii rutina fataverso"
210     RaspunsMesaj = MsgBox("[Eroare in functii].[fataverso] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[fataverso] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
220     If RaspunsMesaj = vbYes Then
230         Resume Next
240     Else
250         GoTo TRATARE_ERORI_iesire
260     End If
270 End Select
    '========================== terminat tratare erori

End Function

Public Function FPERSONA_ACTIVA() As String
    FPERSONA_ACTIVA = PERSOANA_ACTIVA
End Function

'''Public Sub SetFormOpacity(frm As Form, trns As Long, sngOpacity As Single)
'''    Dim lngStyle As Long
'''
'''    ' get the current window style, then set transparency
'''    lngStyle = GetWindowLong(frm.hwnd, GWL_EXSTYLE)
'''    'SetWindowLong frm.hwnd, GWL_EXSTYLE, lngStyle Or WS_EX_LAYERED
'''    SetLayeredWindowAttributes frm.hwnd, trns, (sngOpacity), LWA_ALPHA Or LWA_COLORKEY
'''End Sub


Public Function BonuriDeLucru2(COMANDA As String, strSarja As String, strSarjaprodus As String)
10  On Error GoTo TRATARE_ERORI

    Dim stDocName As String
    Dim rec As DAO.Recordset
    Dim caant As DAO.Recordset
    Dim N, tot, mm

20  Set caant = CurrentDb.OpenRecordset("SELECT COMENZI.cant, COMENZI.NRCOMANDA, COMENZI.NRART, CInt(" & Forms![10 COMENZI SIMPLA].Form.Controls("Sarje").Form.Controls("CANT").Value & "/[7 NORME].[nrPeBon]) AS NUMAR" & _
                                        " FROM (COMENZI LEFT JOIN [Nomenclator articole] ON COMENZI.NRART = [Nomenclator articole].NrArt) LEFT JOIN [7 NORME] ON [Nomenclator articole].produs = [7 NORME].PRODUS" & _
                                        " WHERE (((COMENZI.NRCOMANDA) Like '" & COMANDA & "'));", dbOpenDynaset, dbSeeChanges)

30  If Not caant.EOF Then
40      For N = 1 To caant.RecordCount
50          tot = MesajMod("Cite etichete doriti pentru comanda " & caant.Fields("nrcomanda"), IIf(CStr(Int(caant.Fields("NUMAR"))) <> "0", CStr(Int(caant.Fields("NUMAR"))), 1), "Adauga", "Inlocuieste", "Adauga si tipareste", "Inlocuieste si tipareste", 4)
60          If Casuta.InitText <> 0 Then
70              Select Case tot
                Case 1    ' "Adauga"
80                  Set rec = CurrentDb.OpenRecordset("BONURI DE LUCRU", dbOpenDynaset, dbSeeChanges)
90                  If Casuta.InitText <> "" And val(Casuta.InitText) < 100 Then
100                     For mm = 1 To CInt(Casuta.InitText)
110                         rec.AddNew
120                         rec.Fields("DATA_EMITERE") = Date
130                         rec.Fields("PERSOANA_EMITERE") = "Productie"
140                         rec.Fields("SARJA") = strSarja
150                         rec.Fields("SARJAprodus") = strSarjaprodus
160                         rec.Fields("NRCOMANDA") = COMANDA
170                         rec.Fields("NRART") = caant.Fields("NRART")


180                         rec.Fields("se_tipareste") = -1
190                         rec.Update
200                     Next
210                 End If
220                 rec.Close

230             Case 2    ' "Inlocuieste"
240                 Set rec = CurrentDb.OpenRecordset("BONURI DE LUCRU", dbOpenDynaset, dbSeeChanges)
250                 CurrentDb.Execute "update [bonuri de lucru] set se_tipareste=0 where PERSOANA_PRODUS ='" & PERSOANA_ACTIVA & "'", dbSeeChanges
260                 If Casuta.InitText <> "" And val(Casuta.InitText) < 100 Then
270                     For mm = 1 To CInt(Casuta.InitText)
280                         rec.AddNew
290                         rec.Fields("DATA_EMITERE") = Date
300                         rec.Fields("PERSOANA_EMITERE") = "Productie"
310                         rec.Fields("SARJA") = strSarja
320                         rec.Fields("SARJAprodus") = strSarjaprodus
330                         rec.Fields("NRCOMANDA") = COMANDA
340                         rec.Fields("NRART") = caant.Fields("NRART")

350                         rec.Fields("se_tipareste") = -1
360                         rec.Update
370                     Next
380                 End If
390                 rec.Close

400             Case 3    ' "Adauga si tipareste"
410                 Set rec = CurrentDb.OpenRecordset("BONURI DE LUCRU", dbOpenDynaset, dbSeeChanges)
420                 If Casuta.InitText <> "" And val(Casuta.InitText) < 100 Then
430                     For mm = 1 To CInt(Casuta.InitText)
440                         rec.AddNew
450                         rec.Fields("DATA_EMITERE") = Date
460                         rec.Fields("PERSOANA_EMITERE") = "Productie"
470                         rec.Fields("SARJA") = strSarja
480                         rec.Fields("SARJAprodus") = strSarjaprodus
490                         rec.Fields("NRCOMANDA") = COMANDA
500                         rec.Fields("NRART") = caant.Fields("NRART")

510                         rec.Fields("se_tipareste") = -1
520                         rec.Update
530                     Next
540                 End If
550                 rec.Close
560                 stDocName = "Q-EQ-7-5-3-01 BON DE LUCRU"
570                 DoCmd.OpenReport stDocName, acPreview, , "PERSOANA_PRODUS ='" & PERSOANA_ACTIVA & "'"
580             Case 4    ' "Inlocuieste si tipareste"
590                 Set rec = CurrentDb.OpenRecordset("BONURI DE LUCRU", dbOpenDynaset, dbSeeChanges)
600                 CurrentDb.Execute "update [bonuri de lucru] set se_tipareste=0 where PERSOANA_PRODUS ='" & PERSOANA_ACTIVA & "'", dbSeeChanges
610                 If Casuta.InitText <> "" And val(Casuta.InitText) < 100 Then
620                     For mm = 1 To CInt(Casuta.InitText)
630                         rec.AddNew
640                         rec.Fields("DATA_EMITERE") = Date
650                         rec.Fields("PERSOANA_EMITERE") = "Productie"
660                         rec.Fields("SARJA") = strSarja
670                         rec.Fields("SARJAprodus") = strSarjaprodus
680                         rec.Fields("NRCOMANDA") = COMANDA
690                         rec.Fields("NRART") = caant.Fields("NRART")

700                         rec.Fields("se_tipareste") = -1
710                         rec.Update

720                     Next
730                 End If
740                 rec.Close
750                 stDocName = "Q-EQ-7-5-3-01 BON DE LUCRU"
760                 DoCmd.OpenReport stDocName, acPreview, , "PERSOANA_PRODUS ='" & PERSOANA_ACTIVA & "'"
770             End Select
780         End If
790     Next
800 End If

    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
810 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
820 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
830 Select Case errNumar
    Case 0
840 Case 13
850     Exit Function
860 Case Else
870     ScrieEroare "Eroare in [functii].[BonuriDeLucru2] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata Form_10 COMENZI SIMPLA rutina BonuriDeLucru2"
880     RaspunsMesaj = MsgBox("[Eroare in functii].[BonuriDeLucru2] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[BonuriDeLucru2] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
890     If RaspunsMesaj = vbYes Then
900         Resume Next
910     Else
920         GoTo TRATARE_ERORI_iesire
930     End If
940 End Select
    '========================== terminat tratare erori
End Function





Public Function BackupServer()
    Dim SIR As String
    Dim acum As String
    Dim TIMP As String
    Dim qdf As DAO.QueryDef
10  On Error GoTo TRATARE_ERORI

20  TIMP = Now()
30  acum = Mid(TIMP, 7, 4) & "_" & Mid(TIMP, 4, 2) & "_" & Mid(TIMP, 1, 2) & "_" & Mid(TIMP, 12, 2) & Mid(TIMP, 15, 2) & Mid(TIMP, 18, 2) & "_0000000"
    'D:\Backup\SQLSTOC
    'C:\Program Files\Microsoft SQL Server\MSSQL10_50.MSSQLSERVER\MSSQL\Backup\SQLSTOC
40  SIR = "EXECUTE master.dbo.xp_create_subdir N'D:\Backup\SQLSTOC'" & vbCrLf & _
          " BACKUP DATABASE [SQLSTOC] TO  DISK = N'D:\Backup\SQLSTOC\SQLSTOC_COPIE_" & acum & ".bak' WITH NOFORMAT, NOINIT,  NAME = N'SQLSTOC_backup_" & acum & "', SKIP, REWIND, NOUNLOAD,  STATS = 10" & vbCrLf & _
          " declare @backupSetId as int" & vbCrLf & _
          " select @backupSetId = position from msdb..backupset" & _
          " where database_name = N'SQLSTOC' and backup_set_id=" & _
          "(select max(backup_set_id) from msdb..backupset where database_name=N'SQLSTOC' )" & vbCrLf & _
          " if @backupSetId is null" & _
          " begin raiserror(N'Verify failed. Backup information for database ''SQLSTOC'' not found.', 16, 1) end" & vbCrLf & _
          " RESTORE VERIFYONLY FROM  DISK = N'D:\Backup\SQLSTOC\SQLSTOC_COPIE_" & acum & ".bak'" & _
          " WITH  FILE = @backupSetId,  NOUNLOAD,  NOREWIND"
    'Debug.Print sir
50  Set qdf = CurrentDb.QueryDefs("SQL SERVER BackupTotal")
60  qdf.sql = SIR

    'If lngLinia = 70 then' CORECTEAZA- Case 3146
70  DoCmd.OpenQuery "SQL SERVER BackupTotal"
80  MsgBox "Terminat backup " & vbCrLf & "D:\Backup\SQLSTOC\SQLSTOC_COPIE_" & acum & ".bak"
    '90  DoCmd.OpenQuery "Administrare Backup-uri executate"
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
100 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
110 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
120 Select Case errNumar
    Case 0
130 Case Else
140     ScrieEroare "Eroare in [functii].[BackupServer] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata functii rutina BackupServer"
150     RaspunsMesaj = MsgBox("[Eroare in functii].[BackupServer] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[BackupServer] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
160     If RaspunsMesaj = vbYes Then
170         Resume Next
180     Else
190         GoTo TRATARE_ERORI_iesire
200     End If
210 End Select
    '========================== terminat tratare erori

End Function
Public Function Inlocuire(textul As String, Cauta As String, Inlocuitor As String) As String
    Dim pos

    Dim DREAPTA
    Dim STANGA
    ' lungime = Len(cauta)
    pos = 1
    If textul <> "" Then
        If InStr(pos, textul, Cauta) > 0 Then

10:             pos = InStr(pos, textul, Cauta)
            STANGA = Left$(textul, pos - 1)
            DREAPTA = Right(textul, Len(textul) - pos - Len(Cauta) + 1)
            textul = STANGA & Inlocuitor & DREAPTA
            'DUPA = textul
            If InStr(pos + Len(Inlocuitor), textul, Cauta) > 0 Then pos = pos + Len(Inlocuitor): GoTo 10

        End If

    End If
erori:
    Select Case err.Number

    Case 0
    Case 3258
        Resume Next
    End Select
End Function



Public Function Revizia(element As String) As String
    Dim rec As DAO.Recordset
    Set rec = CurrentDb.OpenRecordset("select max(versiunea) from versiuni where element ='" & element & "'")
    If Not rec.EOF Then

        Revizia = Nz(rec.Fields(0), "-")
    End If
    rec.Close
    Set rec = Nothing
End Function


Public Sub Asteapta(lngMilliSec As Long)
    If lngMilliSec > 0 Then
        Call sapiSleep(lngMilliSec)
    End If
End Sub
Public Function ShellEx( _
       ByVal sFile As String, _
       Optional ByVal sParameters As String = "", _
       Optional ByVal sDefaultDir As String = "", _
       Optional sOperation As String = "open", _
       Optional Owner As Long = 0 _
       ) As Boolean
    Dim lR    ' As Long
    Dim lErr As Long, sErr As Long
    If (InStr(UCase$(sFile), ".EXE") <> 0) Then
        'eShowCmd = 0
    End If
    On Error Resume Next
    If (sParameters = "") And (sDefaultDir = "") Then
        lR = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
    Else
        lR = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, 3)
    End If
    If (lR < 0) Or (lR > 32) Then
        ShellEx = True
    Else
        ' raise an appropriate error:
        lErr = vbObjectError + 1048 + lR
        Select Case lR
        Case 0
            lErr = 7: sErr = "Out of memory"
        Case ERROR_FILE_NOT_FOUND
            lErr = 53: sErr = "File not found"
        Case ERROR_PATH_NOT_FOUND
            lErr = 76: sErr = "Path not found"
        Case ERROR_BAD_FORMAT
            sErr = "The executable file is invalid or corrupt"
        Case SE_ERR_ACCESSDENIED
            lErr = 75: sErr = "Path/file access error"
        Case SE_ERR_ASSOCINCOMPLETE
            sErr = "This file type does not have a valid file association."
        Case SE_ERR_DDEBUSY
            lErr = 285: sErr = "The file could not be opened because the target application is busy.  Please try again in a moment."
        Case SE_ERR_DDEFAIL
            lErr = 285: sErr = "The file could not be opened because the DDE transaction failed.  Please try again in a moment."
        Case SE_ERR_DDETIMEOUT
            lErr = 286: sErr = "The file could not be opened due to time out.  Please try again in a moment."
        Case SE_ERR_DLLNOTFOUND
            lErr = 48: sErr = "The specified dynamic-link library was not found."
        Case SE_ERR_FNF
            lErr = 53: sErr = "File not found"
        Case SE_ERR_NOASSOC
            sErr = "No application is associated with this file type."
        Case SE_ERR_OOM
            lErr = 7: sErr = "Out of memory"
        Case SE_ERR_PNF
            lErr = 76: sErr = "Path not found"
        Case SE_ERR_SHARE
            lErr = 75: sErr = "A sharing violation occurred."
        Case Else
            sErr = "An error occurred occurred whilst trying to open or print the selected file."
        End Select

        err.Raise lErr, , "Eroare .GShell", sErr
        ShellEx = False
    End If

End Function







Function meniuuti()
    meniu utilizator
End Function
Function proprietati()
    On Error GoTo erori
    Dim N
    Dim ret As Boolean
    '    Debug.Print CurrentDb.Name
    For N = 1 To CurrentDb.TableDefs.Count

        If CurrentDb.TableDefs(N).Properties("Connect").Value <> "" Then

            ret = SetAccessProperty(CurrentDb.TableDefs(N), "Description", dbText, CurrentDb.TableDefs(N).Properties("Connect").Value)
            MsgBox CurrentDb.TableDefs(N).name
        End If
    Next
erori:
    Select Case err.Number
    Case 0
    Case 3270
    Case Else
        MsgBox err.Description & err.Number
        Resume Next
    End Select
End Function
Function SetAccessProperty(OBJ As Object, strName As String, _
                           intType As Integer, varSetting As Variant) As Boolean
    Dim prp As Property
    Const conPropNotFound As Integer = 3270

    On Error GoTo ErrorSetAccessProperty
    ' Explicitly refer to Properties collection.
    OBJ.Properties(strName) = varSetting
    OBJ.Properties.Refresh
    SetAccessProperty = True

ExitSetAccessProperty:
    Exit Function

ErrorSetAccessProperty:
    If err = conPropNotFound Then

        ' Create property, denote type, and set initial value.
        Set prp = OBJ.CreateProperty(strName, intType, varSetting)
        ' Append Property object to Properties collection.
        OBJ.Properties.Append prp
        OBJ.Properties.Refresh
        SetAccessProperty = True
        Resume ExitSetAccessProperty
    Else
        MsgBox err & ": " & vbCrLf & err.Description
        SetAccessProperty = False
        Resume ExitSetAccessProperty
    End If
End Function

'The following procedure calls the SetAccessProperty function to set the Subject property.

Sub CallPropertySet()
    Dim dbs As DAO.Database, ctr As CONTAINER, doc As Document
    Dim blnReturn As Boolean

    ' Return reference to current database.
    Set dbs = CurrentDb
    ' Return reference to Databases container.
    Set ctr = dbs.Containers!Databases
    ' Return reference to SummaryInfo document.
    Set doc = ctr.Documents!SummaryInfo
    blnReturn = SetAccessProperty(doc, _
                                  "Subject", dbText, "Business Contacts")
    ' Evaluate return value.
    If blnReturn = True Then

        ''        Debug.Print "Property set successfully."
    Else
        ''        Debug.Print "Property not set successfully."
    End If
End Sub
'"\\Server500gb\MagServer\Magazie\PROGRAM\CLISEE.mdb"
Function Edretete(FILT As String)
    Protocoale.retete (FILT)
End Function

'Function CLIS()

'Module1.CLISEU
'SysCmd acSysCmdSetStatus, "Colectie de clisee."
'End Function


Function IN_JURNAL(baza As String, tip As String, modul As String, MESAJ As String, operatie As String, NumarEroare As String, DescriereEroare As String)
'    Dim ret
'    ret = utilizator
'    CurrentDb.Execute "INSERT INTO operari (data,timp,modul,tip,utilizator,calculator,baza,mesaj,operatie,errnumar,errdesc) VALUES ('" & Date & "','" & time & "','" & modul & "','" & tip & "','" & ret & "','-','" & baza & "','" & MESAJ & "','" & operatie & "','" & numarEroare & "','" & DescriereEroare & "')", dbSeeChanges

End Function

Public Function COMPACT()
    Dim ret
    On Error GoTo erori
    CompactDatabase "C:\DAN\STOCURI.MDB", "C:\dan\COPII\" & month(Date) & "-" & day(Date) & "-" & year(Date) & " " & "STOCURI.MDB"
    Dim baz As DAO.Database
    Set baz = OpenDatabase("C:\dan\COPII\" & month(Date) & "-" & day(Date) & "-" & year(Date) & " " & "STOCURI.MDB")
    baz.Execute "UPDATE SETTINGS SET LIMBA='germana'", dbSeeChanges
    baz.Close
    Set baz = Nothing

    '3356Shell "C:\WinRAR\WinRAR.exe a C:\dan\COPII baza\" & Month(Date) & "-" & Day(Date) & "-" & Year(Date) & " " & "STOCURI.MDB" & ".rar" & "C:\dan\COPII baza\" & Month(Date) & "-" & Day(Date) & "-" & Year(Date) & " " & "STOCURI.MDB", 1
    shell "explorer C:\dan\COPII", vbMaximizedFocus
erori:

    Select Case err.Number
    Case 0
    Case 3356
        MsgBox "Exista copie a fisierului"
        Resume Next
    Case 3204
        Application.Assistant.visible = True
        Application.Assistant.Animation = 7

        ret = MsgBox("Copia bazei de date exista." & vbCrLf & "Doriti inlocuirea ei ?", vbYesNo + vbQuestion)
        Select Case ret
        Case vbYes
            Application.Assistant.visible = False
            Kill "C:\dan\COPII\" & month(Date) & "-" & day(Date) & "-" & year(Date) & " " & "STOCURI.MDB"
            CompactDatabase "C:\DAN\STOCURI.MDB", "C:\dan\COPII\" & month(Date) & "-" & day(Date) & "-" & year(Date) & " " & "STOCURI.MDB"
            Resume Next
        Case Else
            Application.Assistant.visible = False
            Exit Function
        End Select
    Case Else
        MsgBox "Atentie a aparut eroarea " & vbCrLf & err.Description & "  " & err.Number

    End Select
End Function

Function ETICHETA(SIR As String)
    Dim toate As String
    Dim POZITIA_SIR As String
    Dim DECALAJ As Integer
    Dim SIR_REZULTAT As String
    Dim ret As Integer
    Dim N As Integer
    toate = "ABCDEFGHIJKLMNOPQRS TUVWXYZ1234567890.- ABCDEFGHIJKLMNOPQRS TUVWXYZ1234567890.- "

    DECALAJ = 20
    If InStr(1, SIR, "/") Then
        ret = InStr(1, SIR, "/")
        Mid(SIR, ret, 1) = "-"
    End If
    For N = 1 To Len(SIR)
        POZITIA_SIR = Mid(SIR, N, 1)
        SIR_REZULTAT = SIR_REZULTAT & Mid(toate, InStr(1, toate, POZITIA_SIR) + DECALAJ, 1)
    Next
    ETICHETA = SIR_REZULTAT
End Function
Public Function EliminaCrlf(SIR As String) As String
    On Error Resume Next
    Dim start1 As Integer

    Dim part1, part2, LOC
    start1 = 1
    LOC = InStr(start1, SIR, vbCr)
    If LOC = 0 Then
        EliminaCrlf = SIR
        Exit Function
    Else
        Do
            LOC = InStr(start1, SIR, vbCr)
            If LOC <> 0 Then
                start1 = LOC
                part1 = Left(SIR, LOC - 1)
                part2 = Right(SIR, Len(SIR) - LOC)

            End If
            If LOC <> 0 Then
                SIR = part1 & part2

            Else
                'SIR = part1
            End If
        Loop Until LOC = 0
        EliminaCrlf = SIR
    End If
End Function
Public Function EliminaSpatiu(SIR As String) As String
    On Error Resume Next
    Dim start1 As Integer

    Dim part1, part2, LOC
    start1 = 1
    LOC = InStr(start1, SIR, " ")
    If LOC = 0 Then
        EliminaSpatiu = SIR
        Exit Function
    Else
        Do
            LOC = InStr(start1, SIR, " ")
            If LOC <> 0 Then
                start1 = LOC
                part1 = Left(SIR, LOC - 1)
                part2 = Right(SIR, Len(SIR) - LOC)

            End If
            If LOC <> 0 Then
                SIR = part1 & part2

            Else
                'SIR = part1
            End If
        Loop Until LOC = 0
        EliminaSpatiu = SIR
    End If
End Function
'''Function restProductie()
'''On Error Resume Next
'''Dim BAZA As DAO.DATABASE
'''    Dim stDocName As String
'''Set BAZA = CurrentDb
'''BAZA.Execute "DELETE * FROM Urmancom", dbSeeChanges
'''BAZA.Execute "DELETE * FROM urmantip", dbSeeChanges
'''BAZA.Execute "DELETE * FROM urmanumpl", dbSeeChanges
'''BAZA.Execute "DELETE * FROM REST", dbSeeChanges
'''BAZA.Close
'''Set BAZA = Nothing
'''Dim dbs As DAO.DATABASE
''''
'''Set dbs = CurrentDb
'''
'''dbs.Execute " INSERT INTO URMANCOM SELECT COMENZI.NRCOMANDA,COMENZI.DEN,SUM(COMENZI.CANT) AS COMANDA FROM COMENZI WHERE COMENZI.PRODUS LIKE 'P*' AND NRCOMANDA LIKE '*" & Right(CStr(Year(Date)), 2) & "' AND stadiu LIKE 'in lucru' GROUP BY COMENZI.NRCOMANDA,COMENZI.DEN"
'''dbs.Execute " INSERT INTO URMANTIP SELECT INTRARI.nume as nrcomanda,INTRARI.DEN,SUM(INTRARI.CANT) AS TIPARITE FROM intrari WHERE INTRARI.PRODUS LIKE 'S*' AND INTRARI.nume LIKE '*" & Right(CStr(Year(Date)), 2) & "' GROUP BY INTRARI.nume,INTRARI.DEN"
'''dbs.Execute " INSERT INTO URMANUMPL SELECT INTRARI.nume as nrcomanda,INTRARI.DEN,SUM(INTRARI.CANT) AS UMPLUTE FROM intrari WHERE INTRARI.PRODUS LIKE 'P*' AND INTRARI.nume like '*" & Right(CStr(Year(Date)), 2) & "' GROUP BY INTRARI.nume,INTRARI.DEN"
''''''''''BUN
'''dbs.Execute " INSERT INTO REST SELECT Urmancom.nrcomanda, Urmancom.den, Urmancom.COMANDA, urmantip.TIPARITE, urmanumpl.UMPLUTE, iif(urmantip.TIPARITE , urmancom.COMANDA - urmantip.TIPARITE, urmancom.COMANDA) AS detiparit, iif(Urmanumpl.umplute , urmancom.comanda - urmanumpl.umplute,urmancom.COMANDA) AS deumplut FROM ((Urmancom LEFT OUTER JOIN urmanumpl ON Urmancom.den = urmanumpl.den AND " & _
 '''" Urmancom.nrcomanda = urmanumpl.nrcomanda) LEFT OUTER JOIN urmantip ON Urmancom.den = urmantip.den AND Urmancom.nrcomanda = urmantip.nrcomanda) WHERE ((Urmancom.COMANDA > IIF(urmantip.TIPARITE, urmantip.TIPARITE, 0) AND Urmancom.COMANDA > IIF(urmanumpl.UMPLUTE, urmanumpl.UMPLUTE, 0)) OR (Urmancom.COMANDA > IIF(urmantip.TIPARITE, urmantip.TIPARITE, 0) AND" & _
 '''" Urmancom.COMANDA < IIF(urmanumpl.UMPLUTE, urmanumpl.UMPLUTE, 0)) OR (Urmancom.COMANDA < IIF(urmantip.TIPARITE, urmantip.TIPARITE, 0) AND Urmancom.COMANDA > IIF(urmanumpl.UMPLUTE, urmanumpl.UMPLUTE, 0))) ORDER BY Urmancom.NRCOMANDA"
'''dbs.Execute "delete * from urmancom", dbSeeChanges
'''dbs.Execute "delete * from rest where DEumplut=null and DEtiparit=null", dbSeeChanges
'''dbs.Execute "delete * from rest where DEumplut=0 and DEtiparit=0", dbSeeChanges
'''dbs.Execute "delete * from rest where DEumplut<=0 ", dbSeeChanges
'''dbs.Execute "insert into urmancom SELECT distinct nrcomanda from rest "
'''dbs.Execute "UPDATE REST " _
 '''& "SET DETIPARIT = '0'" _
 '''& "WHERE DETIPARIT < 0;"
'''dbs.Execute "UPDATE REST " _
 '''& "SET DEUMPLUT = '0'" _
 '''& "WHERE DEUMPLUT < 0;"
'''
'''dbs.Close
'''Set dbs = Nothing
'''
'''    stDocName = "REST"
'''    DoCmd.OpenReport stDocName, acPreview
'''
'''
'''Err_Command6_Click:
'''  If errNumar <> 0 Then MsgBox err.description & "   " & errNumar & "  FUNCTIA RESTPRODUCTIE "
'''    Resume Next
'''
'''End Function
'''



'''Function restProductievechi()
'''On Error Resume Next
'''Dim BAZA As DAO.DATABASE
'''    Dim stDocName As String
'''Set BAZA = CurrentDb
'''BAZA.Execute "DELETE * FROM Urmancom", dbSeeChanges
'''BAZA.Execute "DELETE * FROM urmantip", dbSeeChanges
'''BAZA.Execute "DELETE * FROM urmanumpl", dbSeeChanges
'''BAZA.Execute "DELETE * FROM REST", dbSeeChanges
'''BAZA.Close
'''Set BAZA = Nothing
'''Dim dbs As DAO.DATABASE
''''
'''Set dbs = CurrentDb
'''
'''dbs.Execute " INSERT INTO URMANCOM SELECT COMENZI.NRCOMANDA,COMENZI.DEN,COMENZI.SARJA,SUM(COMENZI.CANT) AS COMANDA FROM COMENZI WHERE COMENZI.PRODUS LIKE 'P*' AND NRCOMANDA LIKE '*" & Right(CStr(Year(Date)), 2) & "' AND stadiu LIKE 'in lucru' GROUP BY COMENZI.NRCOMANDA,COMENZI.DEN,COMENZI.SARJA"
'''dbs.Execute " INSERT INTO URMANTIP SELECT INTRARI.NRCOMANDA,INTRARI.DEN,INTRARI.SARJA,SUM(INTRARI.CANT) AS TIPARITE FROM intrari WHERE INTRARI.PRODUS LIKE 'S*' AND NRCOMANDA LIKE '*" & Right(CStr(Year(Date)), 2) & "' GROUP BY INTRARI.NRCOMANDA,INTRARI.DEN,INTRARI.SARJA "
'''dbs.Execute " INSERT INTO URMANUMPL SELECT INTRARI.NRCOMANDA,INTRARI.DEN,SARJA,SUM(INTRARI.CANT) AS UMPLUTE FROM intrari WHERE INTRARI.PRODUS LIKE 'P*' AND NRCOMANDA LIKE '*" & Right(CStr(Year(Date)), 2) & "' GROUP BY INTRARI.NRCOMANDA,INTRARI.DEN,INTRARI.SARJA"
''''''''''BUN
'''dbs.Execute " INSERT INTO REST SELECT Urmancom.sarja, Urmancom.nrcomanda, Urmancom.den, Urmancom.COMANDA, urmantip.TIPARITE, urmanumpl.UMPLUTE, iif(urmantip.TIPARITE , urmancom.COMANDA - urmantip.TIPARITE, urmancom.COMANDA) AS detiparit, iif(Urmanumpl.umplute , urmancom.comanda - urmanumpl.umplute,urmancom.COMANDA) AS deumplut FROM ((Urmancom LEFT OUTER JOIN urmanumpl ON Urmancom.den = urmanumpl.den AND " & _
 '''"Urmancom.sarja = urmanumpl.sarja AND Urmancom.nrcomanda = urmanumpl.nrcomanda) LEFT OUTER JOIN urmantip ON Urmancom.den = urmantip.den AND Urmancom.sarja = urmantip.sarja AND Urmancom.nrcomanda = urmantip.nrcomanda) WHERE ((Urmancom.COMANDA > IIF(urmantip.TIPARITE, urmantip.TIPARITE, 0) AND Urmancom.COMANDA > IIF(urmanumpl.UMPLUTE, urmanumpl.UMPLUTE, 0)) OR (Urmancom.COMANDA > IIF(urmantip.TIPARITE, urmantip.TIPARITE, 0) AND" & _
 '''" Urmancom.COMANDA < IIF(urmanumpl.UMPLUTE, urmanumpl.UMPLUTE, 0)) OR (Urmancom.COMANDA < IIF(urmantip.TIPARITE, urmantip.TIPARITE, 0) AND Urmancom.COMANDA > IIF(urmanumpl.UMPLUTE, urmanumpl.UMPLUTE, 0))) ORDER BY Urmancom.NRCOMANDA"
'''dbs.Execute "delete * from urmancom", dbSeeChanges
'''dbs.Execute "delete * from rest where DEumplut=null and DEtiparit=null", dbSeeChanges
'''dbs.Execute "delete * from rest where DEumplut=0 and DEtiparit=0", dbSeeChanges
'''dbs.Execute "delete * from rest where DEumplut<=0 "
'''dbs.Execute "insert into urmancom SELECT distinct nrcomanda from rest "
'''dbs.Execute "UPDATE REST " _
 '''& "SET DETIPARIT = '0'" _
 '''& "WHERE DETIPARIT < 0;"
'''dbs.Execute "UPDATE REST " _
 '''& "SET DEUMPLUT = '0'" _
 '''& "WHERE DEUMPLUT < 0;"
'''
'''dbs.Close
'''Set dbs = Nothing
'''
'''    stDocName = "REST"
'''    DoCmd.OpenReport stDocName, acPreview
'''
'''
'''Err_Command6_Click:
'''  If errNumar <> 0 Then MsgBox err.description & "   " & errNumar & "  FUNCTIA RESTPRODUCTIE "
'''    Resume Next
'''
'''End Function
'''
'''

Function BARASTARE(COD As String) As String

    Dim SIR As String, N As Integer
    BARASTARE = ""
    Dim Data1 As DAO.Recordset
    Dim baz As DAO.Database
    Set baz = CurrentDb
    Set Data1 = baz.OpenRecordset("ABREVIERI")
    Data1.MoveFirst
    For N = 1 To Data1.RecordCount - 1
        If InStr(1, COD, Data1.Fields("PRESCURTARE")) Then SIR = SIR & Data1.Fields("Valoare")
        Data1.MoveNext
    Next
    BARASTARE = SIR
    Data1.Close
    Set Data1 = Nothing
    Set baz = Nothing

End Function

Function EsteGermana()
    On Error Resume Next
    Dim ret

    ret = Dir("c:\dan\stocuri.mdb")
    If ret = "stocuri.mdb" Then
    Else
        Application.Assistant.visible = True
        ret = MsgBox("PROGRAMUL ESTE LEGAT LA BAZA DE DATE(pentru tabele de manevra) 'c:\dan\stocuri.mdb'" & vbCrLf & " Sursa nu exista in directorul specificat", vbOKOnly, "Legatura la baza importata")

        DoCmd.Quit
        Application.Assistant.visible = False
    End If

End Function
Function mareste()

    On Error Resume Next
    DoCmd.MoveSize 0, 0, 14900, 9900
    'DoCmd.MoveSize 0, 0, 15250, 10360
End Function
Function COPIAZA(src As String, dest As String)

    FileCopy src, dest
End Function
'''Function stoca()
'''Dim ret As String
'''
'''On Error GoTo Err_Command55_Click
'''CurrentDb.QueryDefs.Delete ("INTRAT")
'''CurrentDb.QueryDefs.Delete ("INTRATF")
'''CurrentDb.QueryDefs.Delete ("INITIAL")
'''CurrentDb.QueryDefs.Delete ("INITIALF")
'''ret = "SELECT intrari.produs, IIf(intrari.GRUPA,INTRARI.GRUPA,0) AS GRUPA, IIf(intrari.DEN,INTRARI.DEN,0) AS DEN, IIf(intrari.SARJA,INTRARI.SARJA,0) AS SARJA, intrari.um, IIf(intrari.NRCOMANDA,INTRARI.NRCOMANDA,0) AS NRCOMANDA, Sum(intrari.cant) AS intrat " & _
 '''" FROM INTRARI" & _
 '''" WHERE (((intrari.data) between #1/1/" & Right(Year(Date), 2) & "# and #" & Month(Date) & "/" & Day(Date) & "/" & Right(Year(Date), 2) & "#) AND ((intrari.nrcomanda)<>'-'))" & _
 '''" GROUP BY intrari.produs, intrari.um, intrari.grupa, intrari.den, intrari.sarja, intrari.nrcomanda;"
'''CurrentDb.CreateQueryDef "INTRAT", ret
'''
'''ret = "SELECT intrari.produs, IIf(intrari.GRUPA,INTRARI.GRUPA,0) AS GRUPA, IIf(intrari.DEN,INTRARI.DEN,0) AS DEN, IIf(intrari.SARJA,INTRARI.SARJA,0) AS SARJA, intrari.um, IIf(intrari.NRCOMANDA,INTRARI.NRCOMANDA,0) AS NRCOMANDA, Sum(intrari.cant) AS intrat" & _
 '''" FROM INTRARI" & _
 '''" WHERE (((intrari.data) between #1/1/" & Right(Year(Date), 2) & "# and #" & Month(Date) & "/" & Day(Date) & "/" & Right(Year(Date), 2) & "#) AND ((intrari.nrcomanda)='-'))" & _
 '''"GROUP BY intrari.produs, intrari.um, intrari.grupa, intrari.den, intrari.sarja, intrari.nrcomanda;"
'''
'''CurrentDb.CreateQueryDef "INTRATF", ret
'''ret = "SELECT intrari.produs, IIf(intrari.GRUPA,INTRARI.GRUPA,0) AS GRUPA, IIf(intrari.DEN,INTRARI.DEN,0) AS DEN, IIf(intrari.SARJA,INTRARI.SARJA,0) AS SARJA, intrari.um, IIf(intrari.NRCOMANDA,INTRARI.NRCOMANDA,0) AS NRCOMANDA, Sum(intrari.cant) AS initial" & _
 '''" FROM INTRARI" & _
 '''" WHERE (((intrari.data)< #1/1/" & Right(Year(Date), 2) & "#) AND ((intrari.nrcomanda)<>'-'))" & _
 '''"GROUP BY intrari.produs, intrari.um, intrari.grupa, intrari.den, intrari.sarja, intrari.nrcomanda;"
'''
'''
'''CurrentDb.CreateQueryDef "INITIAL", ret
'''
'''ret = "SELECT intrari.produs, IIf(intrari.GRUPA,INTRARI.GRUPA,0) AS GRUPA, IIf(intrari.DEN,INTRARI.DEN,0) AS DEN, IIf(intrari.SARJA,INTRARI.SARJA,0) AS SARJA, intrari.um, IIf(intrari.NRCOMANDA,INTRARI.NRCOMANDA,0) AS NRCOMANDA, Sum(intrari.cant) AS initial" & _
 '''" FROM INTRARI" & _
 '''" WHERE (((intrari.data)< #1/1/" & Right(Year(Date), 2) & "#) AND ((intrari.nrcomanda)='-'))" & _
 '''"GROUP BY intrari.produs, intrari.um, intrari.grupa, intrari.den, intrari.sarja, intrari.nrcomanda;"
'''
'''CurrentDb.CreateQueryDef "INITIALF", ret
'''    Dim stDocName As String
'''
'''    stDocName = "StocCurataTabelStoc"
'''    DoCmd.OpenQuery stDocName, acNormal, acEdit
'''   DoCmd.Close acQuery, stDocName
'''
'''    stDocName = "StocScrieInTabelStoc"
'''    DoCmd.OpenQuery stDocName, acNormal, acEdit
'''    DoCmd.Close acQuery, stDocName
'''
'''
'''Err_Command55_Click:
''''3265
'''
'''Select Case err.number
'''Case 0
'''Case 3265
'''Resume Next
'''Case Else
'''
'''MsgBox err.description & "   Modul Functii rutina Stoc  " & errNumar
'''
'''
'''End Select
'''
'''
'''End Function
'''
'''


Public Sub Proba1()
    meniu "Prod"
End Sub
Public Function meniu(nume As String)
    On Error Resume Next

    Dim bar

    SchimbaProprietate "AllowSpecialKeys", dbBoolean, True
    Dim bari
    Dim rec As DAO.Recordset
    Set rec = CurrentDb.OpenRecordset("SELECT [4 utilizatori].meniu,[4 utilizatori].utilizator,[4 utilizatori].logare FROM [4 utilizatori] where [4 utilizatori].logare like '" & nume & "';")

    If Not rec.EOF Then
        For bari = 1 To rec.RecordCount
            If rec.Fields("logare") = nume Then

                'If Left(nume, 9) = "Ambalator" Then

                'End If



                For Each bar In CommandBars
                    If bar.name = rec.Fields("meniu") Then
                        ''bar.visible = True
                    End If
                Next

            Else
                If Left(nume, 9) = "Ambalator" Then
                    For Each bar In CommandBars
                        If bar.name = "Balante" Then
                            ''bar.visible = True
                        End If
                    Next

                Else

                    For Each bar In CommandBars
                        If bar.name = rec.Fields("meniu") Then
                            bar.visible = False
                        End If
                    Next
                End If

            End If
            rec.MoveNext
        Next
    End If

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "16 Operatori"
    DoCmd.OpenForm stDocName, , , stLinkCriteria

    'End If
    'End Select
End Function
Public Function BLOCARE()
    CurrentDb.Execute "UPDATE [17 BLOCARE] SET CERERE_BLOCARE=-1", dbSeeChanges
End Function

Public Function meniu_pe_nume(nume As String)
    On Error Resume Next
    Dim stDocName As String
    Dim stLinkCriteria As String
    Dim bar

    '  SchimbaProprietate "AllowSpecialKeys", dbBoolean, False
    Dim bari
    Dim rec As DAO.Recordset
    Set rec = CurrentDb.OpenRecordset("SELECT [4 utilizatori].meniu,[4 utilizatori].utilizator,[4 utilizatori].logare FROM [4 utilizatori] where [4 utilizatori].utilizator like '" & nume & "';")

    If Not rec.EOF Then
        For bari = 1 To rec.RecordCount
            If UCase(RTrim(rec.Fields("utilizator"))) = UCase(nume) Or UCase(rec.Fields("utilizator")) = UCase(nume) & "1" Then

                For Each bar In CommandBars
                    If bar.name = RTrim(rec.Fields("meniu")) Then
                        ''bar.visible = True
                        CurrentDb.Execute "update [UsysRibbonActiuni] set vizibil =0 where left(controlid,3)='tab'"
                        Select Case RTrim(rec.Fields("meniu"))
                        Case "CTC"    ' Control -1
                            CurrentDb.Execute "update [UsysRibbonActiuni] set vizibil =-1 where controlid in('tabCTCLogoff','tabCTCControl')"
                            DoCmd.OpenForm "Detectare sesizareCTC"
                        Case "Productie"    '   Claudia -1
                            CurrentDb.Execute "update [UsysRibbonActiuni] set vizibil =-1 where controlid in('tabCTCLogoff','tabPROG_Comenzi','tabNavigare','tabPROG_Administrare_Sistem','tabPROG_Personal')"
                            DoCmd.OpenForm "Detectare sesizarePRODUCTIE"
                        Case "Programator"    '   Dan -1
                            CurrentDb.Execute "update [UsysRibbonActiuni] set vizibil =-1 where controlid in('tabCTCLogoff','tabPROG_Programare','tabPROG_Administrare_Sistem','tabPROG_Comenzi','tabPROG_Stoc','tabPROG_Personal','tabPROG_Mentenanta','tabPROG_Statistica','tabPROG_Calitate','tabNavigare','tabCTCControl','tabDEPOperari','tabDEPDepozitare','tabDEPAmbalare','tabDEPRapoarte','tab_AMB_Ambalatori','tab_ADM1_Administrare','tab_ADM2_Administrare')"
                        Case "Gestionar"    '   Magazioner  -1
                            CurrentDb.Execute "update [UsysRibbonActiuni] set vizibil =-1 where controlid in('tabCTCLogoff','tabDEPOperari','tabDEPDepozitare','tabDEPAmbalare','tabDEPRapoarte','tabPROG_Stoc')"
                        Case "Balante"    ' Ambalatori  -1
                            CurrentDb.Execute "update [UsysRibbonActiuni] set vizibil =-1 where controlid in('tabCTCLogoff','tab_AMB_Ambalatori')"
                        Case "Administrator1"    '  Administrator   -1
                            CurrentDb.Execute "update [UsysRibbonActiuni] set vizibil =-1 where controlid in('tabCTCLogoff','tab_ADM1_Administrare')"
                        Case "Administrator2"    '  Director productie  -1
                            CurrentDb.Execute "update [UsysRibbonActiuni] set vizibil =-1 where controlid in('tabCTCLogoff','tab_ADM2_Administrare')"
                        End Select
                        '''                        gobjRibbon.Invalidate
                        strCtlLabel = "Logat:" & vbCrLf & UCase$(PERSOANA_ACTIVA) & vbCrLf & Date & vbCrLf & time
                        gobjRibbon.InvalidateControl "lblLabel1"

                        gobjRibbon.InvalidateControl "tabCTCLogoff"
                        gobjRibbon.InvalidateControl "tabCTCControl"

                        gobjRibbon.InvalidateControl "tabDEPOperari"
                        gobjRibbon.InvalidateControl "tabDEPDepozitare"
                        gobjRibbon.InvalidateControl "tabDEPAmbalare"
                        gobjRibbon.InvalidateControl "tabDEPRapoarte"

                        gobjRibbon.InvalidateControl "tabPROG_Programare"
                        gobjRibbon.InvalidateControl "tabPROG_Administrare_Sistem"
                        gobjRibbon.InvalidateControl "tabPROG_Comenzi"
                        gobjRibbon.InvalidateControl "tabPROG_Stoc"
                        gobjRibbon.InvalidateControl "tabPROG_Personal"
                        gobjRibbon.InvalidateControl "tabPROG_Mentenanta"
                        gobjRibbon.InvalidateControl "tabPROG_Statistica"
                        gobjRibbon.InvalidateControl "tabPROG_Calitate"
                        gobjRibbon.InvalidateControl "tabNavigare"


                        gobjRibbon.InvalidateControl "tab_AMB_Ambalatori"

                        gobjRibbon.InvalidateControl "tab_ADM1_Administrare"
                        gobjRibbon.InvalidateControl "tab_ADM2_Administrare"


                        If RTrim(rec.Fields("meniu")) = "Balante" Then

                            stDocName = "11 ambalatori"
                            DoCmd.OpenForm stDocName, , , stLinkCriteria
                            Forms("11 ambalatori").Controls("Label58").Properties("Caption").Value = "Schimba utilizator (" & ScoateBonuri & ")"
                        End If

                    End If
                Next
            End If
            rec.MoveNext
        Next
    End If
    'End If
    'End Select
End Function
Function GATA()
    Application.CommandBars("Menu Bar").enabled = True
    Application.CommandBars("Menu Bar").visible = True
    Application.Quit
End Function
'''Function InceputDeAnbun()
'''    Dim dbs As DAO.Database
'''
'''    Set dbs = CurrentDb
'''    dbs.Execute "DROP TABLE SELECTII"
'''    dbs.Execute "create table selectii (produs text(4),grupa text(50),den text(25),sarja text(12),um text(2),cant numeric,data date,nume text(20),prenume text (20),nrcomanda text(15),codprodus text(15),codfurnizor text(15))"
'''    dbs.Execute "INSERT INTO SELECTII SELECT PRODUS,GRUPA,DEN,SARJA,UM,NRCOMANDA,'" & day(Date) & "." & month(Date) & "." & year(Date) - 1 & "' AS DATA,stoc AS CANT FROM [2 stoc general] where nrcomanda like 'stoc';", dbSeeChanges
'''    dbs.Execute " INSERT INTO SELECTII SELECT PRODUS,GRUPA,DEN,SARJA,UM,NRCOMANDA,DATA,CANT,nume,prenume,codprodus,codfurnizor FROM [INTRARI] where nrcomanda like 'retinut';", dbSeeChanges
'''
'''    dbs.Execute "DROP TABLE INTRARI"
'''    dbs.Execute "create table INTRARI (produs text(4),grupa text(50),den text(25),sarja text(10),um text(2),cant numeric,data date,nume text(20),prenume text (20),nrcomanda text(15),codprodus text(15),codfurnizor text(15))"
'''    dbs.Execute " INSERT INTO intrari SELECT PRODUS,GRUPA,DEN,SARJA,UM,NRCOMANDA,DATA,CANT,CODPRODUS,CODFURNIZOR,NUME,PRENUME FROM selectii;", dbSeeChanges
'''    dbs.Execute "update intrari set prenume='PRODUS' where produs like 'P*' or produs like 'S*'", dbSeeChanges
'''    dbs.Execute " delete * FROM SELECTII;", dbSeeChanges
'''    dbs.Execute " deleTE * FROM iesiri;", dbSeeChanges
'''    'dbs.Execute " deleTE * FROM INTRARI WHERE CANT <= 0.09;"
'''    dbs.Close
'''    Set dbs = Nothing
'''    Application.Assistant.visible = True
'''    MsgBox "Fisierul a fost pregatit pentru inceputul de an.", vbOKOnly, "Operatie incheiata"
'''    Application.Assistant.visible = False
'''    'End If
'''
'''
'''    'Else
'''    'Application.Assistant.Visible = True
'''    'MsgBox "TREBUIE SA CREATI BACKUP PENTRU ANUL TRECUT INAINTE DE FORMAREA FISIERULUI."
'''    'Application.Assistant.Visible = False
'''    'End If
'''
'''    'erori:
'''End Function
'''Function InceputDeAn()
'''    Dim ret, RAS
'''    Dim dbs As DAO.Database
'''    ret = Dir("C:\Program Files\STOCURI\Stocuri" & year(Date) - 1 & ".mdb")
'''    If ret = "Stocuri" & year(Date) - 1 & ".mdb" Then
'''        ret = MsgBox("EXISTA BACKUP PE ANUL " & year(Date) - 1 & ". DORITI FORMAREA FISIERULUI", vbYesNo)
'''        If ret = vbYes Then
'''            If month(Date) > 1 Then
'''                RAS = MsgBox("Atentie nu sinteti in luna ianuarie " & vbCrLf & vbCrLf & "        Doriti executia ?", vbInformation + vbYesNo, "ATENTIE")
'''                If RAS = vbYes Then
'''                    Set ret = Nothing
'''
'''                    Set dbs = CurrentDb
'''                    dbs.Execute "DROP TABLE SELECTII"
'''                    dbs.Execute "create table selectii (ID text(20),NrArt text(15),produs text(4),grupa text(50),den text(25),sarja text(8),um text(2),cant numeric,data date,nume text(20),prenume text (20),nrcomanda text(7))"
'''                    dbs.Execute " INSERT INTO SELECTII SELECT nrart,PRODUS,GRUPA,DEN,SARJA,UM,NRCOMANDA,'31.12." & year(Date) - 1 & "' AS DATA,stoc AS CANT FROM QSTOC;", dbSeeChanges
'''                    dbs.Execute "delete * from INTRARI", dbSeeChanges
'''                    dbs.Execute " INSERT INTO intrari SELECT nrart,PRODUS,GRUPA,DEN,SARJA,UM,NRCOMANDA,DATA,CANT,'31.12." & year(Date) - 1 & "' AS id FROM selectii;", dbSeeChanges
'''                    dbs.Execute "update intrari set prenume='MATERIAL'", dbSeeChanges
'''                    dbs.Execute "update intrari set prenume='PRODUS' where produs like 'P*' or produs like 'S*'", dbSeeChanges
'''                    dbs.Execute " delete * FROM SELECTII;", dbSeeChanges
'''                    dbs.Execute " deleTE * FROM iesiri;", dbSeeChanges
'''                    dbs.Execute " deleTE * FROM INTRARI WHERE CANT <= 0.09;", dbSeeChanges
'''                    dbs.Close
'''                    Set dbs = Nothing
'''                    MsgBox "Fisierul a fost pregatit pentru inceputul de an.", vbOKOnly, "Operatie incheiata"
'''                End If
'''            End If
'''            If month(Date) = 1 Then
'''                Set ret = Nothing
'''
'''                Set dbs = CurrentDb
'''                dbs.Execute "DROP TABLE SELECTII"
'''                dbs.Execute "create table selectii (ID text(20),NrArt text(15),produs text(4),grupa text(50),den text(25),sarja text(12),um text(2),cant numeric,data date,nume text(20),prenume text (20),nrcomanda text(10))"
'''                dbs.Execute " INSERT INTO SELECTII SELECT nrart,PRODUS,GRUPA,DEN,SARJA,UM,NRCOMANDA,'31.12." & year(Date) - 1 & "' AS DATA,stoc AS CANT FROM QSTOC;", dbSeeChanges
'''                dbs.Execute "delete * from INTRARI", dbSeeChanges
'''                dbs.Execute "INSERT INTO intrari SELECT nrart,PRODUS,GRUPA,DEN,SARJA,UM,cant,NRCOMANDA,data,'-' AS id,'-' as nume,'-' as prenume FROM selectii;", dbSeeChanges
'''                dbs.Execute "update intrari set prenume='MATERIAL'", dbSeeChanges
'''                dbs.Execute "update intrari set prenume='PRODUS' where produs like 'P*' or produs like 'S*'", dbSeeChanges
'''                dbs.Execute " delete * FROM SELECTII;", dbSeeChanges
'''                dbs.Execute " deleTE * FROM iesiri;", dbSeeChanges
'''                dbs.Execute " deleTE * FROM INTRARI WHERE CANT <= 0.09;", dbSeeChanges
'''                dbs.Close
'''                Set dbs = Nothing
'''                MsgBox "Fisierul a fost pregatit pentru inceputul de an.", vbOKOnly, "Operatie incheiata"
'''            End If
'''        End If
'''
'''
'''    Else
'''        MsgBox "TREBUIE SA CREATI BACKUP PENTRU ANUL TRECUT INAINTE DE FORMAREA FISIERULUI."
'''    End If
'''
'''    'erori:
'''End Function
'''
'''
'''
'''
Sub CreeazaInterogare(fisier As String, fraza As String, DENUMIRE As String)
    On Error GoTo erori
    Dim dbsbaza As DAO.Database
    Dim qdfraport As DAO.QueryDef
    Set dbsbaza = CurrentDb
    Set qdfraport = dbsbaza.CreateQueryDef(DENUMIRE, fraza)
    qdfraport.Close
    dbsbaza.Close
    Set qdfraport = Nothing
    Set dbsbaza = Nothing
erori:
    Select Case err.Number
    Case 0
    Case 3012    'interogarea exista
        Resume Next
    Case 91
        Resume Next
    Case Else
        MsgBox err.Number, "start1-CreeazaInterogare(Fisier As String, Fraza As String, Denumire As String)"
    End Select
End Sub
Public Sub StergeInterogare(DENUMIRE As String)
    On Error GoTo erori
    Dim dbsbaza As DAO.Database
    Set dbsbaza = CurrentDb
    dbsbaza.QueryDefs.Delete (DENUMIRE)
    dbsbaza.Close
    Set dbsbaza = Nothing
erori:
    Select Case err.Number
    Case 0
    Case 3265    'interogarea nu exista
        Resume Next
    Case Else
        MsgBox err.Number, "Start1- StergeInterogare"
        Resume Next
    End Select
End Sub



'''Function productia()
'''On Error Resume Next
'''Dim BAZA As DAO.DATABASE
'''Set BAZA = CurrentDb
''''baza.Execute "create table INTREZ (produs text(4),grupa text(15),den text(25),sarja text(8),um text(2),data date,nume text(20),prenume text (20),SCANT numeric, nrcomanda text(7),LMIN NUMERIC,LMAX NUMERIC)"
''''baza.Execute "create table IESREZ  (produs text(4),grupa text(15),den text(25),sarja text(8),um text(2),data date,nume text(20),prenume text (20),SCANT numeric, nrcomanda text(7),LMIN NUMERIC,LMAX NUMERIC)"
''''baza.Execute "create table PRODUCTIEIN (produs text(4),grupa text(15),den text(25),sarja text(8),um text(2),FINITE NUMERIC,SEMIF NUMERIC,SCANT numeric, nrcomanda text(7),VALOARE TEXT(6),ORDIN NUMERIC,LMIN NUMERIC,LMAX NUMERIC,inR NUMERIC)"
''''baza.Execute "create table PRODUCTIEIE  (produs text(4),grupa text(15),den text(25),sarja text(8),um text(2),FINITE NUMERIC,SEMIF NUMERIC,SCANT numeric,nrcomanda text(7),VALOARE TEXT(6),ORDIN NUMERIC,LMIN NUMERIC,LMAX NUMERIC,inR NUMERIC)"
''''baza.Execute "alter table PRODUCTIEIE add column CONSUMATE NUMERIC"
''''baza.Execute "alter table PRODUCTIEIE add column DIFERENTA NUMERIC"
''''baza.Execute "alter table PRODUCTIEIE add column VALOARE TEXT(6)"
''''baza.Execute "alter table PRODUCTIEIE add column PLECAT NUMERIC"
''''baza.Execute "alter table PRODUCTIEIE add column INTORS NUMERIC"
''''baza.Execute "alter table PRODUCTIEIN add column CONSUMATE NUMERIC"
''''baza.Execute "alter table PRODUCTIEIN add column DIFERENTA NUMERIC"
''''baza.Execute "alter table PRODUCTIEIN add column VALOARE TEXT(6)"
''''baza.Execute "alter table PRODUCTIEIN add column PLECAT NUMERIC"
''''baza.Execute "alter table PRODUCTIEIN add column INTORS NUMERIC"
''''',CONSUMATE NUMERIC,DIFERENTA NUMERIC,VALOARE TEXT(6),PLECAT NUMERIC,INTORS NUMERIC
''''baza.Execute "create table Intraripregatite (produs text(4),grupa text(15),den text(25),sarja text(8),um text(2),FINITE NUMERIC,SEMIF NUMERIC,SCANT numeric,nrcomanda text(7),VALOARE TEXT(6),ORDIN NUMERIC,LMIN NUMERIC,LMAX NUMERIC,inR NUMERIC)"
''''baza.Execute "create table Iesiripregatite (produs text(4),grupa text(15),den text(25),sarja text(8),um text(2),FINITE NUMERIC,CONSUMATE NUMERIC,SCANT numeric,nrcomanda text(7),VALOARE TEXT(6),ORDIN NUMERIC,LMIN NUMERIC,LMAX NUMERIC,inR NUMERIC)"
'''
''''baza.Execute "create table INTREZ (produs text(4),grupa text(15),den text(25),sarja text(8),um text(2),data date,nume text(20),prenume text (20),SCANT numeric, nrcomanda text(7),LMIN NUMERIC,LMAX NUMERIC)"
''''baza.Execute "create table IESREZ  (produs text(4),grupa text(15),den text(25),sarja text(8),um text(2),data date,nume text(20),prenume text (20),SCANT numeric, nrcomanda text(7),LMIN NUMERIC,LMAX NUMERIC)"
''''baza.Execute "create table PRODUCTIEIN (produs text(4),grupa text(15),den text(25),sarja text(8),um text(2),FINITE NUMERIC,SEMIF NUMERIC,SCANT numeric, nrcomanda text(7),VALOARE TEXT(6),ORDIN NUMERIC,LMIN NUMERIC,LMAX NUMERIC,inR NUMERIC)"
''''baza.Execute "create table PRODUCTIEIE  (produs text(4),grupa text(15),den text(25),sarja text(8),um text(2),FINITE NUMERIC,SEMIF NUMERIC,SCANT numeric,nrcomanda text(7),VALOARE TEXT(6),ORDIN NUMERIC,LMIN NUMERIC,LMAX NUMERIC,inR NUMERIC)"
''''baza.Execute "alter table PRODUCTIEIE add column CONSUMATE NUMERIC"
''''baza.Execute "alter table PRODUCTIEIE add column DIFERENTA NUMERIC"
''''baza.Execute "alter table PRODUCTIEIE add column VALOARE TEXT(6)"
''''baza.Execute "alter table PRODUCTIEIE add column PLECAT NUMERIC"
''''baza.Execute "alter table PRODUCTIEIE add column INTORS NUMERIC"
''''baza.Execute "alter table PRODUCTIEIN add column CONSUMATE NUMERIC"
''''baza.Execute "alter table PRODUCTIEIN add column DIFERENTA NUMERIC"
''''baza.Execute "alter table PRODUCTIEIN add column VALOARE TEXT(6)"
''''baza.Execute "alter table PRODUCTIEIN add column PLECAT NUMERIC"
''''baza.Execute "alter table PRODUCTIEIN add column INTORS NUMERIC"
''' Dim numele, prenumele
''' Dim DTDATA As Date
''' prenumele = Forms![PRODUCTIA LUNARA A UNEI PERSOANE]!nume.Column(1)
'''    numele = Forms![PRODUCTIA LUNARA A UNEI PERSOANE]!nume.Column(0)
'''
'''DTDATA = Forms![PRODUCTIA LUNARA A UNEI PERSOANE]!Data.Value
''''baza.Execute "create table Intraripregatite (produs text(4),grupa text(15),den text(25),sarja text(8),um text(2),FINITE NUMERIC,SEMIF NUMERIC,SCANT numeric,nrcomanda text(7),VALOARE TEXT(6),ORDIN NUMERIC,LMIN NUMERIC,LMAX NUMERIC,inR NUMERIC)"
''''baza.Execute "create table Iesiripregatite (produs text(4),grupa text(15),den text(25),sarja text(8),um text(2),FINITE NUMERIC,CONSUMATE NUMERIC,SCANT numeric,nrcomanda text(7),VALOARE TEXT(6),ORDIN NUMERIC,LMIN NUMERIC,LMAX NUMERIC,inR NUMERIC)"
'''    BAZA.Execute "delete * from intrez", dbSeeChanges
'''    BAZA.Execute "delete * from iesrez", dbSeeChanges
'''    BAZA.Execute "delete * from PRODUCTIEIE", dbSeeChanges
'''    BAZA.Execute "delete * from PRODUCTIEIn"
'''    BAZA.Execute "delete * from Iesiripregatite", dbSeeChanges
'''    BAZA.Execute "delete * from Intraripregatite", dbSeeChanges
'''    BAZA.Execute "INSERT INTO IESREZ SELECT DISTINCT IESIRI.PRODUS,IESIRI.GRUPA,IESIRI.DEN,IESIRI.SARJA,IESIRI.UM,IESIRI.CANT AS SCANT,IESIRI.DATA,IESIRI.NUME,IESIRI.PRENUME,IESIRI.NRCOMANDA,COMENZI.LMIN,COMENZI.LMAX" & _
 '''    " FROM IESIRI LEFT JOIN COMENZI ON (IESIRI.NRCOMANDA = COMENZI.NRCOMANDA) WHERE IESiri.nume like'" & numele & "' and IESiri.prenume like'" & prenumele & "' AND iesiri.Data >= " & "#" & Month(DTDATA) & "/01/" & Year(DTDATA) & "#" & " AND IESIRI.Data <= " & "#" & Month(DTDATA) & "/" & Day(DTDATA) & "/" & Year(DTDATA) & "# "
'''
'''
'''    BAZA.Execute "INSERT INTO INTREZ SELECT DISTINCT INTRARI.PRODUS,INTRARI.GRUPA,INTRARI.DEN,INTRARI.SARJA,INTRARI.UM,INTRARI.CANT AS SCANT,INTRARI.DATA,INTRARI.NUME,INTRARI.PRENUME,INTRARI.NRCOMANDA,COMENZI.LMIN,COMENZI.LMAX" & _
 '''    " FROM INTRARI LEFT JOIN COMENZI ON (INTRARI.NRCOMANDA = COMENZI.NRCOMANDA) WHERE intrari.nume like'" & numele & "' and intrari.prenume like'" & prenumele & "' AND intrari.Data >= " & "#" & Month(DTDATA) & "/01/" & Year(DTDATA) & "#" & " AND INTRARI.Data <= " & "#" & Month(DTDATA) & "/" & Day(DTDATA) & "/" & Year(DTDATA) & "# "
'''
'''
'''    BAZA.Execute "INSERT INTO PRODUCTIEIE SELECT PRODUS,GRUPA,DEN,SARJA,UM,LMIN,LMAX,SUM(SCANT) AS finite,NRCOMANDA" & _
 '''    "  FROM IESREZ GROUP BY PRODUS,GRUPA,DEN,SARJA,UM,NRCOMANDA,LMIN,LMAX "
'''    BAZA.Execute "INSERT INTO PRODUCTIEIN SELECT PRODUS,GRUPA,DEN,SARJA,UM,LMIN,LMAX,SUM(SCANT) AS finite,NRCOMANDA" & _
 '''    "  FROM INTREZ  GROUP BY PRODUS,GRUPA,DEN,SARJA,UM,NRCOMANDA,LMIN,LMAX"
'''
'''
'''
'''
'''    BAZA.Execute "delete * from intrez"
'''    BAZA.Execute "delete * from iesrez"
'''    BAZA.Execute "delete * from PRODUCTIEIE"
'''    BAZA.Execute "delete * from PRODUCTIEIn"
'''    BAZA.Execute "delete * from Iesiripregatite"
'''    BAZA.Execute "delete * from Intraripregatite"
'''    BAZA.Execute "INSERT INTO IESREZ SELECT DISTINCT IESIRI.PRODUS,IESIRI.GRUPA,IESIRI.DEN,IESIRI.SARJA,IESIRI.UM,IESIRI.CANT AS SCANT,IESIRI.DATA,IESIRI.NUME,IESIRI.PRENUME,IESIRI.NRCOMANDA,COMENZI.LMIN,COMENZI.LMAX" & _
 '''    " FROM IESIRI LEFT JOIN COMENZI ON (IESIRI.NRCOMANDA = COMENZI.NRCOMANDA) WHERE IESiri.nume like'" & numele & "' and IESiri.prenume like'" & prenumele & "' AND iesiri.Data >= " & "#" & Month(DTDATA) & "/01/" & Year(DTDATA) & "#" & " AND IESIRI.Data <= " & "#" & Month(DTDATA) & "/" & Day(DTDATA) & "/" & Year(DTDATA) & "# "
'''
'''
'''    BAZA.Execute "INSERT INTO INTREZ SELECT DISTINCT INTRARI.PRODUS,INTRARI.GRUPA,INTRARI.DEN,INTRARI.SARJA,INTRARI.UM,INTRARI.CANT AS SCANT,INTRARI.DATA,INTRARI.NUME,INTRARI.PRENUME,INTRARI.NRCOMANDA,COMENZI.LMIN,COMENZI.LMAX" & _
 '''    " FROM INTRARI LEFT JOIN COMENZI ON (INTRARI.NRCOMANDA = COMENZI.NRCOMANDA) WHERE intrari.nume like'" & numele & "' and intrari.prenume like'" & prenumele & "' AND intrari.Data >= " & "#" & Month(DTDATA) & "/01/" & Year(DTDATA) & "#" & " AND INTRARI.Data <= " & "#" & Month(DTDATA) & "/" & Day(DTDATA) & "/" & Year(DTDATA) & "# "
'''
'''
'''    BAZA.Execute "INSERT INTO PRODUCTIEIE SELECT PRODUS,GRUPA,DEN,SARJA,UM,LMIN,LMAX,SUM(SCANT) AS finite,NRCOMANDA" & _
 '''    "  FROM IESREZ GROUP BY PRODUS,GRUPA,DEN,SARJA,UM,NRCOMANDA,LMIN,LMAX "
'''    BAZA.Execute "INSERT INTO PRODUCTIEIN SELECT PRODUS,GRUPA,DEN,SARJA,UM,LMIN,LMAX,SUM(SCANT) AS finite,NRCOMANDA" & _
 '''    "  FROM INTREZ  GROUP BY PRODUS,GRUPA,DEN,SARJA,UM,NRCOMANDA,LMIN,LMAX"
'''
'''
'''
'''
'''
'''
'''
'''
'''    Dim B, C, L1, L2, L3
'''      Dim prim As Integer
'''    Dim fir As String
'''    Dim T As Integer
'''    Dim max As Integer
'''    Dim DATA1 As DAO.Recordset
'''    Dim baz As DAO.DATABASE
'''
'''    Set DATA1 = BAZA.OpenRecordset("variabile")
'''    DATA1.MoveLast
'''    max = DATA1.RecordCount
'''    DATA1.MoveFirst
'''    For T = 1 To max
'''        If LCase(DATA1.Fields("VAriabila")) = "b" Then
'''            If InStr(1, DATA1.Fields("VALOARE"), ",") Then
'''                prim = Val(InStr(1, DATA1.Fields("VALOARE"), ",") - 1)
'''                fir = Left(prim, CStr(DATA1.Fields("VALOARE"))) & "."
'''                B = fir & Right(DATA1.Fields("VALOARE"), Len(DATA1.Fields("VALOARE")) - InStr(1, DATA1.Fields("VALOARE"), ","))
'''            Else
'''                B = DATA1.Fields("VALOARE")
'''            End If
'''        End If
'''        If LCase(DATA1.Fields("VAriabila")) = "c" Then
'''            If InStr(1, DATA1.Fields("VALOARE"), ",") Then
'''                prim = Val(InStr(1, DATA1.Fields("VALOARE"), ",") - 1)
'''                fir = Left(prim, CStr(DATA1.Fields("VALOARE"))) & "."
'''                C = fir & Right(DATA1.Fields("VALOARE"), Len(DATA1.Fields("VALOARE")) - InStr(1, DATA1.Fields("VALOARE"), ","))
'''            Else
'''                C = DATA1.Fields("VALOARE")
'''            End If
'''        End If
'''        If LCase(DATA1.Fields("VAriabila")) = "l1" Then
'''            If InStr(1, DATA1.Fields("VALOARE"), ",") Then
'''                prim = Val(InStr(1, DATA1.Fields("VALOARE"), ",") - 1)
'''                fir = Left(prim, CStr(DATA1.Fields("VALOARE"))) & "."
'''                L1 = fir & Right(DATA1.Fields("VALOARE"), Len(DATA1.Fields("VALOARE")) - InStr(1, DATA1.Fields("VALOARE"), ","))
'''            Else
'''                L1 = DATA1.Fields("VALOARE")
'''            End If
'''        End If
'''        If LCase(DATA1.Fields("VAriabila")) = "l2" Then
'''            If InStr(1, DATA1.Fields("VALOARE"), ",") Then
'''                prim = Val(InStr(1, DATA1.Fields("VALOARE"), ",") - 1)
'''                fir = Left(prim, CStr(DATA1.Fields("VALOARE"))) & "."
'''                L2 = fir & Right(DATA1.Fields("VALOARE"), Len(DATA1.Fields("VALOARE")) - InStr(1, DATA1.Fields("VALOARE"), ","))
'''            Else
'''                L2 = DATA1.Fields("VALOARE")
'''            End If
'''        End If
'''        If LCase(DATA1.Fields("VAriabila")) = "l3" Then
'''            If InStr(1, DATA1.Fields("VALOARE"), ",") Then
'''                prim = Val(InStr(1, DATA1.Fields("VALOARE"), ",") - 1)
'''                fir = Left(prim, CStr(DATA1.Fields("VALOARE"))) & "."
'''                L3 = fir & Right(DATA1.Fields("VALOARE"), Len(DATA1.Fields("VALOARE")) - InStr(1, DATA1.Fields("VALOARE"), ","))
'''            Else
'''                L3 = DATA1.Fields("VALOARE")
'''            End If
'''        End If
'''    DATA1.MoveNext
'''Next
'''
'''DATA1.Close
'''Set DATA1 = Nothing
'''
'''''                                                                adauga  consumabilele
'''' punctul 1
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM, finite as Scant,NRCOMANDA,'CONS' AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE 'C*'"
'''BAZA.Execute "DELETE * FROM Iesiripregatite WHERE PRODUS LIKE 'CE' "
'''BAZA.Execute "DELETE * FROM Iesiripregatite WHERE  PRODUS LIKE 'CA' "
'''BAZA.Execute "DELETE * FROM Iesiripregatite WHERE  PRODUS LIKE 'CF'"
'''' punctul 2
''''                                                       adauga  cantitatile pentru TB-S(SB)        >>>TB
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,grupa,den,sarja,nrcomanda,UM,sum(finite) as scant,'TB' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'TB' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,grupa,den,sarja,nrcomandA,UM,sum(finite) as scant,'TB'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'SB' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'sTB' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'TB' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'sTB'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'SB' group by valoare"
''''marcat
'''BAZA.Execute "update productiein set productiein.inR=1 where productiein.produs like 'SB';"
''''punctul 3
''''''                                                       adauga  cantitatile pentru SB-PB        >>>SB
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,grupa,den,sarja,nrcomanda,UM,sum(finite) as scant,'SB' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'SB' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,grupa,den,sarja,nrcomandA,UM,sum(finite) as scant,'SB'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PB' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'ssB' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'SB' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'ssB'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PB' group by valoare"
''''marcat
'''BAZA.Execute "update productieie set productieiE.inR='1' where productieiE.produs like 'SB'"
'''BAZA.Execute "update productieiN set productieiN.inR='1' where productieiN.produs like 'PB'"
''''punctul 4
''''''                                                       adauga  cantitatile pentru MB-vb*PB        >>>PB
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,grupa,den,sarja,UM,sum(finite) as scant,'PB' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MB' group by PRODUS,grupa,den,SARJA,UM,valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,grupa,den,sarja,UM,sum(finite) as scant,'PB'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PB' group by PRODUS,grupa,den,sarja,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'spB' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MB' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'spB'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PB' group by valoare"
''''marcat
'''BAZA.Execute "update productieiN set productieiN.inR='1' where productieiN.produs like 'PB'"
''''punctul 5
''''''                                                       adauga  cantitatile pentru TC-S(SC)       >>>TC
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,grupa,den,sarja,nrcomanda,UM,sum(finite) as scant,'TC' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'TC' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,grupa,den,sarja,nrcomandA,UM,sum(finite) as scant,'TC'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'SC' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'STC' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'TC' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'STC'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'SC' group by valoare"
''''marcat
'''BAZA.Execute "update productiein set productieiN.inR='1' where productieiN.produs like 'SC'"
''''punctul 6
''''''                                                       adauga  cantitatile pentru sC-S(PC)       >>>SC
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,grupa,den,sarja,nrcomanda,UM,sum(finite) as scant,'SC' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'SC' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,grupa,den,sarja,nrcomandA,UM,sum(finite) as scant,'SC'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PC' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'SSC' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'SC' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'SSC'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PC' group by valoare"
''''marcat
'''BAZA.Execute "update productieie set productieiE.inR='1' where productieiE.produs like 'SC'"
'''BAZA.Execute "update productiein set productiein.inR='1' where productiein.produs like 'PC'"
''''punctul 7
''''''                                                       adauga  cantitatile pentru PC-S(MC) AK      >>>MCAK
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,NRCOMANDA,'MCAK'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MC' and grupa like 'AKTIVIERKAPPE'  group by PRODUS,grupa,den,sarja,nrcomanda," & _
 '''"UM,valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,NRCOMANDA,'MCAK'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PC' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'SMCAK' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MC'  and grupa like 'AKTIVIERKAPPE' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'SMCAK'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PC' group by valoare"
''''''                                                       adauga  cantitatile pentru PC-S(MC) PIN      >>>MCPI
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,NRCOMANDA,'MCPI'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MC' and grupa like 'PIN'  group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,NRCOMANDA,'MCPI'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PC' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'SMCPI' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MC'  and grupa like 'PIN' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'SMCPI'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PC' group by valoare"
''''''                                                       adauga  cantitatile pentru PC-S(MC) TORPEDO      >>>MCPI
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,NRCOMANDA,'MCTO'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MC' and grupa like 'TORPEDO'  group by PRODUS,grupa,den,sarja,nrcomanda," & _
 '''"UM,valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,NRCOMANDA,'MCTO'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PC' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'SMCTO' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MC'  and grupa like 'TORPEDO' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'SMCTO'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PC' group by valoare"
''''PUNCTUL 8
''''''                                                       adauga  cantitatile pentru CA*VC-CC   FOL,BULE      >>>CA
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'CA'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'CA' and grupa like 'FOLIE' AND DEN LIKE 'BULE'  group by PRODUS,grupa,den," & _
 '''"sarja,UM,valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'CA'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'CC'  and grupa like 'FOLIE' AND DEN LIKE 'BULE' AND SARJA LIKE 'GES*' " & _
 '''"group by PRODUS,grupa,den,sarja,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'SCA' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'CA' and grupa like 'FOLIE' AND DEN LIKE 'BULE'  group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'SCA'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'CC'  and grupa like 'FOLIE' AND DEN LIKE 'BULE' AND SARJA LIKE 'GES*' " & _
 '''"group by valoare"
'''' PUNCTUL 9
''''''                                                       adauga  cantitatile pentru PE  ,PB     >>>PE
'''
''''PENTRU TOTALURI FINALE
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'PE'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PE'  group by PRODUS,grupa,den,sarja,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'PE'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'PB' group by PRODUS,grupa,den,sarja,UM,valoare"
''''Subtotaluri PE si PB
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) AS SCANT,'SPE'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PE'  group by valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) AS SCANT,'SPE'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'PB' group by valoare"
''''marcat
'''BAZA.Execute "update productiein set productiein.inR='1' where productiein.produs like 'PE'"
'''BAZA.Execute "update productieie set productieiE.inR='1' where productieiE.produs like 'PB'"
'''' PUNCTUL 10
'''' selecteaza nituri si saibe
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'CEN'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'CE' and grupa like 'nituri' group by PRODUS,grupa,den,sarja,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'CES'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'CE' and grupa like 'nituri' AND DEN LIKE 'sai*'  " & _
 '''"group by PRODUS,grupa,den,sarja,UM,valoare"
''''subtotaluri pentru PE-CE
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) AS SCANT,'SCEN'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PE'  group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) AS SCANT,'SCES'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PE'  group by valoare"
''''Subtotaluri CE
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) AS SCANT,'SCEN'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'CE' and grupa like 'nituri'  and den not like 'sai*' group by valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) AS SCANT,'SCES'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'CE' and grupa like 'nituri' AND DEN LIKE 'sai*'  group by valoare"
''''punctul 11
''''selecteaza CF ,grupa=Klo,den=Ges
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'CFa'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIn WHERE PRODUCTIEIn.PRODUS LIKE  'CF' and grupa like 'FOL*' AND DEN LIKE 'GES*'  " & _
 '''"group by PRODUS,grupa,den,sarja,UM,valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'CFb'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIn WHERE PRODUCTIEIn.PRODUS LIKE  'CF' and grupa like 'KLO*' AND DEN LIKE 'GES*'  " & _
 '''"group by PRODUS,grupa,den,sarja,UM,valoare"
''''subtotal pentru CF-CF
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) AS SCANT,'SCFa'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIn WHERE PRODUCTIEIn.PRODUS LIKE  'CF' and grupa like 'FOL*' AND DEN LIKE 'GES*'  group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) AS SCANT,'SCFb'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIn WHERE PRODUCTIEIn.PRODUS LIKE  'CF' and grupa like 'KLO*' AND DEN LIKE 'GES*'  group by valoare"
''''selecteaza TLa ,Grupa=Fol
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'TLA'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'TL' and grupa like 'FOL*' and sarja like 'SCH*'  " & _
 '''"group by PRODUS,grupa,den,sarja,UM,valoare"
''''subtotal
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) AS SCANT,'SCFa'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'TL' and grupa like 'FOL*'  and sarja like 'SCH*' group by valoare"
''''selecteaza TLb ,Grupa=Klo
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'TLB'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'TL' and grupa like 'Klo*'   group by PRODUS,grupa,den,sarja,UM,valoare"
''''subtotal
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) AS SCANT,'SCFb'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'TL' and grupa like 'Klo*'   group by valoare"
''''   Punctul  12
''''selecteaza PL
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'PL'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIn WHERE PRODUCTIEIn.PRODUS LIKE  'PL' group by PRODUS,grupa,den,sarja,UM,valoare"
''''marcat
'''BAZA.Execute "update productiein set productiein.inR='1' where productiein.produs like 'PL'"
''''selecteaza CFa  PRODUS LIKE  'CF' and grupa like 'FOL*' AND DEN LIKE 'GES*'
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'CFaa'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUCTIEIe.PRODUS LIKE  'CF' and grupa like 'FOL*' AND DEN LIKE 'GES*'  " & _
 '''"group by PRODUS,grupa,den,sarja,UM,valoare"
''''Subtotal CFa-PL
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) AS SCANT,'SCFAa'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUCTIEIe.PRODUS LIKE  'CF' and grupa like 'FOL*' AND DEN LIKE 'GES*'  group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) AS SCANT,'SCFAa'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIn WHERE PRODUCTIEIn.PRODUS LIKE  'PL' group by valoare"
''''marcat
'''BAZA.Execute "update productiein set productiein.inR='1' where productiein.produs like 'PL'"
''''selecteaza CFb  PRODUS LIKE  'CF' and grupa like 'KLO*' AND DEN LIKE 'GES*'
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'CFbb'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUCTIEIe.PRODUS LIKE  'CF' and grupa like 'KLO*' AND DEN LIKE 'GES*'  " & _
 '''"group by PRODUS,grupa,den,sarja,UM,valoare"
''''subtotal pentru CFb-PL
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) AS SCANT,'SCFbb'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUCTIEIe.PRODUS LIKE  'CF' and grupa like 'KLO*' AND DEN LIKE 'GES*' " & _
 '''" group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) AS SCANT,'SCFbb'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIn WHERE PRODUCTIEIn.PRODUS LIKE  'PL' group by valoare"
''''     selecteaza   CLa          PRODUS LIKE  'CL and grupa like 'BOX*' AND sarja LIKE 'BOX*'
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'CLa'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUCTIEIe.PRODUS LIKE  'CL' and grupa like 'BOX*' AND sarja LIKE 'BOX*'  " & _
 '''"group by PRODUS,grupa,den,sarja,UM,valoare"
''''subtotal pentru CLa-PL
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) AS SCANT,'SCLa'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUCTIEIe.PRODUS LIKE  'CL' and grupa like 'BOX*' AND sarja LIKE 'BOX*'  group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) AS SCANT,'SCLa'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIn WHERE PRODUCTIEIn.PRODUS LIKE  'PL' group by valoare"
''''     selecteaza   CLb          PRODUS LIKE  'CL and grupa like 'BOX*' AND sarja LIKE 'BOX*'
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'CLb'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUCTIEIe.PRODUS LIKE  'CL' and grupa like 'BOX*' AND sarja LIKE 'DEC*'  " & _
 '''"group by PRODUS,grupa,den,sarja,UM,valoare"
''''subtotal pentru CLa-PL
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) AS SCANT,'SCLb'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUCTIEIe.PRODUS LIKE  'CL' and grupa like 'BOX*' AND sarja LIKE 'DEC*'  group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) AS SCANT,'SCLb'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIn WHERE PRODUCTIEIn.PRODUS LIKE  'PL' group by valoare"
''''     selecteaza   TL          PRODUS LIKE  'TL and grupa like 'KLO*' AND DEN LIKE 'GRA*' AND sarja LIKE '50'
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'TL' AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUS LIKE 'TL' and grupa like 'KLO*' AND DEN LIKE 'GRA*' AND sarja LIKE '50' " & _
 '''"group by PRODUS,grupa,den,sarja,UM,valoare"
''''subtotal pentru TL-PL
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) AS SCANT,'STL' AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE  PRODUS LIKE  'TL' and grupa like 'KLO*' AND DEN LIKE 'GRA*' AND sarja LIKE '50' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) AS SCANT,'STL'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIn WHERE PRODUCTIEIn.PRODUS LIKE  'PL' group by valoare"
''''Selecteaza ML
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,GRUPA,DEN,SARJA,UM,sum(finite) AS SCANT,'ML'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUS LIKE 'ML' group by PRODUS,grupa,den,sarja,UM,valoare"
''''subtotal pentru ML-PL
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) AS SCANT,'SML'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUCTIEIe.PRODUS LIKE  'ML' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) AS SCANT,'SML'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIn WHERE PRODUCTIEIn.PRODUS LIKE  'PL' group by valoare"
'''' Punctul 13
''''''                                                       adauga  cantitatile pentru TT-ST      >>>TT
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,grupa,den,sarja,nrcomanda,UM,sum(finite) as scant,'TT' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'TT' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,grupa,den,sarja,nrcomandA,UM,sum(finite) as scant,'ST'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'ST' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'STT' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'TT' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'STT'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'ST' group by valoare"
''''marcat
'''BAZA.Execute "update productiein set productiein.inR='1' where productiein.produs like 'ST'"
'''' Punctul 14
''''''                                                       adauga  cantitatile pentru PT-ST       >>>PT
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,grupa,den,sarja,nrcomanda,UM,sum(finite) as scant,'ST1' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'ST' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''
''''HF SCOS DIN LISTA
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,grupa,den,sarja,nrcomandA,UM,sum(finite) as scant,'PT1'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE (PRODUCTIEIN.PRODUS LIKE  'PT'  and PRODUCTIEIN.GRUPA NOT LIKE '*.HF') group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
''''HF IN LISTA
''''baza.Execute "INSERT INTO Intraripregatite SELECT PRODUS,grupa,den,sarja,nrcomandA,UM,sum(finite) as scant,'PT1'  AS VALOARE" & _
 ''''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PT' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''
'''''20.10.2005
''''EXTRAGE HF
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,grupa,den,sarja,nrcomandA,UM,sum(finite) as scant,'PT2'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PT'  and PRODUCTIEIN.GRUPA LIKE '*HF' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''
'''
'''
'''
'''
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'SPT1' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'ST' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'SPT1'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE (PRODUCTIEIN.PRODUS LIKE  'PT') group by valoare"
''''SUBTOTAL  HF
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'SPT2'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE (PRODUCTIEIN.PRODUS LIKE  'PT' and PRODUCTIEIN.GRUPA LIKE '*.HF') group by valoare"
''''marcat
'''BAZA.Execute "update productieie set productieiE.inR='1' where productieiE.produs like 'ST'"
'''BAZA.Execute "update productiein set productiein.inR='1' where productiein.produs like 'PT'"
''''Punctul 15 pentru linia 241
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,grupa,den,sarja,nrcomanda,LMIN,LMAX,UM,sum(finite) as scant,'MT' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MT' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare,LMIN,LMAX"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,grupa,den,sarja,nrcomandA,LMIN,LMAX,UM,sum(finite) as scant,'MT'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PT' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare,LMIN,LMAX"
'''
'''
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'sMT' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MT' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'SMT'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PT' group by valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,grupa,den,sarja,nrcomanda,UM,sum(finite) as scant,'MAA' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MA' and grupa like '*PPE' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'sMAA' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MA' and grupa like '*PPE' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'sMAA'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PT' group by valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,grupa,den,sarja,nrcomanda,UM,sum(finite) as scant,'MAB' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MA' and grupa like '*BEN' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'sMAB' AS VALOARE " & _
 '''"  FROM PRODUCTIEIE WHERE PRODUCTIEIE.PRODUS LIKE  'MA' and grupa like '*BEN' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'sMAB'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PT' group by valoare"
''''punctul 16
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS, grupa,den,sarja,um, sum(finite) as SCANT ,nrcomanda, 'PFIN' AS VALOARE" & _
 '''" FROM PRODUCTIEIn " & _
 '''"WHERE PRODUCTIEIn.inR<>1 and PRODUCTIEIn.produs like 'P*' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,grupa,den, sarja, um, sum(finite)  as SCANT,nrcomanda, 'PCONS' AS VALOARE" & _
 '''" FROM PRODUCTIEIe " & _
 '''"WHERE PRODUCTIEIE.inR<>1 and PRODUCTIEIE.produs like 'P*' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT PRODUS,grupa,den, sarja, um,sum(finite)  as SCANT,nrcomanda, 'sFIN' AS VALOARE" & _
 '''" FROM PRODUCTIEIn " & _
 '''"WHERE PRODUCTIEIn.inR<>1 and PRODUCTIEIn.produs like 'S*' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT PRODUS,grupa,den, sarja, um, sum(finite)  as SCANT,nrcomanda, 'sCONS' AS VALOARE" & _
 '''" FROM PRODUCTIEIe " & _
 '''"WHERE PRODUCTIEIE.inR<>1 and PRODUCTIEIE.produs like 'S*' group by PRODUS,grupa,den,sarja,nrcomanda,UM,valoare"
'''' Punctul 17totaluri
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT produs ,sum(finite) as scant,'TOTp'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'P*' AND GRUPA NOT LIKE '*.HF' group by produs,valoare"
'''
'''
'''
'''
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT produs ,sum(finite) as scant,'TOTs'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'S*' group by produs,valoare"
'''
''''02.11.2005
'''
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT produs ,sum(finite) as scant,'TOTbt'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'SB' or PRODUCTIEIN.PRODUS LIKE  'SC'group by produs,valoare"
'''
'''
'''
'''''''''gata
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT sum(finite) as scant,'TOTt'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'S*' group by valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT sum(finite) as scant,'TOTt'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUCTIEIe.PRODUS LIKE  'T*' group by valoare"
'''BAZA.Execute "INSERT INTO Intraripregatite SELECT produs ,sum(finite) as scant,'TOTm'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIN WHERE PRODUCTIEIN.PRODUS LIKE  'PB' group by produs,valoare"
'''BAZA.Execute "INSERT INTO Iesiripregatite SELECT produs ,sum(finite) as scant,'TOTm'  AS VALOARE" & _
 '''"  FROM PRODUCTIEIe WHERE PRODUCTIEIe.PRODUS LIKE  'MB' group by produs,valoare"
'''BAZA.Execute ("delete * from productieie")
'''BAZA.Execute ("delete * from productiein")
''''ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
''''ooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
'''BAZA.Execute "UPDATE Iesiripregatite " _
 '''& "SET scant = 0 " _
 '''& "WHERE isnull(scant);"
'''BAZA.Execute "UPDATE Intraripregatite " _
 '''& "SET scant = 0 " _
 '''& "WHERE isnull(scant);"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja," & _
 '''" Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'cons' AS VALOARE,10 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='cons'))"
''''PUNCTUL 2
'''''''''''                                                                            TB  consumat +DIFERENTA ordin 61,62,63,64
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja," & _
 '''" Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'TB' AS VALOARE,61 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='TB'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja, " & _
 '''"Intraripregatite.um, Intraripregatite.SCANT AS semif,Intraripregatite.nrcomanda, 'TB' AS VALOARE,62 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='TB'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'TB' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS SEMIF,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif((Iesiripregatite.SCANT-Intraripregatite.SCANT)=0,null,Iesiripregatite.SCANT-Intraripregatite.SCANT) AS DIFERENTA, " & _
 '''"'TB' AS VALOARE,63 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='STB')) "
''''baza.Execute "INSERT into PRODUCTIEIn ( PRODUS,grupa, den, sarja,  VALOARE,ORDIN) values ('---','---','----','---','----',64)"
''''PUNCTUL 3
'''''''''''                                                                            SB  - PB ordin 71,72,73,74
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'SB' AS VALOARE,71 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SB'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja, " & _
 '''"Intraripregatite.um, Intraripregatite.SCANT AS finite,Intraripregatite.nrcomanda, 'SB' AS VALOARE,72 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='SB'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'SB' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite," & _
 '''"Iesiripregatite.SCANT AS CONSUMATE,iif((Iesiripregatite.SCANT-Intraripregatite.SCANT)=0,null,Iesiripregatite.SCANT-Intraripregatite.SCANT) " & _
 '''"AS DIFERENTA, 'SB' AS VALOARE,73 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SSB')) "
''''baza.Execute "INSERT into PRODUCTIEIn ( PRODUS,grupa, den, sarja,  VALOARE,ORDIN) values ('---','---','----','---','----',74)"
''''PUNCTUL 4
'''''''''''                                                                          MB-VB* PB cu variabila VB ordin 81,82,83,84
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS,Iesiripregatite.DEN," & _
 '''"Iesiripregatite.SCANT AS CONSUMATE,iif((Iesiripregatite.SCANT-  " & B & " *Intraripregatite.SCANT)=0,null,Iesiripregatite.SCANT-  " & B & " *" & _
 '''"Intraripregatite.SCANT) AS DIFERENTA, 'PB' AS VALOARE,81 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) and " & _
 '''"(Iesiripregatite.den = Intraripregatite.den)" & _
 '''"WHERE (((Iesiripregatite.VALOARE)='pb')) "
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'SB' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite," & _
 '''"Iesiripregatite.SCANT AS CONSUMATE,SUM(Iesiripregatite.SCANT)- " & B & " *SUM(Intraripregatite.SCANT)AS DIFERENTA, " & _
 '''"'PB' AS VALOARE,82 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SPB')) " & _
 '''" group by Intraripregatite.scant,Iesiripregatite.scant"
'''
'''
'''BAZA.Execute "INSERT into PRODUCTIEIn ( grupa, den, sarja,  VALOARE,ORDIN) values ('===============','=====','=======','===',83)"
''''PUNCTUL 5
''''                                                                                TC - SC  ordin 91,92,93,94
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'TC' AS VALOARE,91 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='TC'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja," & _
 '''" Intraripregatite.um, Intraripregatite.SCANT AS semif,Intraripregatite.nrcomanda, 'TC' AS VALOARE,92 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='TC'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'TC' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS semif," & _
 '''"Iesiripregatite.SCANT AS CONSUMATE,iif((Iesiripregatite.SCANT-Intraripregatite.SCANT)=0,null,Iesiripregatite.SCANT-Intraripregatite.SCANT) " & _
 '''"AS DIFERENTA, 'TC' AS VALOARE,93 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='StC')) "
''''baza.Execute "INSERT into PRODUCTIEIn ( PRODUS,grupa, den, sarja,  VALOARE,ORDIN) values ('---','---','----','---','----',94)"
''''PUNCTUL 6
''''                                                                              SC-PC  ordin 101,102,103,104
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'SC' AS VALOARE,101 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SC'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja, " & _
 '''"Intraripregatite.um, Intraripregatite.SCANT AS finite,Intraripregatite.nrcomanda, 'SC' AS VALOARE,102 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='SC'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'SC' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif((Iesiripregatite.SCANT-Intraripregatite.SCANT)=0,null,Iesiripregatite.SCANT-Intraripregatite.SCANT) AS DIFERENTA, 'SC' AS" & _
 '''" VALOARE,103 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SSC')) "
''''baza.Execute "INSERT into PRODUCTIEIn ( PRODUS,grupa, den, sarja,  VALOARE,ORDIN) values ('---','---','----','---','----',104)"
'''' PUNCTUL 7
''''                                                                            MC AK 111,112,113,114
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'MCAK' AS VALOARE,112 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='MCAK'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja, " & _
 '''"Intraripregatite.um, Intraripregatite.SCANT AS finite,Intraripregatite.nrcomanda, 'mcak' AS VALOARE,111 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='mcak'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'MC' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif((Intraripregatite.SCANT-Iesiripregatite.SCANT)=0,null,Iesiripregatite.SCANT-Intraripregatite.SCANT) AS DIFERENTA, 'MCAK' AS VALOARE,113 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SMCAK')) "
''''                                                                            MC PIN 121,122,123,124
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'MCPI' AS VALOARE,121 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='MCPI'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'MC' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif((Intraripregatite.SCANT-Iesiripregatite.SCANT)=0,null,- Intraripregatite.SCANT+Iesiripregatite.SCANT) AS DIFERENTA, " & _
 '''"'MCPI' AS VALOARE,123 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SMCPI')) "
''''                                                                            MC PIN 131,132,133,134
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'MCTO' AS VALOARE,131 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='MCTO'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'MC' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif((Intraripregatite.SCANT-Iesiripregatite.SCANT)=0,null,- Intraripregatite.SCANT+Iesiripregatite.SCANT) AS DIFERENTA, " & _
 '''"'MCTO' AS VALOARE,133 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SMCTO')) "
''''PUNCTUL 8
''''                                                                            CA FOLIE,BULE 141,142,143,144
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'CA' AS VALOARE,142 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='CA'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja," & _
 '''" Intraripregatite.um, Intraripregatite.SCANT AS finite,Intraripregatite.nrcomanda, 'CA' AS VALOARE,141 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='CA'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'CA' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT AS " & _
 '''"CONSUMATE,iif(((Iesiripregatite.SCANT * " & C & ") - Intraripregatite.SCANT)=0,null,((Iesiripregatite.SCANT * " & C & ") - Intraripregatite.SCANT)) " & _
 '''"AS " & _
 '''"DIFERENTA, 'CA' AS VALOARE,143 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SCA')) "
''''PUNCTUL 9
''''                                                                           PE       ,151
'''
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'PE' AS VALOARE,151 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='PE'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja, " & _
 '''"Intraripregatite.um, Intraripregatite.SCANT AS finite,Intraripregatite.nrcomanda, 'CA' AS VALOARE,152 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='PE'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'PE' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT AS " & _
 '''"CONSUMATE, " & _
 '''"'MCPI' AS VALOARE,153 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SPE')) "
''''PUNCTUL 10
''''                                                                       PE-CENit  160 ,161,162
'''
'''
'''
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'PE' AS VALOARE,161 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='CEN'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'CE' as PRODUS,'Subtotal' as grupa," & _
 '''"Iesiripregatite.SCANT AS CONSUMATE,iif((Intraripregatite.SCANT-Iesiripregatite.SCANT)=0,null,-(Intraripregatite.SCANT-Iesiripregatite.SCANT)) " & _
 '''"AS DIFERENTA, 'MCPI' AS VALOARE,162 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SCEN')) "
''''                                                                       PE-CENit   ,171
'''
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'CE' as PRODUS,'Subtotal' as grupa," & _
 '''"Iesiripregatite.SCANT AS CONSUMATE,iif((Intraripregatite.SCANT-Iesiripregatite.SCANT)=0,null,-(Intraripregatite.SCANT-Iesiripregatite.SCANT)) " & _
 '''"AS DIFERENTA, 'MCPI' AS VALOARE,172 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SCES')) "
'''BAZA.Execute "INSERT into PRODUCTIEIn (grupa, den, sarja,  VALOARE,ORDIN) values ('===============','=====','=======','===',179)"
'''
''''Punctul 11               a                                           CFa  -TLa    181
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja, " & _
 '''"Intraripregatite.um, Intraripregatite.SCANT AS CONSUMATE,Intraripregatite.nrcomanda, 'CF' AS VALOARE,181 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='CFa'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja," & _
 '''" Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'TL' AS VALOARE,182 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='TLA'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'TL' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif(((Intraripregatite.SCANT * " & L1 & ") - Iesiripregatite.SCANT)=0,null,((Intraripregatite.SCANT * " & L1 & ") - Iesiripregatite.SCANT))" & _
 '''" AS DIFERENTA, 'TL' AS VALOARE,183 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='STLA')) "
''''Punctul 11     b                                                    CFb  -TLb  191
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja, " & _
 '''"Intraripregatite.um, Intraripregatite.SCANT AS finite,Intraripregatite.nrcomanda, 'CF' AS VALOARE,191 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='CFb'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'TL' AS VALOARE,192 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='TLb'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'TL' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT" & _
 '''" AS CONSUMATE,iif((Intraripregatite.SCANT * " & L2 & " - Iesiripregatite.SCANT)=0,null,(Intraripregatite.SCANT * " & L2 & " - Iesiripregatite.SCANT)) " & _
 '''"AS DIFERENTA, 'TL' AS VALOARE,193 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='STLb')) "
''''Punctul 12                      listeaza PL        200
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja, " & _
 '''"Intraripregatite.um, Intraripregatite.SCANT AS finite,Intraripregatite.nrcomanda, 'PL' AS VALOARE,200 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='PL'))"
''''                                       listeaza CFa
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS finite,Iesiripregatite.nrcomanda, 'CF' AS VALOARE,201 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='CFAa'))"
''''                                      subtotal CFa-PL
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'CF' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif((Intraripregatite.SCANT- Iesiripregatite.SCANT)=0,null,(Intraripregatite.SCANT - Iesiripregatite.SCANT)) AS DIFERENTA, 'TL' AS VALOARE,202 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SCFaA')) "
''''                                       listeaza CFb
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS finite,Iesiripregatite.nrcomanda, 'CF' AS VALOARE,203 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='CFbb'))"
''''                                      subtotal CFb-PL
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'CF' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite," & _
 '''"Iesiripregatite.SCANT AS CONSUMATE,iif((Intraripregatite.SCANT  - Iesiripregatite.SCANT)=0,null,(Intraripregatite.SCANT - Iesiripregatite.SCANT)) " & _
 '''"AS DIFERENTA, 'TL' AS VALOARE,204 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SCFbb')) "
''''                                       listeaza CLa
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS finite,Iesiripregatite.nrcomanda, 'CF' AS VALOARE,205 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='CLa'))"
''''                                      subtotal Cla-PL
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'CL' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite," & _
 '''"Iesiripregatite.SCANT AS CONSUMATE,iif((Intraripregatite.SCANT  - Iesiripregatite.SCANT)=0,null,(Intraripregatite.SCANT - Iesiripregatite.SCANT)) " & _
 '''"AS DIFERENTA, 'TL' AS VALOARE,206 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SCLA')) "
''''                                       listeaza CLb
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS finite,Iesiripregatite.nrcomanda, 'CF' AS VALOARE,207 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='CLb'))"
''''                                      subtotal Clb-PL
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'CL' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT" & _
 '''" AS CONSUMATE,iif((Intraripregatite.SCANT  - Iesiripregatite.SCANT)=0,null,(Intraripregatite.SCANT - Iesiripregatite.SCANT)) AS DIFERENTA, " & _
 '''"'TL' AS VALOARE,208 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SCLb')) "
''''                                       listeaza TL
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS finite,Iesiripregatite.nrcomanda, 'CF' AS VALOARE,209 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='TL'))"
''''                                      subtotal TL-PL
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'TL' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif((Intraripregatite.SCANT - Iesiripregatite.SCANT)=0,null,(Intraripregatite.SCANT - Iesiripregatite.SCANT)) AS DIFERENTA," & _
 '''" 'TL' AS VALOARE,210 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='STL')) "
''''                                       listeaza ML
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja," & _
 '''" Iesiripregatite.um, Iesiripregatite.SCANT AS finite,Iesiripregatite.nrcomanda, 'CF' AS VALOARE,211 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='TL'))"
''''                                      subtotal ML-PL
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'ML' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif((Intraripregatite.SCANT * " & L3 & " - Iesiripregatite.SCANT)=0,null,(Intraripregatite.SCANT * " & L3 & " - Iesiripregatite.SCANT)) AS DIFERENTA, 'TL' AS VALOARE,212 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SML')) "
''''punctul 13    220
'''' listeaza TT
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'SB' AS VALOARE,220 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='TT'))"
'''' listeaza ST
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja, " & _
 '''"Intraripregatite.um, Intraripregatite.SCANT AS semif,Intraripregatite.nrcomanda, 'SB' AS VALOARE,221 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='ST'))"
'''' calculeza subtotaluri si diferenta
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'ST' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS semif,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif((Iesiripregatite.SCANT-Intraripregatite.SCANT)=0,null,Iesiripregatite.SCANT-Intraripregatite.SCANT) AS DIFERENTA, " & _
 '''"'SB' AS VALOARE,222 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='STT')) "
''''baza.Execute "INSERT into PRODUCTIEIn ( PRODUS,grupa, den, sarja,  VALOARE,ORDIN) values ('---','---','----','---','----',224)"
''''punctul 14   230
'''' listeaza ST
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'ST' AS VALOARE,230 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='ST1'))"
'''' listeaza PT
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja, " & _
 '''"Intraripregatite.um, Intraripregatite.SCANT AS finite,Intraripregatite.nrcomanda, 'PT' AS VALOARE,231 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='PT1'))"
'''
'''
'''''29.10.2005
'''
''''PUNE HF DUPA PT-URI
'''
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja, " & _
 '''"Intraripregatite.um, Intraripregatite.SCANT AS finite,Intraripregatite.nrcomanda, 'PT' AS VALOARE,232 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='PT2'))"
'''
'''
'''
'''
'''' calculeza subtotaluri si diferenta
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'ST' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif((Iesiripregatite.SCANT-Intraripregatite.SCANT)=0,null,Iesiripregatite.SCANT-Intraripregatite.SCANT) AS DIFERENTA," & _
 '''" 'SB' AS VALOARE,233 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SPT1')) "
'''
''''PUNE SUBTOTAL HF LA 234
'''
'''
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT '-' as PRODUS,'Subtotal HF' as grupa,Intraripregatite.SCANT AS SEMIF,'-' " & _
 '''"AS CONSUMATE,'-' AS DIFERENTA," & _
 '''" '-' AS VALOARE,233 AS ORDIN" & _
 '''" FROM Intraripregatite  " & _
 '''"WHERE (((INTRARIpregatite.VALOARE)='SPT2')) "
''''PUNE SUBTOTAL HF LA 262
'''
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT '-' as PRODUS,'        HF' as grupa,Intraripregatite.SCANT AS finite,'-' " & _
 '''"AS CONSUMATE,'-' AS DIFERENTA," & _
 '''" '-' AS VALOARE,262 AS ORDIN" & _
 '''" FROM Intraripregatite  " & _
 '''"WHERE (((INTRARIpregatite.VALOARE)='SPT2')) "
'''
''''baza.Execute "INSERT into PRODUCTIEIn ( PRODUS,grupa, den, sarja,  VALOARE,ORDIN) values ('---','---','----','---','----',234)"
''''Punctul 15
'''''baza.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'pe' AS VALOARE,240 AS ORDIN" & _
 '''''" FROM Iesiripregatite " & _
 '''''"WHERE (((Iesiripregatite.VALOARE)='MT')
'''
'''
'''
'''
'''
'''
'''
''''original
'''
''''baza.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.produs, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 ''''"Iesiripregatite.nrcomanda, Iesiripregatite.SCANT AS CONSUMATE, IIf((Iesiripregatite.SCANT <= Intraripregatite.SCANT*Intraripregatite.LMAX/100)," & _
 ''''"null,Iesiripregatite.SCANT- " & _
 ''''"(INTRARIpregatite.Lmax / 100)* IIf(Intraripregatite.SCANT,Intraripregatite.SCANT,0)) AS DIFERENTA, 'SB' AS VALOARE" & _
 ''''", 241 AS ORDIN " & _
 ''''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) AND " & _
 ''''"(Iesiripregatite.nrcomanda = Intraripregatite.nrcomanda) AND (Iesiripregatite.sarja = Intraripregatite.sarja) AND" & _
 ''''" (Iesiripregatite.den = Intraripregatite.den) " & _
 ''''"WHERE (((Iesiripregatite.VALOARE) Like 'mt'));"
'''
''''modificat 05.02.2006
'''
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.produs, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, Iesiripregatite.nrcomanda, Iesiripregatite.SCANT AS CONSUMATE, IIf((Iesiripregatite.SCANT <= Intraripregatite.SCANT*Intraripregatite.LMAX),null,Iesiripregatite.SCANT- (INTRARIpregatite.Lmax )* IIf(Intraripregatite.SCANT,Intraripregatite.SCANT,0)) AS DIFERENTA, 'SB' AS VALOARE, 241 AS ORDIN FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.den = Intraripregatite.den) AND (Iesiripregatite.sarja = Intraripregatite.sarja) AND (Iesiripregatite.nrcomanda = Intraripregatite.nrcomanda) AND (Iesiripregatite.VALOARE = Intraripregatite.VALOARE)WHERE (((Iesiripregatite.VALOARE) Like 'mt'));"
'''
'''
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'pe' AS VALOARE,243 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='MAA'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'MA' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif((Iesiripregatite.SCANT  - Intraripregatite.SCANT)=0,null,(Iesiripregatite.SCANT - Intraripregatite.SCANT)) AS DIFERENTA," & _
 '''"'ma' AS VALOARE,244 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SMAA')) "
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Iesiripregatite.PRODUS, Iesiripregatite.grupa, Iesiripregatite.den, Iesiripregatite.sarja, " & _
 '''"Iesiripregatite.um, Iesiripregatite.SCANT AS CONSUMATE,Iesiripregatite.nrcomanda, 'pe' AS VALOARE,245 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='MAB'))"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'MA' as PRODUS,'Subtotal' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT " & _
 '''"AS CONSUMATE,iif((Iesiripregatite.SCANT  - Intraripregatite.SCANT)=0,null,(Iesiripregatite.SCANT - Intraripregatite.SCANT)) AS DIFERENTA, " & _
 '''"'ma' AS VALOARE,246 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='SMAB')) "
'''' punctul 16
'''BAZA.Execute "INSERT into PRODUCTIEIn ( PRODUS,grupa, den, sarja,  VALOARE,ORDIN) values ('---','PROD NERAPORTAT','----','---','----',251)"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS, Intraripregatite.grupa, Intraripregatite.den, Intraripregatite.sarja, " & _
 '''"Intraripregatite.um, Intraripregatite.SCANT AS finite,Intraripregatite.nrcomanda, 'P' AS VALOARE,252 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE Intraripregatite.VALOARE LIKE 'PFIN' "
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT PRODUS,grupa,den, sarja, um, SCANT AS consumatE,nrcomanda," & _
 '''" 'P' AS VALOARE,253 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE Iesiripregatite.VALOARE LIKE 'PCONS' "
'''BAZA.Execute "INSERT into PRODUCTIEIn ( PRODUS,grupa, den, sarja,  VALOARE,ORDIN) values ('---','SEMIF NERAPORTAT','----','---','----',247)"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT PRODUS,grupa,den, sarja, um, SCANT AS semif,nrcomanda, 'P' AS VALOARE,248 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE Intraripregatite.VALOARE LIKE 'SFIN' "
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT PRODUS,grupa,den, sarja, um, SCANT AS consumatE,nrcomanda, " & _
 '''"'P' AS VALOARE,248 AS ORDIN" & _
 '''" FROM Iesiripregatite " & _
 '''"WHERE Iesiripregatite.VALOARE LIKE 'SCONS' "
''''Punctul 17 Totaluri
'''BAZA.Execute "INSERT into PRODUCTIEIn ( PRODUS,grupa, den, sarja,  VALOARE,ORDIN) values ('---','TOTAL PRODUSE','----','---','----',260)"
''''total produse diferite de  PT
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS,Intraripregatite.SCANT AS finite, '-' AS grupa,261 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE ((Intraripregatite.VALOARE)='TOTp' and Intraripregatite.PRODUS not like 'PT')"
'''
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS,Intraripregatite.SCANT AS finite, 'PT-HF' AS grupa,261 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='TOTp'and Intraripregatite.PRODUS like 'PT'))"
'''
'''
'''
'''BAZA.Execute "INSERT into PRODUCTIEIn ( PRODUS,grupa, den, sarja,  VALOARE,ORDIN) values ('---','TOTAL SEMIF','----','---','----',254)"
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT Intraripregatite.PRODUS,Intraripregatite.SCANT AS semif, 's' AS VALOARE,255 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='TOTs'))"
'''
'''
''''02.11.2005
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'SB+SC' as grupa ,SUM(Intraripregatite.SCANT) AS semif,256 AS ORDIN" & _
 '''" FROM Intraripregatite " & _
 '''"WHERE (((Intraripregatite.VALOARE)='TOTbt')) GROUP BY GRUPA,ORDIN"
'''
'''
''''265
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT  'T*-S*' as grupa,Intraripregatite.SCANT AS semif,Iesiripregatite.SCANT AS CONSUMATE," & _
 '''"iif((Iesiripregatite.SCANT - Intraripregatite.SCANT)=0,null,(Iesiripregatite.SCANT - Intraripregatite.SCANT)) AS DIFERENTA, " & _
 '''"'TL' AS VALOARE,265 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='TOTt')) "
'''BAZA.Execute "INSERT INTO PRODUCTIEIn SELECT 'MB-Vb*PB' as grupa,Intraripregatite.SCANT AS finite,Iesiripregatite.SCANT AS CONSUMATE," & _
 '''"iif((Iesiripregatite.SCANT-Intraripregatite.SCANT * " & B & ")=0,null,-(Intraripregatite.SCANT * " & B & " - Iesiripregatite.SCANT)) AS DIFERENTA, " & _
 '''"'TL' AS VALOARE,270 AS ORDIN" & _
 '''" FROM Iesiripregatite LEFT JOIN Intraripregatite ON (Iesiripregatite.VALOARE = Intraripregatite.VALOARE) " & _
 '''"WHERE (((Iesiripregatite.VALOARE)='TOTm')) "
'''BAZA.Execute "INSERT into PRODUCTIEIe SELECT distinct * from productiein order by ordin"
'''BAZA.Execute "DELETE * FROM PRODUCTIEIN"
'''BAZA.Execute "INSERT into PRODUCTIEIn SELECT * from productieie order by ordin"
'''BAZA.Close
'''Set BAZA = Nothing
'''
'''End Function
'''


'''Function lunas(numar As Integer)
'''    Dim a As String
'''    Dim an
'''    an = Right([Forms]![FISA CUMULATIVA SEMIFABRICATE]![DATA].Value, 4)
'''    Select Case numar
'''    Case 1
'''        a = "ianuarie " & an
'''    Case 2
'''        a = "februarie " & an
'''    Case 3
'''        a = "martie " & an
'''    Case 4
'''        a = "aprilie " & an
'''    Case 5
'''        a = "mai " & an
'''    Case 6
'''        a = "iunie " & an
'''    Case 7
'''        a = "iulie " & an
'''    Case 8
'''        a = "august " & an
'''    Case 9
'''        a = "septembrie " & an
'''    Case 10
'''        a = "octombrie " & an
'''    Case 11
'''        a = "noiembrie " & an
'''    Case 12
'''        a = "decembrie " & an
'''    End Select
'''    lunas = a
'''End Function

Function titluprod()
End Function
'Function INITIAL()
'
'
'
'    Dim N As Integer
'
'    Dim baza As DAO.Database
'    Set baza = CurrentDb
'    baza.Execute "create table selectii (produs text(4),grupa text(50),den text(25),sarja text(8),um text(2),cant numeric,data date,nume text(20),prenume text (20),nrcomanda text(7),intrat numeric)"
'    baza.Execute "create table manevra (produs text(4),grupa text(50),den text(25),sarja text(8),um text(2),cant numeric,data date,nume text(20),prenume text (20),nrcomanda text(7),initial numeric)"
'    baza.Execute "create table materialselectat (produs text(4),grupa text(50),UM TEXT(2),den text(25),sarja text(8),nrcomanda text(7))"
'    baza.Execute "create table tiesit (produs text(4),grupa text(50),den text(25),sarja text(8),um text(2),cant numeric,data date,nume text(20),prenume text (20),nrcomanda text(7),iesit numeric)"
'    baza.Execute "create table tintrat (produs text(4),grupa text(50),den text(25),sarja text(8),um text(2),cant numeric,data date,nume text(20),prenume text (20),nrcomanda text(7),intrat numeric)"
'    baza.Execute "create table tstoc (produs text(4),grupa text(50),den text(25),sarja text(8),um text(2),cant numeric,data date,nume text(20),prenume text (20),nrcomanda text(7),initial numeric)"
'    baza.Execute "create table CUMULAT (produs text(4),nume text(20),prenume text (20))"
'    For N = 1 To 31
'        baza.Execute "alter table cumulAT add column " & CStr(N) & " numeric"
'    Next
'    baza.Execute "alter table cumulAT add column GOL TEXT(1)"
'    baza.Execute "create table exista (produs text(4),grupa text(50),den text(25),sarja text(8),um text(2),cant numeric,data date,nume text(20),prenume text (20),nrcomanda text(7),intrat numeric)"
'    baza.Execute "create table stoc (produs text(4),grupa text(50),den text(25),sarja text(8),um text(2),nrcomanda text(7),initial numeric,intrat numeric,iesit numeric ,stoc numeric)"
'    baza.Close
'
'    Set baza = Nothing
'
'    StergeInterogare "intrat"
'    StergeInterogare "iesit"
'    StergeInterogare "initial"
'    StergeInterogare "qselectii"
'    StergeInterogare "qstoc"
'    CreeazaInterogare "", "SELECT intrari.PRODUS, IIF(intrari.GRUPA,INTRARI.GRUPA,0) AS GRUPA, IIF(intrari.DEN,INTRARI.DEN,0) AS DEN, " & _
     '                          "IIF(intrari.SARJA,INTRARI.SARJA,0) AS SARJA, intrari.um, IIF(intrari.NRCOMANDA,INTRARI.NRCOMANDA,0) AS NRCOMANDA, sum( intrari.CANT) AS intrat " & _
     '                          " From intrari " & _
     '                          " Where(((intrari.Data) > #1/1/" & year(Date) & "#) And ((intrari.Data) <= #" & month(Date) & "/" & day(Date) & "/" & year(Date) & "#) )" & _
     '                          " GROUP BY intrari.PRODUS, intrari.GRUPA, intrari.DEN, intrari.SARJA, intrari.um, intrari.NRCOMANDA;", "intrat"
'
'
'    CreeazaInterogare "", "SELECT IESIRI.PRODUS, IIF(IESIRI.GRUPA,IESIRI.GRUPA,0) AS GRUPA, IIF(IESIRI.DEN,IESIRI.DEN,0) AS DEN, " & _
     '                          "IIF(IESIRI.SARJA,IESIRI.SARJA,0) AS SARJA, IESIRI.um, IIF(IESIRI.NRCOMANDA,IESIRI.NRCOMANDA,0) AS NRCOMANDA, Sum(IESIRI.CANT) AS iesit" & _
     '                          " from IESIRI " & _
     '                          " GROUP BY IESIRI.PRODUS, IESIRI.GRUPA, IESIRI.DEN, IESIRI.SARJA, IESIRI.um, IESIRI.NRCOMANDA;", "iesit"
'
'
'    CreeazaInterogare "", "SELECT intrari.PRODUS, IIF( intrari.GRUPA,INTRARI.GRUPA,0) AS GRUPA, IIF(intrari.DEN,INTRARI.DEN,0) AS DEN, " & _
     '                          "IIF(intrari.SARJA,INTRARI.SARJA,0) AS SARJA, intrari.um, IIF(intrari.NRCOMANDA,INTRARI.NRCOMANDA,0) AS NRCOMANDA, Sum(intrari.CANT) AS initial" & _
     '                          " From intrari" & _
     '                          " Where ((intrari.Data) < #1/1/" & year(Date) & "#)" & _
     '                          " GROUP BY intrari.PRODUS, intrari.GRUPA, intrari.DEN, intrari.SARJA, intrari.um, intrari.NRCOMANDA;", "initial"
'
'    CreeazaInterogare "", "SELECT intrari.PRODUS, iif(intrari.GRUPA,intrari.grupa,0) AS grupa, iif(intrari.DEN,intrari.den,0) AS den, " & _
     '                          "iif(intrari.SARJA,intrari.sarja,0) AS sarja, intrari.um, iif(intrari.NRCOMANDA,intrari.nrcomanda,0) AS nrcomanda " & _
     '                          " From intrari " & _
     '                          " GROUP BY intrari.PRODUS, intrari.GRUPA, intrari.DEN, intrari.SARJA, intrari.um, intrari.NRCOMANDA;", "qselectii"
'
'
'    CreeazaInterogare "", "SELECT DISTINCT qselectii.PRODUS, qselectii.grupa, qselectii.den, qselectii.sarja, qselectii.um, qselectii.nrcomanda, " & _
     '                          "initial.initial, intrat.intrat, iesit.iesit, IIf((initial.initial),(initial.initial),0)+IIf((intrat.intrat),(intrat.intrat),0)-IIf((iesit.iesit),(iesit.iesit),0) AS stoc " & _
     '                          " FROM ((qselectii LEFT JOIN intrat ON (qselectii.PRODUS = intrat.PRODUS) AND (qselectii.grupa = intrat.GRUPA) AND (qselectii.den = intrat.DEN) " & _
     '                          "AND (qselectii.sarja = intrat.SARJA) AND (qselectii.nrcomanda = intrat.NRCOMANDA)) LEFT JOIN iesit ON (qselectii.PRODUS = iesit.PRODUS) " & _
     '                          "AND (qselectii.grupa = iesit.GRUPA) AND (qselectii.den = iesit.DEN) AND (qselectii.sarja = iesit.SARJA) " & _
     '                          "AND (qselectii.nrcomanda = iesit.NRCOMANDA)) " & _
     '                          "LEFT JOIN initial ON (qselectii.PRODUS = initial.PRODUS) AND (qselectii.grupa = initial.GRUPA) AND (qselectii.den = initial.DEN) AND" & _
     '                          " (qselectii.sarja = initial.SARJA) AND (qselectii.nrcomanda = initial.NRCOMANDA) " & _
     '                          "GROUP BY qselectii.PRODUS, qselectii.grupa, qselectii.den, qselectii.sarja, qselectii.um, qselectii.nrcomanda,qselectii.um, initial.initial, intrat.intrat, iesit.iesit;", "qstoc"
'''
'''erori:
'''    Select Case errNumar
'''    Case 0
'''    Case 3010
'''        Resume Next
'''    Case 3078
'''        Resume Next
'''    Case 3380
'''        Resume Next
'''    Case Else
'''        MsgBox err.number & err.description
'''
'''    End Select
'''End Function
'''
'''
'''
Sub SeteazaProprietati()
'SchimbaProprietate "StartupForm", dbText, "Customers"
    SchimbaProprietate "StartupShowDBWindow", dbBoolean, False
    SchimbaProprietate "StartupShowStatusBar", dbBoolean, False
    SchimbaProprietate "AllowBuiltinToolbars", dbBoolean, False
    SchimbaProprietate "AllowFullMenus", dbBoolean, True
    SchimbaProprietate "AllowBreakIntoCode", dbBoolean, False
    SchimbaProprietate "AllowSpecialKeys", dbBoolean, True
    SchimbaProprietate "AllowBypassKey", dbBoolean, True

End Sub

Function SchimbaProprietate(strPropName As String, varPropType As Variant, varPropValue As Variant) As Integer
    Dim dbs As DAO.Database, prp As Property
    Const conProprietateNegasita = 3270

    Set dbs = CurrentDb
    On Error GoTo InlocuireEronata
    dbs.Properties(strPropName) = varPropValue
    SchimbaProprietate = True

IesireFunctie:
    Exit Function

InlocuireEronata:
    If err = conProprietateNegasita Then
        Set prp = dbs.CreateProperty(strPropName, _
                                     varPropType, varPropValue)
        dbs.Properties.Append prp
        Resume Next
    Else

        SchimbaProprietate = False
        Resume IesireFunctie
    End If
End Function

Function importa()
    On Error GoTo erori
    Dim dbs As DAO.Database

    Set dbs = CurrentDb
    dbs.Execute "DROP TABLE INTRARI"
    dbs.Execute "DROP TABLE IESIRI"
    dbs.Execute "DROP TABLE COMENZI"
    dbs.Execute "DROP TABLE PERSONAL"
    dbs.Execute "DROP TABLE LIMITE"
    dbs.Execute "DROP TABLE VARIABILE"
    dbs.Close
    Set dbs = Nothing
    Application.RunCommand acCmdImport
erori:
    Select Case err.Number
    Case 0
    Case 3376
        Resume Next
    Case 2501
        Resume Next
    Case Else
        MsgBox err.Number & "    " & err.Description & "FUNCTIA IMPORTA"
        Resume Next
    End Select

End Function

'C:\Program Files\Microsoft Office\Office\MSACCESS.EXE  "C:\dan\PROGRAM\incheiere.mdb"
Function etich()
    Dim AP As Access.Application
    Set AP = New Access.Application
    AP.OpenCurrentDatabase ("\\Server500gb\MagServer\Magazie\PROGRAM\ETICHETE1.mdb")
    AP.visible = False

    AP.visible = True
End Function

Function Bon()
    Dim AP As Access.Application
    Set AP = New Access.Application
    AP.OpenCurrentDatabase ("\\Server500gb\MagServer\Magazie\PROGRAM\bonuri de lucru.mdb")
    AP.visible = False

    AP.visible = True
End Function
Function ASISTENTA()
    Dim SIR
    SIR = Nume_scurt(ExtrageCale$(Application.CurrentDb.name) & "\Asistenta\Asistenta Stocuri.chm")
    shell "C:\WINDOWS\hh.exe " & SIR, vbMaximizedFocus
End Function
Function CONTROL_TUB()

    On Error Resume Next
    Dim AP As Access.Application
    AP.CurrentDb
    Set AP = New Access.Application
    AP.OpenCurrentDatabase ("\\Server500gb\MagServer\Magazie\PROGRAM\CONTROALE FINALE.MDB")
    '\\Server500gb\MagServer\Magazie\PROGRAM\
    AP.CurrentDb.Execute ("STERGE_TABEL")
    AP.CurrentDb.Execute ("ATRIBUIRE_DESCRIERE tubulet")
    AP.visible = False
    AP.visible = True

End Function



'C:\Program Files\Microsoft Office\Office\MSACCESS.EXE  "C:\dan\PROGRAM\SITUATII.mdb"

''Function SITUATII()
''Dim ap As Access.Application
''Set ap = New Access.Application
''ap.OpenCurrentDatabase ("\\Server500gb\MagServer\Magazie\PROGRAM\SITUATII.mdb")
''ap.visible = False
''ap.visible = True
''ap.DoCmd.Maximize
''
''End Function

Public Function operator() As String

    Dim Strcn As String
    Dim lngs As Long
    Dim lngres As Long
    Dim rec As DAO.Recordset
    Strcn = String(1024, 0)
    lngs = 1024
    lngres = GetUserName(Strcn, lngs)
    If lngres <> 0 Then
        operator = Environ("computername")    'Mid(Strcn, 1, InStr(Strcn, Chr(0)) - 1)
        Set rec = CurrentDb.OpenRecordset("select utilizator from [4 utilizatori] where logare like '" & utilizator & "'")
        If Not rec.EOF Then operator = rec.Fields(0)
    Else
        operator = ""
    End If

End Function
Public Function CTC_ist() As String


    CTC_ist = PERSOANA_ACTIVA
    CTC_ist = Replace(CTC_ist, "DEP", "")
    CTC_ist = Replace(CTC_ist, "CTC", "")
    CTC_ist = Replace(CTC_ist, "PROD", "")

End Function

Public Function utilizator() As String

    Dim Strcn As String
    Dim lngs As Long
    Dim lngres As Long
    Dim rec As DAO.Recordset
    Strcn = String(1024, 0)
    lngs = 1024
    lngres = GetUserName(Strcn, lngs)
    If lngres <> 0 Then
        utilizator = Environ("username")    'Mid(Strcn, 1, InStr(Strcn, Chr(0)) - 1)
        Set rec = CurrentDb.OpenRecordset("select logare from [4 utilizatori] where logare like '" & utilizator & "'")
        If Not rec.EOF Then utilizator = rec.Fields(0)
    Else
        utilizator = "-"
    End If

End Function

Public Function actualizare_balante()
    Dim MMM
    For MMM = 1 To 10
        DoCmd.TransferDatabase acLink, "Microsoft Access", "\\Balante\program balante\Balante\curente.mdb", acTable, "Balanta " & MMM, "Balanta " & MMM & " curente"
    Next
    DoCmd.TransferDatabase acLink, "Microsoft Access", "\\Balante\program balante\Balante\" & CStr(year(Date)) & " comenzi.mdb", acTable, "comenzi", "comenzi balante"
End Function


Public Sub IdentificaNumereLipsaRapid()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim dict As Object
    Dim I As Long
    Dim strLipsa As String
    Dim totalLipsa As Long
    Dim startTime As Double
    Dim N As Long, M As Long
    Dim sql As String
    Dim retval As Variant
10  On Error GoTo TRATARE_ERORI
    Dim SirEroare As String
20  M = DMax("NR_COMANDA_INTERNA", "comenzi")
30  N = DMax("NR_COMANDA_INTERNA", "comenzi", "NR_COMANDA_INTERNA<>" & M)
    '''40  startTime = Timer
40  Set dict = CreateObject("Scripting.Dictionary")
50  Set db = CurrentDb
60  sql = "SELECT NR_COMANDA_INTERNA FROM COMENZI WHERE NR_COMANDA_INTERNA BETWEEN " & N & " AND " & M
70  Set rs = db.OpenRecordset(sql, dbOpenSnapshot)
80  Do While Not rs.EOF
90      If Not dict.Exists(rs!NR_COMANDA_INTERNA.Value) Then
100         dict.Add rs!NR_COMANDA_INTERNA.Value, True
110     End If
120     rs.MoveNext
130 Loop
140 rs.Close
150 Set rs = Nothing
160 strLipsa = ""
170 totalLipsa = 0
180 For I = N To M
190     If Not dict.Exists(I) Then
200         strLipsa = strLipsa & I & ", "
210         totalLipsa = totalLipsa + 1
220     End If
230 Next I
240 If totalLipsa > 0 Then
250     strLipsa = Left(strLipsa, Len(strLipsa) - 2)
260     retval = SysCmd(acSysCmdSetStatus, "Lipsesc " & totalLipsa & " numere interne")
270     ScrieEroare "Verificare [functii].[IdentificaNumereLipsaRapid] linia " & 0, "Lipsesc " & totalLipsa & " numere: " & vbCrLf & strLipsa, 0, False, "Lipsesc " & totalLipsa & " numere: " & vbCrLf & strLipsa
280     MsgBox "ATENTIE !!!  retineti ce operatii ati efectuat anterior acestui mesaj." & vbCrLf & "Lipsesc " & totalLipsa & " numere: " & vbCrLf & strLipsa, vbExclamation + vbOKOnly, "Verificare numere interne"
290 Else
300     retval = SysCmd(acSysCmdSetStatus, "Nu lipsesc numere interne")
310 End If
320 Set dict = Nothing
330 Set db = Nothing
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
340 Exit Sub

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
350 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
360 Select Case errNumar
    Case 0
370 Case Else
380     ScrieEroare "Eroare in [functii].[IdentificaNumereLipsaRapid] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
390     RaspunsMesaj = MsgBox("[Eroare in functii].[IdentificaNumereLipsaRapid] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[IdentificaNumereLipsaRapid] linia " & lngLinia & " cu numarul " & Errnumar & vbCrLf & Errdescriere"
400     If RaspunsMesaj = vbYes Then
410         Resume Next
420     Else
430         GoTo TRATARE_ERORI_iesire
440     End If
450 End Select
    '========================== terminat tratare erori
End Sub



Function dacaExistaTabel(strTableName As String) As Integer
    Dim db As DAO.Database
    Dim I As Integer
    Set db = CurrentDb
    dacaExistaTabel = False
    For I = 0 To db.TableDefs.Count - 1
        If UCase(strTableName) = UCase(db.TableDefs(I).name) Then
            'Table Exists
            dacaExistaTabel = True
            Exit For
        End If
    Next I
    Set db = Nothing
End Function

Public Function DURATA(dela As String, pinala As String) As String
    Dim PAUZ
    'inceput terminat in inainte pauza10
1   If CDate(dela) < CDate("10:00") And CDate(pinala) < CDate("10:20") Then _
       PAUZ = 0
    'inceput terminat dupa pauza10
2   If CDate(dela) < CDate("10:00") And CDate(pinala) < CDate("10:20") And CDate(pinala) > CDate("10:00") Then _
       PAUZ = DateDiff("n", CDate("10:00"), CDate(pinala))
    ' inceput pauza10 terminat inainte pauza13
3   If CDate(dela) < CDate("10:00") And CDate(pinala) > CDate("10:20") And CDate(pinala) < CDate("13:00") Then _
       PAUZ = 20
    'inceput pauza10 terminat in pauza13
4   If CDate(dela) < CDate("10:00") And CDate(pinala) < CDate("13:10") And CDate(pinala) > CDate("13:00") Then _
       PAUZ = 20 + DateDiff("n", CDate("13:00"), CDate(pinala))
    'inceput pauza10 terminat dupa pauza13
5   If CDate(dela) < CDate("10:00") And CDate(pinala) > CDate("13:10") Then _
       PAUZ = 30
    'inceput terminat dupa pauza10
6   If CDate(dela) > CDate("10:00") And CDate(dela) < CDate("10:20") And CDate(pinala) > CDate("10:20") And CDate(pinala) < CDate("13:00") Then _
       PAUZ = DateDiff("n", CDate(dela), CDate("10:20"))

7   If CDate(dela) > CDate("10:00") And CDate(dela) < CDate("10:20") And CDate(pinala) > CDate("13:00") And CDate(pinala) < CDate("13:10") Then _
       PAUZ = DateDiff("n", CDate(dela), CDate("10:20")) + DateDiff("n", CDate("13:00"), CDate(pinala))

8   If CDate(dela) > CDate("10:00") And CDate(dela) < CDate("10:20") And CDate(pinala) > CDate("13:10") Then _
       PAUZ = DateDiff("n", CDate(dela), CDate("10:20")) + 10

9   If CDate(dela) > CDate("10:20") And CDate(dela) < CDate("13:00") And CDate(pinala) > CDate("10:20") And CDate(pinala) < CDate("13:00") Then _
       PAUZ = 0
       
        If CDate(dela) > CDate("10:20") And CDate(dela) < CDate("13:00") And CDate(pinala) > CDate("13:10") Then _
       PAUZ = 10

10  If CDate(dela) > CDate("10:20") And CDate(dela) < CDate("13:00") And CDate(pinala) > CDate("13:00") And CDate(pinala) < CDate("13:10") Then _
       PAUZ = DateDiff("n", CDate("13:00"), CDate(pinala))

11  If CDate(dela) > CDate("10:20") And CDate(dela) < CDate("13:00") And CDate(pinala) > CDate("13:10") And CDate(pinala) < CDate("23:10") Then _
       PAUZ = 10

12  If CDate(dela) > CDate("13:00") And CDate(dela) < CDate("13:10") And CDate(pinala) > CDate("13:10") And CDate(pinala) < CDate("23:10") Then _
       PAUZ = DateDiff("n", CDate(dela), CDate("13:10"))


13  If CDate(dela) > CDate("13:10") _
       And CDate(pinala) > CDate("13:10") Then _
       PAUZ = 0




    DURATA = DateDiff("n", CDate(dela), (pinala)) - PAUZ

End Function
Function TEST1()
MsgBox DURATA_ZZHHMM("01.08.2025 15:27:10", "01.08.2025 15:27:22")
End Function
Public Function DURATA_ZZHHMM(dela As String, pinala As String) As String
    Dim PAUZ
    Dim Interval
 Dim ore, minutee, ZILEE
10     dela = Format(dela, "hh:mm")
20      pinala = Format(pinala, "hh:mm")
30      DURATA_ZZHHMM = "0m"
    'inceput terminat in inainte pauza10
40    On Error GoTo TRATARE_ERORI
    Dim SirEroare As String
50    If CDate(dela) < CDate("10:00") And CDate(pinala) < CDate("10:20") Then _
       PAUZ = 0
    'inceput terminat dupa pauza10
60    If CDate(dela) < CDate("10:00") And CDate(pinala) < CDate("10:20") And CDate(pinala) > CDate("10:00") Then _
       PAUZ = DateDiff("n", CDate("10:00"), CDate(pinala))
    ' inceput pauza10 terminat inainte pauza13
70    If CDate(dela) < CDate("10:00") And CDate(pinala) > CDate("10:20") And CDate(pinala) < CDate("13:00") Then _
       PAUZ = 20
    'inceput pauza10 terminat in pauza13
80    If CDate(dela) < CDate("10:00") And CDate(pinala) < CDate("13:10") And CDate(pinala) > CDate("13:00") Then _
       PAUZ = 20 + DateDiff("n", CDate("13:00"), CDate(pinala))
    'inceput pauza10 terminat dupa pauza13
90    If CDate(dela) < CDate("10:00") And CDate(pinala) > CDate("13:10") Then _
       PAUZ = 30
    'inceput terminat dupa pauza10
100   If CDate(dela) > CDate("10:00") And CDate(dela) < CDate("10:20") _
       And CDate(pinala) > CDate("10:20") And CDate(pinala) < CDate("13:00") Then _
       PAUZ = DateDiff("n", CDate(dela), CDate("10:20"))

110   If CDate(dela) > CDate("10:00") And CDate(dela) < CDate("10:20") _
       And CDate(pinala) > CDate("13:00") And CDate(pinala) < CDate("13:10") Then _
       PAUZ = DateDiff("n", CDate(dela), CDate("10:20")) + DateDiff("n", CDate("13:00"), CDate(pinala))

120   If CDate(dela) > CDate("10:00") And CDate(dela) < CDate("10:20") _
       And CDate(pinala) > CDate("13:10") Then _
       PAUZ = DateDiff("n", CDate(dela), CDate("10:20")) + 10


130   If CDate(dela) > CDate("10:20") And CDate(dela) < CDate("13:00") _
       And CDate(pinala) > CDate("10:20") And CDate(pinala) < CDate("13:00") Then _
       PAUZ = 0

140   If CDate(dela) > CDate("10:20") And CDate(dela) < CDate("13:00") _
       And CDate(pinala) > CDate("13:00") And CDate(pinala) < CDate("13:10") Then _
       PAUZ = DateDiff("n", CDate("13:00"), CDate(pinala))

150   If CDate(dela) > CDate("10:20") And CDate(dela) < CDate("13:00") _
       And CDate(pinala) > CDate("13:10") And CDate(pinala) < CDate("23:10") Then _
       PAUZ = 10


160   If CDate(dela) > CDate("13:00") And CDate(dela) < CDate("13:10") _
       And CDate(pinala) > CDate("13:10") And CDate(pinala) < CDate("23:10") Then _
       PAUZ = DateDiff("n", CDate(dela), CDate("13:10"))


170   If CDate(dela) > CDate("13:10") _
       And CDate(pinala) > CDate("13:10") Then _
       PAUZ = 0



180   Interval = DateDiff("s", CDate(dela), (pinala)) - PAUZ * 60
190   If IsEmpty(Interval) Then Interval = 0
200   ZILEE = Int(Interval / 28800)
210   ore = Int(Interval / 3600)
220   ore = ore - ZILEE * 8
230   Interval = Interval - ZILEE * 28800 - ore * 3600
240   minutee = Format(CStr(Int((Interval) / 60)), "0#")
250   DURATA_ZZHHMM = IIf(ZILEE > 0, ZILEE & "z:", "") & IIf(ore > 0, ore & "h:", "") & IIf(minutee > 0, minutee & "m", "")
260   If Right(DURATA_ZZHHMM, 1) = ":" Then DURATA_ZZHHMM = Left(DURATA_ZZHHMM, Len(DURATA_ZZHHMM) - 1)
270    If DURATA_ZZHHMM = "" Then DURATA_ZZHHMM = "0m"

    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
280   Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
290   lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
300   Select Case errNumar
    Case 0
310   Case Else
320     ScrieEroare "Eroare in [functii].[DURATA_ZZHHMM] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, SirEroare
        ''280     RaspunsMesaj = MsgBox("[Eroare in functii].[DURATA_ZZHHMM] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        ''        'Executasiraspunde="INFORMATIE "[Eroare in functii].[DURATA_ZZHHMM] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
        ''290     If RaspunsMesaj = vbYes Then
        ''300         Resume Next
        ''310     Else
330     DURATA_ZZHHMM = "erZZHHMM"
340     GoTo TRATARE_ERORI_iesire
        ''330     End If
350   End Select
    '========================== terminat tratare erori

End Function



Function ExecutaInlocuire(ByVal sInput As String, Inlocuitor As String) As String
    On Error Resume Next
    If Len(sInput) > 1 Then
        Dim a, ENC1, ENC2, J, TEMP, LENG, I, B
        a = Left$(Inlocuitor, 2)
        J = 2
        While J > 0
            ENC1 = val(ENC1) + Mid$(a, J, 1)
            J = J - 1
        Wend
        B = Right$(Inlocuitor, 2)
        J = 2
        While J > 0
            ENC2 = val(ENC2) + Mid$(B, J, 1)
            J = J - 1
        Wend
        LENG = Len(sInput)
        I = 1
        While I > 0
            TEMP = 0
            TEMP = Asc(Mid$(sInput, I, 1))
            If I Mod 2 = 0 Then
                TEMP = TEMP + ENC1
            Else
                TEMP = TEMP - ENC2
            End If
            ExecutaInlocuire = ExecutaInlocuire + Chr$(TEMP)
            I = I + 1
            If I = LENG + 1 Then I = 0
        Wend
    Else
        ExecutaInlocuire = "-"
    End If
End Function

Function RevenireInlocuire(ByVal sInput As String, Inlocuitor As String) As String
    On Error Resume Next
    Dim a, ENC1, ENC2, J, TEMP, LENG, I, B
    If Len(sInput) > 1 Then
        a = Left$(Inlocuitor, 2)
        J = 2
        While J > 0
            ENC1 = val(ENC1) + Mid$(a, J, 1)
            J = J - 1
        Wend
        ENC1 = ENC1
        'Assigning for even places
        B = Right$(Inlocuitor, 2)
        J = 2
        While J > 0
            ENC2 = val(ENC2) + Mid$(B, J, 1)
            J = J - 1
        Wend

        LENG = Len(sInput)
        I = 1
        While I > 0
            TEMP = 0
            TEMP = Asc(Mid$(sInput, I, 1))
            If I Mod 2 = 0 Then
                TEMP = TEMP - ENC1
            Else
                TEMP = TEMP + ENC2
            End If
            RevenireInlocuire = RevenireInlocuire + Chr$(TEMP)
            I = I + 1
            If I = LENG + 1 Then I = 0
        Wend
    Else
        RevenireInlocuire = ""
    End If


End Function

Private Function ValidateConnectStringLocal() As Boolean
    On Error Resume Next
    Dim qdfPUBS As QueryDef
    err.Clear
    'DoCmd.Hourglass True
    ValidateConnectStringLocal = True
    GasitServerulActiv = True
    Set qdfPUBS = CurrentDb.CreateQueryDef("")
    qdfPUBS.Connect = DLookup("valoare", "configurari", "proprietate='conectarelocala'")
    qdfPUBS.ReturnsRecords = False
    qdfPUBS.ODBCTimeout = 5
    qdfPUBS.sql = "SELECT MAX(COD_MATERIAL) FROM [dbo].[1 APROVIZIONARE]"
    qdfPUBS.Execute
    If err.Number Then ValidateConnectStringLocal = False: GasitServerulActiv = False
    Set qdfPUBS = Nothing
    'DoCmd.Hourglass False

End Function
Private Function ValidateConnectStringInternet() As Boolean
    On Error Resume Next
    Dim qdfPUBS As QueryDef
    err.Clear
    'DoCmd.Hourglass True
    ValidateConnectStringInternet = True
    Set qdfPUBS = CurrentDb.CreateQueryDef("")
    qdfPUBS.Connect = DLookup("valoare", "configurari", "proprietate='conectareinternet'")
    qdfPUBS.ReturnsRecords = False
    qdfPUBS.ODBCTimeout = 5
    qdfPUBS.sql = "SELECT MAX(COD_MATERIAL) FROM [dbo].[1 APROVIZIONARE]"
    qdfPUBS.Execute
    If err.Number Then ValidateConnectStringInternet = False
    Set qdfPUBS = Nothing
    'DoCmd.Hourglass False

End Function
Public Function StabilesteTimp()

    Dim rec As DAO.Recordset
    Dim DATA, ora
    Dim diferenta As String
10  On Error GoTo TRATARE_ERORI
    '''20  If blnINRETEA Then
30  DATA = Date
40  ora = time
50  Set rec = CurrentDb.OpenRecordset("Sincronizare ora")

60  If DateDiff("d", Date, CDate(rec.Fields("data"))) <> 0 Then
70      MsgBox "ATENTIE ANUNTATI PROGRAMATORUL -timpul serverului este foarte diferit de timpul statiei " & DateDiff("d", Date, CDate(rec.Fields("data"))) & " zile. NU ESTE PERMIS LUCRUL IN ACESTE CONDITII"
        If TrimiteLogare Then ClipBoard_SetData "LOGARE- " & utilizator & " " & Now() & " " & PERSOANA_ACTIVA & "' INTARZIERE FATA DE SERVER'"

80      DoCmd.Quit
90  End If

100 If DATA <> CDate(rec.Fields("data")) Or ora <> CDate(Left(rec.Fields("ora"), 8)) Then
110     DATA = CDate(rec.Fields("data"))
120     ora = CDate(Left(rec.Fields("ora"), 8))
130     diferenta = "Verificati daca este diferenta intre ora server=" & DATA & " " & ora & vbCrLf & " si ora calculator=" & Date & " " & time & vbCrLf & vbCrLf & "Daca exista diferenta anuntati administratorul de sistem."
140     Date = DATA
150     time = ora
160     rec.Close
170 End If
    '''180 End If
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
190 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
200 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
210 Select Case errNumar
    Case 0
220     Resume Next
230 Case 70
        Resume Next
        ''240             MsgBox "Nu se poate modifica ora din lipsa de drept de administrator" & vbCrLf & diferenta
        ''250             GoTo TRATARE_ERORI_iesire
260 Case Else
270     ScrieEroare "Eroare in [Module2].[StabilesteTimp] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata Module2 rutina StabilesteTimp"
280     RaspunsMesaj = MsgBox("[Eroare in Module2].[StabilesteTimp] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in Module2].[StabilesteTimp] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
290     If RaspunsMesaj = vbYes Then
300         Resume Next
310     Else
320         GoTo TRATARE_ERORI_iesire
330     End If
340 End Select
    '========================== terminat tratare erori
End Function

Public Function VerificaSincronizareCuServer(DATA As String)


    Dim qdf As QueryDef
    Dim rec As DAO.Recordset
10  On Error GoTo TRATARE_ERORI
    Dim SIR As String
    '''20  If blnINRETEA Then
30  If DATA <> "" Then
40      If InStr(1, DATA, ".") Then
50      Else
60          Exit Function
70      End If
80  End If

90  If ValidateConnectStringLocal = True Then SIR = DLookup("valoare", "configurari", "proprietate='conectarelocala'")
100 If GasitServerulActiv = False Then
110     If ValidateConnectStringInternet = True Then SIR = DLookup("valoare", "configurari", "proprietate='conectareinternet'")
120 End If
    Dim SemneOmise As String
130 SemneOmise = Nz(DLookup("valoare", "setari", "parametru='Semne omise'"), "")

140 If SemneOmise <> "" Then
        dtScadentaParola = CDate(Replace(RevenireInlocuire(SemneOmise, "4825"), "4825", ""))
150     If SIR <> "" Then
            '150         DoCmd.Hourglass True

            'CurrentDb.QueryDefs.Delete "aaaa"
160         Set qdf = CurrentDb.CreateQueryDef("aaaa")
170         qdf.Connect = SIR
180         qdf.ReturnsRecords = True
190         qdf.ODBCTimeout = 5
200         qdf.sql = "SELECT getdate() as sss "
210         Set rec = CurrentDb.OpenRecordset("aaaa")



230         Select Case DATA
            Case ""    'data server
240             If CDate(Format(rec.Fields(0), "dd.mm.yyyy")) <= dtScadentaParola Then
                    ' functioneaza normal
250             Else
260                 MsgBox "NO ACTIVE LICENSE Unable to start session (reason: Microsoft.SqlServer.Management.Server.Common.BaseException: Configuration information could not be read from the database. ---> System.Data.SqlClient.SqlException: profile name is not valid" & vbCrLf & _
                           " at System.Data.SqlClient.SqlConnection.OnError(SqlException exception, Boolean breakConnection)" & vbCrLf & _
                           " at System.Data.SqlClient.SqlInternalConnection.OnError(SqlException exception, Boolean breakConnection)" & vbCrLf & _
                           " at System.Data.SqlClient.TdsParser.ThrowExceptionAndWarning(TdsParserStateObject stateObj)"
270                 CurrentDb.Execute "update setari set valoare=valoare & '#%' where parametru='Semne omise'"
280                 CurrentDb.Execute "update setari set parametru='Semne eliminate' where parametru='Semne omise'"
290                 CurrentDb.QueryDefs.Delete "aaaa"
300                 DoCmd.Quit
310             End If

320             Set qdf = Nothing
                '320                 DoCmd.Hourglass False
330         Case Else    'data din program
340             If CDate(Format(DATA, "dd.mm.yyyy")) <= dtScadentaParola Then
                    ' functioneaza normal

350             Else
                    'probleme data calendaristica server
360                 MsgBox "Unable to start session (reason: Microsoft.SqlServer.Management.Server.Common.BaseException: Configuration information could not be read from the database. ---> System.Data.SqlClient.SqlException: profile name is not valid" & vbCrLf & _
                           " at System.Data.SqlClient.SqlConnection.OnError(SqlException exception, Boolean breakConnection)" & vbCrLf & _
                           " at System.Data.SqlClient.SqlInternalConnection.OnError(SqlException exception, Boolean breakConnection)" & vbCrLf & _
                           " at System.Data.SqlClient.TdsParser.ThrowExceptionAndWarning(TdsParserStateObject stateObj)"
370                 CurrentDb.Execute "update setari set valoare=valoare & '#%' where parametru='Semne omise'"
380                 CurrentDb.Execute "update setari set parametru='Semne eliminate' where parametru='Semne omise'"
390                 CurrentDb.QueryDefs.Delete "aaaa"
400                 DoCmd.Quit
410             End If
420             Set qdf = Nothing
                '430                 DoCmd.Hourglass False
430         End Select
440     Else
450         CurrentDb.QueryDefs.Delete "aaaa"
460         'DoCmd.Quit
470     End If    'If SIR <> "" Then
480 Else
490     'DoCmd.Quit
500 End If
510 CurrentDb.QueryDefs.Delete "aaaa"
    '''520 End If
530 Exit Function

    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
540 DoCmd.Quit
550 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
560 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
570 Select Case errNumar
    Case 0
580 Case 3265
590     Resume Next
600 Case 3012

610 Case Else

        '290         ScrieEroare "Eroare in [functii].[VerificaSincronizareCuServer] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata functii rutina VerificaSincronizareCuServer"
620     RaspunsMesaj = MsgBox("[Unable to start session] (Check date format dd.mm.yyyy for right acces SQL Server)  " & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)

        '''630     MsgBox "linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere
640     If RaspunsMesaj = vbYes Then
            shell "control intl.cpl", vbNormalFocus
650         CurrentDb.QueryDefs.Delete "aaaa"
660
670
680     Else

690         GoTo TRATARE_ERORI_iesire
700     End If
710 End Select
    '========================== terminat tratare erori
End Function
Sub PrelungireVerificareServer()
    VerificaCaractereNepermise "15.03.2026"
End Sub
Public Function VerificaCaractereNepermise(Analizat As String)
    Dim SemneOmise As String
    If InStr(1, Analizat, ".") Then

        CurrentDb.Execute "update setari set valoare='" & Replace(ExecutaInlocuire(Analizat, "4825"), "'", "''") & "' where parametru='Semne omise'"
        SemneOmise = Nz(DLookup("valoare", "setari", "parametru='Semne omise'"), "")
        MsgBox "Stabilit " + RevenireInlocuire(SemneOmise, "4825")
    End If
End Function

Public Function VerificaCaractereNepermise1()
    Dim SemneOmise As String
    SemneOmise = Nz(DLookup("valoare", "setari", "parametru='Semne omise'"), "")
    MsgBox "Verificat " + RevenireInlocuire(SemneOmise, "4825")
End Function
Public Function interna() As String
    On Error Resume Next
    interna = Forms![10 COMENZI SIMPLA]!INTERN
End Function

Public Function VerificTiparitura()
    Dim rec As DAO.Recordset
    Dim X As Integer
    Dim N As Integer
    Dim sxnr
    Dim tipariturasablon
    Dim TIPARITURA
    Dim ex
    Dim pos, pos2, pos1
10  On Error GoTo TRATARE_ERORI

20  For X = 1 To 4
30      Set rec = CurrentDb.OpenRecordset("SELECT distinct Email_issues.Response, Sarje.tiparitura, Sarje.Comanda, Sarje.data_livrare, CLISEE.TEXT, CLISEE.COD" & _
                                          " FROM ((Email_issues LEFT JOIN Sarje ON Email_issues.Order = Sarje.Comanda) LEFT JOIN Retete ON Email_issues.Product_Article = Retete.ArticolProdus) LEFT JOIN CLISEE ON Retete.ArticolMaterial = CLISEE.COD" & _
                                          " WHERE (((Email_issues.Response) Like '*" & 2020 + X & " *') AND ((CLISEE.TEXT) Like '*jjjj*'))" & _
                                          " ORDER BY Sarje.data_livrare DESC;")
40      If Not rec.EOF Then
50          rec.MoveLast
60          rec.MoveFirst
70          For N = 1 To rec.RecordCount
80              sxnr = rec.Fields("cod")
90              tipariturasablon = rec.Fields("text")
100             TIPARITURA = rec.Fields("tiparitura")
110             ex = "-"
120             If Not IsNull(TIPARITURA) Then
130                 pos = 1
140                 Do
150                     pos = InStr(pos + 1, TIPARITURA, " ")
160                     If pos > 0 Then pos2 = pos
170                 Loop Until pos = 0
180                 If pos2 > 0 Then
190                     ex = Right(TIPARITURA, Len(TIPARITURA) - pos2)
200                 End If
210             End If
220             pos1 = InStr(1, tipariturasablon, "JJJJ")
230             If pos1 > 0 Then
240                 tipariturasablon = Right(tipariturasablon, Len(tipariturasablon) - pos1 + 1)
250             Else
260                 tipariturasablon = ""
270             End If
280             Select Case InStr(1, tipariturasablon, "-")
                Case 0
290                 If InStr(1, ex, "-") Then
300                     MsgBox "Atentie valabilitatea " & ex & " nu trebuie sa contina semnul (-)."
310                 End If
320             Case 5
330                 If Not InStr(1, ex, "-") Then
340                     MsgBox "Atentie valabilitatea " & ex & " trebuie sa contina semnul (-)."
350                 End If
360             End Select
370         Next
380     End If
390 Next
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
400 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
410 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
420 Select Case errNumar
    Case 0
430 Case Else
440     ScrieEroare "Eroare in [functii].[VerificTiparitura] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata functii rutina VerificTiparitura"
450     RaspunsMesaj = MsgBox("[Eroare in functii].[VerificTiparitura] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[VerificTiparitura] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
460     If RaspunsMesaj = vbYes Then
470         Resume Next
480     Else
490         GoTo TRATARE_ERORI_iesire
500     End If
510 End Select
    '========================== terminat tratare erori

End Function
Public Function DeschideChecklist(faza As String, Optional ARATA As String = "NU", Optional Left As Long = 0, Optional filtru As String = "", Optional COMANDA As String)
10  On Error GoTo TRATARE_ERORI

    Dim inaltime
    Dim UMPLERE As String
    Dim Tiparire As String
    Dim Protocol As String
    Dim Material As String
20  If ARATA = "DA" Then
30  Else
40      ARATA = Nz(DLookup("VALOARE", "SETARI", "PARAMETRU='CheckList " & faza & "'"), "NU")
50  End If
60  UMPLERE = Nz(DLookup("articolmaterial", "retete", "articolprodus='" & filtru & "' and tipmaterial='Fullgut'"), "")
70  If UMPLERE = "" Then
80      UMPLERE = " and grupa not like 'U'"
90  Else
100     UMPLERE = ""
110 End If
120 Tiparire = Nz(DLookup("articolmaterial", "retete", "articolprodus='" & filtru & "' and tipmaterial='Druckbild'"), "")
130 If Tiparire = "" Then
140     Tiparire = " and grupa not like 'T'"
150 Else
160     Tiparire = ""
170 End If

180 Protocol = Nz(DLookup("PRODUS", "NOMENCLATOR ARTICOLE", "NRART='" & filtru & "' AND PRODUS LIKE 'P*'"), "")
190 If Protocol = "" Then
200     Protocol = " and grupa not like 'P'"
210 Else
220     Protocol = ""
230 End If

240 Material = Nz(DLookup("PRODUS", "NOMENCLATOR ARTICOLE", "NRART='" & filtru & "' AND PRODUS NOT LIKE 'P*' AND PRODUS NOT LIKE 'S*'"), "")
250 If Material <> "" Then
        If COMANDA <> "RECEPTIE" Then
            UMPLERE = ""
            Tiparire = ""
            Protocol = ""
            Material = " and grupa like 'R'"
        Else
260         Material = " and grupa not like 'M'"

        End If
270 Else
280     Material = " and grupa not like 'R'"
290 End If

300 filtru = UMPLERE & Tiparire & Protocol & Material
310 filtru = "faza = '" & faza & "' " & filtru
320 If ARATA = "DA" Then
330     DoCmd.OpenForm "CheckList", , , filtru
340     inaltime = DCount("faza", "checklist", filtru)
350     Forms("CheckList").InsideHeight = 2300 + inaltime * 360
360     Forms("CheckList").Move Left, 0

370 End If
    '========================== incep tratare erori
TRATARE_ERORI_iesire:
    'DBEngine.Rollback
380 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
390 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
400 Select Case errNumar
    Case 0
410 Case Else
420     ScrieEroare "Eroare in [functii].[DeschideChecklist] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netratata functii rutina DeschideChecklist"
430     RaspunsMesaj = MsgBox("[Eroare in functii].[DeschideChecklist] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[DeschideChecklist] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
440     If RaspunsMesaj = vbYes Then
450         Resume Next
460     Else
470         GoTo TRATARE_ERORI_iesire
480     End If
490 End Select
    '========================== terminat tratare erori

End Function
Sub InlocuireLegaturi()
    Dim baza As Database
10  On Error GoTo TRATARE_ERORI
    '''
    '''20  Set baza = CurrentDb
    '''30  On Error Resume Next
    '''    Dim N
    '''    Debug.Print baza.TableDefs("001 DATE CALCULATOARE").Properties("Connect").Value
    '''   baza.TableDefs("001 DATE CALCULATOARE").Properties("Connect").Value = "ODBC;driver=SQL Server ; Server=SERVERK8;UID=sa;PWD=comenzi2012#;DATABASE=SQLSTOC"
    '''    baza.TableDefs("001 DATE CALCULATOARE").Connect = baza.TableDefs("001 DATE CALCULATOARE").Properties("Connect").Value
    '''       Debug.Print baza.TableDefs("001 DATE CALCULATOARE").Properties("Connect").Value
    'baza.TableDefs("001 DATE CALCULATOARE").RefreshLink
    ''40  For N = 1 To baza.TableDefs.Count
    ''50      If InStr(1, baza.TableDefs(N).Name, "MSys") Then
    ''60      Else
    ''70          Select Case baza.TableDefs(N).Name
    ''            Case "4 utilizatori"
    ''
    ''80          Case Else
    ''
    ''90              If baza.TableDefs(N).Properties("Connect").Value <> "" And InStr(1, baza.TableDefs(N).Properties("Connect").Value, "ODBC;") Then
    ''100                 Debug.Print baza.TableDefs(N).Name & "   |   " & baza.TableDefs(N).Properties("Connect").Value
    ''110             Else
    ''120             End If
    ''130         End Select
    ''140     End If
    ''150 Next
    '========================== incep tratare erori
TRATARE_ERORI_iesire:

160 Exit Sub

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
170 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
180 Select Case errNumar
    Case 0
190 Case Else
200     ScrieEroare "Eroare in [functii].[InlocuireLegaturi] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netrarata functii rutina InlocuireLegaturi1"
210     RaspunsMesaj = MsgBox("[Eroare in functii].[InlocuireLegaturi] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[InlocuireLegaturi] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
220     If RaspunsMesaj = vbYes Then
230         Resume Next
240     Else
250         GoTo TRATARE_ERORI_iesire
260     End If
270 End Select
    '========================== terminat tratare erori
End Sub

Public Function cauta_text1(cautat As String)
    Dim mdl As MODULE


    Dim N
    Dim lngSLine As Long, lngSCol As Long '
    Dim lngELine As Long, lngECol As Long
    Dim aremodul As Boolean
10  On Error GoTo TRATARE_ERORI

20  For N = 1 To Application.vbe.ActiveVBProject.VBComponents.Count
30      Application.vbe.ActiveVBProject.VBComponents(N).CodeModule.codePane.Show
40      Select Case Application.vbe.ActiveVBProject.VBComponents(N).type
        Case 1, 2

50          Set mdl = Modules(Application.vbe.ActiveVBProject.VBComponents(N).name)


60      Case 100
70          If InStr(1, Application.vbe.ActiveVBProject.VBComponents(N).name, "Form_") Then
80              If Forms(N).Form.HasModule Then
90                  aremodul = True
100                 Set mdl = Forms(N).Form.MODULE
110
120             Else
130                 aremodul = False
140             End If
150         End If
160         If InStr(1, Application.vbe.ActiveVBProject.VBComponents(N).name, "Report_") Then
170             If Reports(N).Report.HasModule Then
180                 aremodul = True
190                 Set mdl = Reports(Replace(Application.vbe.ActiveVBProject.VBComponents(N).name, "Report_", "")).Report.MODULE
200             Else
210                 aremodul = False
220             End If
230         End If
240     End Select
250     If aremodul Then
260         If mdl.Find(cautat, lngSLine, lngSCol, lngELine, lngECol) Then

270             Application.vbe.ActiveVBProject.VBComponents(Application.vbe.ActiveVBProject.VBComponents(N).name).CodeModule.codePane.SetSelection lngSLine, lngSCol, lngELine, lngECol
                '280             Debug.Print Application.VBE.ActiveVBProject.VBComponents(N).Name & "    Linia " & lngSLine & " Coloana " & lngSCol
                '290             Exit Function
300         End If
310     End If
320 Next
    '========================== incep tratare erori
TRATARE_ERORI_iesire:

330 Exit Function

TRATARE_ERORI:
    Dim lngLinia As Long, errNumar As Long, errDescriere As String
340 lngLinia = Erl: errNumar = err.Number: errDescriere = err.Description
    Dim RaspunsMesaj As String
350 Select Case errNumar
    Case 0
360 Case 2456
370     Resume Next
380 Case 2457
390     Resume Next
400 Case 2451
410     Resume Next
420 Case Else
        '190              ScrieEroare "Eroare in [functii].[cauta_text] linia " & lngLinia, errNumar & " " & errDescriere, errNumar, False, "Netrarata functii rutina cauta_text"
430     RaspunsMesaj = MsgBox("[Eroare in functii].[cauta_text] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere & vbCrLf & "DORITI SA CONTINUATI EXECUTIA", vbYesNo)
        'Executasiraspunde="INFORMATIE "[Eroare in functii].[cauta_text] linia " & lngLinia & " cu numarul " & errNumar & vbCrLf & errDescriere"
440     If RaspunsMesaj = vbYes Then
450         Resume Next
460     Else
470         GoTo TRATARE_ERORI_iesire
480     End If
490 End Select
    '========================== terminat tratare erori
End Function


Public Function EsteSarbatoareRO(D As Date) As Boolean
    Dim Y As Integer
    Y = year(D)

    Select Case Format(D, "mm-dd")
        Case "01-01", "01-02" ' Anul Nou
        Case "01-24"         ' Unirea Principatelor
        Case "05-01"         ' Ziua Muncii
        Case "08-15"         ' Adormirea Maicii Domnului
        Case "11-30"         ' Sfantul Andrei
        Case "12-01"         ' Ziua Nationala
        Case "12-25", "12-26" ' Craciun
            EsteSarbatoareRO = True
            Exit Function
    End Select

    ' Sarbatori Pascale (calculate dinamic)
    Dim Paste As Date
    Paste = DataPaste(Y)

    If D = Paste _
       Or D = Paste + 1 _
       Or D = Paste + 49 _
       Or D = Paste + 50 Then
        EsteSarbatoareRO = True
        Exit Function
    End If

    EsteSarbatoareRO = False
End Function

'==========================
' Calculeaza data Pastelui Ortodox
'==========================
Public Function DataPaste(Y As Integer) As Date
    Dim a As Integer, B As Integer, C As Integer
    Dim D As Integer, E As Integer

    a = Y Mod 19
    B = Y Mod 4
    C = Y Mod 7
    D = (19 * a + 16) Mod 30
    E = (2 * B + 4 * C + 6 * D) Mod 7

    DataPaste = DateSerial(Y, 4, 3) + D + E
End Function

'==========================
' Functie principala: zile lucratoare intre doua date
'==========================
Public Function ZileLucratoareRO(datastart As Date, datastop As Date) As Integer
    Dim D As Date
    Dim cnt As Integer
    Dim lInainte, lDupa
    For D = datastart To datastop
        lInainte = month(D)
        If lDupa <> "" Then
            If lInainte <> lDupa Then
                Debug.Print Format(D, "m mmm yyyy") & "  " & cnt
                cnt = 0
            End If
        End If
        If Weekday(D, vbMonday) <= 5 Then
            If Not EsteSarbatoareRO(D) Then
                cnt = cnt + 1
            End If
        End If


        lDupa = month(D)
    Next D

    ZileLucratoareRO = cnt
End Function
Function testZileLucratoareRO()
ZileLucratoareRO "30.11.2023", "31.12.2025"
End Function