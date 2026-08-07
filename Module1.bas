Option Compare Database
Option Explicit
Public blnINRETEA As Boolean
Type str_DEVNAMES
    PGB As String * 94
End Type

Type type_DEVNAMES
   strDriverOffset  As String
   strDeviceOffset   As String
   strOutputOffset    As String
   intDefault  As Integer
End Type




Type str_DEVMODE
    RGB As String * 94
End Type

Type type_DEVMODE
    strDeviceName As String * 16
    intSpecVersion As Integer
    intDriverVersion As Integer
    intSize As Integer
    intDriverExtra As Integer
    lngFields As Long
    intOrientation As Integer
    intPaperSize As Integer
    intPaperLength As Integer
    intPaperWidth As Integer
    intScale As Integer
    intCopies As Integer
    intDefaultSource As Integer
    intPrintQuality As Integer
    intColor As Integer
    intDuplex As Integer
    intResolution As Integer
    intTTOption As Integer
    intCollate As Integer
    strFormName As String * 16
    lngPad As Long
    lngBits As Long
    lngPW As Long
    lngPH As Long
    lngDFI As Long
    lngDFr As Long
End Type

Sub rap()
'Dim rap As Report
'Set rap = Reports("Lucrari de reparatii")
'CheckCustomPage ("Lucrari de reparatii")
End Sub

'Sub CheckCustomPage(rptName As String)
'    Dim DevString As str_DEVMODE
'        Dim Devnam As str_DEVNAMES
'
'    Dim DM As type_DEVMODE
'    Dim imprimanta As type_DEVNAMES
'
'    Dim strDevModeExtra As String
'    Dim strDevnameExtra As String
'    Dim rpt As Report
'    Dim intResponse As Integer
'    ' Opens report in Design view.
'    DoCmd.OpenReport rptName, acDesign
'
'Set rpt = Reports(rptName)
'    If Not IsNull(rpt.PrtDevMode) Then
'        strDevModeExtra = rpt.PrtDevMode   ' Gets current DEVMODE structure.
'        strDevnameExtra = rpt.PrtDevNames
'
'
'        DevString.RGB = strDevModeExtra
'        Devnam.PGB = strDevnameExtra
'
'        LSet DM = DevString
'        'LSet imprimanta = Devnam
'        imprimanta.strDeviceOffset = "PDFCreator"
'        imprimanta.intDefault = 0
'        imprimanta.strDriverOffset = "PDFCreator"
'        imprimanta.strOutputOffset = "PDFCreator"
'
'
'       DM.strDeviceName = "PDFCreator"
'       DM.intCopies = 2
'
'
'        LSet Devnam = imprimanta
'            LSet DevString = DM         ' Update property.
'            Mid(strDevnameExtra, 1, 94) = Devnam.PGB
'
'            Mid(strDevModeExtra, 1, 94) = DevString.RGB
'            rpt.PrtDevMode = strDevModeExtra
'            rpt.PrtDevNames = strDevnameExtra
'       End If
'
'
'
'        'If DM.intPaperSize = 256 Then
'        '    ' Display user-defined size.
'        '    intResponse = MsgBox("The current custom page size is " _
'         '       & DM.intPaperWidth / 254 & " inches wide by " _
'         '       & DM.intPaperLength / 254 & " inches long. Do you want " _
''& "to change the settings?", 4)
''        Else
''            ' Currently not user-defined.
''            intResponse = MsgBox("The report does not have a custom page size. " _
''                & "Do you want to define one?", 4)
''        End If
''        If intResponse = 6 Then
''            ' User wants to change settings.
''            ' Initialize fields.
''            DM.lngFields = DM.lngFields Or DM.intPaperSize Or DM.intPaperLength _
''                Or DM.intPaperWidth
''            DM.intPaperSize = 256       ' Set custom page.
''            ' Prompt for length and width.
'
''DM.intPaperLength = InputBox("Please enter page length " _
''                & "in inches.") * 254
''            DM.intPaperWidth = InputBox("Please enter page width " _
''                & "in inches.") * 254
''            LSet DevString = DM         ' Update property.
''            Mid(strDevModeExtra, 1, 94) = DevString.RGB
''            rpt.PrtDevMode = strDevModeExtra
''        End If
''    End If
'End Sub
Function PE(CT As label)
CT.fontunderline = True
CT.ForeColor = 8388863
End Function
Function LINGA(CT As label)
CT.fontunderline = False
CT.ForeColor = 0
End Function

Function aa()
  Dim wrkJet As Workspace
    Dim dbsPubs As DAO.Database
    Dim dbsLoop As DAO.Database
    Dim prpLoop As Property

    Set wrkJet = CreateWorkspace("", "admin", "", dbUseJet)
    Set dbsPubs = wrkJet.OpenDatabase("comenzi", dbDriverNoPrompt, True, "ODBC;DSN=SQL11STOCURI;UID=sa;PWD=comenzi2012#;LANGUAGE=us_english;DATABASE=SQL_STOC")

    For Each dbsLoop In wrkJet.Databases
'        Debug.Print "Database properties for " & _
            dbsLoop.Name & ":"

        On Error Resume Next
        ' Enumerate the Properties collection of each Database
        ' object.
        For Each prpLoop In dbsLoop.Properties
            If prpLoop.name = "Connection" Then
                ' Property actually returns a Connection object.

'Debug.Print "    Connection[.Name] = " & _
                    dbsLoop.Connection.Name
            Else
'                Debug.Print "    " & prpLoop.Name & " = " & _
                    prpLoop
            End If
        Next prpLoop
        On Error GoTo 0

    Next dbsLoop

     dbsPubs.Close
 
    wrkJet.Close

      
        
        
End Function


'===Start Generat AI===
Function TestAntigravity() As String
    TestAntigravity = "Antigravity Active"
End Function
'===Final Generat AI===