VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
' Cellule Office

Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.UseSystemSeparators = True
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
'   Cancel = Not CBool(Application.UserName Like "*E592622*")
End Sub
Sub wProtect()
    If ThisWorkbook.ReadOnly = False Then
        If Not CBool(Application.UserName Like "*E592622*") = True Then
            ThisWorkbook.Close False
            Exit Sub
        End If
     End If
    If Not CBool(Application.UserName Like "*E592622*") = True Then
        Application.ShowDevTools = False
        Application.OnKey "%{F11}", " "
        On Error Resume Next
        ThisWorkbook.VBProject.VBE.MainWindow.Visible = False
        On Error GoTo 0
    End If
    
End Sub

Private Sub Workbook_Open()
'    If Not CBool(Application.UserName Like "*E592622*") = True Then ActiveWorkbook.ChangeFileAccess Mode:=xlReadOnly
'    wProtect
'    Message_du_jour_wrk False

    Call FiltresOff
    Call HideFeuille(2)
    Call DBStructure.getEntete
    Call DBStructure.getInterval
    'ThisWorkbook.Sheets("DNT").Range("A9").Value = ThisWorkbook.Sheets("DATA").Range("A65000").End(xlUp).Row
    ThisWorkbook.sheets("HOME").Range("Version").Value = "VERSION" & " " & Left(Mid(ThisWorkbook.Name, 7), Len(Mid(ThisWorkbook.Name, 7)) - 5)
    ThisWorkbook.sheets("HOME").Activate

   ' forcer separateur decimal
    Application.DecimalSeparator = "."
    Application.UseSystemSeparators = False
    Application.ThousandsSeparator = " "
    AutoBackupWorkbook
End Sub
Private Sub Workbook_SheetActivate(ByVal sh As Object)
    HideFeuille (-1)
End Sub

Function HideFeuille(options As Integer)
   Dim tables(13) As String
   Dim i As Integer
    
    tables(1) = "CONFIGURATIONS SEETINGS"
    tables(2) = "SETTINGS"
    tables(3) = "structure"
    tables(4) = "POWERTRAIN"
    tables(5) = "CONFIGURATIONS"
    tables(6) = "Target vehicle"
    tables(7) = "TARGETS"
    tables(8) = "Calculs"
    tables(9) = "Graph_status"
    tables(10) = "DEFINITION SDV"
    tables(11) = "PARAMETRES GRAPH"
    tables(12) = "ENTETE_COLONNE"
    tables(13) = "sdv MANAGER"
    
    For i = 1 To UBound(tables)
             If sheets(tables(i)).Visible <> options Then
                   sheets(tables(i)).Visible = 2
             End If
    Next i
'
   
End Function

