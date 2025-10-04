VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} delPowertrain 
   Caption         =   "Settings"
   ClientHeight    =   3810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4980
   OleObjectBlob   =   "delPowertrain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "delPowertrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function chargeVal()
    Dim v
    Dim i As Integer
    v = ThisWorkbook.sheets("POWERTRAIN").UsedRange.Value
    For i = 3 To UBound(v, 1)
        If v(i, 1) = "Titre config" Then
                Me.ComboBox3.AddItem v(i, 2)
        End If
    Next i
    
    Erase v
End Function
Private Sub CommandButton3_Click()
        If Len(Me.ComboBox3.Value) > 0 Then
            Call Dels
            MsgBox "Suppression Réussie", vbInformation, "ODRIV"
            Unload Me
        Else
            MsgBox "REMPLIR ", vbCritical, "ODRIV"
        End If
End Sub

Private Sub UserForm_Initialize()
        Call chargeVal
End Sub

Function Dels()
    Dim v
    Dim i As Integer
    Dim found As Long
    
    found = 0
    Application.EnableEvents = False
    v = ThisWorkbook.sheets("POWERTRAIN").UsedRange.Value
    For i = 3 To UBound(v, 1)
        If v(i, 1) = "Titre config" And UCase(CStr(v(i, 2))) = UCase(Me.ComboBox3) Then found = i
        If UCase(v(i, 1)) = "SOMME" And found <> 0 Then
                ThisWorkbook.sheets("POWERTRAIN").Rows(found & ":" & i).EntireRow.Delete
'                 ThisWorkbook.Sheets("POWERTRAIN").Rows(found & ":" & i).Select
                 Application.EnableEvents = True
                Exit Function
        End If
    Next i
    Application.EnableEvents = True
    Erase v
End Function
