VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} defaultMode 
   Caption         =   "Mode par defaut"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4755
   OleObjectBlob   =   "defaultMode.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "defaultMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function getMode() As String
   Dim GearBx As String
   Dim r As Range
   getMode = ""
   GearBx = ThisWorkbook.sheets("HOME").Range("Gears")
    With ThisWorkbook.sheets("CONFIGURATIONS")
           Set r = .Cells(.Range("GEARBOX").row + 1, .Range("GEARBOX").Column)
           While Len(r.Value) > 0
                If UCase(r.Value) = UCase(GearBx) Then
                    getMode = r.Offset(0, 3)
                End If
                Set r = r.Offset(1, 0)
           Wend
    End With
End Function



Private Sub CommandButton1_Click()
  If Len(Me.ComboBox1) > 0 Then
     Call Load_Data.MKdefMode(Me.ComboBox1)
     Unload Me
   Else
      MsgBox "Entrer une valeur correcte", vbCritical, "ODRIV"
   End If
End Sub

Private Sub UserForm_Initialize()
 Dim StrS As String
 Dim v() As String
 Dim i As Integer
 
 StrS = getMode
 If StrS = "" Then
        Me.ComboBox1.AddItem "AUCUN MODE"
 Else
        If InStr(1, StrS, "-") <> 0 Then
            v = Split(StrS, "-")
            For i = 0 To UBound(v)
                Me.ComboBox1.AddItem UCase(v(i))
            Next i
        Else
            Me.ComboBox1.AddItem UCase(StrS)
        End If
 End If
 
End Sub
