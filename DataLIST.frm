VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataLIST 
   Caption         =   "R�cup�ration des SDV"
   ClientHeight    =   6900
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8295.001
   OleObjectBlob   =   "DataLIST.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub CommandButton1_Click()
    If SelA = 999 Then
        MsgBox "SELECTION VIDE", vbCritical, "ODRIV"
        Exit Sub
    Else
        DataLoad.Show
    End If
 
End Sub

Private Sub CommandButton2_Click()
     Dim j As Long
      Dim v
      Dim t
      Dim i As Integer
      Dim r As Range
      
      t = SelA
     If t <> 999 Then
            v = ThisWorkbook.sheets("DEFINITION SDV").UsedRange.Columns("A:E").Value
            With ThisWorkbook.Worksheets("DEFINITION SDV")
'                     .Outline.ShowLevels RowLevels:=2
                      For j = 1 To UBound(v, 1)
                            Set r = .Range("A" & j)
                            If (v(j, 1)) & "--" & (v(j, 2)) = Me.ListeValeur.list(t) Then
                               Set r = .Range("A" & j)
                               i = j + 1
                               While .Range("A" & i) = .Range("A" & j)
                                 Set r = Union(r, .Range("A" & i))
                                 i = i + 1
                               Wend
                               r.EntireRow.Delete
                               Unload Me
                               Exit Sub
                           End If
                      Next j
        
             End With
     Else
            MsgBox "SELECTION VIDE", vbCritical, "ODRIV"
     End If

    
End Sub

Private Sub ListeValeur_Click()
    If SelA <> 999 Then
        Me.code.Caption = Me.ListeValeur.list(SelA)
    End If
End Sub

Private Sub UserForm_Initialize()
      Dim j As Long
      Dim v
      
       v = ThisWorkbook.sheets("DEFINITION SDV").UsedRange.Columns("A:E").Value
      With ThisWorkbook.Worksheets("DEFINITION SDV")
'             .Outline.ShowLevels RowLevels:=2
              For j = 1 To UBound(v, 1)
                    If IsNumeric(v(j, 1)) = True And Len(v(j, 1)) > 0 And Len(v(j, 3)) = 0 Then
                       Me.ListeValeur.AddItem v(j, 1) & "--" & v(j, 2)
                   End If
              Next j

     End With

End Sub



Function SelA() As Integer
Dim i As Long
 SelA = 999
For i = 0 To Me.ListeValeur.ListCount - 1
        If Me.ListeValeur.Selected(i) = True Then
            SelA = i
            Exit Function
        End If
Next i
End Function


