VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} EditPowertrain 
   Caption         =   "Settings"
   ClientHeight    =   3465
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   OleObjectBlob   =   "EditPowertrain.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "EditPowertrain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CommandButton4_Click()
            If Len(Me.ComboBox4) = 0 Then
                    MsgBox "Choisir SDV", vbCritical, "ODRIV"
            Else
                     AddPowertrain.TextBox22 = Me.ComboBox4
                     AddPowertrain.TextBox22.Locked = True
                     AddPowertrain.editing = "OK"
                     RemplissageDonnee
                     Unload Me
                     AddPowertrain.Show
            End If
End Sub

Private Sub UserForm_Initialize()
        Call chargeVal
End Sub

Function chargeVal()
    Dim v
    Dim i As Integer
    v = ThisWorkbook.sheets("POWERTRAIN").UsedRange.Value
    For i = 3 To UBound(v, 1)
        If v(i, 1) = "Titre config" Then
                Me.ComboBox4.AddItem v(i, 2)
        End If
    Next i
    
    Erase v
End Function

Function RemplissageDonnee()
        Dim r As Long
        Dim i As Integer
        Dim j As Integer
       
        With ThisWorkbook.sheets("POWERTRAIN")

             r = FTNum
            For i = 1 To 7 Step 2
                       For j = 2 To .Cells(r + i, .Columns.Count).End(xlToLeft).Column
                           CompareV (.Cells(r + i, j))
                     Next j
            Next i
            
        End With
End Function

Function CompareV(r As Range)
        Dim i As Integer

        With ThisWorkbook.sheets("POWERTRAIN")
       
              If .Cells(r.row, 1) = "Engine type" Then
                       For i = 0 To AddPowertrain.ListView15.ListCount - 1
                            If r.Offset(1, 0) = "X" And UCase(AddPowertrain.ListView15.list(i)) = UCase(r.Value) Then
                                 AddPowertrain.ListView15.Selected(i) = True
                            ElseIf r.Offset(1, 0) = "" And UCase(AddPowertrain.ListView15.list(i)) = UCase(r.Value) Then
                                AddPowertrain.ListView15.Selected(i) = False
                            End If
                       Next i
                       
            ElseIf .Cells(r.row, 1) = "Gearbox type" Then
                       For i = 0 To AddPowertrain.ListView16.ListCount - 1
                            If r.Offset(1, 0) = "X" And UCase(AddPowertrain.ListView16.list(i)) = UCase(r.Value) Then
                                 AddPowertrain.ListView16.Selected(i) = True
                            ElseIf r.Offset(1, 0) = "" And UCase(AddPowertrain.ListView16.list(i)) = UCase(r.Value) Then
                                AddPowertrain.ListView16.Selected(i) = False
                            End If
                       Next i
                       
                       
             ElseIf .Cells(r.row, 1) = "Number of gears" Then
                       For i = 0 To AddPowertrain.ListView17.ListCount - 1
                            If r.Offset(1, 0) = "X" And UCase(AddPowertrain.ListView17.list(i)) = UCase(r.Value) Then
                                 AddPowertrain.ListView17.Selected(i) = True
                            ElseIf r.Offset(1, 0) = "" And UCase(AddPowertrain.ListView17.list(i)) = UCase(r.Value) Then
                                AddPowertrain.ListView17.Selected(i) = False
                            End If
                       Next i
                       
                       
            ElseIf .Cells(r.row, 1) = "Area" Then
                       For i = 0 To AddPowertrain.ListView18.ListCount - 1
                            If r.Offset(1, 0) = "X" And UCase(AddPowertrain.ListView18.list(i)) = UCase(r.Value) Then
                                 AddPowertrain.ListView18.Selected(i) = True
                            ElseIf r.Offset(1, 0) = "" And UCase(AddPowertrain.ListView18.list(i)) = UCase(r.Value) Then
                                AddPowertrain.ListView18.Selected(i) = False
                            End If
                       Next i
                       
                       
            End If
        End With
End Function

Function FTNum() As Long
    Dim v
    Dim i As Integer
    
    FTNum = 0
  
    v = ThisWorkbook.sheets("POWERTRAIN").UsedRange.Value
    For i = 3 To UBound(v, 1)
        If v(i, 1) = "Titre config" And UCase(CStr(v(i, 2))) = UCase(Me.ComboBox4) Then
                FTNum = i
                Exit Function
        End If
    Next i
    
    Erase v
End Function
