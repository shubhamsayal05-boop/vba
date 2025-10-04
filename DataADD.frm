VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataADD 
   ClientHeight    =   4890
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7590
   OleObjectBlob   =   "DataADD.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataADD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton6_Click()
    Dim i As Long
    
    If Len(Me.POrdre) > 0 And Len(Me.Pcolonne) > 0 Then
         With ThisWorkbook.sheets("DEFINITION SDV")
                .Range("A2:E3").Copy Destination:=ThisWorkbook.Worksheets("DEFINITION SDV").Cells(ThisWorkbook.Worksheets("DEFINITION SDV").Range("A65000").End(xlUp).row + 1, 1)
                
                i = ThisWorkbook.Worksheets("DEFINITION SDV").Range("A65000").End(xlUp).row
                .Range("A" & i - 1 & ":A" & i) = Me.POrdre
                .Range("B" & i - 1) = Me.Pcolonne
                DataLoad.code.Caption = Me.POrdre & "--" & Me.Pcolonne
                Unload Me
                DataLoad.Show
                
        End With
    End If
End Sub

Private Sub UserForm_Initialize()
   Me.POrdre = _
  ThisWorkbook.Worksheets("DEFINITION SDV").Range("A" & ThisWorkbook.Worksheets("DEFINITION SDV").Range("A65000").End(xlUp).row) + 1
  addSDVList
End Sub

Function addSDVList()
  
    Dim v
    Dim i As Long
    
    v = ThisWorkbook.sheets("structure").UsedRange.Columns(2).Value
    For i = 2 To UBound(v, 1)
        If Len(v(i, 1)) > 0 Then
          Me.Pcolonne.AddItem UCase(v(i, 1))
        End If
    Next i
    Erase v
  
   
End Function
