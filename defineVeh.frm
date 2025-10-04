VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} defineVeh 
   Caption         =   "Vehicule"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5220
   OleObjectBlob   =   "defineVeh.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "defineVeh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub CommandButton1_Click()
  Dim c As Range
  If ThisWorkbook.sheets("HOME").Range("Fuel").Value = "" Or ThisWorkbook.sheets("HOME").Range("Gears").Value = "" Or ThisWorkbook.sheets("HOME").Range("Software") = "" Or ThisWorkbook.sheets("HOME").Range("Prestation").Value = "" Or ThisWorkbook.sheets("HOME").Range("DriveVersion").Value = "" Or ThisWorkbook.sheets("HOME").Range("Milestone").Value = "" Or ThisWorkbook.sheets("HOME").Range("Area").Value = "" Then
        With ThisWorkbook.Worksheets("CONFIGURATIONS")
                addVehRating (Me.tEMPS)
                addVehGraphStatus (Me.tEMPS)
                addVehTotPoint
                Set c = .Range("VEHICLE")
                Set c = c.Offset(1, 0)
                While c.Value <> ""
                    Set c = c.Offset(1, 0)
                Wend
                If .Range("A" & c.row & ":B" & c.row).MergeCells = False Then
                    .Rows(c.row + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    .Range("A" & c.row & ":B" & c.row).Borders(1).LineStyle = xlContinuous
                    .Range("A" & c.row & ":B" & c.row).Borders(2).LineStyle = xlContinuous
                    .Range("A" & c.row & ":B" & c.row).Borders(3).LineStyle = xlContinuous
                    .Range("A" & c.row & ":B" & c.row).Borders(4).LineStyle = xlContinuous
                    .Range("A" & c.row & ":B" & c.row).Merge
                    .Range("A" & c.row & ":B" & c.row) = Me.tEMPS
                End If
           End With
           
         Unload Me
         MsgBox "Terminé", vbInformation, "ODRIV"
 Else
         Unload Me
         MsgBox "Erase Project", vbCritical, "ODRIV"
 End If

End Sub

Function addVehRating(nameVeh As String)
    Dim r
    Dim lastC
    Dim colD
    With sheets("RATING")
'        Application.ScreenUpdating = False
        .Activate
        colD = .Rows("21:22").Find(What:="Drivability Lowest Events", lookat:=xlWhole).Column
        .Columns(colD).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range(.Cells(21, colD), .Cells(22, colD)).Merge
        .Cells(21, colD) = nameVeh
        
        colD = .Rows("21:22").Find(What:="Dynamism Lowest Events", lookat:=xlWhole).Column
        .Columns(colD).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
        .Range(.Cells(21, colD), .Cells(22, colD)).Merge
        .Cells(21, colD) = nameVeh
        
        lastC = .Cells(10, .Columns.Count).End(xlToLeft).Column
        .Columns(lastC).Copy Destination:=.Cells(1, lastC + 1)
        Application.CutCopyMode = False
        .Cells(10, lastC + 1) = nameVeh
        .Cells(16, lastC + 1) = nameVeh
    End With
    sheets("CONFIGURATIONS").Activate
'    Application.ScreenUpdating = True
End Function

'Function addVehRating(nameVeh As String)
'    Dim r
'    Dim lastC
'    Dim cold
'
'    With sheets("RATING")
'        cold = .Rows("21:22").Find(What:="Drivability Lowest Events", LookAt:=xlWhole).Column
'        .Columns(cold).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'        .Cells(21, cold) = nameVeh
'
'        cold = .Rows("21:22").Find(What:="Dynamism Lowest Events", LookAt:=xlWhole).Column
'        .Columns(cold).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
'        .Cells(22, cold) = nameVeh
'
'        lastC = .Cells(10, .Columns.Count).End(xlToLeft).Column
'        .Columns(lastC).Copy Destination:=.Cells(1, lastC + 1)
'        Application.CutCopyMode = False
'        .Cells(10, lastC + 1) = nameVeh
'        .Cells(16, lastC + 1) = nameVeh
'    End With
'
'End Function

Function addVehGraphStatus(nameVeh As String)
    Dim r
    Dim lastr
    Dim i As Integer
    
    With sheets("Graph_status")
        lastr = .Cells(.Rows.Count, 1).End(xlUp).row
        For i = 1 To lastr
           If .Cells(i, 1) = "index rouge" Then
             .Rows(i - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
             .Cells(i - 1, 1) = nameVeh
             i = i + 1
           End If
        Next i
        Application.CutCopyMode = False
    End With

End Function

Function addVehTotPoint()
    Dim r
    Dim lastr, lastC
    Dim i As Integer
    
    With sheets("totalPoint")
        .Columns("S:DK").EntireColumn.Delete
        lastr = sheets("RATING").Cells(.Rows.Count, 4).End(xlUp).row
        lastC = sheets("RATING").Cells(lastr, sheets("RATING").Columns.Count).End(xlToLeft).Column
        sheets("RATING").Range(sheets("RATING").Cells(lastr, 2), sheets("RATING").Cells(lastr, lastC)).Copy Destination:=.Range("S1")
        Application.CutCopyMode = False
       
    End With

End Function

