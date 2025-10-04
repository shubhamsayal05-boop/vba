Attribute VB_Name = "Rating_Distribution"
Option Explicit

Sub Distributions(ByVal onglet As String, Prt As String)
    
    Dim NEVENTS As Integer
    Dim derniereLigne As Long
    
    derniereLigne = TotEventSheet(onglet)
    If Prt = "driv" Then ThisWorkbook.sheets(onglet).Range("G8") = derniereLigne - 6
    If Prt = "dyn" Then ThisWorkbook.sheets(onglet).Range("BN8") = derniereLigne - 6
    NEVENTS = ThisWorkbook.sheets(onglet).Range("G8").Value

    With ThisWorkbook.sheets(onglet)
        If Prt = "driv" Then
        .Range("G11") = Application.WorksheetFunction.CountIf(.Range("Q7:O" & NEVENTS + 6), "GREEN")
        .Range("G14") = Application.WorksheetFunction.CountIf(.Range("Q7:O" & NEVENTS + 6), "YELLOW")
        .Range("G17") = Application.WorksheetFunction.CountIf(.Range("Q7:O" & NEVENTS + 6), "RED")
       ElseIf Prt = "dyn" Then
        .Range("BN11") = Application.WorksheetFunction.CountIf(.Range("BX7:BV" & NEVENTS + 6), "GREEN")
        .Range("BN14") = Application.WorksheetFunction.CountIf(.Range("BX7:BV" & NEVENTS + 6), "YELLOW")
        .Range("BN17") = Application.WorksheetFunction.CountIf(.Range("BX7:BV" & NEVENTS + 6), "RED")
       End If
    End With

End Sub








