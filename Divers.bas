Attribute VB_Name = "Divers"
Option Explicit
Sub GoHome()
    ThisWorkbook.sheets("home").Activate
    ThisWorkbook.sheets("DocVersions").Visible = False
End Sub

Sub Moniteur(ByVal votre_message As String)
    ThisWorkbook.Worksheets("HOME").Range("Moniteur").Interior.color = RGB(255, 255, 255)
    ThisWorkbook.Worksheets("HOME").Range("Moniteur") = Format(Now, "dd/mm/yyyy hh:mm") & "  -  " & votre_message & vbCrLf & "------------------------------" & vbCrLf & ThisWorkbook.Worksheets("HOME").Range("Moniteur").Value
End Sub



