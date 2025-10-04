VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Private Sub Worksheet_Activate()
    Dim RS As Double
'    RS = GetNoteGlobalTarget("driv")
'    ActiveSheet.Range("EQ3").Value = IIf(RS = -555, "", RS)
'
'    RS = GetNoteGlobalTarget("dyn")
'    ActiveSheet.Range("EQ4").Value = IIf(RS = -555, "", RS)
    
    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    ThisWorkbook.sheets("RATING").Shapes("ZoneTexte 37").TextFrame2.TextRange.Characters.text = (val(ThisWorkbook.sheets("Calculs").Range("seuilvA").Value) * 100) & "%"
    ThisWorkbook.sheets("RATING").Shapes("ZoneTexte 39").TextFrame2.TextRange.Characters.text = (val(ThisWorkbook.sheets("Calculs").Range("seuilrA").Value) * 100) & "%"
    On Error GoTo 0
    Call getTotalColor
End Sub

