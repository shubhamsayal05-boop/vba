Attribute VB_Name = "ColorGlobalByStatus"
Function GlobalDrivability()
    Dim r As Range
    Dim v
    Dim i As Long
    Dim iCol As Integer
    Dim sh As Worksheet
    Dim vCol As String
    Dim pr As Integer
    Dim rt As Integer, ir
    
    
      ir = ThisWorkbook.Worksheets("RATING").Rows("10:10").Find(What:="Tested vehicle", lookat:=xlWhole).Column
     'ThisWorkbook.Sheets("rating").Range("L10").Formula = "=calculs!M39"
    If Not colorGlobalDriv Is Nothing Then
            If colorGlobalDriv.Count <> 0 Then
                    If ThisWorkbook.Worksheets("RATING").Cells(12, ir) < ThisWorkbook.sheets("calculs").Range("seuilvA") Then 'si taux pts bas vert
                        If (ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") / 100) >= ThisWorkbook.sheets("calculs").Range("seuilrB") Then ' et si index jaune ou vert
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "GREEN" 'alors vert
                        ElseIf (ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") / 100) < ThisWorkbook.sheets("calculs").Range("seuilrB") Then 'sinon si index rouge
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "YELLOW" 'alors jaune
                        Else
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "GREEN" 'sinon vert
                        End If
                    ElseIf ThisWorkbook.Worksheets("RATING").Cells(12, ir) > ThisWorkbook.sheets("calculs").Range("seuilrA") Then 'sinon si taux pts bas rouge
                        If (ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") / 100) >= ThisWorkbook.sheets("calculs").Range("seuilvB") Then 'si index vert
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "RED" 'alors rouge
                        ElseIf (ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") / 100) < ThisWorkbook.sheets("calculs").Range("seuilrB") Then 'sinon si index rouge
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "RED" 'alors rouge
                        Else
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "RED" 'sinon rouge
                        End If
                    Else 'sinon, donc si taux jaune
                        If (ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") / 100) >= ThisWorkbook.sheets("calculs").Range("seuilvA") Then 'et si index vert
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "YELLOW" 'alors jaune
                        ElseIf (ThisWorkbook.sheets("RATING").Range("RESULTATGLOBAL1") / 100) < ThisWorkbook.sheets("calculs").Range("seuilrB") Then 'sinon si index rouge
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "RED" 'alors rouge
                        Else
                            ThisWorkbook.sheets("RATING").Range("E11").Value = "YELLOW" 'sinon jaune
                        End If
                    End If
                    
            End If
   End If

End Function
