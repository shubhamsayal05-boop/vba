Attribute VB_Name = "SuppressionSDV"
Option Explicit

Sub positionOrder()
    OrdreS.Show
End Sub
Sub delAll()
        On Error GoTo Ers
        Application.DisplayAlerts = False
        EventAndScreen (False)
    
        If Len(ActiveCell.Value) = 0 Or ActiveCell.row = 1 Or ActiveCell.Column <> 1 Then
            MsgBox "Selectionner SDV", vbCritical, "ODRIV"
        ElseIf (ActiveCell.Interior.color) = 11851260 Then
            MsgBox "Selectionner SDV", vbCritical, "ODRIV"
        ElseIf sheetExists(ActiveCell.Value) = True Then
            MsgBox "SDV chargée Erase All Data", vbCritical, "ODRIV"
        ElseIf MsgBox("Voulez Vous Supprimer '" & ActiveCell.Value & "'", vbCritical + vbYesNo, "ODRIV") = vbYes Then
             Call DelConfigurationSetting(ActiveCell.Value)
             Call DelCalculs(ActiveCell.Value)
             Call DelStructure(ActiveCell.Value)
             Call DelRating(ActiveCell.Value)
             Call delSettings(ActiveCell.Value)
             Call delTargets(ActiveCell.Value)
             Call delTargetVehicle(ActiveCell.Value)
             Call delDefinitionSdv(ActiveCell.Value)
             Call delGraph(ActiveCell.Value)
             Call delPowertrain(ActiveCell.Value)
             ActiveCell.EntireRow.Delete
             MsgBox "Opération Réussie", vbInformation, "ODRIV"
        End If
        EventAndScreen (True)
        Application.DisplayAlerts = True

Ers:
        If ERR.Number <> 0 Then
                MsgBox ERR.description, vbCritical, "ODRIV"
                Application.DisplayAlerts = True
                EventAndScreen (True)
        End If
End Sub
Function DelConfigurationSetting(sdv As String) As Long
        Dim r As Long
        Dim i As Integer
        Dim n As Integer
        Dim IsConfig As Boolean
        IsConfig = False
        
        ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=2
        With ThisWorkbook.sheets("CONFIGURATIONS SEETINGS")
                r = .Range("A65000").End(xlUp).row
                For i = 3 To r
                            If UCase(.Cells(i, 1)) = UCase(sdv) Then
                                    If Len(.Cells(i + 1, 1)) = 0 And i + 1 <= .Range("B65000").End(xlUp).row Then n = i + 1 Else n = i
                                    While Len(.Cells(n, 1)) = 0 And n <= .Range("B65000").End(xlUp).row
                                             n = n + 1
                                             IsConfig = True
                                    Wend
                                    
                                    If IsConfig = True Then .Rows(i & ":" & n - 1).Delete Shift:=xlUp Else .Rows(i & ":" & n).Delete Shift:=xlUp
                            End If
               Next i
           ThisWorkbook.sheets("CONFIGURATIONS SEETINGS").Outline.ShowLevels RowLevels:=1
     End With
End Function

Function DelCalculs(sdv As String) As Long
    Dim i As Long
    i = rowStartSDV("Calculs", sdv, 2)
    If i <> 0 Then
        With ThisWorkbook.sheets("Calculs")
                .Rows(i).Delete Shift:=xlUp
        End With
    End If
End Function
Function delPowertrain(sdv As String) As Long
    Dim i As Long
    
    i = rowStartSDV("POWERTRAIN", sdv, 1)
    While i <> 0
        With ThisWorkbook.sheets("POWERTRAIN")
                .Rows(i).Delete Shift:=xlUp
        End With
        i = rowStartSDV("POWERTRAIN", sdv, 1)
    Wend
    
   
End Function
Function delSettings(sdv As String)
    Dim i As Long
    i = rowStartSDV("SETTINGS", sdv, 1)
    If i <> 0 Then
        With ThisWorkbook.sheets("SETTINGS")
'                .Rows(i).Delete Shift:=xlUp
                .Rows(i & ":" & i + 14).EntireRow.Delete
        End With
    End If
End Function
Function DelStructure(sdv As String) As Long
    Dim st As String
    Dim n As Long, o As Long
    
    st = getNumberRow(sdv)
    If st = "" Then Exit Function
    o = val(Split(st, ";")(1))
    n = val(Split(st, ";")(0))
    With ThisWorkbook.sheets("Structure")
         .Rows(n & ":" & o).Delete Shift:=xlUp
    End With
   
End Function
Function DelRating(sdv As String) As Long
    Dim i As Long
    i = rowStartSDV("RATING", sdv, 4)
    If i <> 0 Then
        With ThisWorkbook.sheets("RATING")
                .Rows(i).Delete Shift:=xlUp
        End With
    End If
End Function
'Function delSetting(SDV As String) As Long
'    Exit Function
''    Dim i As Long, j As Long
''    Dim r As Range
''
''    i = rowStartSDV("SETTINGS", SDV, 1)
''    If i <> 0 Then
''        j = i
''        With ThisWorkbook.Sheets("SETTINGS")
''                Set r = .Range("B" & i)
''                While Application.CountA(.Rows(r.Row).EntireRow) > 0 Or r.Interior.Color = 14277081
''                    Set r = r.offset(1, 0)
''                Wend
''                .Rows(i).Delete Shift:=xlUp
''        End With
''    End If
'
'End Function

Function delTargets(sdv As String)
    Dim derLigne As Long
    
    
    With ThisWorkbook.Worksheets("TARGETS")
            .AutoFilterMode = False
            derLigne = .Range("A65000").End(xlUp).row
            .Rows(1).AutoFilter Field:=1, Criteria1:=sdv
            On Error Resume Next
            .Range("A2:A" & derLigne).SpecialCells(xlCellTypeVisible).Delete
            .AutoFilterMode = False
            If ERR.Number <> 0 Then ERR.Clear
    End With

End Function

Function delTargetVehicle(sdv As String)
    Dim derLigne As Long
    
    
    With ThisWorkbook.Worksheets("TARGET VEHICLE")
            .AutoFilterMode = False
            derLigne = .Range("A65000").End(xlUp).row
            .Rows(1).AutoFilter Field:=1, Criteria1:=sdv
            On Error Resume Next
            .Range("A2:A" & derLigne).SpecialCells(xlCellTypeVisible).Delete
            .AutoFilterMode = False
            If ERR.Number <> 0 Then ERR.Clear
    End With

End Function
Function delDefinitionSdv(sdv As String)
      Dim j As Long
      Dim v
      Dim i As Integer
      Dim r As Range, delSup As Range

       v = ThisWorkbook.sheets("DEFINITION SDV").UsedRange.Columns("A:E").Value
       With ThisWorkbook.Worksheets("DEFINITION SDV")
                  For j = 1 To UBound(v, 1)
                        If UCase(v(j, 2)) = UCase(sdv) Then
                           Set r = .Range("A" & j)
                           i = j + 1
                           While .Range("A" & i) = .Range("A" & j)
                             Set r = Union(r, .Range("A" & i))
                             i = i + 1
                           Wend
                          If delSup Is Nothing Then Set delSup = r Else Set delSup = Union(delSup, r)
                          Set r = Nothing
                       End If
                  Next j
                  If Not delSup Is Nothing Then delSup.EntireRow.Delete
       End With
End Function
Function delGraph(sdv As String)
    Dim i As Integer, j As Long
    Dim p As Integer
    Dim v
    
    With ThisWorkbook.Worksheets("PARAMETRES GRAPH")
            v = .UsedRange.Value
            j = rowStartSDV("PARAMETRES GRAPH", sdv, 1)
            i = j + 1
            If i = 0 Then Exit Function
            p = i
            While i <= UBound(v, 1) And Len(v(p, 1)) = 0
                If i <= UBound(v, 1) Then p = i
                i = i + 1
            Wend
            If p < UBound(v, 1) Then p = p - 1
            If j <> 0 And p <> 0 Then .Rows(j & ":" & p).EntireRow.Delete
    End With
    
End Function
Function rowStartSDV(onglet As String, sdv As String, col As Integer) As Long
        Dim i As Long
        Dim v As Variant
        
        rowStartSDV = 0
        With ThisWorkbook.sheets(onglet)
            v = .UsedRange.Value
            v = .Range(.Cells(1, 1), .Cells(UBound(v, 1), UBound(v, 2))).Value
        End With
        For i = 1 To UBound(v, 1)
            If UCase(CStr(v(i, col))) = UCase(sdv) Then
                   rowStartSDV = i
                   Exit Function
            End If
       Next i
       
End Function







