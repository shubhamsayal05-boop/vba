Attribute VB_Name = "UpdateGraph"
Option Explicit

Function updateGraphSdv(sdv As String)
    Dim getParamSdv As String
    getParamSdv = checkGraphEnable(sdv)
    deleteAll (sdv)
    If getParamSdv <> "" Then
         Call placeGraphe(getParamSdv, sdv)
         Call populateGraph(getParamSdv, sdv)
    End If
End Function

Function checkGraphEnable(sdv As String) As String
        Dim i As Integer
        Dim v As Variant
        
        checkGraphEnable = ""
        v = ThisWorkbook.Worksheets("PARAMETRES GRAPH").UsedRange.Value
        For i = 1 To UBound(v, 1)
            If UCase(v(i, 1)) = UCase(sdv) Then
                       If v(i + 1, 4) = "Activ�" Then
                            If checkGraphEnable = "" Then checkGraphEnable = "Graphique_0" Else checkGraphEnable = checkGraphEnable & ";" & "Graphique_0"
                            checkGraphEnable = checkGraphEnable & ":" & positionCell(v(i + 3, 4), 1)
                            checkGraphEnable = checkGraphEnable & ":" & i
                       End If
                       
                       If v(i + 8, 4) = "Activ�" Then
                            If checkGraphEnable = "" Then checkGraphEnable = "Graphique_1" Else checkGraphEnable = checkGraphEnable & ";" & "Graphique_1"
                           checkGraphEnable = checkGraphEnable & ":" & positionCell(v(i + 10, 4), 2)
                           checkGraphEnable = checkGraphEnable & ":" & i
                       End If
                       
                       If v(i + 15, 4) = "Activ�" Then
                            If checkGraphEnable = "" Then checkGraphEnable = "Cold_Hot" Else checkGraphEnable = checkGraphEnable & ";" & "Cold_Hot"
                            checkGraphEnable = checkGraphEnable & ":" & positionCell(v(i + 17, 2), 3)
                            checkGraphEnable = checkGraphEnable & ":" & i
                       End If
                       
                       If v(i + 20, 4) = "Activ�" Then
                            If checkGraphEnable = "" Then checkGraphEnable = "Frein_FSE" Else checkGraphEnable = checkGraphEnable & ";" & "Frein_FSE"
                           checkGraphEnable = checkGraphEnable & ":" & positionCell(v(i + 22, 2), 4)
                           checkGraphEnable = checkGraphEnable & ":" & i
                       End If
                       
                        If v(i + 25, 4) = "Activ�" Then
                            If checkGraphEnable = "" Then checkGraphEnable = "Decel_Frein" Else checkGraphEnable = checkGraphEnable & ";" & "Decel_Frein"
                            checkGraphEnable = checkGraphEnable & ":" & positionCell(v(i + 27, 2), 5)
                            checkGraphEnable = checkGraphEnable & ":" & i
                       End If
                       Exit Function
            End If
        Next i
        
End Function
Function placeGraphe(graph As String, sdv As String)
    Dim i As Long
    Dim v() As String
 
    v = Split(graph, ";")
    For i = 0 To UBound(v)
          If Split(v(i), ":")(0) = "Graphique_0" Or Split(v(i), ":")(0) = "Graphique_1" Then
                If Split(v(i), ":")(0) = "Graphique_0" Then
                   ThisWorkbook.Worksheets(sdv).Shapes(Split(v(i), ":")(0)).Visible = True
                   ThisWorkbook.Worksheets(sdv).Shapes("Keys").Visible = True
                ElseIf Split(v(i), ":")(0) = "Graphique_1" Then
                    ThisWorkbook.Worksheets(sdv).Shapes(Split(v(i), ":")(0)).Visible = True
                End If
                Call nameGraph(CStr(Split(v(i), ":")(0)), sdv, val(Split(v(i), ":")(2)))
          Else
                ThisWorkbook.Worksheets("GRAPHIQUES").ListObjects(Split(v(i), ":")(0)).Range.Copy Destination:=ThisWorkbook.Worksheets(sdv).Range(Split(v(i), ":")(1))
          End If
    Next i
    
  
End Function
Function nameGraph(graph As String, sdv As String, ligne As Integer)
    With ThisWorkbook.Worksheets("PARAMETRES GRAPH")
        If graph = "Graphique_0" Then
            ThisWorkbook.Worksheets(sdv).ChartObjects(graph).Chart.ChartTitle.text = _
                "Population : " & .Cells(ligne + 3, 3) & " / " & .Cells(ligne + 3, 2)
            ThisWorkbook.Worksheets(sdv).ChartObjects(graph).Chart.Axes(xlValue, xlPrimary).AxisTitle.text = _
                .Cells(ligne + 3, 3)
            ThisWorkbook.Worksheets(sdv).ChartObjects(graph).Chart.Axes(xlCategory).AxisTitle.text = _
                .Cells(ligne + 3, 2)
        ElseIf graph = "Graphique_1" Then
            ThisWorkbook.Worksheets(sdv).ChartObjects(graph).Chart.ChartTitle.text = _
                "Population : " & .Cells(ligne + 10, 3) & " / " & .Cells(ligne + 10, 2)
            ThisWorkbook.Worksheets(sdv).ChartObjects(graph).Chart.Axes(xlValue, xlPrimary).AxisTitle.text = _
                .Cells(ligne + 10, 3)
            ThisWorkbook.Worksheets(sdv).ChartObjects(graph).Chart.Axes(xlCategory).AxisTitle.text = _
                .Cells(ligne + 10, 2)
            
        End If
    End With
        
End Function
Function positionCell(c As Variant, GraphP As Integer)
   On Error Resume Next
   Dim cr As String
   cr = ThisWorkbook.Worksheets("GRAPHIQUES").Range(c).Value
    If ERR.Number <> 0 Then
            ERR.Clear
            If GraphP = 1 Then positionCell = "B24"
            If GraphP = 2 Then positionCell = "B70"
            If GraphP = 3 Then positionCell = "B24"
            If GraphP = 4 Then positionCell = "B24"
            If GraphP = 5 Then positionCell = "B24"
            
    Else
            positionCell = c
    End If
    
    
End Function

Function deleteAll(sdv As String)
   Dim s As shape
   Dim o As ListObject
  
   
   For Each s In ThisWorkbook.Worksheets(sdv).Shapes
      If s.Name = "Graphique_0" Or s.Name = "Graphique_1" Then
        s.Visible = False
        ThisWorkbook.Worksheets(sdv).Shapes("Keys").Visible = False
      End If
   Next s
       
   For Each o In ThisWorkbook.Worksheets(sdv).ListObjects
     o.Delete
  Next o
   
End Function


Function populateGraph(graph As String, sdv As String)
    Dim i As Integer
    Dim v() As String
    Dim x As Range, y As Range
    Dim totX As Long, totY As Long
    Dim c As Integer
  
   
    v = Split(graph, ";")
    For i = 0 To UBound(v)
          If Split(v(i), ":")(0) = "Graphique_0" Or Split(v(i), ":")(0) = "Graphique_1" Then
                If Split(v(i), ":")(0) = "Graphique_0" Then c = 3 Else c = 10
                With ThisWorkbook.sheets(sdv)
                        If Not .Rows(6).Cells.Find(What:=ThisWorkbook.sheets("PARAMETRES GRAPH").Cells(val(Split(v(i), ":")(2)) + c, 2).Value, lookat:=xlWhole) Is Nothing And _
                            Not .Rows(6).Cells.Find(What:=ThisWorkbook.sheets("PARAMETRES GRAPH").Cells(val(Split(v(i), ":")(2)) + c, 3).Value, lookat:=xlWhole) Is Nothing Then
                            Set x = .Rows(6).Cells.Find(What:=ThisWorkbook.sheets("PARAMETRES GRAPH").Cells(val(Split(v(i), ":")(2)) + c, 2).Value, lookat:=xlWhole)
                            Set y = .Rows(6).Cells.Find(What:=ThisWorkbook.sheets("PARAMETRES GRAPH").Cells(val(Split(v(i), ":")(2)) + c, 3).Value, lookat:=xlWhole)
                            totX = .Cells(.Rows.Count, x.Column).End(xlUp).row
                            totY = .Cells(.Rows.Count, y.Column).End(xlUp).row
                            If totX < 8 Then totX = 8
                            If totY < 8 Then totY = 8
                            If totY > totX Then totX = totY
                            .Shapes(Split(v(i), ":")(0)).Chart.FullSeriesCollection(1).Values = "='" & .Name & "'!" & .Cells(7, y.Column).Address & ":" & .Cells(totX, y.Column).Address
                            .Shapes(Split(v(i), ":")(0)).Chart.FullSeriesCollection(1).XValues = "='" & .Name & "'!" & .Cells(7, x.Column).Address & ":" & .Cells(totX, x.Column).Address
                            Call ChartParam(sdv, val(Split(v(i), ":")(2)), CStr(Split(v(i), ":")(0)))
                            Call Visu_Graphe(sdv, CStr(Split(v(i), ":")(0)))
                        End If
                End With
          Else
                With ThisWorkbook.Worksheets(sdv).ListObjects(checkObject(CStr(Split(v(i), ":")(0)), sdv))
                        Call Visu_Tableau(sdv, CStr(Split(v(i), ":")(0)), .Range.Cells(1, 1).row + 1, .Range.Cells(1, 1).Column + 1, val(Split(v(i), ":")(2)))
                End With
          End If
    Next i
    
End Function

Function valueAxe(valAxe As String)
        If UCase(valAxe) = "AUTOMATIQUE" Or UCase(valAxe) = "AUTO" Or IsNumeric(valAxe) = False Then
            valueAxe = "Auto"
        Else
            If Application.DecimalSeparator = "," Then
                valueAxe = replace(valAxe, ".", ",")
            Else
                valueAxe = replace(valAxe, ",", ".")
            End If
        End If
End Function

Function ChartParam(sdv As String, rowParam As Long, GraphName As String)
   Dim i As Integer
   Dim v(2) As Integer
   If GraphName = "Graphique_0" Then
    v(1) = 7
    v(2) = 5
  Else
     v(1) = 14
     v(2) = 12
  End If
    
    With ThisWorkbook.sheets("PARAMETRES GRAPH")
        For i = 1 To 2
             'Minimum
           
             If valueAxe(.Cells(rowParam + v(i), 2).Value) = "Auto" Then
                ThisWorkbook.sheets(sdv).Shapes(GraphName).Chart.Axes(i).MinimumScaleIsAuto = True
            Else
                ThisWorkbook.sheets(sdv).Shapes(GraphName).Chart.Axes(i).MinimumScale = val(valueAxe(.Cells(rowParam + v(i), 2).Value))
            End If
            
             'Maximum
             If valueAxe(.Cells(rowParam + v(i), 3).Value) = "Auto" Then
                ThisWorkbook.sheets(sdv).Shapes(GraphName).Chart.Axes(i).MaximumScaleIsAuto = True
             Else
                ThisWorkbook.sheets(sdv).Shapes(GraphName).Chart.Axes(i).MaximumScale = val(valueAxe(.Cells(rowParam + v(i), 3).Value))
             End If
             
              'Intervalle
             If valueAxe(.Cells(rowParam + v(i), 4).Value) = "Auto" Then
                ThisWorkbook.sheets(sdv).Shapes(GraphName).Chart.Axes(i).MajorUnitIsAuto = True
             Else
                ThisWorkbook.sheets(sdv).Shapes(GraphName).Chart.Axes(i).MajorUnit = val(valueAxe(.Cells(rowParam + v(i), 4).Value))
             End If
        
        Next i
    End With
   
    
End Function

Function Visu_Graphe(sdv As String, graph As String)
    Dim MaxsV As Long
    Dim i As Integer
    Dim colPriori As Integer
    Dim t As Integer
                         
    With sheets(sdv).ChartObjects(graph).Chart.SeriesCollection(1)
                t = 6
              For i = 1 To .Points.Count
                   t = t + 1
                   If i > (ThisWorkbook.Worksheets(sdv).Range("O65000").End(xlUp).row - 6) Then Exit For
                   
                   With .Points(i)
                        colPriori = ThisWorkbook.sheets(sdv).Rows(6).Find(What:="Event Rating", lookat:=xlWhole).Column
                        If ThisWorkbook.sheets(sdv).Cells(t, colPriori) = "GREEN" Then
                           .MarkerBackgroundColorIndex = 10
                           .MarkerForegroundColorIndex = 1
                         ElseIf ThisWorkbook.sheets(sdv).Cells(t, colPriori) = "YELLOW" Then
                           .MarkerBackgroundColorIndex = 6
                           .MarkerForegroundColorIndex = 1
                         ElseIf ThisWorkbook.sheets(sdv).Cells(t, colPriori) = "RED" Or ThisWorkbook.sheets(sdv).Cells(t, colPriori) = "RED +" Then
                           .MarkerBackgroundColorIndex = 3
                           .MarkerForegroundColorIndex = 1
                         End If
       
                       'Mettre un symbole diff�rent pour chaque LDP
                      
                       
                        If InStr(1, UCase(sdv), "DOWNSH") <> 0 Then
                             If Not ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="gear new", lookat:=xlWhole) Is Nothing Then _
                             MaxsV = CNum(ThisWorkbook.sheets(sdv).Cells(t, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="gear new", lookat:=xlWhole).Column))
                       ElseIf Not ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="gear old", lookat:=xlWhole) Is Nothing And _
                       Not ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="gear new", lookat:=xlWhole) Is Nothing Then
                           MaxsV = _
                           Application.Max(ThisWorkbook.sheets(sdv).Cells(t, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="gear new", lookat:=xlWhole).Column), _
                            ThisWorkbook.sheets(sdv).Cells(t, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="gear old", lookat:=xlWhole).Column))
                       ElseIf Not ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="gear new", lookat:=xlWhole) Is Nothing Then
                            MaxsV = CNum(ThisWorkbook.sheets(sdv).Cells(t, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="gear new", lookat:=xlWhole).Column))
                       ElseIf Not ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="gear old", lookat:=xlWhole) Is Nothing Then
                            MaxsV = CNum(ThisWorkbook.sheets(sdv).Cells(t, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="gear old", lookat:=xlWhole).Column))
                       ElseIf Not ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="Gear", lookat:=xlWhole) Is Nothing Then
                            MaxsV = CNum(ThisWorkbook.sheets(sdv).Cells(t, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="Gear", lookat:=xlWhole).Column))
                       Else
                           MaxsV = 1
                       End If
       
       
                       
       
                       If MaxsV = 2 Then
                           .MarkerStyle = 1
                           .MarkerSize = 8
                       ElseIf MaxsV = 3 Then
                           .MarkerStyle = 8
                           .MarkerSize = 9
                       ElseIf MaxsV = 4 Then
                           .MarkerStyle = 4
                           .MarkerSize = 9
                       ElseIf MaxsV = 5 Then
                           .MarkerStyle = 2
                           .MarkerSize = 11
                       ElseIf MaxsV = 6 Then
                           .MarkerStyle = 3
                           .MarkerSize = 11
                       ElseIf MaxsV = 7 Then
                           .MarkerStyle = 8
                           .MarkerSize = 11
                       ElseIf MaxsV = 8 Then
                           .MarkerStyle = 1
                           .MarkerSize = 11
                       ElseIf MaxsV = 1 Then
                           .MarkerStyle = 3
                           .MarkerSize = 8
                       End If
       
                   End With
       
               Next i
       
           End With
 
   
End Function
Sub Visu_Tableau(sdv As String, typeTableau As String, rowTab As Long, colTab As Long, ligne As Long)

    Dim ColdArray(), HotArray(), TotalArray()
    Dim lastRow
    Dim a As Integer
    Dim b As Integer
    Dim i As Integer
    Dim ShColor As String
    Dim Temperature
    Dim ColColor As Integer
    Dim Shift As String
    Dim k As Integer
    Dim clim, frein, FSE
    Dim vitesse
     
    If typeTableau = "Cold_Hot" Then

        With ThisWorkbook.sheets(sdv)
            lastRow = .Range("Q65000").End(xlUp).row
          
            ReDim ColdArray(11, 2), HotArray(11, 2), TotalArray(11, 2)
            For a = 0 To 11
                For b = 0 To 2
                    ColdArray(a, b) = 0
                    HotArray(a, b) = 0
                    TotalArray(a, b) = 0
                Next
            Next

            For i = 7 To lastRow
                ShColor = .Cells(i, .Rows(6).Cells.Find(What:="Event Rating", lookat:=xlWhole).Column)
                If paramToFind(ligne, "Cold_Hot", 2, sdv) <> "" Then
                    Temperature = .Cells(i, .Rows(6).Cells.Find(What:=paramToFind(ligne, "Cold_Hot", 2, sdv), lookat:=xlWhole).Column)
                Else
                    Exit Sub
                End If
                
                If ShColor = "GREEN" Then
                    ColColor = 0
                ElseIf ShColor = "YELLOW" Then
                    ColColor = 1
                ElseIf ShColor = "RED" Or ShColor = "RED +" Then
                    ColColor = 2
                End If
                
                If paramToFind(ligne, "Cold_Hot", 3, sdv) <> "" And paramToFind(ligne, "Cold_Hot", 4, sdv) <> "" Then
                    Shift = .Cells(i, .Rows(6).Cells.Find(What:=paramToFind(ligne, "Cold_Hot", 3, sdv), lookat:=xlWhole).Column) & "-" & .Cells(i, .Rows(6).Cells.Find(What:=paramToFind(ligne, "Cold_Hot", 4, sdv), lookat:=xlWhole).Column)
                Else
                    Exit Sub
                End If
                
                
               ' Shift = .Cells(i, .Rows(6).Cells.Find(What:=paramToFind(ligne, "Cold_Hot", 3), LookAt:=xlWhole).Column) & "-" & .Cells(i, ThisWorkbook.sheets(SDV).Rows(6).Cells.Find(What:="New Selector_Lever_Position", LookAt:=xlWhole).Column)

                If Temperature < 60 And Temperature <> "" Then

                    If Shift = "D-R" Then
                        ColdArray(0, ColColor) = ColdArray(0, ColColor) + 1
                        TotalArray(0, ColColor) = TotalArray(0, ColColor) + 1
                    ElseIf Shift = "R-D" Then
                        ColdArray(1, ColColor) = ColdArray(1, ColColor) + 1
                        TotalArray(1, ColColor) = TotalArray(1, ColColor) + 1
                    ElseIf Shift = "N-D" Then
                        ColdArray(2, ColColor) = ColdArray(2, ColColor) + 1
                        TotalArray(2, ColColor) = TotalArray(2, ColColor) + 1
                    ElseIf Shift = "D-N" Then
                        ColdArray(3, ColColor) = ColdArray(3, ColColor) + 1
                        TotalArray(3, ColColor) = TotalArray(3, ColColor) + 1
                    ElseIf Shift = "N-R" Then
                        ColdArray(4, ColColor) = ColdArray(4, ColColor) + 1
                        TotalArray(4, ColColor) = TotalArray(4, ColColor) + 1
                    ElseIf Shift = "R-N" Then
                        ColdArray(5, ColColor) = ColdArray(5, ColColor) + 1
                        TotalArray(5, ColColor) = TotalArray(5, ColColor) + 1
                    ElseIf Shift = "P-D" Then
                        ColdArray(6, ColColor) = ColdArray(6, ColColor) + 1
                        TotalArray(6, ColColor) = TotalArray(6, ColColor) + 1
                    ElseIf Shift = "D-P" Then
                        ColdArray(7, ColColor) = ColdArray(7, ColColor) + 1
                        TotalArray(7, ColColor) = TotalArray(7, ColColor) + 1
                    ElseIf Shift = "P-R" Then
                        ColdArray(8, ColColor) = ColdArray(8, ColColor) + 1
                        TotalArray(8, ColColor) = TotalArray(8, ColColor) + 1
                    ElseIf Shift = "R-P" Then
                        ColdArray(9, ColColor) = ColdArray(9, ColColor) + 1
                        TotalArray(9, ColColor) = TotalArray(9, ColColor) + 1
                    ElseIf Shift = "M-D" Then
                        ColdArray(10, ColColor) = ColdArray(10, ColColor) + 1
                        TotalArray(10, ColColor) = TotalArray(10, ColColor) + 1
                    ElseIf Shift = "D-M" Then
                        ColdArray(11, ColColor) = ColdArray(11, ColColor) + 1
                        TotalArray(11, ColColor) = TotalArray(11, ColColor) + 1
                    End If

                ElseIf Temperature >= 60 Then

                    If Shift = "D-R" Then
                        HotArray(0, ColColor) = HotArray(0, ColColor) + 1
                        TotalArray(0, ColColor) = TotalArray(0, ColColor) + 1
                    ElseIf Shift = "R-D" Then
                        HotArray(1, ColColor) = HotArray(1, ColColor) + 1
                        TotalArray(1, ColColor) = TotalArray(1, ColColor) + 1
                    ElseIf Shift = "N-D" Then
                        HotArray(2, ColColor) = HotArray(2, ColColor) + 1
                        TotalArray(2, ColColor) = TotalArray(2, ColColor) + 1
                    ElseIf Shift = "D-N" Then
                        HotArray(3, ColColor) = HotArray(3, ColColor) + 1
                        TotalArray(3, ColColor) = TotalArray(3, ColColor) + 1
                    ElseIf Shift = "N-R" Then
                        HotArray(4, ColColor) = HotArray(4, ColColor) + 1
                        TotalArray(4, ColColor) = TotalArray(4, ColColor) + 1
                    ElseIf Shift = "R-N" Then
                        HotArray(5, ColColor) = HotArray(5, ColColor) + 1
                        TotalArray(5, ColColor) = TotalArray(5, ColColor) + 1
                    ElseIf Shift = "P-D" Then
                        HotArray(6, ColColor) = HotArray(6, ColColor) + 1
                        TotalArray(6, ColColor) = TotalArray(6, ColColor) + 1
                    ElseIf Shift = "D-P" Then
                        HotArray(7, ColColor) = HotArray(7, ColColor) + 1
                        TotalArray(7, ColColor) = TotalArray(7, ColColor) + 1
                    ElseIf Shift = "P-R" Then
                        HotArray(8, ColColor) = HotArray(8, ColColor) + 1
                        TotalArray(8, ColColor) = TotalArray(8, ColColor) + 1
                    ElseIf Shift = "R-P" Then
                        HotArray(9, ColColor) = HotArray(9, ColColor) + 1
                        TotalArray(9, ColColor) = TotalArray(9, ColColor) + 1
                    ElseIf Shift = "M-D" Then
                        HotArray(10, ColColor) = HotArray(10, ColColor) + 1
                        TotalArray(10, ColColor) = TotalArray(10, ColColor) + 1
                    ElseIf Shift = "D-M" Then
                        HotArray(11, ColColor) = HotArray(11, ColColor) + 1
                        TotalArray(11, ColColor) = TotalArray(11, ColColor) + 1
                    End If

                ElseIf Temperature = "" Then

                    If Shift = "D-R" Then
                        TotalArray(0, ColColor) = TotalArray(0, ColColor) + 1
                    ElseIf Shift = "R-D" Then
                        TotalArray(1, ColColor) = TotalArray(1, ColColor) + 1
                    ElseIf Shift = "N-D" Then
                        TotalArray(2, ColColor) = TotalArray(2, ColColor) + 1
                    ElseIf Shift = "D-N" Then
                        TotalArray(3, ColColor) = TotalArray(3, ColColor) + 1
                    ElseIf Shift = "N-R" Then
                        TotalArray(4, ColColor) = TotalArray(4, ColColor) + 1
                    ElseIf Shift = "R-N" Then
                        TotalArray(5, ColColor) = TotalArray(5, ColColor) + 1
                    ElseIf Shift = "P-D" Then
                        TotalArray(6, ColColor) = TotalArray(6, ColColor) + 1
                    ElseIf Shift = "D-P" Then
                        TotalArray(7, ColColor) = TotalArray(7, ColColor) + 1
                    ElseIf Shift = "P-R" Then
                        TotalArray(8, ColColor) = TotalArray(8, ColColor) + 1
                    ElseIf Shift = "R-P" Then
                        TotalArray(9, ColColor) = TotalArray(9, ColColor) + 1
                     ElseIf Shift = "M-D" Then
                        TotalArray(10, ColColor) = TotalArray(10, ColColor) + 1
                    ElseIf Shift = "D-M" Then
                        TotalArray(11, ColColor) = TotalArray(11, ColColor) + 1
                    End If

                End If
            Next

            For a = 0 To 2
                For b = 0 To 11

                    If TotalArray(b, a) <> 0 Then
                        For k = 1 To TotalArray(b, a)
                            .Cells(3 * b + a + rowTab, colTab) = .Cells(3 * b + a + rowTab, colTab) & ChrW(&H25CF)
                        Next
                    End If

                    If ColdArray(b, a) <> 0 Then
                        For k = 1 To ColdArray(b, a)
                            .Cells(3 * b + a + rowTab, colTab + 1) = .Cells(3 * b + a + rowTab, colTab + 1) & ChrW(&H25CF)
                        Next
                    End If

                    If HotArray(b, a) <> 0 Then
                        For k = 1 To HotArray(b, a)
                            .Cells(3 * b + a + rowTab, colTab + 2) = .Cells(3 * b + a + rowTab, colTab + 2) & ChrW(&H25CF)
                        Next
                    End If

                Next
            Next
        End With

    ElseIf typeTableau = "Frein_FSE" Then

        With ThisWorkbook.sheets(sdv)
            lastRow = .Range("Q65000").End(xlUp).row
            ReDim ColdArray(1, 2), HotArray(1, 2), TotalArray(1, 2)
            For a = 0 To 1
                For b = 0 To 2
                    ColdArray(a, b) = 0
                    HotArray(a, b) = 0
                    TotalArray(a, b) = 0
                Next
            Next

            For i = 7 To lastRow
                ShColor = .Cells(i, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="Event Rating", lookat:=xlWhole).Column)
                
                If paramToFind(ligne, "Frein_FSE", 2, sdv) <> "" Then
                     clim = .Cells(i, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:=paramToFind(ligne, "Frein_FSE", 2, sdv), lookat:=xlWhole).Column)
                Else
                    Exit Sub
                End If
                

                If ShColor = "GREEN" Then
                    ColColor = 0
                ElseIf ShColor = "YELLOW" Then
                    ColColor = 1
                ElseIf ShColor = "RED" Or ShColor = "RED +" Then
                    ColColor = 2
                End If

'                 Shift = .Cells(i, ThisWorkbook.sheets(SDV).Rows(6).Cells.Find(What:="Old Selector_Lever_Position", LookAt:=xlWhole).Column) & "-" & .Cells(i, ThisWorkbook.sheets(SDV).Rows(6).Cells.Find(What:="New Selector_Lever_Position", LookAt:=xlWhole).Column)
                 If paramToFind(ligne, "Frein_FSE", 3, sdv) <> "" Then
                     frein = .Cells(i, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:=paramToFind(ligne, "Frein_FSE", 3, sdv), lookat:=xlWhole).Column)
                Else
                    Exit Sub
                End If
                
                 If paramToFind(ligne, "Frein_FSE", 4, sdv) <> "" Then
                      FSE = .Cells(i, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:=paramToFind(ligne, "Frein_FSE", 4, sdv), lookat:=xlWhole).Column)
                Else
                    Exit Sub
                End If
                
               

                If clim = 0 And clim <> "" Then

                    If frein = 1 And FSE = 0 Then
                        ColdArray(0, ColColor) = ColdArray(0, ColColor) + 1
                        TotalArray(0, ColColor) = TotalArray(0, ColColor) + 1
                    ElseIf FSE = 1 And frein = 0 Then
                        ColdArray(1, ColColor) = ColdArray(1, ColColor) + 1
                        TotalArray(1, ColColor) = TotalArray(1, ColColor) + 1
                    End If

                ElseIf clim = 1 Then

                    If frein = 1 And FSE = 0 Then
                        HotArray(0, ColColor) = HotArray(0, ColColor) + 1
                        TotalArray(0, ColColor) = TotalArray(0, ColColor) + 1
                    ElseIf FSE = 1 And frein = 0 Then
                        HotArray(1, ColColor) = HotArray(1, ColColor) + 1
                        TotalArray(1, ColColor) = TotalArray(1, ColColor) + 1
                    End If


                ElseIf clim = "" Then

                    If frein = 1 And FSE = 0 Then
                        TotalArray(0, ColColor) = TotalArray(0, ColColor) + 1
                    ElseIf FSE = 1 And frein = 0 Then
                        TotalArray(1, ColColor) = TotalArray(1, ColColor) + 1
                    End If

                End If
            Next


            For a = 0 To 2
                For b = 0 To 1

                    If TotalArray(b, a) <> 0 Then
                        For k = 1 To TotalArray(b, a)
                            .Cells(3 * b + a + rowTab, colTab) = .Cells(3 * b + a + rowTab, colTab) & ChrW(&H25CF)
                        Next
                    End If

                    If ColdArray(b, a) <> 0 Then
                        For k = 1 To ColdArray(b, a)
                            .Cells(3 * b + a + rowTab, colTab + 1) = .Cells(3 * b + a + rowTab, colTab + 1) & ChrW(&H25CF)
                        Next
                    End If

                    If HotArray(b, a) <> 0 Then
                        For k = 1 To HotArray(b, a)
                            .Cells(3 * b + a + rowTab, colTab + 2) = .Cells(3 * b + a + rowTab, colTab + 2) & ChrW(&H25CF)
                        Next
                    End If

                Next
            Next
        End With

    ElseIf typeTableau = "Decel_Frein" Then

        With ThisWorkbook.sheets(sdv)
            lastRow = .Range("Q65000").End(xlUp).row
            ReDim ColdArray(2, 2), HotArray(2, 2), TotalArray(2, 2)
            For a = 0 To 2
                For b = 0 To 2
                    ColdArray(a, b) = 0
                    HotArray(a, b) = 0
                    TotalArray(a, b) = 0
                Next
            Next

            For i = 7 To lastRow
                ShColor = .Cells(i, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:="Event Rating", lookat:=xlWhole).Column)
                
                
                If paramToFind(ligne, "Decel_Frein", 2, sdv, 2) <> "" Then
                      clim = .Cells(i, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:=paramToFind(ligne, "Decel_Frein", 2, sdv, 2), lookat:=xlWhole).Column)
                Else
                    Exit Sub
                End If

                If ShColor = "GREEN" Then
                    ColColor = 0
                ElseIf ShColor = "YELLOW" Then
                    ColColor = 1
                ElseIf ShColor = "RED" Or ShColor = "RED +" Then
                    ColColor = 2
                End If

                 If paramToFind(ligne, "Decel_Frein", 2, sdv) <> "" And paramToFind(ligne, "Decel_Frein", 3, sdv) <> "" Then
                      Shift = .Cells(i, .Rows(6).Cells.Find(What:=paramToFind(ligne, "Decel_Frein", 2, sdv), lookat:=xlWhole).Column) & "-" & .Cells(i, .Rows(6).Cells.Find(What:=paramToFind(ligne, "Decel_Frein", 3, sdv), lookat:=xlWhole).Column)
                Else
                    Exit Sub
                End If
                
                 If paramToFind(ligne, "Decel_Frein", 3, sdv, 2) <> "" Then
                       frein = .Cells(i, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:=paramToFind(ligne, "Decel_Frein", 3, sdv, 2), lookat:=xlWhole).Column)
                Else
                    Exit Sub
                End If
                
                If paramToFind(ligne, "Decel_Frein", 4, sdv) <> "" Then
                      vitesse = .Cells(i, ThisWorkbook.sheets(sdv).Rows(6).Cells.Find(What:=paramToFind(ligne, "Decel_Frein", 4, sdv), lookat:=xlWhole).Column)
                Else
                    Exit Sub
                End If
               
                

                If clim = 0 Then
                    If frein = 1 Then
                        ColdArray(0, ColColor) = ColdArray(0, ColColor) + 1
                        TotalArray(0, ColColor) = TotalArray(0, ColColor) + 1
                    ElseIf frein = 0 And Shift = "D-N" And vitesse > 0 Then
                        ColdArray(1, ColColor) = ColdArray(1, ColColor) + 1
                        TotalArray(1, ColColor) = TotalArray(1, ColColor) + 1
                    ElseIf frein = 0 And Shift = "D-N" And vitesse = 0 And vitesse <> "" Then
                        ColdArray(2, ColColor) = ColdArray(2, ColColor) + 1
                        TotalArray(2, ColColor) = TotalArray(2, ColColor) + 1
                    End If

                ElseIf clim = 1 Then

                    If frein = 1 Then
                        HotArray(0, ColColor) = HotArray(0, ColColor) + 1
                        TotalArray(0, ColColor) = TotalArray(0, ColColor) + 1
                    ElseIf frein = 0 And Shift = "D-N" And vitesse > 0 Then
                        HotArray(1, ColColor) = HotArray(1, ColColor) + 1
                        TotalArray(1, ColColor) = TotalArray(1, ColColor) + 1
                    ElseIf frein = 0 And Shift = "D-N" And vitesse = 0 And vitesse <> "" Then
                        HotArray(2, ColColor) = HotArray(2, ColColor) + 1
                        TotalArray(2, ColColor) = TotalArray(2, ColColor) + 1
                    End If


                ElseIf clim = "" Then

                    If frein = 1 Then
                        TotalArray(0, ColColor) = TotalArray(0, ColColor) + 1
                    ElseIf frein = 0 And Shift = "D-N" And vitesse > 0 Then
                        TotalArray(1, ColColor) = TotalArray(1, ColColor) + 1
                    ElseIf frein = 0 And Shift = "D-N" And vitesse = 0 And vitesse <> "" Then
                        TotalArray(2, ColColor) = TotalArray(2, ColColor) + 1
                    End If

                End If
            Next


            For a = 0 To 2
                For b = 0 To 2

                    If TotalArray(b, a) <> 0 Then
                        For k = 1 To TotalArray(b, a)
                            .Cells(3 * b + a + rowTab, colTab) = .Cells(3 * b + a + rowTab, colTab) & ChrW(&H25CF)
                        Next
                    End If

                    If ColdArray(b, a) <> 0 Then
                        For k = 1 To ColdArray(b, a)
                            .Cells(3 * b + a + rowTab, colTab + 1) = .Cells(3 * b + a + rowTab, colTab + 1) & ChrW(&H25CF)
                        Next
                    End If

                    If HotArray(b, a) <> 0 Then
                        For k = 1 To HotArray(b, a)
                            .Cells(3 * b + a + rowTab, colTab + 2) = .Cells(3 * b + a + rowTab, colTab + 2) & ChrW(&H25CF)
                        Next
                    End If

                Next
            Next
        End With

    End If

End Sub



Function CNum(num As Long)
  CNum = val(num)
End Function

Function copyGraph(graph As shape, sdv As String)
    Dim i As Integer
    
    On Error Resume Next
    For i = 1 To 10
        ERR.Clear
        graph.Copy
        ThisWorkbook.Worksheets(sdv).Paste
        If ERR.Number = 0 Then Exit Function Else Application.Wait Now + TimeValue("0:00:02")
    Next i

End Function

Function checkObject(nomTable As String, sdv As String) As String
  Dim o As ListObject
  Dim suffix As String
  
  For Each o In ThisWorkbook.Worksheets(sdv).ListObjects
     If Left(o.Name, Len(nomTable)) = nomTable Then
            suffix = Mid(o.Name, Len(nomTable) + 1, 1)
            If suffix = "" Or (suffix >= "0" And suffix <= "9") Then
                checkObject = o.Name
                Exit Function
            End If
     End If
  Next o
End Function

Function paramToFind(ligne As Long, typeTable As String, id As Integer, sdv As String, Optional rPlus As Integer)
    paramToFind = ""
    With ThisWorkbook.sheets("PARAMETRES GRAPH")
        If typeTable = "Cold_Hot" Then
            If Not ThisWorkbook.sheets(sdv).Rows(6).Find(What:=.Cells(ligne + 19, id), lookat:=xlWhole) Is Nothing Then paramToFind = .Cells(ligne + 19, id)
        ElseIf typeTable = "Frein_FSE" Then
            If Not ThisWorkbook.sheets(sdv).Rows(6).Find(What:=.Cells(ligne + 24, id), lookat:=xlWhole) Is Nothing Then paramToFind = .Cells(ligne + 24, id)
        ElseIf typeTable = "Decel_Frein" Then
            If rPlus > 0 Then
                 If Not ThisWorkbook.sheets(sdv).Rows(6).Find(What:=.Cells(ligne + 29 + rPlus, id), lookat:=xlWhole) Is Nothing Then paramToFind = .Cells(ligne + 29 + rPlus, id)
            Else
                If Not ThisWorkbook.sheets(sdv).Rows(6).Find(What:=.Cells(ligne + 29, id), lookat:=xlWhole) Is Nothing Then paramToFind = .Cells(ligne + 29, id)
            End If
        End If
    End With
End Function

























