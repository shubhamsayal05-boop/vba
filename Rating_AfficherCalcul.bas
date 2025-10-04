Attribute VB_Name = "Rating_AfficherCalcul"
Option Explicit

Public colorGlobalDriv As Object
Public colorGlobalDyn  As Object
Public colorGlobalDrivPred As Object
Public colorGlobalDynPred  As Object
Public found As Boolean
Sub PrepaCalcul_Rating()
    Dim msg_err
    Dim k, col, endcol, startcol As Integer
    Dim Target_Veh As String
    Dim userResponse As VbMsgBoxResult
    
    
    
    
    
    On Error GoTo gEr
    EventAndScreen (False)
 
    Set colorGlobalDyn = CreateObject("Scripting.Dictionary")
    Set colorGlobalDynPred = CreateObject("Scripting.Dictionary")
    Set colorGlobalDriv = CreateObject("Scripting.Dictionary")
    Set colorGlobalDrivPred = CreateObject("Scripting.Dictionary")
    
    
    Target_Veh = ThisWorkbook.Worksheets("HOME").Cells(23, 3)
    startcol = 14
    
    endcol = 0
    For col = startcol To ThisWorkbook.Worksheets("RATING").Columns.Count
        If ThisWorkbook.Worksheets("RATING").Cells(21, col).Value = "Drivability Lowest Events" Then
            endcol = col
            Exit For
        End If
    Next col
    found = False
    k = startcol
    Do While found = False And k <= endcol
        If Target_Veh = ThisWorkbook.Worksheets("RATING").Cells(21, k) Then
            found = True
        Else
            k = k + 1
        End If
    Loop
    
   
    
        If ThisWorkbook.sheets("HOME").Range("Mode").Value = "" Or ThisWorkbook.sheets("HOME").Range("DriveVersion").Value = "" Or ThisWorkbook.sheets("HOME").Range("Milestone").Value = "" Or ThisWorkbook.sheets("HOME").Range("Gears") = "" Or ThisWorkbook.sheets("HOME").Range("Fuel").Value = "" Then
            MsgBox "Information missing. Please complete the ""PROJECT SUMMARY"" section.", vbExclamation, "Warning!"
        Else
        
            ' si la feuille tampon DATA existe pas donc faut reimpoter tous les donnees du projet courant
            If sheetExists("DATA") = True Then
                remplirFormule
                If found = False Then
                    userResponse = MsgBox("Target vehicle not found in RATING, continue?", vbYesNo + vbQuestion, "ODRIV")
                    If userResponse = vbYes Then
                        Call Afficher_Calcul
                    Else
                        MsgBox "The calculation is stopped", vbInformation
                    End If
                Else
                    Call Afficher_Calcul
                End If
            End If
        
        End If
   
   Set colorGlobalDyn = Nothing
   Set colorGlobalDynPred = Nothing
   Set colorGlobalDriv = Nothing
   Set colorGlobalDrivPred = Nothing
   
    EventAndScreen (True)
gEr:
    If ERR.Number <> 0 Then
        EventAndScreen (True)
        Application.Calculation = xlCalculationAutomatic
        msg_err = ERR.description
        ProgressExit
        MsgBox msg_err, vbCritical, "ODRIV"
    End If
End Sub

Sub NoteGlobale_2()
    Dim r As Range
    Dim v
    Dim i As Long
    Dim iCol As Integer
    Dim sh As Worksheet
    Dim vCol As String
    Dim pr As Integer
    Dim rt As Integer, ir
    
    
    ir = ThisWorkbook.Worksheets("RATING").Rows("10:10").Find(What:="Tested vehicle", lookat:=xlWhole).Column
    v = getColDtD("driv")
    i = getColDtD("dyn")
     'ThisWorkbook.Sheets("rating").Range("L10").Formula = "=calculs!M39"
   With ThisWorkbook.sheets("Graph_status")
             If Not colorGlobalDriv Is Nothing Then
                     If colorGlobalDriv.Count <> 0 Then
                             If ThisWorkbook.Worksheets("RATING").Cells(12, ir) < .Cells(.Range("tDriv").row + 2, v) Then
                                 If (ThisWorkbook.sheets("RATING").Cells(11, ir)) > .Cells(.Range("iDriv").row + 2, v) Then
                                     ThisWorkbook.sheets("RATING").Range("E11").Value = "GREEN"
                                 ElseIf (ThisWorkbook.sheets("RATING").Cells(11, ir)) < .Cells(.Range("iDriv").row + 3, v) Then
                                     ThisWorkbook.sheets("RATING").Range("E11").Value = "YELLOW"
                                 Else
                                     ThisWorkbook.sheets("RATING").Range("E11").Value = "GREEN"
                                 End If
                             ElseIf ThisWorkbook.Worksheets("RATING").Cells(12, ir) > .Cells(.Range("tDriv").row + 3, v) Then
                                 If (ThisWorkbook.sheets("RATING").Cells(11, ir)) > .Cells(.Range("iDriv").row + 2, v) Then
                                     ThisWorkbook.sheets("RATING").Range("E11").Value = "RED"
                                 ElseIf (ThisWorkbook.sheets("RATING").Cells(11, ir)) < .Cells(.Range("iDriv").row + 3, v) Then
                                     ThisWorkbook.sheets("RATING").Range("E11").Value = "RED"
                                 Else
                                     ThisWorkbook.sheets("RATING").Range("E11").Value = "RED"
                                 End If
                             Else
                                 If (ThisWorkbook.sheets("RATING").Cells(11, ir)) > .Cells(.Range("tDriv").row + 2, v) Then
                                     ThisWorkbook.sheets("RATING").Range("E11").Value = "YELLOW"
                                 ElseIf (ThisWorkbook.sheets("RATING").Cells(11, ir)) < .Cells(.Range("iDriv").row + 3, v) Then
                                     ThisWorkbook.sheets("RATING").Range("E11").Value = "RED"
                                 Else
                                     ThisWorkbook.sheets("RATING").Range("E11").Value = "YELLOW"
                                 End If
                             End If
                             
                     End If
            End If
            
            If Not colorGlobalDyn Is Nothing Then
                     If colorGlobalDyn.Count <> 0 Then
                              If ThisWorkbook.Worksheets("RATING").Cells(18, ir) < .Cells(.Range("tDYN").row + 2, v) Then
                                  If (ThisWorkbook.sheets("RATING").Cells(17, ir)) > .Cells(.Range("iDYN").row + 2, v) Then
                                      ThisWorkbook.sheets("RATING").Range("E17").Value = "GREEN"
                                  ElseIf (ThisWorkbook.sheets("RATING").Cells(17, ir)) < .Cells(.Range("iDYN").row + 3, v) Then
                                      ThisWorkbook.sheets("RATING").Range("E17").Value = "YELLOW"
                                  Else
                                      ThisWorkbook.sheets("RATING").Range("E17").Value = "GREEN"
                                  End If
                              ElseIf ThisWorkbook.Worksheets("RATING").Cells(18, ir) > .Cells(.Range("tDYN").row + 3, v) Then
                                  If (ThisWorkbook.sheets("RATING").Cells(17, ir)) > .Cells(.Range("iDYN").row + 2, v) Then
                                      ThisWorkbook.sheets("RATING").Range("E17").Value = "RED"
                                  ElseIf (ThisWorkbook.sheets("RATING").Cells(17, ir)) < .Cells(.Range("iDYN").row + 3, v) Then
                                      ThisWorkbook.sheets("RATING").Range("E17").Value = "RED"
                                  Else
                                      ThisWorkbook.sheets("RATING").Range("E17").Value = "RED"
                                  End If
                              Else
                                  If (ThisWorkbook.sheets("RATING").Cells(17, ir)) > .Cells(.Range("tDYN").row + 2, v) Then
                                      ThisWorkbook.sheets("RATING").Range("E17").Value = "YELLOW"
                                  ElseIf (ThisWorkbook.sheets("RATING").Cells(17, ir)) < .Cells(.Range("iDYN").row + 3, v) Then
                                      ThisWorkbook.sheets("RATING").Range("E17").Value = "RED"
                                  Else
                                      ThisWorkbook.sheets("RATING").Range("E17").Value = "YELLOW"
                                  End If
                              End If
                     End If
             End If
             
        End With
End Sub

Sub NoteGlobale_Pred()
    Dim r As Range
    Dim v
    Dim i As Long
    Dim iCol As Integer
    Dim sh As Worksheet
    Dim vCol As String
    Dim pr As Integer
    Dim rt As Integer, ir
    
    
    ir = ThisWorkbook.Worksheets("RATING").Rows("10:10").Find(What:="Tested vehicle", lookat:=xlWhole).Column
    v = getColDtD("driv", 4)
    i = getColDtD("dyn", 4)
     'ThisWorkbook.Sheets("rating").Range("L10").Formula = "=calculs!M39"
   With ThisWorkbook.sheets("Graph_status")
             If Not colorGlobalDriv Is Nothing Then
                     If colorGlobalDriv.Count <> 0 Then
                             If ThisWorkbook.Worksheets("RATING").Cells(12, ir) < .Cells(.Range("tDriv").row + 2, v) Then
                                 If (ThisWorkbook.sheets("RATING").Cells(11, ir)) > .Cells(.Range("iDriv").row + 2, v) Then '.Range("RESULTATGLOBAL1")
                                     ThisWorkbook.sheets("RATING").Range("F11").Value = "GREEN"
                                 ElseIf (ThisWorkbook.sheets("RATING").Cells(11, ir)) < .Cells(.Range("iDriv").row + 3, v) Then
                                     ThisWorkbook.sheets("RATING").Range("F11").Value = "YELLOW"
                                 Else
                                     ThisWorkbook.sheets("RATING").Range("F11").Value = "GREEN"
                                 End If
                             ElseIf ThisWorkbook.Worksheets("RATING").Cells(12, ir) > .Cells(.Range("tDriv").row + 3, v) Then
                                 If (ThisWorkbook.sheets("RATING").Cells(11, ir)) > .Cells(.Range("iDriv").row + 2, v) Then
                                     ThisWorkbook.sheets("RATING").Range("F11").Value = "RED"
                                 ElseIf (ThisWorkbook.sheets("RATING").Cells(11, ir)) < .Cells(.Range("iDriv").row + 3, v) Then
                                     ThisWorkbook.sheets("RATING").Range("F11").Value = "RED"
                                 Else
                                     ThisWorkbook.sheets("RATING").Range("F11").Value = "RED"
                                 End If
                             Else
                                 If (ThisWorkbook.sheets("RATING").Cells(11, ir)) > .Cells(.Range("iDriv").row + 2, v) Then
                                     ThisWorkbook.sheets("RATING").Range("F11").Value = "YELLOW"
                                 ElseIf (ThisWorkbook.sheets("RATING").Cells(11, ir)) < .Cells(.Range("iDriv").row + 3, v) Then
                                     ThisWorkbook.sheets("RATING").Range("F11").Value = "RED"
                                 Else
                                     ThisWorkbook.sheets("RATING").Range("F11").Value = "YELLOW"
                                 End If
                             End If
                             
                     End If
            End If
            
            If Not colorGlobalDyn Is Nothing Then
                     If colorGlobalDyn.Count <> 0 Then
                              If ThisWorkbook.Worksheets("RATING").Cells(18, ir) < .Cells(.Range("tDYN").row + 2, v) Then
                                  If (ThisWorkbook.sheets("RATING").Cells(17, ir)) > .Cells(.Range("iDYN").row + 2, v) Then
                                      ThisWorkbook.sheets("RATING").Range("F17").Value = "GREEN"
                                  ElseIf (ThisWorkbook.sheets("RATING").Cells(17, ir)) < .Cells(.Range("iDYN").row + 3, v) Then
                                      ThisWorkbook.sheets("RATING").Range("F17").Value = "YELLOW"
                                  Else
                                      ThisWorkbook.sheets("RATING").Range("F17").Value = "GREEN"
                                  End If
                              ElseIf ThisWorkbook.Worksheets("RATING").Cells(18, ir) > .Cells(.Range("tDYN").row + 3, v) Then
                                  If (ThisWorkbook.sheets("RATING").Cells(17, ir)) > .Cells(.Range("iDYN").row + 2, v) Then
                                      ThisWorkbook.sheets("RATING").Range("F17").Value = "RED"
                                  ElseIf (ThisWorkbook.sheets("RATING").Cells(17, ir)) < .Cells(.Range("iDYN").row + 3, v) Then
                                      ThisWorkbook.sheets("RATING").Range("F17").Value = "RED"
                                  Else
                                      ThisWorkbook.sheets("RATING").Range("F17").Value = "RED"
                                  End If
                              Else
                                  If (ThisWorkbook.sheets("RATING").Cells(17, ir)) > .Cells(.Range("iDYN").row + 2, v) Then
                                      ThisWorkbook.sheets("RATING").Range("F17").Value = "YELLOW"
                                  ElseIf (ThisWorkbook.sheets("RATING").Cells(17, ir)) < .Cells(.Range("iDYN").row + 3, v) Then
                                      ThisWorkbook.sheets("RATING").Range("F17").Value = "RED"
                                  Else
                                      ThisWorkbook.sheets("RATING").Range("F17").Value = "YELLOW"
                                  End If
                              End If
                     End If
             End If
             
        End With
End Sub
Function getColDtD(part As String, Optional ml As String)
    Dim mil
    getColDtD = 0
    With sheets("Graph_status")
        If ml = "" Then
            mil = sheets("HOME").Range("MILESTONE")
        Else
            mil = ml
        End If
        If part = "driv" Then
            If Not .Rows(.Range("iDriv").row + 1).Find(What:=mil, lookat:=xlWhole) Is Nothing Then
                getColDtD = .Rows(.Range("iDriv").row + 1).Find(What:=mil, lookat:=xlWhole).Column
            End If
        Else
            If Not .Rows(.Range("iDyn").row + 1).Find(What:=mil, lookat:=xlWhole) Is Nothing Then
                getColDtD = .Rows(.Range("iDyn").row + 1).Find(What:=mil, lookat:=xlWhole).Column
            End If
        End If
    End With
End Function
Sub Afficher_Calcul()
    Dim NbSDV As Integer
    Dim i As Integer
    Dim test As Double
    Dim v
    Dim ii As Long
    Dim lastColonnes As Long
    Dim lastRow As Long, lastr As Long
    Dim j
    
    
    
    
    
    
    ProgressLoad
    ProgressTitle ("Nettoyage Des SDV")
    Call Erase_All2(True)
     ThisWorkbook.Worksheets("HOME").Range("Moniteur").Interior.color = RGB(255, 0, 0)
     Call ResetOptions
    v = ThisWorkbook.sheets("structure").UsedRange.Columns(2).Value
    
    
    Call CompteSumPriority.initList
    Call CompteSumPriorityDyn.initListDyn
    For ii = 2 To UBound(v, 1)
        If Len(v(ii, 1)) > 0 And sheetExists(v(ii, 1)) = True Then
            If (checkCriteria(v(ii, 1)) = True And checkCorrespondancePriority(v(ii, 1)) = True) _
            Or (checkCriteria(v(ii, 1)) = True And checkCorrespondancePriorityDyn(v(ii, 1))) Then
                 Call UpdateTab(v(ii, 1))
                 ThisWorkbook.sheets(CStr((v(ii, 1)))).Visible = -1
            Else
                With ThisWorkbook.sheets(CStr((v(ii, 1))))
                    For i = 13 To 15
                        If .Cells(.Rows.Count, i).End(xlUp).row > lastr Then lastr = .Cells(.Rows.Count, i).End(xlUp).row
                    Next i
                    If lastr > 6 Then
                        Call RAZ_onglet(CStr(v(ii, 1)))
                        Call RAZ_ongletDyn(CStr(v(ii, 1)))
                    End If
                    .Visible = 2
                End With
            End If
            
       End If
    Next ii
   
    
     NoteGlobale
     NoteGlobaleDyn
     Call Moniteur("Rating has been calculated.")
     Call calculTauxPointBas
     Call calculTauxPointBasDyn
     Call hyperlinkAdd
     Call MaskEmptySdv
     
     ThisWorkbook.RefreshAll
     ThisWorkbook.Worksheets("RATING").Activate
     ThisWorkbook.Worksheets("RATING").Calculate
    
     'pour la couleur_rating
     ProgressTitle ("MAJ RATING")
     Call NoteGlobale_2
      If sheets("HOME").Range("Milestone") <> 4 Then Call NoteGlobale_Pred
     MajTargets
'     test = GetNoteGlobalTarget("driv")
'    ThisWorkbook.sheets("RATING").Range("EQ3").Value = IIf(test = "-555", "", test)
'
'    test = GetNoteGlobalTarget("dyn")
'    ThisWorkbook.sheets("RATING").Range("EQ4").Value = IIf(test = "-555", "", test)
    CalculIndexTarget
    getTotalColor
    
    
    If found = True Then
        Call summIndice("driv")
        Call summIndice("dyn")
    End If
        
   
    
    
    ThisWorkbook.Worksheets("RATING").Shapes("UpdateTargetButton").Visible = True
    
    Erase v
    ProgressExit
    MsgBox "DONE", vbInformation, "ODRIV"
End Sub

Sub OkaZ()
    Dim test As Double
      Call calculTauxPointBas
      Call calculTauxPointBasDyn
    
     ProgressTitle ("MAJ RATING")
     Call NoteGlobale_2
'     MajTargets
'     test = GetNoteGlobalTarget("driv")
'    ThisWorkbook.sheets("RATING").Range("EQ3").Value = IIf(test = "-555", "", test)
'
'    test = GetNoteGlobalTarget("dyn")
'    ThisWorkbook.sheets("RATING").Range("EQ4").Value = IIf(test = "-555", "", test)
    CalculIndexTarget
    getTotalColor
    Call summIndice("driv")
    Call summIndice("dyn")
   
End Sub

Function getColumnVeh(pr As String)
        Dim i
        Dim colname
        Dim jConcat As String
        colname = ThisWorkbook.sheets("HOME").Range("C23").Value
        
        jConcat = ""
        With ThisWorkbook.sheets("RATING")
                If pr = "driv" Then
                        i = .Rows("21:22").Find(What:="Driveability Index", lookat:=xlWhole).Column
                        i = i + 1
                        While Len(.Cells(21, i)) > 0 And .Cells(21, i) <> "Drivability Lowest Events"
                            If InStr(1, "," & colname & ",", "," & .Cells(21, i) & ",") <> 0 Then
                               jConcat = IIf(jConcat = "", i, jConcat & "," & i)
                            End If
                            i = i + 1
                        Wend
                Else
                        i = .Rows("21:22").Find(What:="Dynamism Index", lookat:=xlWhole).Column
                        i = i + 1
                        While Len(.Cells(21, i)) > 0 And .Cells(21, i) <> "Dynamism Lowest Events"
                           If InStr(1, "," & colname & ",", "," & .Cells(21, i) & ",") <> 0 Then
                               jConcat = IIf(jConcat = "", i, jConcat & "," & i)
                            End If
                            i = i + 1
                        Wend
                End If
        End With
        getColumnVeh = jConcat
            
End Function

Function summIndice(part As String)
    Dim r As Range
    Dim i As Long
    Dim tot As Variant, totTable()  As Variant
    Dim cols As String, Co() As String
    Dim dynamicDriv As Integer
    Dim j, z, w, y
    Dim taux1, tauxTable() As Variant
    Dim DL As Integer
    Dim getOkTarget() As Boolean
    
    i = 23
    
    j = i
    
    cols = getColumnVeh(part)
    If InStr(1, cols, ",") = 0 Then
        ReDim Co(0)
        Co(0) = cols
    Else
        Co = Split(cols, ",")
    End If
    
    ReDim getOkTarget(UBound(Co))
    ReDim totTable(UBound(Co))
    ReDim tauxTable(UBound(Co))
    
    For z = 0 To UBound(Co)
         getOkTarget(z) = checkEmptyDD(part, 23, Co(z))
         totTable(z) = 0
         tauxTable(z) = 0
    Next z
    
    If part = "driv" Then dynamicDriv = 1
    If part = "dyn" Then dynamicDriv = 0
    
    
    With ThisWorkbook.sheets("RATING")
    DL = getLastRowRating
    
    While i <= DL
        If Len(.Cells(i, 2)) > 0 Then
            If j <> i Then
                If dynamicDriv = 1 Then
                    If taux1 > 0 Then .Range("M" & j).Value = Round(tot / taux1, 1)
                    For z = 0 To UBound(Co)
                       If getOkTarget(z) <> True Then
                            If tauxTable(z) > 0 Then .Cells(j, val(Co(z))) = Round(totTable(z) / tauxTable(z), 1)
'                            If taux2 > 0 Then .Range("N" & j).Value = Round(tot(1) / taux2, 1)
                        End If
                    Next z
                    
                   '.Range("V" & j).Value = tot(1)
                Else
                    If taux1 > 0 Then .Cells(j, .Rows("21:22").Find(What:="Dynamism Index", lookat:=xlWhole).Column).Value = Round(tot / taux1, 1)
'                    If getOkTarget <> True Then
'                        If taux2 > 0 Then .Range("W" & j).Value = Round(tot(1) / taux2, 1)
'                        '.Range("W" & j).Value = tot(1)
'                    End If
                     For z = 0 To UBound(Co)
                       If getOkTarget(z) <> True Then
                            If tauxTable(z) > 0 Then .Cells(j, val(Co(z))) = Round(totTable(z) / tauxTable(z), 1)
'                            If taux2 > 0 Then .Range("N" & j).Value = Round(tot(1) / taux2, 1)
                        End If
                    Next z
                End If
                j = i
'                getOkTarget = checkEmptyDD(part, i)
                tot = 0
                taux1 = 0
                 For z = 0 To UBound(Co)
                     getOkTarget(z) = checkEmptyDD(part, i, Co(z))
                     totTable(z) = 0
                     tauxTable(z) = 0
                Next z
            End If
        Else
            If dynamicDriv = 1 Then
                If Len(.Range("M" & i).Value) > 0 Then
                    taux1 = taux1 + Weight(.Range("D" & i).Value)
'                    taux2 = taux1
                    tot = tot + (Weight(.Range("D" & i).Value) * .Range("M" & i).Value)
'                    tot(1) = tot(1) + (Weight(.Range("D" & i).Value) * .Range("N" & i).Value)
                    For z = 0 To UBound(Co)
                         tauxTable(z) = taux1
                         totTable(z) = totTable(z) + (Weight(.Range("D" & i).Value) * .Cells(i, val(Co(z))).Value)
                    Next z
                End If
            Else
                 If Len(Cells(i, .Rows("21:22").Find(What:="Dynamism Index", lookat:=xlWhole).Column).Value) > 0 Then
                    taux1 = taux1 + Weight(.Range("D" & i).Value)
'                    taux2 = taux1
                    tot = tot + (Weight(.Range("D" & i).Value) * .Cells(i, .Rows("21:22").Find(What:="Dynamism Index", lookat:=xlWhole).Column).Value)
'                    tot(1) = tot(1) + (Weight(.Range("D" & i).Value) * .Range("W" & i).Value)
                    For z = 0 To UBound(Co)
                         tauxTable(z) = taux1
                         totTable(z) = totTable(z) + (Weight(.Range("D" & i).Value) * .Cells(i, val(Co(z))).Value)
                    Next z
                 End If
            End If
        End If
        i = i + 1
    Wend
    If j <> i And taux1 <> 0 Then
                If dynamicDriv = 1 Then
                   .Range("M" & j).Value = Round(tot / taux1, 1)
                    For z = 0 To UBound(Co)
                        If getOkTarget(z) <> True Then
                             If tauxTable(z) > 0 Then .Cells(j, val(Co(z))) = Round(totTable(z) / tauxTable(z), 1)
                         End If
                    Next z
'                   If getOkTarget <> True Then
'                        .Range("N" & j).Value = Round(tot(1) / taux2, 1)
'                   End If
'                   .Range("V" & j).Value = tot(1) / taux1
                Else
                  .Cells(j, .Rows("21:22").Find(What:="Dynamism Index", lookat:=xlWhole).Column).Value = Round(tot / taux1, 1)
                  For z = 0 To UBound(Co)
                        If getOkTarget(z) <> True Then
                             If tauxTable(z) > 0 Then .Cells(j, val(Co(z))) = Round(totTable(z) / tauxTable(z), 1)
                         End If
                   Next z
'                  If getOkTarget <> True Then
'                     .Range("W" & j).Value = Round(tot(1) / taux2, 1)
'                     .Range("W" & j).Value = tot(1) / taux1
'                   End If
                End If
              
    End If
            
    End With
End Function


Function checkEmptyDD(part As String, JS As Long, cL As String) As Boolean
    Dim r As Range
    Dim i As Long
    Dim j
    Dim k, l
    Dim DL As Integer
    i = JS + 1
    
    j = i
    
    If part = "driv" Then
        k = 13
        l = val(cL)
    End If
    If part = "dyn" Then
        k = sheets("RATING").Rows("21:22").Find(What:="Dynamism Index", lookat:=xlWhole).Column
        l = val(cL)
    End If
    checkEmptyDD = False
    
    With ThisWorkbook.sheets("RATING")
            DL = getLastRowRating
            
            While i <= DL And Len(.Cells(i, 2)) = 0
                If Len(.Cells(i, 4)) > 0 Then
                    If Len(.Cells(i, k)) > 0 And Len(.Cells(i, l)) = 0 Then
                        checkEmptyDD = True
                        Exit Function
                    ElseIf Len(.Cells(i, l)) > 0 And Len(.Cells(i, k)) = 0 Then
                        checkEmptyDD = True
                        Exit Function
                    End If
                End If
                i = i + 1
          Wend
  End With
End Function














