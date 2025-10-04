VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ediitSDVName 
   Caption         =   "Settings"
   ClientHeight    =   3270
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4605
   OleObjectBlob   =   "ediitSDVName.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ediitSDVName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
       Dim i As Long
       Dim j As Long
       On Error GoTo Es
       Application.EnableEvents = False
       Application.ScreenUpdating = False
        With ThisWorkbook.Worksheets("SDV Manager")
                i = .Range("A65000").End(xlUp).row
                If Len(Me.TextBox2.Value) > 0 Then
                    For j = 2 To i
                      If UCase(.Range("A" & j).Value) = UCase(nom) Then
                          MsgBox "Nom D�j� Attribu�", vbCritical, "ODRIV"
                          Exit Sub
                      End If
                    Next j
                    
                    .Rows(i + 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                    .Rows(i).Copy Destination:=.Range("A" & i + 1)
                    .Range("A" & i + 1) = Me.TextBox2.Value
                    If InStr(1, .Range("b" & i + 1), ".") = 0 Then
                        .Range("a" & i + 1 & ":" & "b" & i + 1).Interior.color = RGB(255, 255, 255)
                        .Range("b" & i + 1) = .Range("b" & i + 1) & ".1"
                    Else
                        .Range("b" & i + 1) = Split(.Range("b" & i + 1), ".")(0) & "." & val(Split(.Range("b" & i + 1), ".")(1)) + 1
                    End If
                    Call CreateNew.NewSDVCalcul(Me.TextBox2.Value)
                    Call CreateNew.NewSDVStructure(Me.TextBox2.Value)
                    Call CreateNew.NewSDVConfigurationSetting(Me.TextBox2.Value)
                    Call CreateNew.NewSDVDefinitionSDV(Me.TextBox2.Value)
                    Call CreateNew.NewSDVPowertrain(Me.TextBox2.Value)
                    Call CreateNew.NewSDVRating(Me.TextBox2.Value)
                    Call CreateNew.NewSDVSeetings(Me.TextBox2.Value)
                    Call convertGraph.CreateNew(Me.TextBox2.Value)
                    MsgBox "Op�ration R�ussie", vbInformation, "ODRIV"
                    Unload Me
                    
                Else
                      MsgBox "Nom Vide", vbCritical, "ODRIV"
                End If
        End With
       Application.EnableEvents = True
       Application.ScreenUpdating = True
       
Es:
       If ERR.Number <> 0 Then
            MsgBox ERR.description
              Application.EnableEvents = True
              Application.ScreenUpdating = True
       End If
     
End Sub

