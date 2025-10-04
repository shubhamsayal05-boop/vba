VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Feuil4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Deactivate()
     On Error Resume Next
    Application.EnableEvents = False
    checkChange
    Application.EnableEvents = True
    On Error GoTo 0
End Sub
