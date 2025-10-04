VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} unlocksheet 
   Caption         =   "SELECTION"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6345
   OleObjectBlob   =   "unlocksheet.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "unlocksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CommandButton1_Click()
    
        
            If UCase(Me.TextBox2.Value) = "UNLOCK" Then
                    If SelectionActuel = False Then
                        MsgBox "COCHER POUR SELECTIONNER", vbCritical, "ODRIV"
                    Else
                        Unload Me
                    End If
            Else
                    MsgBox "MOT DE PASSE INCORRECT", vbCritical, "ODRIV"
            End If
      
End Sub



Private Sub ListBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
       
        
End Sub



Private Sub ListeValeurS_ItemClick(ByVal Item As MSComctlLib.ListItem)
        Me.ListeValeurS.ListItems(Item.index).Selected = False
End Sub

Private Sub UserForm_Initialize()
        With Me.ListeValeurS
               .View = lvwReport   'affichage en mode Rapport
               .Gridlines = True   'affichage d'un quadrillage
               .FullRowSelect = False   'Sélection des lignes comlètes
               .LabelEdit = lvwManual  'desactive edition du listview
               .MultiSelect = True
               .HideSelection = True
               .HoverSelection = False
               .CheckBoxes = True
        End With
       
        Call loadME
       
End Sub

Function loadME()
    Dim tables(13) As String
    Dim i As Integer
    
    tables(1) = "CONFIGURATIONS SEETINGS"
    tables(2) = "SETTINGS"
    tables(3) = "structure"
    tables(4) = "POWERTRAIN"
    tables(5) = "CONFIGURATIONS"
'    TabLes(6) = "COVERAGE Rate"
    tables(6) = "Target vehicle"
    tables(7) = "TARGETS"
    tables(8) = "Calculs"
    tables(9) = "Graph_status"
    tables(10) = "DEFINITION SDV"
    tables(11) = "PARAMETRES GRAPH"
    tables(12) = "ENTETE_COLONNE"
    tables(13) = "SDV MANAGER"
    
    For i = 1 To UBound(tables)
             If sheets(tables(i)).Visible = 2 Then
                   Me.ListeValeurS.ListItems.Add , , UCase(tables(i))
             End If
    Next i

   
End Function
Function SelectionActuel() As Boolean
Dim i As Long
 SelectionActuel = False
 
For i = 1 To Me.ListeValeurS.ListItems.Count
        If Me.ListeValeurS.ListItems(i).Checked = True Then
             sheets(Me.ListeValeurS.ListItems(i).text).Visible = -1
             SelectionActuel = True
        End If
Next i
End Function

