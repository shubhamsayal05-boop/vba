VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataEdit 
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8025
   OleObjectBlob   =   "DataEdit.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DataEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub CommandButton6_Click()
    Dim i As Integer
    Dim li As Variant
    i = val(DataLoad.ids.Caption)
    If i <> 0 Then
        If Me.ComboBox2 Like "*SUPERIEUR*" Or Me.ComboBox2 Like "*INFERIEUR*" Then
              If IsNumeric(Me.PValeur) = False Then
                        MsgBox "VALEUR NUMERIQUE", vbCritical, "ODRIV"
                        Exit Sub
             End If
        End If
        If Len(Me.Pcolonne.Value) > 0 And Len(Me.PValeur.Value) > 0 And Len(Me.ComboBox2.Value) > 0 And Len(Me.POrdre.Value) > 0 Then
                With DataLoad
                    .ListView11.ListItems(i).text = Me.Pcolonne.Value
                    .ListView11.ListItems(i).ListSubItems(1).text = Me.ComboBox2
                    .ListView11.ListItems(i).ListSubItems(2).text = Me.PValeur.Value
                    .ListView11.ListItems(i).ListSubItems(3).text = Me.POrdre.Value
                    On Error Resume Next
                    DataLoad.ListView11.SelectedItem.Selected = False
                    Unload Me
                End With
        Else
                MsgBox "SELECTION", vbCritical, "ODRIV"
        End If
    Else
         If Me.ComboBox2 Like "*SUPERIEUR*" Or Me.ComboBox2 Like "*INFERIEUR*" Then
              If IsNumeric(Me.PValeur) = False Then
                        MsgBox "VALEUR NUMERIQUE", vbCritical, "ODRIV"
                        Exit Sub
             End If
        End If
         If Len(Me.Pcolonne.Value) > 0 And Len(Me.PValeur.Value) > 0 And Len(Me.ComboBox2.Value) > 0 And Len(Me.POrdre.Value) > 0 Then
                With DataLoad
                    Set li = .ListView11.ListItems.Add(, , Me.Pcolonne)
                    li.SubItems(1) = Me.ComboBox2.Value
                    li.SubItems(2) = Me.PValeur.Value
                    li.SubItems(3) = Me.POrdre.Value
                    On Error Resume Next
                   DataLoad.ListView11.SelectedItem.Selected = False
                    Unload Me
                End With
        Else
                MsgBox "SELECTION", vbCritical, "ODRIV"
        End If
    End If
End Sub

Private Sub CommandButton7_Click()
    Me.code.Caption = "1"
    COLONNES.Show
End Sub

Private Sub CommandButton8_Click()
  Me.code.Caption = "2"
  COLONNES.Show
End Sub

Private Sub UserForm_Initialize()
    Me.ComboBox2.AddItem "EGAL A"
    Me.ComboBox2.AddItem "EGALITE AVEC"
    Me.ComboBox2.AddItem "EGALITE AVEC (SI VALEUR NON VIDE)"
    Me.ComboBox2.AddItem "INFERIEUR A"
    Me.ComboBox2.AddItem "INFERIEUR OU EGAL A"
    Me.ComboBox2.AddItem "SUPERIEUR A"
    Me.ComboBox2.AddItem "SUPERIEUR OU EGAL A"
    Me.ComboBox2.AddItem "DIFFERENT DE"
    Me.ComboBox2.AddItem "CONTIENT"
    Me.ComboBox2.AddItem "NE CONTIENT PAS"

End Sub

Private Sub UserForm_Terminate()
   On Error Resume Next
  DataLoad.ListView11.SelectedItem.Selected = False
End Sub
