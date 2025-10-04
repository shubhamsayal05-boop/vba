VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} defineTemp 
   Caption         =   "Température par défaut"
   ClientHeight    =   3435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5220
   OleObjectBlob   =   "defineTemp.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "defineTemp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandButton1_Click()
   If val(Me.tEMPS) > 0 Then
     Call Load_Data.MkTemp(val(Me.tEMPS))
     Unload Me
   Else
      MsgBox "Entrer une valeur correcte", vbCritical, "ODRIV"
   End If
End Sub
