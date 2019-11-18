VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "Bienvenue sur UTB'Quest"
   ClientHeight    =   7155
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9135.001
   OleObjectBlob   =   "Menu.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub UserForm_Initialize()
     Image1.Picture = LoadPicture(ActiveWorkbook.Path & "\pics\title.gif") 'permet de charger une image présente dans le fichier du document'
End Sub

Private Sub CommandButton3_Click() 'Permet de quitter le programme'
    If MsgBox("Etes-vous sur ?", vbYesNo, "Confirmation") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub CommandButton2_Click()
    If MsgBox("placeholder", vbOKOnly, "Infos") = vbOK Then
    End If
End Sub

