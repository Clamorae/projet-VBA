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

Sub UserForm_Initialize() 'Sub qui se lance lorsque l'UserForm est initialisé
    Image1.Picture = LoadPicture(wkb.Path & "\pics\title.gif") 'permet de charger l'image title.gif, wkb.Path étant le chemin du classeur, donc du fichier Excel
    Call PlayMusic("jojo.wav") 'voir le module Functions
End Sub

Private Sub CommandButton3_Click() 'Permet de quitter le programme'
    If MsgBox("Etes-vous sur ?", vbYesNo, "Confirmation") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub CommandButton2_Click()
    If MsgBox("on verras après", vbOKOnly, "Infos") = vbOK Then 'lance un MsgBox donnant des infos sur le jeu'
    End If
End Sub

Private Sub CommandButton1_Click()
    Welcome.Show (0) 'l'agument 0 permet à plusieurs UserForms d'être affiché à la fois.
    Unload Me
End Sub

