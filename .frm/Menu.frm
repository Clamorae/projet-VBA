VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Menu 
   Caption         =   "Bienvenue sur UTB'Quest"
   ClientHeight    =   9420.001
   ClientLeft      =   135
   ClientTop       =   465
   ClientWidth     =   15915
   OleObjectBlob   =   "Menu.frx":0000
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub UserForm_Initialize() 'Sub qui se lance lorsque l'UserForm est initialis�
    Application.DisplayFullScreen = True
     With Me
        .Width = Application.Width
        .Height = Application.Height
    End With
    With Image1
    .Height = Application.Height
    .Width = Application.Width - CommandButton1.Width
    .Picture = LoadPicture(wkb.Path & "\pics\title.jpg") 'permet de charger l'image title.gif, wkb.Path �tant le chemin du classeur, donc du fichier Excel
    End With
    Call PlayMusic("title.wav") 'voir le module Functions
End Sub

Private Sub CommandButton3_Click() 'Permet de quitter le programme'
    If MsgBox("Etes-vous sur ?", vbYesNo, "Confirmation") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub CommandButton2_Click()
    If MsgBox("on verras apr�s", vbOKOnly, "Infos") = vbOK Then 'lance un MsgBox donnant des infos sur le jeu'
    End If
End Sub

Private Sub CommandButton1_Click()
    Welcome.Show (0) 'l'agument 0 permet � plusieurs UserForms d'�tre affich� � la fois.
    Unload Me
End Sub

