VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} StartMessage 
   Caption         =   "/!\1 Nouveau Mail"
   ClientHeight    =   2775
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6690
   OleObjectBlob   =   "StartMessage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "StartMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub UserForm_Initialize()
    Text.Caption = Replace(Text, "NAME", Cname) 'ajoute le nom du joueur dans l'UserForm
End Sub

Private Sub CommandButton1_Click()
    Main.Show (0)
    Unload Me
End Sub
