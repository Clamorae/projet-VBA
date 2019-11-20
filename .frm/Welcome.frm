VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Welcome 
   Caption         =   "Bienvenue à l'UTBM !"
   ClientHeight    =   2895
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6750
   OleObjectBlob   =   "Welcome.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Welcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    If (TName.value = vbNullString Or TSurname.value = vbNullString) Then
        MsgBox "veulliez entrer un nom et prénom valide"
    Else
        name = TName.value
        If girl.value = True Then
            gender = "girl"
        Else
            gender = "boy"
        End If
    End If
End Sub

