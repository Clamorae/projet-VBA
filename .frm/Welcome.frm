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
    If (TName.Value = vbNullString Or TSurname.Value = vbNullString) Then 'Force le joueur à entrer un nom
        MsgBox "veulliez entrer un nom et prénom"
    Else
        Cname = TSurname.Value & " " & TName.Value
        If girl.Value = True Then
            Cgender = "girl"
        Else
            Cgender = "boy" 'asoocie "boy" ou "girl" à la variable Cgender en fonction du choix du joueur
        End If
        wkb.Sheets("Sevenans").Activate 'affiche la map de sevenans
        StartMessage.Show (0)
        Unload Me
    End If
End Sub

