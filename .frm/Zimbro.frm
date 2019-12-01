VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Zimbro 
   Caption         =   "Zimbro"
   ClientHeight    =   9630.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6810
   OleObjectBlob   =   "Zimbro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Zimbro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer
Dim box As Object


Private Sub UserForm_Initialize()
    i = 0
    Pic.Picture = LoadPicture(wkb.Path & "\pics\zimbro.jpg")
'cette boucle fait passer tout les éléments d'un UserForm et, si c'est un TextBox, associe un des mails présent dans la feullie "Data" à la valeur de la TextBox
    For Each box In Me.Controls
        If TypeName(box) = "TextBox" Then
            box.Font.Size = 14
            box.Value = Range("Data!j" & 3 + i).Value '
            i = i + 1
        End If
    Next
    Zimbro.BackColor = RGB(255, 106, 0) ' utilisation de la couleur RGB pour avoir la même couleur d'arrière plan que l'image de Zimbro
    Unread.BackColor = RGB(255, 106, 0)
    Unread.Caption = "vous avez " & Range("Data!j2") & " nouveaux messages"
End Sub

Sub update()
    Call UserForm_Initialize
End Sub


Private Sub updt_Click()
    Call update
End Sub

