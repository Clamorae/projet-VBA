VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9015.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Activate()
     While UserForm1.Visible = True
        DoEvents                                  'le DoEvents permet a l'UserForm de vérifier les events pour ne pas cracher dans la loop'
        test = test + 1
        CommandButton3.Caption = CStr(test)
    Wend
End Sub

Private Sub UserForm_Click()
    MsgBox "nigga"
End Sub

Sub UserForm_Initialize()
     Image1.Picture = LoadPicture("C:\Users\Paul\Desktop\title.gif")
     Dim test As Integer
End Sub



