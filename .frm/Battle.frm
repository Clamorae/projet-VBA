VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Battle 
   Caption         =   "On vous attaque !"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8040
   OleObjectBlob   =   "Battle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Battle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim time As Single
Dim badUnit As Integer
Dim badEnd As Single
Dim isattacking As Boolean
Dim goodUnit As Integer
Dim goodEnd As Single
Dim timeKeeper As Single
Dim infight As Boolean
Dim i As Integer


Sub UserForm_Activate()
    MsgBox mstHP
    BadHP.Caption = mstHP & "/" & mstHPmax
    GoodHP.Caption = hp & "/" & maxHp
    timeKeeper = Timer
    badEnd = timeKeeper + mstSpeed
    badUnit = 0
    goodUnit = 0
    infight = 1
    While infight:
        timeKeeper = Timer
        DoEvents
        Call mainloop
        Wend
End Sub


Sub mainloop()
    If timeKeeper >= badEnd Then
        badUnit = badUnit + 1
        badEnd = timeKeeper + mstSpeed
    End If
    
    If timeKeeper >= goodEnd And isattacking = True Then
        goodUnit = goodUnit + 1
        goodEnd = timeKeeper + 0.1
    End If

'cette section met à jour les barres de progression et arrete l'attaque du joueur l'orsque la bar arrive à sa fin'
    BadTimeBar.Caption = ""
    GoodTimeBar.Caption = ""
    If badUnit = 11 Then
    Call atkPlayer
        badUnit = 0
    End If
    
    If goodUnit = 16 Then
        Call atkMst
        isattacking = False
        goodUnit = 0
    End If
    
    For i = 1 To badUnit:
        BadTimeBar.Caption = BadTimeBar.Caption + "_"
    Next
    
    For i = 1 To goodUnit:
        GoodTimeBar.Caption = GoodTimeBar.Caption + "_"
    Next

    
End Sub


Private Sub BAtk_Click()
    isattacking = True
    goodEnd = timeKeeper + 0.1
End Sub

Sub atkMst()
    mstHP = CInt(mstHP - atk + mstDef - (Rnd * atk))
    BadHP.Caption = mstHP & "/" & mstHPmax
    BadHPBar = ""
    For i = 1 To CInt(mstHP * 10 / mstHPmax)
        BadHPBar.Caption = BadHPBar.Caption + "_"
    Next
    If mstHP <= 0 Then
        GoodHP.Caption = 0 & "/" & mstHPmax
        MsgBox "he ded gg wp ez"
        End
        Unload Me
    End If
End Sub

Sub atkPlayer()
    hp = CInt(hp - mstAtk + def - (Rnd * mstAtk))
    GoodHP.Caption = hp & "/" & maxHp
    GoodHPbar = ""
    For i = 1 To CInt(hp * 16 / maxHp)
        GoodHPbar.Caption = GoodHPbar.Caption + "_"
        Next
    If hp <= 0 Then
        GoodHP.Caption = 0 & "/" & maxHp
        MsgBox "malheuresement, un lama vous as écrasé, sheh"
        Application.Quit
    End If
End Sub



