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
Dim line1, line2, line3, line4, line5 As String
Dim i As Integer
Dim playerDamage As Integer
Dim mstDamage As Integer
Dim isdefending As Integer

Sub UserForm_Activate()
    PlayMusic ("battle.wav")
    BAtk.BackColor = RGB(200, 200, 200)
    BSkill.BackColor = RGB(200, 200, 200)
    Bdeff.BackColor = RGB(200, 200, 200)
    logInit
    addToLog ("Vous êtes attaqué par: " & mstName)
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
        goodEnd = timeKeeper + speed
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
        BAtk.BackColor = RGB(200, 200, 200)
        BSkill.BackColor = RGB(200, 200, 200)
        Bdeff.BackColor = RGB(200, 200, 200)
    End If
    
    For i = 1 To badUnit:
        BadTimeBar.Caption = BadTimeBar.Caption + "_"
    Next
    
    For i = 1 To goodUnit:
        GoodTimeBar.Caption = GoodTimeBar.Caption + "_"
    Next

    
End Sub


Private Sub BAtk_Click()
    If isdefending = 0 Then
        isattacking = True
        goodEnd = timeKeeper + 0.1
        BAtk.BackColor = RGB(0, 150, 200)
        BSkill.BackColor = RGB(100, 100, 100)
        Bdeff.BackColor = RGB(100, 100, 100)
    End If
End Sub

Private Sub Bdeff_Click()
    If isattacking = False Then
        If isdefending = 0 Then
            BAtk.BackColor = RGB(100, 100, 100)
            BSkill.BackColor = RGB(100, 100, 100)
            Bdeff.BackColor = RGB(0, 200, 0)
            isdefending = 1
        Else
            BAtk.BackColor = RGB(200, 200, 200)
            BSkill.BackColor = RGB(200, 200, 200)
            Bdeff.BackColor = RGB(200, 200, 200)
            isdefending = 0
        End If
    End If
End Sub

Sub atkMst()
    mstDamage = CInt(atk - mstDef + (Rnd * atk))
    If mstHP > mstHP - mstDamage Then
        mstHP = mstHP - mstDamage
        BadHP.Caption = mstHP & "/" & mstHPmax
        addToLog ("*SWIING*Vous faites " & mstDamage & " de dégât !")
        BadHPBar = ""
        For i = 1 To CInt(mstHP * 10 / mstHPmax)
            BadHPBar.Caption = BadHPBar.Caption + "_"
        Next
        If mstHP <= 0 Then
            BadHP.Caption = 0 & "/" & mstHPmax
            MsgBox "Vous avez battu l'adversaire, félicitation!"
            Unload Me
            End
        End If
    Else
        addToLog ("Vous attaquez mais, *CRACK*, vous tombez et votre attaque rate")
    End If
End Sub

Sub atkPlayer()
    playerDamage = CInt(mstAtk - (def * isdefending) + (Rnd * mstAtk))
    If hp > hp - playerDamage Then
        hp = hp - playerDamage
        GoodHP.Caption = hp & "/" & maxHp
        addToLog ("*BAM*Vous perdez " & playerDamage & " points de vie !")
        GoodHPbar = ""
        For i = 1 To CInt(hp * 16 / maxHp)
            GoodHPbar.Caption = GoodHPbar.Caption + "_"
            Next
        If hp <= 0 Then
            GoodHP.Caption = 0 & "/" & maxHp
            If MsgBox _
            ("Malheuresement, vous avez perdu et n'avez pas réussi a venir a bout de ce devoir.Mais bon, il n'est pas trop tard pour ce reprendre ! Voulez-vous tentez encore une fois ?", vbYesNo, "GAME OVER") = vbNo Then
                MsgBox "Très bien. Merci d'avoir joué !"
                Application.Quit
            Else
                Unload Me
                MsgBox "test" 'redemarrez le combat
            End If
            Unload Me
            End
        End If
    Else
        addToLog ("l'ennemi attaque, mais vous l'évitez avec agilité !")
    End If
End Sub


Sub logInit()
    line1 = ""
    line2 = ""
    line3 = ""
    line4 = ""
    line5 = ""
    Log.Value = line1 & vbCrLf & line2 & vbCrLf & line3 & vbCrLf & line4 & vbCrLf & line5
End Sub

Sub addToLog(Text As String)
    line1 = line2
    line2 = line3
    line3 = line4
    line4 = line5
    line5 = Text
    Log.Value = line1 & vbCrLf & line2 & vbCrLf & line3 & vbCrLf & line4 & vbCrLf & line5
End Sub

