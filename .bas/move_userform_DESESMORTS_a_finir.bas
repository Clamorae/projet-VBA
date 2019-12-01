Attribute VB_Name = "Module1"
Sub Userform_initialize()
    While 1 = 1
        DoEvents
        Call aleatoire
    Wend
End Sub

Sub aleatoire()
    Randomize
    Rng = Int(5 * Rnd) + 1
    If Rng = 1 Then
    Call centerr
    ElseIf Rng = 2 Then
    Call downn
    ElseIf Rng = 3 Then
    Call upp
    ElseIf Rng = 4 Then
    Call rightt
    Else
    Call leftt
    End If
End Sub
Sub moveform(ByVal v As Integer, ByVal h As Integer)
    With Me
        .Top = .Top + v
        .left = .left + h
    End With
End Sub

Private Sub centerr()
With Me
    .StartUpPosition = 0
    .left = Application.left + (0.5 * Application.Width) - (0.5 * .Width)
    .Top = Application.Top + (0.5 * Application.Height) - (0.5 * .Height)
End With
End Sub
 Private Sub downn()
    Call moveform(0, -25)
 End Sub
 Private Sub upp()
    Call moveform(0, 25)
 End Sub
 Private Sub rightt()
    Call moveform(25, 0)
 End Sub
 Private Sub leftt()
    Call moveform(25, 0)
 End Sub

