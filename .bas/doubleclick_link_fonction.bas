Attribute VB_Name = "Module1"

Sub Worksheet_BeforeDoubleClick(ByVal OwO As Range, Cancel As Boolean)
    If OwO = "1" Then
        MsgBox ("UwU")
    ElseIf OwO = "2" Then
    ElseIf OwO = "3" Then
    ElseIf OwO = "4" Then
    ElseIf OwO = "5" Then
    ElseIf OwO = "6" Then
    ElseIf OwO = "7" Then
    Else
        MsgBox ("Ce n'est pas une case valide")
    End If
    Cancel = True
End Sub
