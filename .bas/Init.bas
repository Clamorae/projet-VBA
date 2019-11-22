Attribute VB_Name = "Init"
Sub init()
    Call assignVariables
    Menu.Show (0)                  'ici, le 0 sert a lancer la fenetre en mode non modal'
End Sub

Sub onGameLaunch()
    ThisWorkbook.Sheets("Sevenans").Activate
    'ici on va assigner la plupart des variables'
End Sub

