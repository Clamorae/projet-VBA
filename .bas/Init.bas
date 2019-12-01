Attribute VB_Name = "Init"
Sub init() 'Première fontion à se lancer, elle met le choses en place pour le lancement du jeu, puis lance le menu'
    Call assignVariables 'apelle la fonction qui va assigner à certaines variables des valeurs spécifique'
    CommandBars.ExecuteMso "HideRibbon" 'cette commande ferme la barre de comande d'excel pour plus d'immersion'
    ThisWorkbook.Sheets("Title Screen").Activate
    Menu.Show (0) 'ici, le 0 sert a lancer la fenetre en mode non modal'
End Sub


