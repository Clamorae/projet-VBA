Attribute VB_Name = "Variables"
Option Explicit
'Valeur int----------------------------------------------'

    Public hp As Integer          'd�finit la variable de la vie du joueur

    
    Public xp As Integer         'definit la variable de credits, qui correspond � l'exp�rience du joueur
    
    
    
    Public atk As Integer             'definit la variable de l'attaque de joueur
    
    
    Public def As Integer             'definit la variable de l'attaque de joueur
    

'Valeur str----------------------------------------------'
    Public Cname As String 'le nom seras assign� lors de la cr�ation de perso'
    Public Cgender As String

'variables autres----------------------------------------'
    Public wkb As Workbook


Sub assignVariables() 'assignation des valeurs aux variables'
    hp = 100
    xp = 0
    atk = 0
    def = 0
    Set wkb = ThisWorkbook 'permet de remplacer le ThisWorkbook par wkb pour plus de simplicit�(le set est n�c�ssaire car il s'agit d'un objet et non d'une variable)'

End Sub
