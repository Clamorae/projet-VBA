Attribute VB_Name = "Variables"
Option Explicit
'Valeur int----------------------------------------------'

    Public hp As Single         'd�finit la variable de la vie du joueur
    Public maxHp As Integer
    
    Public xp As Integer         'definit la variable de credits, qui correspond � l'exp�rience du joueur
    Public maxXp As Integer
    
    
    Public atk As Integer             'definit la variable de l'attaque de joueur
    
    
    Public def As Integer             'definit la variable de l'attaque de joueur
    

'Valeur str----------------------------------------------'
    Public Cname As String 'le nom seras assign� lors de la cr�ation de perso'
    Public Cgender As String

'variables autres----------------------------------------'
    Public wkb As Workbook
    
    Public mstHP As Single
    Public mstHPmax As Integer
    Public mstStrength As Integer
    Public mstSpeed As Single
    Public mstAtk As Integer
    Public mstDef As Integer


Sub assignVariables() 'assignation des valeurs aux variables'
    hp = 100
    maxHp = 100
    xp = 0
    maxXp = 10
    atk = 10
    def = 0
    Set wkb = ThisWorkbook 'permet de remplacer le ThisWorkbook par wkb pour plus de simplicit�(le set est n�c�ssaire car il s'agit d'un objet et non d'une variable)'
    Battle.DrawBuffer = 64000
End Sub
