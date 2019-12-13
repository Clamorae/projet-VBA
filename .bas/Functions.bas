Attribute VB_Name = "Functions"
Option Explicit

'Importation de la DLL "winmm.dll" qui permet la lecture de fichier audio via la fonction "sndPlaySound"
Public Declare PtrSafe Function Playsound _
Lib "winmm.dll" _
Alias "sndPlaySoundA" _
(ByVal Path As String, ByVal Flags As Long) As Long
'Ici, Playsound coresspond au nom donn� � la donction import�e de la DLL (Dynamic Link Librairy) _
PtrSafe permet la compatibilit� avec les syst�mes 64 bits _
winmm.dll est le nom de la DLL dans laquelle la fonction est situ�e. _
sndPlaySoundA est le nom de la fonction dans la DLL. _
Enfin, Path et Flags sont les deux arguments pris par la fonction, Path �tant le chemin du fichier et Flags permettant de pr�ciser la mani�re _
avec laquelle jouer le fichier (le &H1 utilis� dans le code permet de jouer la musique pendant l'�x�cution du programme, _
l'argument &H0 pause le programme jusqu'a la fin du fichier audio) _
Plus de d�tails : https://docs.microsoft.com/en-us/previous-versions/dd798676(v%3Dvs.85)


Sub PlayMusic(Name As String) 'cette fonction permet de jouer le fichier .wav situ� dans le dossier sounds du projet dont le nom est pr�cis� en argument'
    Call Playsound(wkb.Path & "\sounds\" & Name, &H1)
End Sub

Sub StopMusic() 'cette fonction permet d'arr�ter la musique. En effet, un "Path" ayant pour valeur "Null" entra�ne l'arr�t de la musique dans la fonction sndPlaySoundA'
    Call Playsound(vbNullString, &H1)
End Sub

Sub battle_tendency()
    Call assignVariables
    mstHP = CInt("69")
    mstHPmax = 69
    mstAtk = 2
    mstSpeed = 1
    mstDef = 0
    Battle.Show (0)
End Sub

