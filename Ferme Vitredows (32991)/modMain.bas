Attribute VB_Name = "modMain"
Option Explicit

'----------------------------------------------------------------------------------
' Source : http://www.vbfrance.com/code.aspx?ID=32991
' Auteur : Jack
'----------------------------------------------------------------------------------
' Si vous reprenez cette id�e pour votre projet, merci de m'en laisser le cr�dit.
'----------------------------------------------------------------------------------
' If you take theses effects for your own application, thanks to keep these credits
'----------------------------------------------------------------------------------

' ToDo :
' - Intercepter clavier pour d�tecter Echap -> End
' - Module de config pour choix couleur de la transparence
'                         choix degr� de transparence
' - Une fois le programme termin�, envoyer ordre de veille � l'�cran


' --- Gestion temps �coul�
Private Declare Function GetTickCount Lib "kernel32" () As Long
' --- Gestion du son
Private Declare Function PlaySoundmem Lib "winmm.dll" Alias "PlaySoundA" ( _
                                    ByVal lpszName As Long, _
                                    ByVal hModule As Long, _
                                    ByVal dwFlags As Long) As Long
Private Const SND_SYNC = &H0        ' Attend que le son soit jou� pour revenir
Private Const SND_ASYNC = &H1       ' D�marre le son et reviens
Private Const SND_NODEFAULT = &H2   ' Si probl�me, n'�mettra pas de bip
Private Const SND_MEMORY = &H4      ' Le son est en m�moire
Private Const SND_LOOP = &H8        ' Joue en boucle (arr�t = sndPlaySound(Null, SND_SYNC)
Private Const SND_NOSTOP = &H10     ' N'interrompt pas le son en cours
Private Const SND_NOWAIT = &H2000
'

Private Sub main()
    
    ' On va g�rer ici l'application, la forme n'�tant qu'un support graphique

    Dim maForme As frmVitre, Result As Long
    Dim EcranHauteur As Long, EcranLargeur As Long
    Dim Son() As Byte, Opacit� As Long, Incr�ment As Single
    Dim ChronoStart As Long, ChronoPass� As Long
    Dim Dur�eSon As Long, SonOk As Boolean

    ' Degr� de transparence de la forme (0 clair, 255 sombre)
    Opacit� = 155
    
    '---------- On charge la forme
    Set maForme = New frmVitre
    Load maForme
    ' Pour l'instant, forme pas visible car pas Show
    
    ' Positionne la forme en bas de l'�cran sur toute sa largeur
    ' Dimension de l'�cran principal
    EcranLargeur = Screen.Width
    EcranHauteur = Screen.Height
    ' Notre forme en bas de l'�cran, sur toute sa largeur
    With maForme
        .Left = 0
        .Top = EcranHauteur ' en bas de l'�cran
        .Width = EcranLargeur
        .Height = EcranHauteur  ' On dessine en dessous de l'�cran, pas grave
    End With
    maForme.Show    ' Rend la forme visible (dimensions hors �cran)
    ' Rend la forme transparente
    Call Transparence("ON", maForme, Opacit�)
    
    
    '---------- Synchronisation image et son
    ' La dur�e du fichier son est de 3 secondes et quelques
    ' Il faut que la forme parte du bas de l'�cran jusqu'au sommet
    '   en 3 secondes aussi.
    ' Pour �tre pr�cis, on va chronom�trer en milli�mes de secondes
    '   le temps qui passe, et donc on connaitra le temps de Son qu'il
    '   nous reste pour arriver en haut.
    ' Une r�gle de trois et on saura de combien de twips il faut monter
    
    '-- Son initial (pour rigoler)
    ' Les fichiers sons sont dans le fichier de ressources
    ' Extraction pr�vention, lol
    ' Dans ce cas, on attend qu'il ait fini de jouer avant de poursuivre
    Son = LoadResData(4012, "ATTENTION")
    SonOk = True ' par d�faut
    Result = PlaySoundmem(VarPtr(Son(0)), 0, SND_NOWAIT Or SND_NODEFAULT Or SND_MEMORY)
    ' Si impossible d'envoyer le son, pas la peine d'essayer plus tard
    If Result = 0 Then SonOk = False
    ' Attend une demie-seconde
    ChronoStart = GetTickCount
    Do While (GetTickCount - ChronoStart) < 500
        DoEvents
    Loop
    
    
    ' -- Son vitre qui monte : On ne fait que le lancer sans attendre apr�s
    Dur�eSon = 3500 ' Si vous changer le son, changez aussi la dur�e ici
    If SonOk Then
        Son = LoadResData(4013, "LEVAGE_VITRE")
        Call PlaySoundmem(VarPtr(Son(0)), 0, SND_NOWAIT Or SND_NODEFAULT Or SND_MEMORY Or SND_ASYNC)
        DoEvents
    End If
    ' Le son est parti : Lance le Chrono
    ChronoStart = GetTickCount
    
    
    ' On a 3000 millisecondes pour se d�placer de EcranHauteur
    ' On va r�p�ter le calcul pendant toute la dur�e du son
    Do While GetTickCount - ChronoStart < Dur�eSon
        ' Calcule � quelle hauteur on devrait �tre en fonction du temps
        ' Temps �coul�
        ChronoPass� = Dur�eSon - (GetTickCount - ChronoStart)
        If ChronoPass� <= 0 Then ChronoPass� = 0
        ' On positionne la forme en fonction du temps �coul�
        maForme.Top = EcranHauteur * ChronoPass� / Dur�eSon
        DoEvents
        DoEvents
        DoEvents
    Loop
    maForme.Top = 0
    ' -- Une petite pause
    ChronoStart = GetTickCount
    Do While (GetTickCount - ChronoStart) < 500
        DoEvents
    Loop
    
    
    ' -- Dernier son : l'alarme (celle-l� elle me fait bien marrer)
    If SonOk Then
        Son = LoadResData(4014, "MOUIP_MOUIP")
        Call PlaySoundmem(VarPtr(Son(0)), 0, SND_NOWAIT Or SND_NODEFAULT Or SND_MEMORY Or SND_ASYNC)
        DoEvents
    End If
    ReDim Son(0)    ' Vide la variable, plus besoin
    
    
    ' ---------- On fini par noicir compl�tement l'�cran en 3 sec
    ' Chez moi, le PC met 11700 millisecondes pour ex�cuter 100 incr�ments d'opacit�
    ' Si je veux que tout devienne noir en 3 secondes, il faut incr�menter de :
    Incr�ment = (11700! / 3000!) * (100! / (255 - Opacit�))
    Do While Opacit� < 255
        ' Augmente l'opacit� jusqu'� devenir noir
        Opacit� = Opacit� + Incr�ment
        Opacit� = IIf(Opacit� > 255, 255, Opacit�)
        Call Transparence("ON", maForme, CByte(Opacit�))
        ' -- Une petite pause
        ChronoStart = GetTickCount
        Do While (GetTickCount - ChronoStart) < 30
            DoEvents
        Loop
    Loop
    ' -- Une petite pause
    ChronoStart = GetTickCount
    Do While (GetTickCount - ChronoStart) < 500
        DoEvents
    Loop
    
CestLaFin:
    ' Supprime la transparence
    Call Transparence("OFF", maForme, 0)
    DoEvents
    ' Supprime notre forme ador�e
    Unload maForme
    Set maForme = Nothing
    End

End Sub
