Attribute VB_Name = "modMain"
Option Explicit

'----------------------------------------------------------------------------------
' Source : http://www.vbfrance.com/code.aspx?ID=32991
' Auteur : Jack
'----------------------------------------------------------------------------------
' Si vous reprenez cette idée pour votre projet, merci de m'en laisser le crédit.
'----------------------------------------------------------------------------------
' If you take theses effects for your own application, thanks to keep these credits
'----------------------------------------------------------------------------------

' ToDo :
' - Intercepter clavier pour détecter Echap -> End
' - Module de config pour choix couleur de la transparence
'                         choix degré de transparence
' - Une fois le programme terminé, envoyer ordre de veille à l'écran


' --- Gestion temps écoulé
Private Declare Function GetTickCount Lib "kernel32" () As Long
' --- Gestion du son
Private Declare Function PlaySoundmem Lib "winmm.dll" Alias "PlaySoundA" ( _
                                    ByVal lpszName As Long, _
                                    ByVal hModule As Long, _
                                    ByVal dwFlags As Long) As Long
Private Const SND_SYNC = &H0        ' Attend que le son soit joué pour revenir
Private Const SND_ASYNC = &H1       ' Démarre le son et reviens
Private Const SND_NODEFAULT = &H2   ' Si problème, n'émettra pas de bip
Private Const SND_MEMORY = &H4      ' Le son est en mémoire
Private Const SND_LOOP = &H8        ' Joue en boucle (arrêt = sndPlaySound(Null, SND_SYNC)
Private Const SND_NOSTOP = &H10     ' N'interrompt pas le son en cours
Private Const SND_NOWAIT = &H2000
'

Private Sub main()
    
    ' On va gérer ici l'application, la forme n'étant qu'un support graphique

    Dim maForme As frmVitre, Result As Long
    Dim EcranHauteur As Long, EcranLargeur As Long
    Dim Son() As Byte, Opacité As Long, Incrément As Single
    Dim ChronoStart As Long, ChronoPassé As Long
    Dim DuréeSon As Long, SonOk As Boolean

    ' Degré de transparence de la forme (0 clair, 255 sombre)
    Opacité = 155
    
    '---------- On charge la forme
    Set maForme = New frmVitre
    Load maForme
    ' Pour l'instant, forme pas visible car pas Show
    
    ' Positionne la forme en bas de l'écran sur toute sa largeur
    ' Dimension de l'écran principal
    EcranLargeur = Screen.Width
    EcranHauteur = Screen.Height
    ' Notre forme en bas de l'écran, sur toute sa largeur
    With maForme
        .Left = 0
        .Top = EcranHauteur ' en bas de l'écran
        .Width = EcranLargeur
        .Height = EcranHauteur  ' On dessine en dessous de l'écran, pas grave
    End With
    maForme.Show    ' Rend la forme visible (dimensions hors écran)
    ' Rend la forme transparente
    Call Transparence("ON", maForme, Opacité)
    
    
    '---------- Synchronisation image et son
    ' La durée du fichier son est de 3 secondes et quelques
    ' Il faut que la forme parte du bas de l'écran jusqu'au sommet
    '   en 3 secondes aussi.
    ' Pour être précis, on va chronométrer en millièmes de secondes
    '   le temps qui passe, et donc on connaitra le temps de Son qu'il
    '   nous reste pour arriver en haut.
    ' Une règle de trois et on saura de combien de twips il faut monter
    
    '-- Son initial (pour rigoler)
    ' Les fichiers sons sont dans le fichier de ressources
    ' Extraction prévention, lol
    ' Dans ce cas, on attend qu'il ait fini de jouer avant de poursuivre
    Son = LoadResData(4012, "ATTENTION")
    SonOk = True ' par défaut
    Result = PlaySoundmem(VarPtr(Son(0)), 0, SND_NOWAIT Or SND_NODEFAULT Or SND_MEMORY)
    ' Si impossible d'envoyer le son, pas la peine d'essayer plus tard
    If Result = 0 Then SonOk = False
    ' Attend une demie-seconde
    ChronoStart = GetTickCount
    Do While (GetTickCount - ChronoStart) < 500
        DoEvents
    Loop
    
    
    ' -- Son vitre qui monte : On ne fait que le lancer sans attendre après
    DuréeSon = 3500 ' Si vous changer le son, changez aussi la durée ici
    If SonOk Then
        Son = LoadResData(4013, "LEVAGE_VITRE")
        Call PlaySoundmem(VarPtr(Son(0)), 0, SND_NOWAIT Or SND_NODEFAULT Or SND_MEMORY Or SND_ASYNC)
        DoEvents
    End If
    ' Le son est parti : Lance le Chrono
    ChronoStart = GetTickCount
    
    
    ' On a 3000 millisecondes pour se déplacer de EcranHauteur
    ' On va répéter le calcul pendant toute la durée du son
    Do While GetTickCount - ChronoStart < DuréeSon
        ' Calcule à quelle hauteur on devrait être en fonction du temps
        ' Temps écoulé
        ChronoPassé = DuréeSon - (GetTickCount - ChronoStart)
        If ChronoPassé <= 0 Then ChronoPassé = 0
        ' On positionne la forme en fonction du temps écoulé
        maForme.Top = EcranHauteur * ChronoPassé / DuréeSon
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
    
    
    ' -- Dernier son : l'alarme (celle-là elle me fait bien marrer)
    If SonOk Then
        Son = LoadResData(4014, "MOUIP_MOUIP")
        Call PlaySoundmem(VarPtr(Son(0)), 0, SND_NOWAIT Or SND_NODEFAULT Or SND_MEMORY Or SND_ASYNC)
        DoEvents
    End If
    ReDim Son(0)    ' Vide la variable, plus besoin
    
    
    ' ---------- On fini par noicir complètement l'écran en 3 sec
    ' Chez moi, le PC met 11700 millisecondes pour exécuter 100 incréments d'opacité
    ' Si je veux que tout devienne noir en 3 secondes, il faut incrémenter de :
    Incrément = (11700! / 3000!) * (100! / (255 - Opacité))
    Do While Opacité < 255
        ' Augmente l'opacité jusqu'à devenir noir
        Opacité = Opacité + Incrément
        Opacité = IIf(Opacité > 255, 255, Opacité)
        Call Transparence("ON", maForme, CByte(Opacité))
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
    ' Supprime notre forme adorée
    Unload maForme
    Set maForme = Nothing
    End

End Sub
