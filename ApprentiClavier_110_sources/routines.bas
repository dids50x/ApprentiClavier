Attribute VB_Name = "Module_routines"
'***************** ROUTINES NE NECESSITANT PAS DE TRADUCTIONS ******************************
'Ce logiciel libre est disponible sous licence GNU/GPL,
'dont une copie se trouvera dans le fichier gpl.txt,
'avec une traduction fran�aise non officielle gpl-fr.txt.

'*****************  INITIALIZATION  ********************************************************
Public Sub inits()
' FullScreenSwitch = 0 pour debug, FullScreenSwitch = 1 pour mode normal
FullScreenSwitch = 1

' Autres inits
altf4 = 0
avecf2 = 0
avecf3 = 0
bascule = 0
bipinhibit = 0
cadencecara = 300: cadencemot = 260: cadenceligne = 400
consult = 0
derligne = 0
echapbis = 0: echapoff = 0
erepeat = 0
espacevalid = 0
f1msgform = 0
fullscreeninhibit = 0
incomplet = 0
inexo = 0
iter = 0: iiante = 0: iiprec = 0
iwrong = 0: iwrongbis = 0: iwrongbismax = 5: iwrongCR = 0: iwrongCRmax = 5
KeyFirst = 0: KeySecond = 0: KeyThird = 0
keyinhibit = 0
llold = -2
le�onfontsize5 = 18
msgf = 0
nbcaras = 0: nbonscaras = -1
nbexo = 0
nextle�on = 0
noalt = 1
nodoublesono = 0
noF1 = 0
noechapF1 = 0
notab = 1
numpad = 0
pagenum = 0
passb = 0
pctt = 0: pct1 = 85  ' Attention, rester en accord avec le texte cit� dans la page pgic4
quitactive = 0
quitF2 = 0
t2inhibit = 0
timein = 0: timeout = 0
typele�on = 0
zfactor = 1
zoomfactor = 1  ' no zoom 12/2011
zoomvalue = 1.15 ' with zoom 12/2011
le�onfontsize5 = 18 * zoomvalue ' 12/2011

' Taille cara default pour MSGFORM seulement 12/2011
fsizedefault = 11

' Modifications et ajouts 12/2011
' Jeu de couleurs pour MSGFORM et autres formes
f_blanc = &HFFFFFF
f_bleuclair = &HFFFF00
f_bleufonc� = &HC00000
f_bleuvif = &HFFFF00
f_gris = &HC0C0C0
f_grisp�le = &HE0E0E0
f_grisfonc� = &H808080
f_jaunevif = &H80FFFF
f_jaunetr�svif = &HFFFF
f_marronsombre = &H4040
f_noirfonc� = &H80000008
f_noirgris = &H4040
f_noir = &H0
f_noirpresque = &H400000
f_orangeclair = &HC0E0FF
f_rouge = &H40C0
f_rougevif = &HFF
f_vert = &H808000
f_vertp�le = &HC0C000
f_vertsombre = &H404000
f_vertvif = &HFF00
f_violet = &HC000C0
f_violetsombre = &H400040
f_violetvif = &HFF00FF

' Couleurs pour DEFAULT MSGFORM
ffc_default = f_noir
fbc_default = f_gris

' Couleurs pour Aide F1 MSGFORM
ffc_f1 = f_orangeclair
fbc_f1 = f_vert

' Couleurs pour QUITTER MSGFORM
ffc_quit = f_noir
fbc_quit = f_bleufonc�

' Variables Default colors for MSGFORM
ffc = ffc_default
fbc = fbc_default

' Temps et fr�quence des beep sonores, avril 2008
txtTps = 0.05

' Chemin du fichier exe
vpath = App.Path
If Right(vpath, 1) <> "\" Then vpath = vpath & "\"

'Reset des le�ons du rep Standard (+ Set Personnalis�) (parm 0=noforce, 1=Force les le�ons Standard)
nivoRep = "Standard" 'immuable, ne pas traduire
Module_le�ons.reset_standard 1
Module_le�ons2.reset2_standard 1
Module_le�ons3.reset3_standard 1

' M�moriser pour futur Restore en quittant
bNumLockState = Module_routines.IsNumLockOn()
bCapsLockState = Module_routines.IsCapsLockOn()
bScrollLockState = Module_routines.IsScrollLockOn()

' Localiser Jaws, Modifier fichier default.jkm
Module_routines.SonoLocate
'Module_routines.modify_jkm "Basics"

' BIENVENUE, NOM de l'UTILISATEUR
Bienvenue.Show 1

' Reset des CAPSLOCK, NUMLOCK, SCROLLLOCK, � placer APR�S le bienvenue.show
Module_routines.SetKeys "CAPSLOCK_OFF"
Module_routines.SetKeys "NUMLOCK_OFF"
Module_routines.SetKeys "SCROLLLOCK_OFF"
End Sub


'*****************  TRADUIRE la BARRE DE MENU dans les MENUS  *******************************
'modifi� 12/2011 pour zoom et couleurs
Public Sub MenuEditorTrans(menu)
On Error Resume Next
menu.Fichier.Caption = meFichier
On Error Resume Next
menu.Quitter_bm.Caption = meQuitter_bm
On Error Resume Next
menu.Options.Caption = meOptions
On Error Resume Next
menu.Standard.Caption = meStandard
On Error Resume Next
menu.Personnalis�.Caption = mePersonnalis�
On Error Resume Next
menu.DebExpliNormal.Caption = meDebExpliNormal
On Error Resume Next
menu.DebExpliRapide.Caption = meDebExpliRapide
On Error Resume Next
menu.DebGenLent.Caption = meDebGenLent
On Error Resume Next
menu.DebGenMoyen.Caption = meDebGenMoyen
On Error Resume Next
menu.DebGenVite.Caption = meDebGenVite
On Error Resume Next
menu.BipClassique.Caption = meBipClassique
On Error Resume Next
menu.BipDiff�rent.Caption = meBipDiff�rent
On Error Resume Next
menu.NoZoom.Caption = meNoZoom
On Error Resume Next
menu.WithZoom.Caption = meWithZoom
On Error Resume Next
menu.BasicColors.Caption = meBasicColors
On Error Resume Next
menu.OtherColors.Caption = meOtherColors
On Error Resume Next
menu.Aide.Caption = meAide
On Error Resume Next
menu.AideG�n�rale.Caption = meAideG�n�rale
On Error Resume Next
menu.AideM�moire.Caption = meAideM�moire
On Error Resume Next
menu.Enseignant.Caption = meEnseignant
On Error Resume Next
menu.Sonorisation.Caption = meSonorisation
On Error Resume Next
menu.Aproposde.Caption = meAproposde
On Error Resume Next
menu.Reset.Caption = meReset
On Error Resume Next
menu.Restart.Caption = meRestart
End Sub


'*********  SonoLocate : Localiser les sous-r�ps Jaws pour les CONFIG jcf, jsb...  *********
Public Sub SonoLocate()
ii = 0

' ****** D�BUT du CODE IDENTIQUE � SONOLOCATE de SetUpGlobal.bas dans ApprentiClavier_Setup *********
' *** VERSION JAWS3.7 en c:\jfw37
repjaws = ""
ujaws(ii) = "c:\"
repjaws = Dir("c:\jfw37", vbDirectory)
If repjaws <> "" Then
    repj(ii) = "jfw37"
    ii = ii + 1
End If

' *** VERSION JAWS3.7 en d:\jfw37
repjaws = ""
ujaws(ii) = "d:\"
On Error Resume Next
repjaws = Dir("d:\jfw37", vbDirectory)
If repjaws <> "" Then
    repj(ii) = "jfw37"
    ii = ii + 1
End If

' *** VERSIONS JAWS4 et JAWS5 en  c:\jaws???
repjaws = ""
ujaws(ii) = "c:\"
repjaws = Dir("c:\jaws???", vbDirectory)

' BOUCLE en ujaws = C:\ pour REP�RER les r�ps CONFIG de VERSIONS JAWS, si plusieurs r�ps c:\jaws??? !!
If repjaws <> "" Then
    repj(ii) = repjaws  '1er r�sultat
    ujaws(ii) = "c:\"
    
    Do While repj(ii) <> ""
        If repj(ii) <> "." And repj(ii) <> ".." Then
            repjaws = Dir
            ii = ii + 1
            repj(ii) = repjaws
            ujaws(ii) = "c:\"
        End If
    Loop
End If

' *** VERSIONS JAWS4 et JAWS5 en  d:\jaws???
repjaws = ""
ujaws(ii) = "d:\"
On Error Resume Next
repjaws = Dir("d:\jaws???", vbDirectory)

' BOUCLE en ujaws = D:\ pour REP�RER les r�ps CONFIG de VERSIONS JAWS, si plusieurs r�ps c:\jaws???
If repjaws <> "" Then
    repj(ii) = repjaws  '1er r�sultat
    
    Do While repj(ii) <> ""
        If repj(ii) <> "." And repj(ii) <> ".." Then
            repjaws = Dir
            ii = ii + 1
            repj(ii) = repjaws
            If repj(ii) <> "" Then ujaws(ii) = ujaws(ii - 1)
        End If
    Loop
End If
        
' ajout mars 2008
' *** VERSIONS JAWS8 et Suivantes, pour Windows VISTA, en C:\ProgramData\Freedom Scientific\Jaws\?.??\Settings\Fra
repUsers = Dir("c:\Users\", vbDirectory)
'MsgBox repUsers
repjaws = ""
ujaws(ii) = "c:\"
On Error Resume Next
repjaws = Dir("c:\ProgramData\Freedom Scientific\Jaws\?*", vbDirectory)  ' "?*" pour jaws 20xx

' BOUCLE en ujaws = C:\Documents... pour REP�RER les r�ps CONFIG de VERSIONS JAWS, si plusieurs r�ps c:\Documents and Settings\...jaws\?.??
If repjaws <> "" Then
    If repjaws <> "." And repjaws <> ".." Then
        repj(ii) = repjaws  '1er r�sultat
        ujaws(ii) = "c:\"
        ii = ii + 1
    End If
    
    Do While repjaws <> ""
        repjaws = Dir
        If repjaws <> "" And repjaws <> "." And repjaws <> ".." Then
            repj(ii) = "ProgramData\Freedom Scientific\Jaws\" & repjaws
            ujaws(ii) = "c:\"
            ii = ii + 1
        End If
    Loop
End If

' *** VERSIONS JAWS8 et Suivantes, pour Windows VISTA, en D:\ProgramData\Freedom Scientific\Jaws\?.??\Settings\Fra
repjaws = ""
ujaws(ii) = "d:\"
On Error Resume Next
repjaws = Dir("d:\ProgramData\Freedom Scientific\Jaws\??.??", vbDirectory) 'mars 2008

' BOUCLE en ujaws = d:\Documents... pour REP�RER les r�ps CONFIG de VERSIONS JAWS, si plusieurs r�ps d:\Documents and Settings\...jaws\?.??
If repjaws <> "" Then
    If repjaws <> "." And repjaws <> ".." Then
        repj(ii) = repjaws  '1er r�sultat
        ujaws(ii) = "d:\"
        ii = ii + 1
    End If
    
    Do While repjaws <> ""
        repjaws = Dir
        If repjaws <> "" And repjaws <> "." And repjaws <> ".." Then
            repj(ii) = "ProgramData\Freedom Scientific\Jaws\" & repjaws
            ujaws(ii) = "d:\"
            ii = ii + 1
        End If
    Loop
End If

' �viter deux affichages m�me version de jaws
If repUsers = "" Then

' *** VERSIONS JAWS6 et Suivantes, pour Windows XP, en C:\Documents and Settings\All Users\Application Data\Freedom Scientific\Jaws\??.??\Settings\Fra
repjaws = ""
ujaws(ii) = "c:\"
On Error Resume Next
repjaws = Dir("c:\Documents and Settings\All Users\Application Data\Freedom Scientific\Jaws\??.??", vbDirectory) 'septembre 2007
' BOUCLE en ujaws = C:\Documents... pour REP�RER les r�ps CONFIG de VERSIONS JAWS, si plusieurs r�ps c:\Documents and Settings\...jaws\?.??
If repjaws <> "" Then
    If repjaws <> "." And repjaws <> ".." Then
        repj(ii) = repjaws  '1er r�sultat
        ujaws(ii) = "c:\"
        ii = ii + 1
    End If
    
    Do While repjaws <> ""
        repjaws = Dir
        If repjaws <> "" And repjaws <> "." And repjaws <> ".." Then
            repj(ii) = "Documents and Settings\All Users\Application Data\Freedom Scientific\Jaws\" & repjaws
            ujaws(ii) = "c:\"
            ii = ii + 1
        End If
    Loop
End If

' *** VERSIONS JAWS6 et Suivantes, pour Windows XP, en D:\Documents and Settings\All Users\Application Data\Freedom Scientific\Jaws\??.??\Settings\Fra
repjaws = ""
ujaws(ii) = "d:\"
On Error Resume Next
repjaws = Dir("d:\Documents and Settings\All Users\Application Data\Freedom Scientific\Jaws\??.??", vbDirectory) 'septembre 2007

' BOUCLE en ujaws = D:\Documents... pour REP�RER les r�ps CONFIG de VERSIONS JAWS, si plusieurs r�ps c:\Documents and Settings\...jaws\??.??
If repjaws <> "" Then
    If repjaws <> "." And repjaws <> ".." Then
        repj(ii) = repjaws  '1er r�sultat
        ujaws(ii) = "d:\"
        ii = ii + 1
    End If
    
    Do While repjaws <> ""
        repjaws = Dir
        If repjaws <> "" And repjaws <> "." And repjaws <> ".." Then
            repj(ii) = "Documents and Settings\All Users\Application Data\Freedom Scientific\Jaws\" & repjaws
            ujaws(ii) = "d:\"
            ii = ii + 1
        End If
    Loop
End If

End If  ' fin If repUsers
            
' Bilan
'MsgBox "0 " & ujaws(0) & repj(0) & "   1 " & ujaws(1) & repj(1) & "   2 " & ujaws(2) & repj(2) & "   3 " & ujaws(3) & repj(3) & "   4 " & ujaws(4) & repj(4) & "   5 " & ujaws(5) & repj(5) & "   6 " & ujaws(6) & repj(6) & "   7 " & ujaws(7) & repj(7) & "   8 " & ujaws(8) & repj(8)
' ***** FIN du CODE IDENTIQUE � SONOLOCATE de SetUpGlobal.bas dans ApprentiClavier_Setup.vbp ******

' Jaws INTROUVABLE
If repj(0) = "" Then
    ujaws(ii) = ""
    repjawsnames = ""
    Exit Sub
End If

' BOUCLE de COPIE EVENTUELLE vers les r�ps CONFIG JAWS trouv�s
ii = 0
Do While repj(ii) <> ""
    If repj(ii) <> "." And repj(ii) <> ".." Then
        repjawsfra = ujaws(ii) & repj(ii) & repjawscountry
        repjawsjsb = repjawsfra & "ApprentiClavier.jsb"
    
        ' SONOCOPY : Lancer les copies (ou l'effacement) des fichiers de configuration ApprentiClavier
        rrs(ii) = InStr(1, repjawsfra, "Documents and Settings\All Users\Application Data\Freedom Scientific\Jaws", 1)
        rrt(ii) = InStr(1, repjawsfra, "ProgramData\Freedom Scientific\Jaws", 1) 'ajout mars 2008
        Module_routines.sonocopy 'appel sans condition, modifi� juin 2007 pour version Jaws 8 sans r�pertoire factdef
    End If
    
    ' Num�ro de version JAWS suivante
    ii = ii + 1
Loop
End Sub


' *******************  SONOCOPY  **********************************************
Public Sub sonocopy()
'modifi� juin 2007 pour version Jaws 8 sans r�pertoire factdef, il vaut mieux rechercher jfw.exe
'MsgBox "  repj=" & repj(ii) & "  rrs=" & rrs(ii) & "  rrt=" & rrt(ii)

' Versions JAWS 3.7, 4 et 5
If rrs(ii) = 0 And rrt(ii) = 0 Then 'modifi� mars 2008
    repjexe(ii) = repj(ii)
Else
' Versions JAWS 6 et Suivantes, modifi� septembre 2007
    If InStr(1, Right(LCase(repj(ii)), 5), "\") = 0 Then repjexe(ii) = "Program Files\Freedom Scientific\Jaws\" & Right(LCase(repj(ii)), 5)
    If InStr(1, Right(LCase(repj(ii)), 5), "\") = 1 Then repjexe(ii) = "Program Files\Freedom Scientific\Jaws\" & Right(LCase(repj(ii)), 4)
    If InStr(1, Right(LCase(repj(ii)), 5), "\") = 2 Then repjexe(ii) = "Program Files\Freedom Scientific\Jaws\" & Right(LCase(repj(ii)), 3)
End If

' Si l'ex�cutable JAWS est ABSENT, on ne fait rien
' Version Jaws 3.7 avec ex�cutable jaws.exe, autres versions avec jfw.exe
If LCase(repjexe(ii)) = "jfw37" Then
    If Dir(ujaws(ii) & repjexe(ii) & "\jaws.exe") = "" Then Exit Sub
Else
    If Dir(ujaws(ii) & repjexe(ii) & "\jfw.exe") = "" Then Exit Sub
End If

' Si JSB PR�SENT, config pr�sentes, on donnera simplement les noms des r�ps Jaws comme info
If Dir(repjawsjsb) <> "" Then GoTo SONOC1

' Si JSB ABSENT (d� � nouvelle install de jaws): Importation de tous les fichiers de config V-Jaws
' Ce r�pertoire V-Jaws est une sorte de sauvegarde de la configuration ApprentiClavier,
' au cas o� Jaws serait install� post�rieurement � ApprentiClavier.
' Pour repartir de V-Jaws, il suffit de supprimer ApprentiClavier.jsb dans le rep Jawsxxx\settings\fra\

' Fichier JCF : moins bavard que default.jcf. Notamment pas de messages "tuteurs' gr�ce � TUTOR=0|0|0
If Dir("c:\ApprentiClavier\V-Jaws\ApprentiClavier.jcf") <> "" Then FileCopy "c:\ApprentiClavier\V-Jaws\ApprentiClavier.jcf", repjawsfra & "ApprentiClavier.jcf"

' Fichier JDF dictionnaire pour mieux prononcer
If Dir("c:\ApprentiClavier\V-Jaws\ApprentiClavier.jdf") <> "" Then FileCopy "c:\ApprentiClavier\V-Jaws\ApprentiClavier.jdf", repjawsfra & "ApprentiClavier.jdf"

' Fichier JSS qui d�finit le d�bit de la voix et la ponctuation, gr�ce aux titres de fen�tre
If Dir("c:\ApprentiClavier\V-Jaws\ApprentiClavier.jss") <> "" Then
    ' Copie jss de V-Jaws vers repjawsfra
    FileCopy "c:\ApprentiClavier\V-Jaws\ApprentiClavier.jss", repjawsfra & "ApprentiClavier.jss"

    ' Fichier JSB : il faut toujours UTILISER LE COMPILATEUR de la version Jaws appropri�, pb de compatibilit�
    ' Il faudra donner le chemin de l'ex�cutable scompile : on d�finit repjexe
    
    ' Localiser le SCOMPILE.EXE
    tempo = ""
    tempo = Dir(ujaws(ii) & repjexe(ii) & "\scompile.exe")
    If tempo <> "" Then
                
        ' Se placer dans le chemin des CONFIG, ou, dans le chemin du compilateur, selon le type de version JAWS
        ' Y copier le compilateur, ou, le script jss, selon le type de version JAWS
        ChDrive (ujaws(ii))
        
        ' Versions JAWS 3.7, 4 et 5
        If rrs(ii) = 0 And rrt(ii) = 0 Then 'modifi� mars 2008
            ChDir (repjawsfra)
            FileCopy ujaws(ii) & repjexe(ii) & "\scompile.exe", repjawsfra & "scompile.exe"
        
        ' Versions JAWS 6 et Suivantes
        Else
            ChDir (ujaws(ii) & repjexe(ii))
            FileCopy "c:\ApprentiClavier\V-Jaws\ApprentiClavier.jss", ujaws(ii) & repjexe(ii) & "\ApprentiClavier.jss"
        End If
                    
        ' COMPILER le jss
        On Error Resume Next
        Module_exec.ExecAndWait "scompile.exe ApprentiClavier.jss"
                
        ' Versions JAWS 3.7, 4 et 5
        If rrs(ii) = 0 And rrt(ii) = 0 Then
            ' M�nage
            On Error Resume Next
            Kill repjawsfra & "scompile.exe"
        End If
        
        ' Versions JAWS 6 et Suivantes (rrs pour Windows XP, rrt pour Windows VISTA) modifi� mars 2008
        If rrs(ii) > 0 Or rrt(ii) > 0 Then
            If ujaws(ii) & repjexe(ii) & "\" <> repjawsfra Then
                On Error Resume Next
                FileCopy ujaws(ii) & repjexe(ii) & "\ApprentiClavier.jsb", repjawsfra & "ApprentiClavier.jsb"
                ' M�nage
                On Error Resume Next
                Kill ujaws(ii) & repjexe(ii) & "\ApprentiClavier.jss"
                On Error Resume Next
                Kill ujaws(ii) & repjexe(ii) & "\ApprentiClavier.jsb"
            End If
        End If
    
    End If
End If

' Si ECHEC de SCOMPILE, c'est normalement d� � un Jaws trop ancien
' on copie alors le jsb (obtenu sp�cifiquement en Version 4.01 max) de V-Jaws
If Dir(repjawsjsb) = "" Then
        If LCase(repj(ii)) = "jfw37" Or LCase(repj(ii)) = "jaws401" Then
            If Dir("c:\ApprentiClavier\V-Jaws\ApprentiClavier-Jaws401.jsb") <> "" Then FileCopy "c:\ApprentiClavier\V-Jaws\ApprentiClavier-Jaws401.jsb", repjawsfra & "ApprentiClavier.jsb"
        Else
            If Dir("c:\ApprentiClavier\V-Jaws\ApprentiClavier.jsb") <> "" Then FileCopy "c:\ApprentiClavier\V-Jaws\ApprentiClavier.jsb", repjawsfra & "ApprentiClavier.jsb"
        End If
End If

SONOC1:
' Fournit l'INFORMATION sur les r�pertoires D�TECT�S
If Dir(repjawsjsb) <> "" Then
    ' JAWS Versions 3.7, 4 et 5
    If rrs(ii) = 0 And rrt(ii) = 0 Then  'modifi� mars 2008
        repjawsnames = repjawsnames & LCase(ujaws(ii)) & LCase(repj(ii)) & ". "
    End If
    ' JAWS Version 6 et Suivantes modifi� septembre 2007
    If rrs(ii) > 0 Or rrt(ii) > 0 Then 'modifi� mars 2008
        ' Nom de rep Jaws � 5 caras tel que "10.20"
        If InStr(1, Right(LCase(repj(ii)), 5), "\") = 0 Then repjawsnames = repjawsnames & LCase(ujaws(ii)) & "jaws" & Right(LCase(repj(ii)), 5) & ". "
        ' Nom de rep Jaws � 4 caras tel que "6.20"
        If InStr(1, Right(LCase(repj(ii)), 5), "\") = 1 Then repjawsnames = repjawsnames & LCase(ujaws(ii)) & "jaws" & Right(LCase(repj(ii)), 4) & ". "
        ' Nom de rep Jaws � 3 caras tel que "6.0"
        If InStr(1, Right(LCase(repj(ii)), 5), "\") = 2 Then repjawsnames = repjawsnames & LCase(ujaws(ii)) & "jaws" & Right(LCase(repj(ii)), 3) & ". "
    End If
End If
End Sub


'*********  DATA_USER : NOM UTILISATEUR et gestion de son R�PERTOIRE, puis appel du MENU PRINCIPAL  ************
Public Sub data_user()
' modifi� 12/2011 pour zoom et couleurs
' Reset
echapbis = 0: keyinhibit = 0: t2inhibit = 0
If nom = "" Then Exit Sub

' Cr�ation du r�pertoire Utilisateur
On Error Resume Next
MkDir vpath & "Utilisateurs"
vfile = vpath & "Utilisateurs\" & nom
On Error Resume Next
MkDir vfile

' Cr�ation de la table pctok si elle n'existe pas encore pour cet utilisateur (numle�on en derni�re col)
If Dir(vfile & "\pctok.txt") = "" Then
    For jj = 0 To 49
        For kk = 0 To 8
        pctok(jj, kk) = 0
        Next kk
        'Visualise num�ro de le�on en derni�re col
        If jj < 25 Then pctok(jj, kk) = jj - 2
        If jj >= 25 Then pctok(jj, kk) = jj - 27
    Next jj
Else
    ' R�cup�ration des r�sultats ant�rieurs de l'utilisateur (+ numle�on en derni�re col)
    Open vfile & "\pctok.txt" For Input As #1
    For jj = 0 To 49
        Input #1, pctok(jj, 0), pctok(jj, 1), pctok(jj, 2), pctok(jj, 3), pctok(jj, 4), pctok(jj, 5), pctok(jj, 6), pctok(jj, 7), pctok(jj, 8), pctok(jj, 9)
    Next jj
    Close #1
End If

' Cr�ation de la table vitok si elle n'existe pas encore pour cet utilisateur (numle�on en derni�re col)
If Dir(vfile & "\vitok.txt") = "" Then
    For jj = 0 To 49
        For kk = 0 To 8
        vitok(jj, kk) = 0
        Next kk
        'Visualise num�ro de le�on en derni�re col
        If jj < 25 Then vitok(jj, kk) = jj - 2
        If jj >= 25 Then vitok(jj, kk) = jj - 27
    Next jj
Else
    ' R�cup�ration des r�sultats ant�rieurs de l'utilisateur (+ numle�on en derni�re col)
    Open vfile & "\vitok.txt" For Input As #1
    For jj = 0 To 49
        Input #1, vitok(jj, 0), vitok(jj, 1), vitok(jj, 2), vitok(jj, 3), vitok(jj, 4), vitok(jj, 5), vitok(jj, 6), vitok(jj, 7), vitok(jj, 8), vitok(jj, 9)
    Next jj
    Close #1
End If

' Cr�ation du fichier INI de l'utilisateur, s'il n'existe pas
If Dir(vfile & "\" & nom & ".ini") = "" Then
    numle�on = 0: numexo = 0
    nivo = msgStandard
    nivoRep = "Standard"
    debexplilevel = msgNormal
    biplevel = msgClassique
    debgenlevel = msgMoyen
    zoomlevel = msgNoZoom
    colorslevel = msgBasicColors
Else
    ' R�cup�ration de l'avancement de l'utilisateur pour pr�s�lection des menus, fichier ini
    Open vfile & "\" & nom & ".ini" For Input As #2
    Input #2, numle�on, numexo, nivo, debexplilevel, biplevel, debgenlevel, zoomlevel, colorslevel
    Close #2
    If nivo = msgStandard Or nivo = msgPersonnalis� Then
        If nivo = msgStandard Then nivoRep = "Standard" 'immuable, ne pas traduire, juin 2007
        If nivo = msgPersonnalis� Then nivoRep = "Personnalis�" 'immuable, ne pas traduire, juin 2007
        Else
        Module_routines.reset_ini
    End If
    If debexplilevel = msgNormal Or debexplilevel = msgRapide Then
        Else
        Module_routines.reset_ini
    End If
    If biplevel = msgClassique Or biplevel = msgDiff�rent Then
        Else
        Module_routines.reset_ini
    End If
    If debgenlevel = msgLent Or debgenlevel = msgMoyen Or debgenlevel = msgVite Then
        Else
        Module_routines.reset_ini
    End If
    If zoomlevel = msgNoZoom Or zoomlevel = msgWithZoom Then
        Module_routines.ZoomSet
        Else
        Module_routines.reset_ini
    End If
    If colorslevel = msgBasicColors Or colorslevel = msgOtherColors Then
        Else
        Module_routines.reset_ini
    End If
    If numle�on > 25 Or numle�on < 0 Then numle�on = 0
    If numexo > 8 Or numexo < 0 Then numexo = 0
End If

' APPEL du MENU_PRINCIPAL
Unload Bienvenue
Menu_principal.Show 1
End Sub

' ***** ZOOMSET 12/2011 ********
Public Sub ZoomSet()
If zoomlevel = msgNoZoom Then zoomfactor = 1
If zoomlevel = msgWithZoom Then zoomfactor = zoomvalue
End Sub

' ****************  RESET_INI : refait le fichier ini par d�faut  ***************************
' modifi� 12/2011 pour zoom et couleurs
Public Sub reset_ini()
Kill vfile & "\" & nom & ".ini"
nivo = msgStandard
debexplilevel = msgNormal
biplevel = msgClassique
debgenlevel = msgMoyen
zoomlevel = msgNoZoom
colorslevel = msgBasicColors
End Sub

' ************** CARA2LIGNE1 : compare cara de text2 � info ligne de text1 *******************
' **** Routine pour accepter/refuser les r�ponses utilisateur dans la le�on 1 ****************
Public Sub cara2ligne1(le�on)
With le�on

' D�tecter le Alt255 final �ventuel (qui permet sonorisation correcte du dernier cara mis en surbrillance)
If Right(.text1.Text, 1) = "�" Then
    lt1 = Len(.text1.Text) - 1
Else
    lt1 = Len(.text1.Text)
End If

' Associer le code touche attendu
If UCase(Left(.text1.Text, lt1)) = UCase(vvEntr�e) Then KeyExpect = 13
If UCase(Left(.text1.Text, lt1)) = UCase(vvControl) Then KeyExpect = 17
If UCase(Left(.text1.Text, lt1)) = UCase(vvControlGauche) Then KeyExpect = 17
If UCase(Left(.text1.Text, lt1)) = UCase(vvControlDroit) Then KeyExpect = 17
If UCase(Left(.text1.Text, lt1)) = UCase(vvAlt) Then KeyExpect = 18
If UCase(Left(.text1.Text, lt1)) = UCase(vvAltGr) Then KeyExpect = 255 'Cas particulier 17 puis 18
If UCase(Left(.text1.Text, lt1)) = UCase(vv�chap) Then KeyExpect = 27
If UCase(Left(.text1.Text, lt1)) = UCase(vvEchap) Then KeyExpect = 27
If UCase(Left(.text1.Text, lt1)) = UCase(vvEspace) Then KeyExpect = 32
If UCase(Left(.text1.Text, lt1)) = UCase(vvFlecheGauche) Then KeyExpect = 37
If UCase(Left(.text1.Text, lt1)) = UCase(vvFlecheHaut) Then KeyExpect = 38
If UCase(Left(.text1.Text, lt1)) = UCase(vvFlecheDroite) Then KeyExpect = 39
If UCase(Left(.text1.Text, lt1)) = UCase(vvFlecheBas) Then KeyExpect = 40
If UCase(Left(.text1.Text, lt1)) = UCase(vvWindowsGauche) Then KeyExpect = 91
If UCase(Left(.text1.Text, lt1)) = UCase(vvWindowsDroit) Then KeyExpect = 92
If UCase(Left(.text1.Text, lt1)) = UCase(vvMenuContextuel) Then KeyExpect = 93
If UCase(Left(.text1.Text, lt1)) = UCase(vvVerrouillageMajuscules) Then KeyExpect = 20
If UCase(Left(.text1.Text, lt1)) = UCase(vvMaj) Then KeyExpect = 16
If UCase(Left(.text1.Text, lt1)) = UCase(vvMajGauche) Then KeyExpect = 16
If UCase(Left(.text1.Text, lt1)) = UCase(vvMajDroit) Then KeyExpect = 16
If UCase(Left(.text1.Text, lt1)) = UCase(vvInsertion) Then KeyExpect = 45
If UCase(Left(.text1.Text, lt1)) = UCase(vvSuppression) Then KeyExpect = 46
If UCase(Left(.text1.Text, lt1)) = UCase(vvD�but) Then KeyExpect = 36
If UCase(Left(.text1.Text, lt1)) = UCase(vvFin) Then KeyExpect = 35
If UCase(Left(.text1.Text, lt1)) = UCase(vvPagePr�c�dente) Then KeyExpect = 33
If UCase(Left(.text1.Text, lt1)) = UCase(vvPageSuivante) Then KeyExpect = 34
If UCase(Left(.text1.Text, lt1)) = UCase(vvTab) Then KeyExpect = 9
If UCase(Left(.text1.Text, lt1)) = UCase(vvRetourArriere) Or LCase(Left(.text1.Text, lt1)) = LCase(vvRetourArri�re) Then KeyExpect = 8
If LCase(Left(.text1.Text, lt1)) = LCase(vvRetourArri�re) Then KeyExpect = 8
If UCase(Left(.text1.Text, lt1)) = UCase(vvImpression) Then KeyExpect = 44
If Left(.text1.Text, lt1) = vvArr�tD�fil Then KeyExpect = 145
If UCase(Left(.text1.Text, lt1)) = UCase(vvPause) Then KeyExpect = 19

' Touches du pav� num�rique
If UCase(Left(.text1.Text, lt1)) = UCase(vvVerrouillageNum�rique) Then KeyExpect = 144
If UCase(Left(.text1.Text, lt1)) = UCase(vvPlus) Then KeyExpect = 107
If UCase(Left(.text1.Text, lt1)) = UCase(vvMoins) Then KeyExpect = 109
If UCase(Left(.text1.Text, lt1)) = UCase(vvTiret) Then KeyExpect = 109
If UCase(Left(.text1.Text, lt1)) = UCase(vvDiviser) Then KeyExpect = 111
If UCase(Left(.text1.Text, lt1)) = UCase(vvBarreOblique) Then KeyExpect = 111
If UCase(Left(.text1.Text, lt1)) = UCase(vvMultiplier) Then KeyExpect = 106
If UCase(Left(.text1.Text, lt1)) = UCase(vv�toile) Then KeyExpect = 106
If UCase(Left(.text1.Text, lt1)) = UCase(vvAst�risque) Then KeyExpect = 106
If UCase(Left(.text1.Text, lt1)) = UCase(vvPoint) Then KeyExpect = 110

'If numpad = 1 Then
'    If UCase(left(.text1.text,lt1)) = UCase(vvD�but) Then keyexpect = 103
'    If UCase(left(.text1.text,lt1)) = UCase(vvFin) Then keyexpect = 97
'    If UCase(left(.text1.text,lt1)) = UCase(vvPagePr�c�dente) Then keyexpect = 105
'    If UCase(left(.text1.text,lt1)) = UCase(vvPageSuivante) Then keyexpect = 99
'    If UCase(left(.text1.text,lt1)) = UCase(vvFlecheGauche) Then keyexpect = 100
'    If UCase(left(.text1.text,lt1)) = UCase(vvFlecheDroite) Then keyexpect = 102
'    If UCase(left(.text1.text,lt1)) = UCase(vvFlecheHaut) Then keyexpect = 104
'    If UCase(left(.text1.text,lt1)) = UCase(vvFlecheBas) Then keyexpect = 98
'    If UCase(left(.text1.text,lt1)) = UCase(vvInsertion) Then keyexpect = 98
'    If UCase(left(.text1.text,lt1)) = UCase(vvSuppression) Then keyexpect = 98
'End If

' Touches de Fonction
If UCase(Left(.text1.Text, lt1)) = "F1" Then KeyExpect = 112
If UCase(Left(.text1.Text, lt1)) = "F2" Then KeyExpect = 113
If UCase(Left(.text1.Text, lt1)) = "F3" Then KeyExpect = 114
If UCase(Left(.text1.Text, lt1)) = "F4" Then KeyExpect = 115
If UCase(Left(.text1.Text, lt1)) = "F5" Then KeyExpect = 116
If UCase(Left(.text1.Text, lt1)) = "F6" Then KeyExpect = 117
If UCase(Left(.text1.Text, lt1)) = "F7" Then KeyExpect = 118
If UCase(Left(.text1.Text, lt1)) = "F8" Then KeyExpect = 119
If UCase(Left(.text1.Text, lt1)) = "F9" Then KeyExpect = 120
If UCase(Left(.text1.Text, lt1)) = "F10" Then KeyExpect = 121
If UCase(Left(.text1.Text, lt1)) = "F11" Then KeyExpect = 122
If UCase(Left(.text1.Text, lt1)) = "F12" Then KeyExpect = 123

' Lettres
If Left(.text1.Text, lt1) = "a" Or Left(.text1.Text, lt1) = "A" Then KeyExpect = 65
If Left(.text1.Text, lt1) = "b" Or Left(.text1.Text, lt1) = "B" Then KeyExpect = 66
If Left(.text1.Text, lt1) = "c" Or Left(.text1.Text, lt1) = "C" Then KeyExpect = 67
If Left(.text1.Text, lt1) = "d" Or Left(.text1.Text, lt1) = "D" Then KeyExpect = 68
If Left(.text1.Text, lt1) = "e" Or Left(.text1.Text, lt1) = "E" Then KeyExpect = 69
If Left(.text1.Text, lt1) = "f" Or Left(.text1.Text, lt1) = "F" Then KeyExpect = 70
If Left(.text1.Text, lt1) = "g" Or Left(.text1.Text, lt1) = "G" Then KeyExpect = 71
If Left(.text1.Text, lt1) = "h" Or Left(.text1.Text, lt1) = "H" Then KeyExpect = 72
If Left(.text1.Text, lt1) = "i" Or Left(.text1.Text, lt1) = "I" Then KeyExpect = 73
If Left(.text1.Text, lt1) = "j" Or Left(.text1.Text, lt1) = "J" Then KeyExpect = 74
If Left(.text1.Text, lt1) = "k" Or Left(.text1.Text, lt1) = "K" Then KeyExpect = 75
If Left(.text1.Text, lt1) = "l" Or Left(.text1.Text, lt1) = "L" Then KeyExpect = 76
If Left(.text1.Text, lt1) = "m" Or Left(.text1.Text, lt1) = "M" Then KeyExpect = 77
If Left(.text1.Text, lt1) = "n" Or Left(.text1.Text, lt1) = "N" Then KeyExpect = 78
If Left(.text1.Text, lt1) = "o" Or Left(.text1.Text, lt1) = "O" Then KeyExpect = 79
If Left(.text1.Text, lt1) = "p" Or Left(.text1.Text, lt1) = "P" Then KeyExpect = 80
If Left(.text1.Text, lt1) = "q" Or Left(.text1.Text, lt1) = "Q" Then KeyExpect = 81
If Left(.text1.Text, lt1) = "r" Or Left(.text1.Text, lt1) = "R" Then KeyExpect = 82
If Left(.text1.Text, lt1) = "s" Or Left(.text1.Text, lt1) = "S" Then KeyExpect = 83
If Left(.text1.Text, lt1) = "t" Or Left(.text1.Text, lt1) = "T" Then KeyExpect = 84
If Left(.text1.Text, lt1) = "u" Or Left(.text1.Text, lt1) = "U" Then KeyExpect = 85
If Left(.text1.Text, lt1) = "v" Or Left(.text1.Text, lt1) = "V" Then KeyExpect = 86
If Left(.text1.Text, lt1) = "w" Or Left(.text1.Text, lt1) = "W" Then KeyExpect = 87
If Left(.text1.Text, lt1) = "x" Or Left(.text1.Text, lt1) = "X" Then KeyExpect = 88
If Left(.text1.Text, lt1) = "y" Or Left(.text1.Text, lt1) = "Y" Then KeyExpect = 89
If Left(.text1.Text, lt1) = "z" Or Left(.text1.Text, lt1) = "Z" Then KeyExpect = 90

' Taux de r�ussite
pctt = 100 * (nbcaras / (nbcaras + iwrong))

' FOCUS sur text2
On Error Resume Next
.text2.SetFocus
End With
End Sub


' ************** RAC2LIGNE1 : compare raccourci de text2 � info ligne de text1 ***************
' **** Routine pour accepter/refuser les r�ponses utilisateur dans la le�on 13 ***************
Public Sub rac2ligne1(le�on)
With le�on
ii = Len(.text2.Text) ' Pour help_f2

' D�tecter le Alt255 final �ventuel
If Right(.text1.Text, 1) = "�" Then
    lt1 = Len(.text1.Text) - 1
Else
    lt1 = Len(.text1.Text)
End If

' Chiffres en pav� num�rique (AVEC ou SANS Alt255)
If numpad >= 1 Then
    ShiftExpect = 0
    If Left(.text1.Text, lt1) = "0" Then KeyExpect = 96
    If Left(.text1.Text, lt1) = "1" Then KeyExpect = 97
    If Left(.text1.Text, lt1) = "2" Then KeyExpect = 98
    If Left(.text1.Text, lt1) = "3" Then KeyExpect = 99
    If Left(.text1.Text, lt1) = "4" Then KeyExpect = 100
    If Left(.text1.Text, lt1) = "5" Then KeyExpect = 101
    If Left(.text1.Text, lt1) = "6" Then KeyExpect = 102
    If Left(.text1.Text, lt1) = "7" Then KeyExpect = 103
    If Left(.text1.Text, lt1) = "8" Then KeyExpect = 104
    If Left(.text1.Text, lt1) = "9" Then KeyExpect = 105
End If

' Chiffres au clavier principal (AVEC ou SANS Alt255)
If numpad <= 0 Then
    ShiftExpect = 1 'azerty
    'ShiftExpect = 0 'qwertz
    If Left(.text1.Text, lt1) = "0" Then KeyExpect = 48
    If Left(.text1.Text, lt1) = "1" Then KeyExpect = 49
    If Left(.text1.Text, lt1) = "2" Then KeyExpect = 50
    If Left(.text1.Text, lt1) = "3" Then KeyExpect = 51
    If Left(.text1.Text, lt1) = "4" Then KeyExpect = 52
    If Left(.text1.Text, lt1) = "5" Then KeyExpect = 53
    If Left(.text1.Text, lt1) = "6" Then KeyExpect = 54
    If Left(.text1.Text, lt1) = "7" Then KeyExpect = 55
    If Left(.text1.Text, lt1) = "8" Then KeyExpect = 56
    If Left(.text1.Text, lt1) = "9" Then KeyExpect = 57
End If

' Partie finale d'une combinaison (AVEC ou SANS Alt255 final)
' Lettre finale majuscule d'une combinaison
If Asc(Mid(.text1.Text, lt1, 1)) > 64 And Asc(Mid(.text1.Text, lt1, 1)) < 91 Then
    ShiftExpect = 1
    KeyExpect = Asc(Mid(.text1.Text, lt1, 1))
End If
' Lettre finale minuscule d'une combinaison
If Asc(Mid(.text1.Text, lt1, 1)) > 96 And Asc(Mid(.text1.Text, lt1, 1)) < 123 Then
    ShiftExpect = 0
    KeyExpect = Asc(Mid(.text1.Text, lt1, 1)) - 32
End If

' Touches de Fonction (attention au shiftexpect plus bas et modifi� par d�but de la combinaison)
' SANS Alt255 final
If Right(.text1.Text, 1) <> "�" Then
    If Right(UCase(.text1.Text), 2) = "F1" Then KeyExpect = 112
    If Right(UCase(.text1.Text), 2) = "F2" Then KeyExpect = 113
    If Right(UCase(.text1.Text), 2) = "F3" Then KeyExpect = 114
    If Right(UCase(.text1.Text), 2) = "F4" Then KeyExpect = 115
    If Right(UCase(.text1.Text), 2) = "F5" Then KeyExpect = 116
    If Right(UCase(.text1.Text), 2) = "F6" Then KeyExpect = 117
    If Right(UCase(.text1.Text), 2) = "F7" Then KeyExpect = 118
    If Right(UCase(.text1.Text), 2) = "F8" Then KeyExpect = 119
    If Right(UCase(.text1.Text), 2) = "F9" Then KeyExpect = 120
    If Right(UCase(.text1.Text), 3) = "F10" Then KeyExpect = 121
    If Right(UCase(.text1.Text), 3) = "F11" Then KeyExpect = 122
    If Right(UCase(.text1.Text), 3) = "F12" Then KeyExpect = 123
    If (KeyExpect > 111 And KeyExpect < 124) Then ShiftExpect = 0
End If

' Touches de Fonction (attention au shiftexpect plus bas et modifi� par d�but de la combinaison)
' AVEC Alt255 final
If Right(.text1.Text, 1) = "�" Then
    If Right(UCase(.text1.Text), 3) = "F1�" Then KeyExpect = 112
    If Right(UCase(.text1.Text), 3) = "F2�" Then KeyExpect = 113
    If Right(UCase(.text1.Text), 3) = "F3�" Then KeyExpect = 114
    If Right(UCase(.text1.Text), 3) = "F4�" Then KeyExpect = 115
    If Right(UCase(.text1.Text), 3) = "F5�" Then KeyExpect = 116
    If Right(UCase(.text1.Text), 3) = "F6�" Then KeyExpect = 117
    If Right(UCase(.text1.Text), 3) = "F7�" Then KeyExpect = 118
    If Right(UCase(.text1.Text), 3) = "F8�" Then KeyExpect = 119
    If Right(UCase(.text1.Text), 3) = "F9�" Then KeyExpect = 120
    If Right(UCase(.text1.Text), 4) = "F10�" Then KeyExpect = 121
    If Right(UCase(.text1.Text), 4) = "F11�" Then KeyExpect = 122
    If Right(UCase(.text1.Text), 4) = "F12�" Then KeyExpect = 123
    If (KeyExpect > 111 And KeyExpect < 124) Then ShiftExpect = 0
End If

' Partie du d�but d'une combinaison
If UCase(Left(.text1.Text, Len(vvMaj))) = UCase(vvMaj) Then ShiftExpect = 1
If UCase(Left(.text1.Text, Len(vvCtrl))) = UCase(vvCtrl) Then ShiftExpect = 2
If UCase(Left(.text1.Text, Len(vvControl))) = UCase(vvControl) Then ShiftExpect = 2
If UCase(Left(.text1.Text, Len(vvAlt))) = UCase(vvAlt) Then ShiftExpect = 4
If UCase(Left(.text1.Text, Len(vvCtrl) + 1 + Len(vvMaj))) = UCase(vvCtrl) & "+" & UCase(vvMaj) Then ShiftExpect = 3
If UCase(Left(.text1.Text, Len(vvControl) + 1 + Len(vvMaj))) = UCase(vvControl) & "+" & UCase(vvMaj) Then ShiftExpect = 3
If UCase(Left(.text1.Text, Len(vvCtrl) + 1 + Len(vvAlt))) = UCase(vvCtrl) & "+" & UCase(vvAlt) Then ShiftExpect = 6
If UCase(Left(.text1.Text, Len(vvControl) + 1 + Len(vvAlt))) = UCase(vvControl) & "+" & UCase(vvAlt) Then ShiftExpect = 6

'Cas des caract�res command�s par AltGr (AVEC ou SANS Alt255)
If Left(.text1.Text, lt1) = "@" Then KeyExpect = 48 'azerty
'If Left(.text1.Text, lt1) = "@" Then KeyExpect = 50 'qwertz

If Left(.text1.Text, lt1) = "~" Then KeyExpect = 50 'azerty
'If Left(.text1.Text, lt1) = "~" Then KeyExpect = 221 'qwertz

If Left(.text1.Text, lt1) = "#" Then KeyExpect = 51 'azerty et qwertz

If Left(.text1.Text, lt1) = "{" Then KeyExpect = 52 'azerty
'If Left(.text1.Text, lt1) = "{" Then KeyExpect = 220 'qwertz

If Left(.text1.Text, lt1) = "[" Then KeyExpect = 53 'azerty
'If Left(.text1.Text, lt1) = "[" Then KeyExpect = 186 'qwertz

If Left(.text1.Text, lt1) = "|" Then KeyExpect = 54 'azerty
'If Left(.text1.Text, lt1) = "|" Then KeyExpect = 55 'qwertz

If Left(.text1.Text, lt1) = "`" Then KeyExpect = 55 'azerty
'If Left(.text1.Text, lt1) = "`" Then
'    KeyExpect = 221: ShiftExpect = 1 'qwertz
'End If

If Left(.text1.Text, lt1) = "\" Then KeyExpect = 56 'azerty
'If Left(.text1.Text, lt1) = "\" Then KeyExpect = 226 'qwertz

If Left(.text1.Text, lt1) = "}" Then KeyExpect = 187 'azerty
'If Left(.text1.Text, lt1) = "}" Then KeyExpect = 223 'qwertz

If Left(.text1.Text, lt1) = "]" Then KeyExpect = 219 'azerty
'If Left(.text1.Text, lt1) = "]" Then KeyExpect = 192 'qwertz

If Left(.text1.Text, lt1) = "@" Or Left(.text1.Text, lt1) = "~" Or Left(.text1.Text, lt1) = "#" Or Left(.text1.Text, lt1) = "{" Or Left(.text1.Text, lt1) = "[" Then ShiftExpect = 6
If Left(.text1.Text, lt1) = "|" Or Left(.text1.Text, lt1) = "`" Or Left(.text1.Text, lt1) = "\" Or Left(.text1.Text, lt1) = "}" Or Left(.text1.Text, lt1) = "]" Then ShiftExpect = 6

' Autres caract�res (AVEC ou SANS Alt255)
If Left(.text1.Text, lt1) = "�" Then
    KeyExpect = 48: ShiftExpect = 0 'azerty
    'KeyExpect = 220: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "&" Then
    KeyExpect = 49: ShiftExpect = 0 'azerty
    'KeyExpect = 54: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "�" Then
    KeyExpect = 50: ShiftExpect = 0 'azerty
    'KeyExpect = 222: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = """" Then
    KeyExpect = 51: ShiftExpect = 0 'azerty
    'KeyExpect = 50: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "'" Then
    KeyExpect = 52: ShiftExpect = 0 'azerty
    'KeyExpect = 219: ShiftExpect = 0 'qwertz
End If
If Left(.text1.Text, lt1) = "(" Then
    KeyExpect = 53: ShiftExpect = 0 'azerty
    'KeyExpect = 56: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "-" And numpad <= 0 Then
    KeyExpect = 54: ShiftExpect = 0 'azerty
    'KeyExpect = 189: ShiftExpect = 0 'qwertz
End If
If Left(.text1.Text, lt1) = "-" And numpad >= 1 Then
    KeyExpect = 109: ShiftExpect = 0 'azerty et qwertz
End If
If Left(.text1.Text, lt1) = "�" Then
    KeyExpect = 55: ShiftExpect = 0 'azerty
    'KeyExpect = 186: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "_" Then
    KeyExpect = 56: ShiftExpect = 0 'azerty
    'KeyExpect = 189: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "�" Then
    KeyExpect = 57: ShiftExpect = 0 'azerty
    'KeyExpect = 52: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "$" Then
    KeyExpect = 186: ShiftExpect = 0 'azerty
    'KeyExpect = 223: ShiftExpect = 0 'qwertz
End If
If Left(.text1.Text, lt1) = "�" Then
    KeyExpect = 186: ShiftExpect = 1 'azerty
    'KeyExpect = 223: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "+" And numpad <= 0 Then
    KeyExpect = 187: ShiftExpect = 1 'azerty
    'KeyExpect = 49: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "+" And numpad >= 1 Then
    KeyExpect = 107: ShiftExpect = 0 'azerty et qwertz
End If
If Left(.text1.Text, lt1) = "=" Then
    KeyExpect = 187: ShiftExpect = 0 'azerty
    'KeyExpect = 48: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "," Then
    KeyExpect = 188: ShiftExpect = 0 'azerty et qwertz
End If
If Left(.text1.Text, lt1) = "?" Then
    KeyExpect = 188: ShiftExpect = 1 'azerty
    'KeyExpect = 219: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = ";" Then
    KeyExpect = 190: ShiftExpect = 0 'azerty
    'KeyExpect = 188: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "." And numpad <= 0 Then
    KeyExpect = 190: ShiftExpect = 1 'azerty
    'KeyExpect = 190: ShiftExpect = 0 'qwertz
End If
If Left(.text1.Text, lt1) = "." And numpad >= 1 Then
    KeyExpect = 110: ShiftExpect = 0 'azerty et qwertz
End If
If Left(.text1.Text, lt1) = "/" And numpad <= 0 Then
    KeyExpect = 191: ShiftExpect = 1 'azerty
    'KeyExpect = 55: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "/" And numpad >= 1 Then
    KeyExpect = 111: ShiftExpect = 0 'azerty et qwertz
End If
If Left(.text1.Text, lt1) = ":" Then
    KeyExpect = 191: ShiftExpect = 0 'azerty
    'KeyExpect = 190: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "�" Then
    KeyExpect = 192: ShiftExpect = 0 'azerty, n'existe pas sur qwertz
End If
If Left(.text1.Text, lt1) = "%" Then
    KeyExpect = 192: ShiftExpect = 1 'azerty
    'KeyExpect = 53: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = ")" Then
    KeyExpect = 219: ShiftExpect = 0 'azerty
    'KeyExpect = 57: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "�" Then
    KeyExpect = 219: ShiftExpect = 1 'azerty
    'KeyExpect = 191: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "*" And numpad <= 0 Then
    KeyExpect = 220: ShiftExpect = 0 'azerty
    'KeyExpect = 51: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "*" And numpad >= 1 Then
    KeyExpect = 106: ShiftExpect = 0 'azerty et qwertz
End If
If Left(.text1.Text, lt1) = "�" Then
    KeyExpect = 220: ShiftExpect = 1 'azerty, n'existe pas sur qwertz
End If
If Left(.text1.Text, lt1) = "^" Then
    KeyExpect = 221: ShiftExpect = 0 'azerty et qwertz
End If
If Left(.text1.Text, lt1) = "�" Then
    KeyExpect = 221: ShiftExpect = 1 'azerty
    'KeyExpect = 192: ShiftExpect = 0 'qwertz
End If
If Left(.text1.Text, lt1) = "�" Then
    KeyExpect = 222: ShiftExpect = 0 'azerty, n'existe pas sur qwertz
End If
If Left(.text1.Text, lt1) = "!" Then
    KeyExpect = 223: ShiftExpect = 0 'azerty
    'KeyExpect = 192: ShiftExpect = 1 'qwertz
End If
If Left(.text1.Text, lt1) = "�" Then
    KeyExpect = 223: ShiftExpect = 1 'azerty
    'KeyExpect = 191: ShiftExpect = 0 'qwertz
End If
If Left(.text1.Text, lt1) = "<" Then
    KeyExpect = 226: ShiftExpect = 0 'azerty et qwertz
End If
If Left(.text1.Text, lt1) = ">" Then
    KeyExpect = 226: ShiftExpect = 1 'azerty et qwertz
End If

'If .text1.Text = "�" Then keyexpect = 129
'If .text1.Text = "�" Then keyexpect = 131
'If .text1.Text = "�" Then keyexpect = 132
'If .text1.Text = "�" Then keyexpect = 136
'If .text1.Text = "�" Then keyexpect = 137
'If .text1.Text = "�" Then keyexpect = 139
'If .text1.Text = "�" Then keyexpect = 140
'If .text1.Text = "�" Then keyexpect = 144
'If .text1.Text = "�" Then keyexpect = 147
'If .text1.Text = "�" Then keyexpect = 148
'If .text1.Text = "�" Then keyexpect = 150

' Touches sp�ciales  (AVEC ou SANS Alt255) (attention au shiftexpect plus bas)
If UCase(Left(.text1.Text, lt1)) = UCase(vvEntr�e) Then KeyExpect = 13
If UCase(Left(.text1.Text, lt1)) = UCase(vvControl) Then KeyExpect = 17
If UCase(Left(.text1.Text, lt1)) = UCase(vvControlGauche) Then KeyExpect = 17
If UCase(Left(.text1.Text, lt1)) = UCase(vvControlDroit) Then KeyExpect = 17
If UCase(Left(.text1.Text, lt1)) = UCase(vvAlt) Then KeyExpect = 18
If UCase(Left(.text1.Text, lt1)) = UCase(vvAltGr) Then KeyExpect = 255 'Cas particulier 17 puis 18
If UCase(Left(.text1.Text, lt1)) = UCase(vv�chap) Then KeyExpect = 27
If UCase(Left(.text1.Text, lt1)) = UCase(vvEchap) Then KeyExpect = 27
If UCase(Left(.text1.Text, lt1)) = UCase(vvEspace) Then KeyExpect = 32
If UCase(Left(.text1.Text, lt1)) = UCase(vvFlecheGauche) Then KeyExpect = 37
If UCase(Left(.text1.Text, lt1)) = UCase(vvFlecheHaut) Then KeyExpect = 38
If UCase(Left(.text1.Text, lt1)) = UCase(vvFlecheDroite) Then KeyExpect = 39
If UCase(Left(.text1.Text, lt1)) = UCase(vvFlecheBas) Then KeyExpect = 40
If UCase(Left(.text1.Text, lt1)) = UCase(vvWindowsGauche) Then KeyExpect = 91
If UCase(Left(.text1.Text, lt1)) = UCase(vvWindowsDroit) Then KeyExpect = 92
If UCase(Left(.text1.Text, lt1)) = UCase(vvMenuContextuel) Then KeyExpect = 93
If UCase(Left(.text1.Text, lt1)) = UCase(vvVerrouillageMajuscules) Then KeyExpect = 20
If UCase(Left(.text1.Text, lt1)) = UCase(vvMaj) Then KeyExpect = 16
If UCase(Left(.text1.Text, lt1)) = UCase(vvMajGauche) Then KeyExpect = 16
If UCase(Left(.text1.Text, lt1)) = UCase(vvMajDroit) Then KeyExpect = 16
If UCase(Left(.text1.Text, lt1)) = UCase(vvInsertion) Then KeyExpect = 45
If UCase(Left(.text1.Text, lt1)) = UCase(vvSuppression) Then KeyExpect = 46
If UCase(Left(.text1.Text, lt1)) = UCase(vvD�but) Then KeyExpect = 36
If UCase(Left(.text1.Text, lt1)) = UCase(vvFin) Then KeyExpect = 35
If UCase(Left(.text1.Text, lt1)) = UCase(vvPagePr�c�dente) Then KeyExpect = 33
If UCase(Left(.text1.Text, lt1)) = UCase(vvPageSuivante) Then KeyExpect = 34
If UCase(Left(.text1.Text, lt1)) = UCase(vvTab) Then KeyExpect = 9
If UCase(Left(.text1.Text, lt1)) = UCase(vvRetourArriere) Or LCase(Left(.text1.Text, lt1)) = LCase(vvRetourArri�re) Then KeyExpect = 8
If KeyExpect < 47 Then ShiftExpect = 0

' Touches du pav� num�rique (AVEC ou SANS Alt255)
If UCase(Left(.text1.Text, lt1)) = UCase(vvVerrouillageNum�rique) Then KeyExpect = 144
If KeyExpect = 144 Then ShiftExpect = 0
If UCase(Left(.text1.Text, lt1)) = UCase(vvPlus) Then KeyExpect = 107
If UCase(Left(.text1.Text, lt1)) = UCase(vvMoins) Then KeyExpect = 109
If UCase(Left(.text1.Text, lt1)) = UCase(vvTiret) Then KeyExpect = 109
If UCase(Left(.text1.Text, lt1)) = UCase(vvDiviser) Then KeyExpect = 111
If UCase(Left(.text1.Text, lt1)) = UCase(vvBarreOblique) Then KeyExpect = 111
If UCase(Left(.text1.Text, lt1)) = UCase(vvMultiplier) Then KeyExpect = 106
If UCase(Left(.text1.Text, lt1)) = UCase(vv�toile) Then KeyExpect = 106
If UCase(Left(.text1.Text, lt1)) = UCase(vvAst�risque) Then KeyExpect = 106
If UCase(Left(.text1.Text, lt1)) = UCase(vvPoint) Then KeyExpect = 110
If KeyExpect > 105 And KeyExpect < 112 Then ShiftExpect = 0

' Taux de r�ussite
pctt = 100 * (nbcaras / (nbcaras + iwrong))

' Faire appara�tre �ventuellement l'indication MAJUSCULE (SANS Alt255)
If Right(.text1.Text, 1) <> "�" Then
    On Error Resume Next
    If Len(.text1.Text) = 1 And Asc(Mid(.text1.Text, ii + 1, 1)) >= 65 And Asc(Mid(.text1.Text, ii + 1, 1)) <= 90 Then
        .text3.Text = "�" & .text1.Text & vvMajuscule ' (Alt255 devant pour visibilit�)
        .text3.Width = 0.32 * .Width
        .text3.SelStart = 0
        .text3.SelLength = Len(.text3.Text)
        .text3.Visible = True
    End If
End If

' Faire appara�tre �ventuellement l'indication MAJUSCULE (AVEC Alt255)
If Right(.text1.Text, 1) = "�" Then
    On Error Resume Next
    If Len(.text1.Text) = 2 And Asc(Mid(.text1.Text, ii + 1, 1)) >= 65 And Asc(Mid(.text1.Text, ii + 1, 1)) <= 90 Then
        .text3.Text = "�" & .text1.Text & vvMajuscule ' (Alt255 devant pour visibilit�)
        .text3.Width = 0.32 * .Width
        .text3.SelStart = 0
        .text3.SelLength = Len(.text3.Text)
        .text3.Visible = True
    End If
End If

' FOCUS sur text2
On Error Resume Next
.text2.SetFocus
End With
End Sub


' ***************** TEXT2TEXT1 : compare text2 � text1 **************************************
' **** Routine pour accepter/refuser les r�ponses utilisateur dans toutes le�ons sauf 1 et 13 *******
Public Sub text2text1(indif, sonocara, alea, timevalid, concat, pass)
' indif=1 signifie indiff�rent � majuscule/minuscule
' sonocara=1 signifie sonorisation lettre par lettre
' alea=1 signifie que le text1 est charg� en mode random
' timevalid=1 ou 2 signifie que l'on compte le temps elapsed
' concat=1 signifie qu'on n'efface pas l'ancien text1
' pass=1 signifie qu'on passe � la ligne suivante d�s la premi�re erreur de frappe
' pass=2 signifie qu'on redonne la m�me ligne si l'utilisateur a fait 2 erreurs dans la ligne

' Pour une meilleure sonorisation par Jaws de la derni�re lettre d'une ligne,
' ApprentiClavier rajoute un blanc dur Alt255 � la fin de la ligne ou apr�s le caract�re demand�.

If t2inhibit = 1 Then Exit Sub
concatf = concat
With le�on_courante

' ll est le nombre de caract�res dans la ligne envoy�e � l'utilisateur
ll = Len(.text1.Text)

' ii est le nombre de caract�res d�j� envoy� par l'utilisateur
ii = Len(.text2.Text)
If ii > 0 Then

' Si la nouvelle ligne est un simple "� la ligne", et si on tape un cara, c'est une faute � d�tecter
If Len(currentline) = 0 And Mid(.text2.Text, ii, 1) <> Chr(10) Then GoTo LS4

' Simple ENTR�E UTILISATEUR, ce qui cr�e 2 caras 13 10 d'un coup
    If Mid(.text2.Text, ii, 1) = Chr(10) Then
        ii = ii - 2
        If ii < 0 Then ii = 0
    
        ' CAS "� la ligne"
        If concat = 1 And ii = zz Then
            GoTo LS1
        
        ' CAS erreur utilisateur
        Else
            Module_routines.bip le�on_courante
            If pass = 1 Then
                pass = 0
                GoTo LS1
            End If
            mm = .text1.SelLength
            .text1.SelLength = 0
            Call Sleep(cadencecara)
            .text1.SelLength = mm
            t2inhibit = 1
            .text2.Text = Left(.text1.Text, ii)
            .text2.SelStart = ii
            t2inhibit = 0
            Exit Sub
        End If
    End If
End If

' text2 doit R�PLIQUER text1
On Error Resume Next
cara1 = Mid(.text1.Text, ii, 1)
On Error Resume Next
cara2 = Mid(.text2.Text, ii, 1)

'Cas PAV� NUM�RIQUE, attention astuce, cara2 ESPACE doit permettre de R�P�TER
If numpad = 1 And cara2 <> " " Then  ' Ainsi numpad = 2 accepte clavier principal et pav� num�rique
    If cara1 = "0" And keyforce <> 96 Then cara2 = "@"
    If cara1 = "1" And keyforce <> 97 Then cara2 = "@"
    If cara1 = "2" And keyforce <> 98 Then cara2 = "@"
    If cara1 = "3" And keyforce <> 99 Then cara2 = "@"
    If cara1 = "4" And keyforce <> 100 Then cara2 = "@"
    If cara1 = "5" And keyforce <> 101 Then cara2 = "@"
    If cara1 = "6" And keyforce <> 102 Then cara2 = "@"
    If cara1 = "7" And keyforce <> 103 Then cara2 = "@"
    If cara1 = "8" And keyforce <> 104 Then cara2 = "@"
    If cara1 = "9" And keyforce <> 105 Then cara2 = "@"
End If

'Cas d'une dict�e import�e contenant des apostrophes de code 146 (vari�t� d'apostrophe)
If Asc(cara1) = 146 And cara2 = "'" Then cara2 = cara1

' BONNE R�PONSE !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
On Error Resume Next
If (cara2 = cara1 And indif = 0) Or (UCase(cara1) = UCase(cara2) And indif = 1) Then
    
    ' RESET
    ff = 0 ' Flag faute � 0
    iwrongbis = 0
    .text4.Text = ""
    .text4.Visible = False
    .Picture1.Visible = False
    If .text1.Text <> " " Then ' lignesuivante contenant un simple espace : non compt�e
        nbonscaras = nbonscaras + 1
    End If
    If alea = 0 Then nbcaras = nbcaras + 1
    
    ' TAUX de r�ussite
    If alea = 0 Then Module_routines.scoreaffich le�on_courante, alea
    If alea = 1 And nbcaras > 0 And nbonscaras >= 0 Then pctt = 100 * (nbonscaras / nbcaras)
    
    If concat = 1 And ii = ll - 1 Then
        Module_routines.text3fill le�on_courante
    End If
    
    ' PASSAGE FORC� � la LIGNE si on arrive au cara Alt255
    If Mid(.text1.Text, ii + 1, 1) = "�" Then GoTo LS1
    
    ' PASSAGE �ventuel � la PARTIE SUIVANTE de la MEME CURRENTLINE
    If concat = 1 And ii = ll And ll < zz Then
        iistart = iistop + 1
        
        'D�tecter la fin de la partie suivante par un ".", un "...", un "!", un "?".
        iistop = InStr(iistart, currentline, ".")
        iistop2 = InStr(iistart, currentline, "!")
        iistop3 = InStr(iistart, currentline, "?")
        
        'D�tecter la fin de la phrase
        If typele�on >= 3 Then
            Module_routines.DetectPhraseEnd
        Else
            If iistop = 0 Then iistop = Len(currentline)
        End If
        
        'D�finir le texte de la partie suivante
        .text1.Text = .text1.Text & Mid(currentline, iistart, iistop - iistart + 1)
        If alea = 1 Then nbcaras = nbcaras + Len(Mid(currentline, iistart, iistop - iistart))
        llold = ll
        ll = Len(.text1.Text)
        iistartp = ii
        .text1.SelStart = ii
        .text1.SelLength = ll - ii
        Call Sleep(cadenceligne)
        GoTo LS2
    End If
            
    ' AVANT le PASSAGE � la LIGNE SUIVANTE
    ' Au lieu de donner un CR sur "� la ligne", on a donn� un cara, revenir en arri�re
    If ii > ll Then
        ' Espace n'est pas une erreur
        If Mid(.text2.Text, ii, 1) <> " " Then iwrongCR = iwrongCR + 1
        If iwrongCR < iwrongCRmax Then
            .text2.Text = Left(.text2.Text, Len(.text2.Text) - 1)
        Else
            .text2.Text = Left(.text1.Text, ii) & Chr(13) & Chr(10)
            iiante = 0: iiprec = 0: iwrongCR = 0
        End If
        ii = Len(.text2.Text)
        .text2.SelStart = ii
    End If
    
    If ii = ll Then
        'En mode concat, prononcer "� la ligne" et attendre le CR
        If concat = 1 Then
            Module_routines.text3fill le�on_courante
            Exit Sub
        End If
    
LS1:
        'En mode pass=2, une double faute fait redonner la ligne en cours (jusqu'� irecur fois)
        If pass = 2 And irecur < 2 Then
            If iwrongl >= 2 Then
                irecur = irecur + 1
                iwrongl = 0: iwrongbis = 0
                .text2.Text = ""
                tempo = .text1.Text
                .text1.Text = ""
                le�on_courante.Cls
                Call Sleep(cadenceligne)
                .text1.Text = tempo
                t2inhibit = 0
                .text1.SelStart = 0
                .text1.SelLength = Len(.text1.Text)
                Exit Sub
            End If
        End If
        
        'Prononcer et dimensionner auparavant l'�ventuel MESSAGE "Continuez�"
        If Not msgtext1(iter) = "" Then
            .label1.Visible = False
            .text1.Text = msgtext1(iter)
            currentline = .text1.Text
            Module_routines.AdjustWidthAndSize le�on_courante, 0
            .text1.SelStart = 0
            .text1.SelLength = Len(msgtext1(iter))
            If Len(msgtext1(iter)) < 10 Then
                Call Sleep(800)
            Else
                Call Sleep(1200)
            End If
        End If
        .label1.Visible = True
        msgtext1(iter) = ""
        
        'Laisser voir bri�vement la frappe text2 si text2 fait un seul caract�re
        If Len(.text2.Text) = 1 Then Call Sleep(100)
        
        'Prononcer auparavant l'�ventuelle BOITE de DIALOGUE
        If Not msgtext2(iter) = "" Then
T1:
            pagenum = 0
            msgtext0 = msgtext2(iter) + pressez
            Msgform.Show 1
            If msgf = 2 Then GoTo T1
            If msgf = 1 Then
                msgtext2(iter) = ""
                'Shell vpath & "sonbip2.exe", 0
                'If Dir(vpath & "sonbip2.exe") <> "" Then Module_exec.ExecAndWait vpath & "sonbip2.exe"
                Module_routines.sonbip2tons 'avril 2008
            End If
            If msgf = 0 Then
                msgtext2(iter) = ""
                .Quitter_Click
                'SendKeys "{ESC}"
                'Sendkeys est remplac� par des actions keybd_event pour Windows Vista juin 2007
                keybd_event VK_ESCAPE, 0, 0, 0
                keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
                Exit Sub
            End If
        End If
        
        ' PASSAGE � la LIGNE SUIVANTE
        Module_routines.lignesuivante alea, timevalid, concat
        iter = iter + 1
        If derligne = 2 Then
            derligne = 0
            Exit Sub
        End If
        llold = ll
        ll = Len(.text1.Text): ii = Len(.text2.Text)
    End If

    ' PREPARER la SELECTION du MOT ou CARA text1
    ' Ne pas DUPLIQUER la SONO quand on se trouve au d�but d'une "nouvelle" ligne (attention � Alt255)
    If ((Len(.text1.Text) > 1 And Right(.text1.Text, 1) <> "�") Or (Len(.text1.Text) > 2 And Right(.text1.Text, 1) = "�")) And sonocara = 0 Then
        If iter = 0 Then
            .text1.SelStart = ii
        End If
        If iter > 0 And concat = 0 And ii > 0 Then
            .text1.SelStart = ii
        End If
        If iter > 0 And concat = 1 And ii = llold + 2 Then
            .text1.SelStart = ii
            .text1.SelLength = 0: Call Sleep(cadencemot)
        End If
        If iter > 0 And concat = 1 And ii > llold + 2 Then
            .text1.SelStart = ii
        End If
    End If
        
    ' SELECTION du CARA text1 en cours
    ' SELECTION CARA UTILE seulement s'il y a PLUSIEURS CARAS (attention � Alt255)
    If ((Len(.text1.Text) > 1 And Right(.text1.Text, 1) <> "�") Or (Len(.text1.Text) > 2 And Right(.text1.Text, 1) = "�")) And sonocara = 1 Then
        Call Sleep(cadencecara)   ' D�finit vraiment la cadence du prochain cara
        If ii = 0 And Mid(.text1.Text, 2, 1) = " " Then
        Else
            .text1.SelStart = ii
            If ii > 0 Then
                If Mid(.text1.Text, ii + 2, 1) = "�" Then
                    .text1.SelLength = 2
                Else
                    .text1.SelLength = 1
                End If
            End If
        End If
    End If

    ' SELECTION MOT de text1, si on est APR�S UN ESPACE ou en DEBUT de LIGNE(Mid avec ii=0 resume next)
    If ll > 0 Then
        On Error Resume Next
        If Mid(.text1.Text, ii, 1) = " " Or ii = llold + 2 Then
LS2:
            ' SELECTION MOT par MOT UTILE seulement s'il y a PLUSIEURS MOTS
            If InStr(currentline, " ") > 0 Then Module_routines.nextspace le�on_courante
            
            ' SELECTION du CARA text1 apr�s avoir s�lectionn� le MOT
            ' SELECTION CARA UTILE seulement s'il y a PLUSIEURS VRAIS CARAS (attention � Alt255)
            If ((Len(.text1.Text) > 1 And Mid(.text1.Text, 2, 1) <> "�") Or (Len(.text1.Text) > 2 And Mid(.text1.Text, 2, 1) = "�")) And sonocara = 1 Then
                
                ' Pas de s�lection si le mot a un seul vrai cara (kk=2 sans un Alt255)
                If kk = 2 And Mid(.text1.Text, 2, 1) <> "�" Then
                ' Pas de s�lection si le mot a un seul vrai cara (kk=3 avec un Alt255)
                ElseIf kk = 3 And Mid(.text1.Text, 2, 1) = "�" Then
                ' Sinon on s�lectionne le cara
                Else
                    .text1.SelLength = 0: Call Sleep(50)  ' Attention pas cadencecara en Win 98
                    .text1.SelLength = 1:
                End If
            End If
            
            ' FOCUS sur text2
            On Error Resume Next
            .text2.SetFocus
        End If
    End If
    
    ' SIGNALER caras non sonoris�s ESPACE, MAJUSCULE... (et compter les ESPACES)
    Module_routines.text3fill le�on_courante

Else
        ' TOUCHES d'AIDE pour l'utilisateur
        ' "TOUCHE" ESPACE = RESELECTION du CARA de text1 en cours
    If cara2 = " " Then
        ii = ii - 1
        .text1.SelStart = ii: .text1.SelLength = 0
        Call Sleep(cadencecara)
        If .text3.Text = "�" & vvAlaligne Then
            .text3.Text = ""
            .text3.Visible = False
            Call Sleep(cadencemot)
            .text3.Width = 0.2 * .Width
            .text3.Text = "�" & vvAlaligne ' (Alt255 devant pour visibilit�)
            .text3.SelStart = 0
            .text3.SelLength = 10
            .text3.Visible = True
        End If
        
        
        ' "TOUCHE" MAJ+ESPACE = RESELECTION de la (FIN de la) LIGNE de text1 en cours
        If lrepeat = 1 Then
            lrepeat = 0
            If concatf = 0 Then
                iistartp = ii
                '.text1.SelStart = 0
                .text1.SelStart = ii
                .text1.SelLength = Len(.text1.Text): Call Sleep(cadenceligne)
                .text1.SelStart = iistartp
                .text1.SelLength = 0
            End If
            If concatf = 1 Then
                '.text1.SelStart = iistartp
                '.text1.SelLength = Len(.text1.Text) - iistartp: Call Sleep(cadenceligne)
                .text1.SelStart = ii
                .text1.SelLength = Len(.text1.Text) - ii: Call Sleep(cadenceligne)
                .text1.SelStart = Len(.text1.Text)
                .text1.SelLength = 0
            End If
        GoTo LS3
        End If
        
        '"TOUCHE" CONTROL+ESPACE = RESELECTION du MOT de text1 en cours
        If wrepeat = 1 Then
            wrepeat = 0
            Module_routines.nextspace le�on_courante
            .text1.SelLength = kk - 1: Call Sleep(cadencemot)
        
        ' SUITE d'un APPUI sur ESPACE
        Else
            .text1.SelStart = ii
            ' Surbrillance �largie si Alt255 suit
            If Mid(.text1.Text, ii + 2, 1) <> "�" Then .text1.SelLength = 1
            If Mid(.text1.Text, ii + 2, 1) = "�" Then .text1.SelLength = 2
        End If
        
LS3:
        ' SUITE pour toutes combinaisons avec ESPACE
        t2inhibit = 1
        .text2.Text = Left(.text1.Text, ii)
        t2inhibit = 0
        If ll > 0 Then .text2.SelStart = Len(.text2.Text)
        Exit Sub
    End If
    
    ' !!! FAUTE DE FRAPPE (sauf retour chariot) !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    Module_routines.bip le�on_courante
    
    ' Enregistrement des fautes
    If iwrong < 150 And cara1 <> " " Then
        fautesur(iwrong) = cara1
        nboccur(iwrong) = 1
    End If
        
    ' CAS RANDOM (le�on6), Faute de frappe apr�s le premier cara
    If alea = 1 And ii > 1 Then
        Module_routines.lignesuivante 1, timevalid, concat
        iter = iter + 1
        If derligne = 2 Then
            derligne = 0
            Exit Sub
        End If
        ll = Len(.text1.Text): ii = Len(.text2.Text)
        Exit Sub
    End If
    
LS4:
    ' CAS NORMAL, FAUTE de frappe
    iwrong = iwrong + 1: iwrongbis = iwrongbis + 1: iwrongl = iwrongl + 1
    ff = 1 ' Flag qu'on vient de faire une faute
    
    ' Pour affichage imm�diat, sauf pour le�ons 6, 7, 13, et dict�es 14, 15, 17
    If alea = 0 Then Module_routines.scoreaffich le�on_courante, alea
        
    ' Reset de iwrongbis si l'utilisateur progresse
    iiprec = ii
    If iiprec > iiante Then
        iiante = iiprec
        iwrongbis = 0
    End If
    
    ' Faute sur ESPACE, prononce � nouveau ESPACE
    If cara1 = " " Then
        On Error Resume Next
        .text2.SetFocus
        .text3.Text = "�" & vvEspace ' (Alt255 devant pour visibilit�)
        .text3.Width = 0.18 * .Width
        .text3.SelStart = 0
        .text3.SelLength = Len(.text3.Text)
        .text3.Visible = True
    Else
        
        ' Faute sur "� la ligne"
        If .text3.Text = "�" & vvAlaligne Then ' (Alt255 devant pour visibilit�)
            On Error Resume Next
            .text2.SetFocus
            .text3.Text = ""
            .text3.Visible = False
            Call Sleep(cadencemot)
            .text3.Width = 0.2 * .Width
            .text3.Text = "�" & vvAlaligne ' (Alt255 devant pour visibilit�)
            .text3.SelStart = 0
            .text3.SelLength = Len(.text3.Text)
            .text3.Visible = True
        Else
        
            ' Faute sur cara par erreur Majuscule/Minuscule
            If Asc(cara2) = Asc(cara1) + 32 And indif = 0 Then
                On Error Resume Next
                .text2.SetFocus
                .text3.Text = "�" & cara1 & vvMajuscule ' (Alt255 devant pour visibilit�)
                .text3.Width = 0.32 * .Width
                .text3.SelStart = 0
                .text3.SelLength = Len(.text3.Text)
                .text3.Visible = True
            Else
            
                ' Faute sur cara par erreur Majuscule/Minuscule
                If (Asc(cara2) = Asc(cara1) - 32 And indif = 0) Or (cara2 = "1" And cara1 = "&") Or (cara2 = "2" And cara1 = "�") Or (cara2 = "3" And cara1 = """") Or (cara2 = "4" And cara1 = "'") Or (cara2 = "5" And cara1 = "(") Or (cara2 = "6" And cara1 = "-") Or (cara2 = "7" And cara1 = "�") Or (cara2 = "8" And cara1 = "_") Or (cara2 = "9" And cara1 = "�") Or (cara2 = "0" And cara1 = "�") Or (cara2 = "?" And cara1 = ",") Or (cara2 = "." And cara1 = ";") Or (cara2 = "/" And cara1 = ":") Then
                    On Error Resume Next
                    .text2.SetFocus
                    .text3.Text = "�" & cara1 & vvMinuscule ' (Alt255 devant pour visibilit�)
                    .text3.Width = 0.32 * .Width
                    .text3.SelStart = 0
                    .text3.SelLength = Len(.text3.Text)
                    .text3.Visible = True
                Else
            
                    ' AUTRE FAUTE sur cara, prononce � nouveau le cara � taper ou le mot
                    .text1.SelLength = 0
                    
                    ' Erreur sur la premi�re lettre mot le�on6, r�p�ter le mot entier
                    If alea = 1 And ii = 1 Then
                        .text1.SelLength = Len(.text1.Text)
                        .text3.Text = "�" & cara1 ' (Alt255 devant pour visibilit�)
                        .text3.Width = 0.06 * .Width
                        .text3.Visible = True
                        
                    ' Cas g�n�ral, prononce le cara � taper, illumine l'�ventuel Alt255
                    Else
                        Call Sleep(cadencecara)
                        If Len(.text1.Text) = 2 And Right(.text1.Text, 1) = "�" Then
                            .text1.SelLength = 2
                        Else
                            .text1.SelLength = 1
                        End If
                        .text3.Text = "�" & cara1 ' (Alt255 devant pour visibilit�)
                        .text3.Width = 0.06 * .Width
                        .text3.Visible = True
                    End If
                End If
            End If
        End If
    End If
        
    ' RESET de text2 APR�S le NOMBRE MAX de FAUTES sur le cara !!!
    If iwrongbis >= iwrongbismax - 1 Then
        If .text3.Text <> vvAlaligne Then
            .text2.Text = Left(.text1.Text, ii)
            iiante = 0: iiprec = 0
        Else
            .text2.Text = Left(.text1.Text, ii) & Chr(13) & Chr(10)
            iiante = 0: iiprec = 0
        End If
    Else
        t2inhibit = 1
        .text2.Text = Left(.text1.Text, ii - 1)
        t2inhibit = 0
    End If
    
    ' Fin du reset
    If ll > 0 Then .text2.SelStart = Len(.text2.Text)
End If
End With
End Sub


' ******* LIGNESUIVANTE : Am�ne ligne suivante text1, jusqu'au SCORE et � QUITTER *******
Public Sub lignesuivante(alea, timevalid, concat)
' alea=1 signifie que le text1 est charg� en mode random
' timevalid=1 ou 2 signifie que l'on compte le temps elapsed
' concat=1 signifie qu'on n'efface pas l'ancien text1
With le�on_courante

    'Affichage status avant chargement de la ligne, pour la le�on6 o� alea=1
    If .text1.Text <> " " Then ' lignesuivante contenant un simple espace : non compt�e
        nbmots = nbmots + 1
    End If
    If alea = 1 And Len(.text1.Text) > 1 Then
        If Right(.text1.Text, 1) <> "�" Then
            nbcaras = nbcaras + Len(.text1.Text)
        Else
            nbcaras = nbcaras + Len(.text1.Text) - 1
        End If
        If nbcaras > 0 And nbonscaras >= 0 Then pctt = 100 * (nbonscaras / nbcaras)
        If nbonscaras >= 0 Then scorecourant = " " & CInt(pctt) & " %.    " & nbmots & msgMots & nbcaras & " caras.   " & nbonscaras & " bons caras."
        If nbonscaras < 0 Then scorecourant = " " & CInt(pctt) & " %.    " & nbmots & msgMots & nbcaras & " caras."
        .text3.Text = ""
        On Error Resume Next
        .text5.Text = scorecourant
    End If
    If alea = 0 Then
        On Error Resume Next
        pctt = 100 * (nbcaras / (nbcaras + iwrong))
        
        'Affichage status avant chargement de la ligne, sauf pour le�ons 6, 7, 13, et dict�es 14, 15, 17
        If (typele�on <= 3) And timevalid = 0 Then
            scorecourant = CInt(pctt) & " %."
            On Error Resume Next
            .text5.Text = scorecourant
        End If
    End If
       
    'RESET text2 et text1
    If concat = 0 Then
        t2inhibit = 1
        .text2.Text = "": .text1.Text = ""
        le�on_courante.Cls
        t2inhibit = 0
    End If
    
    'Autres RESET
    iiante = 0: iiprec = 0: iwrongl = 0: irecur = 0
    .text4.Visible = False
    erepeat = 0
    
    'DATER le MOMENT o� la nouvelle ligne arrive
    startline = Now
    elapsedtot = elapsedtot + elapsed
    elapsed = 0  ' indispensable tous cas
    
    ' CAS NORMAL, AMENER la LIGNE SUIVANTE
    If alea = 0 Then
        If Not EOF(1) Then
            Line Input #1, currentline
            If typele�on <= 3 And Len(currentline) = 1 Then .text1.Visible = False
            Call Sleep(cadenceligne)
            
            ' Si la ligne est vide, signaler "� la ligne"
            If Len(currentline) = 0 And concatf = 0 Then Line Input #1, currentline
            
            'Ajuster width, left, font.size pour le cas du text1 de 1 cara (sauf pour dict�es)
            If typele�on <> 14 Then Module_routines.AdjustWidthAndSize le�on_courante, 1
        
            ' AMENER ou RAJOUTER le TEXTE de la LIGNE SUIVANTE
            iwrongCR = 0
            iistart = 0
            
            'D�tecter la fin de la partie suivante par un ".", un "...", un "!", un "?".
            iistop = InStr(currentline, ".")
            iistop2 = InStr(currentline, "!")
            iistop3 = InStr(currentline, "?")
            
            'D�tecter la fin de la phrase
            If typele�on >= 3 Then
                Module_routines.DetectPhraseEnd
            Else
                If iistop = 0 Then iistop = Len(currentline)
            End If
            
            'D�finir le texte de la partie suivante
            If concat = 1 And ii = zz Then
                .text1.Text = .text1.Text & CRLF & Left(currentline, iistop)
            Else
                If pasdepoint = 0 Then .text1.Text = .text1.Text & Left(currentline, iistop)
                If pasdepoint = 1 Then .text1.Text = .text1.Text & currentline
            End If
            
            ' MISE � JOUR (ne pas oublier caras 10 13 dans zz et ii)
            zz = zz + 2 + Len(currentline)
            iistartp = ii + 2
            If concat = 0 Then .text1.SelStart = 0
            If concat = 1 Then .text1.SelStart = iistartp
            .text1.SelLength = iistop
            
            ' �VITER DOUBLE SONO sur le�on1 en Win 98, et le�ons 10A, 10B, 10C, 18C, 18D !!!
            If ((Len(currentline) > 1 And Right(currentline, 1) <> "�") Or (Len(currentline) > 2 And Right(currentline, 1) = "�")) And typele�on > 1 And nodoublesono = 0 Then Call Sleep(cadenceligne)
            
            ' �VITER DOUBLE SONO sur les dict�es lorsque la ligne suivante n'a qu'un mot !!!
            If typele�on = 14 And InStr(currentline, " ") > 0 And nodoublesono = 0 Then Call Sleep(cadenceligne)
        End If
    
        ' CAS NORMAL, QUITTER � la FIN de la DERNI�RE LIGNE du text1
        If EOF(1) And derligne = 1 Then
            If .text3.Text = "�" & vvAlaligne Then keyinhibit = 2
            Module_routines.score timevalid
            Exit Sub
        End If
    
        ' CAS NORMAL, D�tecter qu'on a charg� la derni�re ligne
        If EOF(1) Then
        derligne = 1
        End If
    End If
    
    ' CAS RANDOM, AMENER la LIGNE SUIVANTE de la TABLE datatext1 [1+(nblignes/4) fois]
    If alea = 1 Then
        Call Sleep(cadenceligne)
        If iter < Int(1 + (nbli / 2.5)) Then
            If bascule = 0 Then
                .text1.Text = " "
                .text3.Text = "�" & vvEspace ' (Alt255 devant pour visibilit�)
                .text3.Width = 0.18 * .Width
                .text3.SelStart = 0
                .text3.SelLength = Len(.text3.Text)
                .text3.Visible = True
                bascule = 1
            Else
                .text1.Text = datatext1(Int((nbli * Rnd) + 1))
                .text3.Visible = False
                bascule = 0
            End If
            .text1.SelStart = 0
            .text1.SelLength = Len(.text1.Text)
        End If
    
        ' CAS RANDOM, QUITTER la TABLE datatext1
        If iter >= Int(1 + (nbli / 2.5)) Then Module_routines.score timevalid
    End If
    
    ' RESET FINAL
    If typele�on <= 3 Then .text1.Visible = True
End With
End Sub


' ******************  SCORE  *************************************************************
Public Sub score(timevalid)
Unload le�on_courante

' Compter les fautes
For kk = 0 To 149
    For jj = 0 To 149
        If fautesur(jj) <> "" And jj <> kk Then
            vcomp = StrComp(fautesur(kk), fautesur(jj), 0)
            If vcomp = 0 Then
                nboccur(kk) = nboccur(kk) + 1
                fautesur(jj) = ""
            End If
        End If
    Next jj
Next kk

' Trier les 5 fautes les plus nombreuses, les mettre dans fautes.txt, avec des "f" et "j"
msgtext0 = ""
nn = 0
Close #3
Open vfile & "\fautes.txt" For Output As #3
fauteprec = ""
For jj = 1 To 50
    mm = 51 - jj
    For kk = 0 To 149
        If nboccur(kk) > mm And nn < 5 Then
            fautecourante = fautesur(kk)
            If fauteprec <> "" Then
                Print #3, "f" & CRLF & fautesur(kk) & CRLF & fauteprec & CRLF & fautesur(kk) & CRLF & "j" & CRLF & fautesur(kk) & CRLF & fauteprec
            Else
                Print #3, "f" & CRLF & fautesur(kk) & CRLF & "j" & CRLF & fautesur(kk)
            End If
            
            ' Les mettre aussi dans msgtext0
            If fautesur(kk) = "." Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvPoint
            ElseIf fautesur(kk) = "," Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvVirgule
            ElseIf fautesur(kk) = ";" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvPointVirgule
            ElseIf fautesur(kk) = ":" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvDeuxPoints
            ElseIf fautesur(kk) = "?" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvPointInterrogation
            ElseIf fautesur(kk) = "!" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvPointExclamation
            ElseIf fautesur(kk) = "/" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvBarreOblique
            ElseIf fautesur(kk) = "|" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvBarreVerticale
            ElseIf fautesur(kk) = "\" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvBarreObliqueInvers�e
            ElseIf fautesur(kk) = """" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvGuillemet
            ElseIf fautesur(kk) = "'" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvApostrophe
            ElseIf fautesur(kk) = "`" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvAccentGrave
            ElseIf fautesur(kk) = "(" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvParenth�seGauche
            ElseIf fautesur(kk) = ")" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvParenth�seDroite
            ElseIf fautesur(kk) = "[" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvCrochetGauche
            ElseIf fautesur(kk) = "]" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvCrochetDroit
            ElseIf fautesur(kk) = "{" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvAccoladeGauche
            ElseIf fautesur(kk) = "}" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvAccoladeDroite
            ElseIf fautesur(kk) = "-" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvTiret
            ElseIf fautesur(kk) = "_" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvSoulign�
            ElseIf fautesur(kk) = "~" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvTilde
            ElseIf fautesur(kk) = "<" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvInf�rieur
            ElseIf fautesur(kk) = ">" Then
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & vvSup�rieur
            Else
                msgtext0 = msgtext0 & CRLF & nboccur(kk) & msgFautesSur & fautesur(kk) & "�"  'avril 2008 ajout d'un blanc dur
            End If
            nboccur(kk) = 0
            nn = nn + 1
        End If
        fauteprec = fautecourante
    Next kk
Next jj
Print #3, fautecourante
Close #3
If noF1 = 1 Then msgtext0 = ""
If msgtext0 <> "" Then
    msgtext0 = msgtext0 & CRLF & msgPressezF1 & CRLF
    f1msgform = 1
End If

' Le score lui-m�me
If nivo = msgStandard Then pctok(numle�on, numexo + 1) = pctt
If nivo = msgPersonnalis� Then pctok(numle�on + 25, numexo + 1) = pctt
If timevalid = 0 Then
    If pctt >= pct1 Then msgtext0 = msgtext0 + CRLF + msgR�ussi� & CInt(pctt) & msgPourcent & ". "
    If pctt < pct1 Then msgtext0 = msgtext0 + CRLF + msgR�ussi� & CInt(pctt) & msgPourcentSeulement & ". "
End If
If timevalid > 0 Then
    If nivo = msgStandard Then vitok(numle�on, numexo + 1) = CInt(nbmots * 60 / elapsedtot)
    If nivo = msgPersonnalis� Then vitok(numle�on + 25, numexo + 1) = CInt(nbmots * 60 / elapsedtot)
    If pctt >= pct1 And timevalid = 1 Then msgtext0 = msgtext0 + CRLF + msgR�ussi� & CInt(pctt) & msgPourcent & "," + CRLF + msgAvec & nbmots & msgMotsEn & elapsedtot & msgSecondes + CRLF2 + msgExoSuivant
    If pctt < pct1 And timevalid = 1 Then msgtext0 = msgtext0 + CRLF + msgR�ussi� & CInt(pctt) & msgPourcentSeulement & "," + CRLF + msgAvec & nbmots & msgMotsEn & elapsedtot & msgSecondes + CRLF2 + msgExoIdem
    If pctt >= pct1 And timevalid = 2 Then msgtext0 = msgtext0 + CRLF + msgR�ussi� & CInt(pctt) & msgPourcent & "," + CRLF + msgAvec & nbmots & msgCommandesEn & elapsedtot & msgSecondes + CRLF2 + msgExoSuivant
    If pctt < pct1 And timevalid = 2 Then msgtext0 = msgtext0 + CRLF + msgR�ussi� & CInt(pctt) & msgPourcentSeulement & "," + CRLF + msgAvec & nbmots & msgCommandesEn & elapsedtot & msgSecondes + CRLF2 + msgExoIdem
End If

'Enregistrer dans des fichiers texte les fautes et le score
'If numexo = 0 Then letter = "A"
'If numexo = 1 Then letter = "B"
'If numexo = 2 Then letter = "C"
'If numexo = 3 Then letter = "D"
'If numexo = 4 Then letter = "E"
'If numexo = 5 Then letter = "F"
'If numexo = 6 Then letter = "G"
'If numexo = 7 Then letter = "H"
'If numexo = 8 Then letter = "I"
'Open vfile & "\" & numle�on - 2 & letter & ".txt" For Output As #3
'Print #3, msgtext0
'Close #3

' Montrer les fautes et le score
If pctt >= pct1 Then msgtext0 = msgtext0 & pressez_suivant
If pctt < pct1 Then msgtext0 = msgtext0 & pressez_pr�c�dent
SC10:
If forcepause = 1 Then forcepause = 2
pagenum = 0
Msgform.Show 1
If msgf = 2 Then GoTo SC10

' Exercice sur les fautes
If msgf = 3 Then
    Close #1
    exo_courant = "fautes.txt"
    Module_routines.resetmsg
    derligne = 2  'pour pouvoir r�pondre � la derni�re ligne de la le�on
    Le�on2et3.Caption = msgExoFautes
    Le�on2et3.Show 1
End If
f1msgform = 0

' Retour
If msgf >= 1 And pctt >= pct1 Then numexo = numexo + 1
If numexo >= menucount And pctt >= pct1 Then nextle�on = 1
Module_routines.quit_l
End Sub


' ***********************  AFFICHAGE IMMEDIAT du SCORE  ****************************
Public Sub scoreaffich(le�on_courante, alea)
' Pour affichage imm�diat, sauf pour le�ons 6, 7, 13, et dict�es 14, 15, 17
With le�on_courante
    If alea = 0 Then
        On Error Resume Next
        pctt = 100 * (nbcaras / (nbcaras + iwrong))
        If (typele�on <= 3) And timevalid = 0 Then
            scorecourant = CInt(pctt) & " %."
            On Error Resume Next
            .text5.Text = scorecourant
        End If
    End If
End With
End Sub


' ***********************  NEXTSPACE  Cherche espace suivant  ****************************
Public Sub nextspace(le�on_courante)
With le�on_courante
    pp = InStr(1, .text1.Text, " ", 1)
    If pp > 0 Then
        'recherche l'espace suivant ou fin de ligne
        jj = 0: kk = 0
        Do
            jj = jj + 1
            If Mid(.text1.Text, ii + jj, 1) = " " Or ii + jj = ll + 1 Then
                kk = jj
            End If
            If jj >= Len(.text1.Text) Then kk = Len(.text1.Text)
        Loop While kk = 0
        .text1.SelLength = kk - 1
        If sonocara = 1 And erepeat = 0 Then Call Sleep(cadencemot)
    Else
        kk = Len(.text1.Text) + 1  ' Pour traitement par Control+Espace
    End If
End With
End Sub



' **** TEXT3FILL : SONORISE caras non sonoris�s, gr�ce � text3, et COMPTE les MOTS  ******
Public Sub text3fill(le�on_courante)
With le�on_courante

'RESET text3
.text3.Text = ""
.text3.Visible = False
Call Sleep(10)

' Ligne vide ou "� la ligne"
If ii = zz And ii > 0 Then
    .text1.SelLength = 0
    .text3.Text = "�" & vvAlaligne ' (Alt255 devant pour visibilit�)
    .text3.Width = 0.2 * .Width
    .text3.SelStart = 0
    .text3.SelLength = Len(.text3.Text)
    .text3.Visible = True
End If

' Cara ESPACE
On Error Resume Next
If Mid(.text1.Text, ii + 1, 1) = " " Then
    If concatf = 0 Then
        .text3.Text = "�" & vvEspace ' (Alt255 devant pour visibilit�)
        .text3.Width = 0.18 * .Width
        .text3.SelStart = 0
        .text3.SelLength = Len(.text3.Text)
        .text3.Visible = True
    End If
    If concatf = 1 And ii = ll - 1 Then
        .text3.Text = "�" & vvEspace ' (Alt255 devant pour visibilit�)
        .text3.Width = 0.18 * .Width
        .text3.SelStart = 0
        .text3.SelLength = Len(.text3.Text)
        .text3.Visible = True
    End If
    
    'COMPTER ainsi les MOTS
    If .text1.Text <> " " Then ' lignesuivante contenant un simple espace : non compt�e
        nbmots = nbmots + 1
    End If
End If

' Cara POINT
On Error Resume Next
If Mid(.text1.Text, ii + 1, 1) = "." And concatf = 1 Then
    .text3.Text = "�" & vvPoint ' (Alt255 devant pour visibilit�)
    .text3.Width = 0.15 * .Width
    .text3.SelStart = 0
    .text3.SelLength = Len(.text3.Text)
    .text3.Visible = True
End If

' Cara VIRGULE
On Error Resume Next
If Mid(.text1.Text, ii + 1, 1) = "," And concatf = 1 Then
    .text3.Text = "�" & vvVirgule ' (Alt255 devant pour visibilit�)
    .text3.Width = 0.21 * .Width
    .text3.SelStart = 0
    .text3.SelLength = Len(.text3.Text)
    .text3.Visible = True
End If

' Cara "("
On Error Resume Next
If Mid(.text1.Text, ii + 1, 1) = "(" And concatf = 1 Then
    .text3.Text = "�" & vvParenth�seGauche ' (Alt255 devant pour visibilit�)
    .text3.Width = 0.6 * .Width
    .text3.SelStart = 0
    .text3.SelLength = Len(.text3.Text)
    .text3.Visible = True
End If

' Cara ")"
On Error Resume Next
If Mid(.text1.Text, ii + 1, 1) = ")" And concatf = 1 Then
    .text3.Text = "�" & vvParenth�seDroite ' (Alt255 devant pour visibilit�)
    .text3.Width = 0.6 * .Width
    .text3.SelStart = 0
    .text3.SelLength = Len(.text3.Text)
    .text3.Visible = True
End If

' Cara "["
On Error Resume Next
If Mid(.text1.Text, ii + 1, 1) = "[" And concatf = 1 Then
    .text3.Text = "�" & vvCrochetGauche ' (Alt255 devant pour visibilit�)
    .text3.Width = 0.6 * .Width
    .text3.SelStart = 0
    .text3.SelLength = Len(.text3.Text)
    .text3.Visible = True
End If

' Cara "]"
On Error Resume Next
If Mid(.text1.Text, ii + 1, 1) = "]" And concatf = 1 Then
    .text3.Text = "�" & vvCrochetDroit ' (Alt255 devant pour visibilit�)
    .text3.Width = 0.6 * .Width
    .text3.SelStart = 0
    .text3.SelLength = Len(.text3.Text)
    .text3.Visible = True
End If

' Cara "{"
On Error Resume Next
If Mid(.text1.Text, ii + 1, 1) = "{" And concatf = 1 Then
    .text3.Text = "�" & vvAccoladeGauche ' (Alt255 devant pour visibilit�)
    .text3.Width = 0.6 * .Width
    .text3.SelStart = 0
    .text3.SelLength = Len(.text3.Text)
    .text3.Visible = True
End If

' Cara "}"
On Error Resume Next
If Mid(.text1.Text, ii + 1, 1) = "}" And concatf = 1 Then
    .text3.Text = "�" & vvAccoladeDroite ' (Alt255 devant pour visibilit�)
    .text3.Width = 0.6 * .Width
    .text3.SelStart = 0
    .text3.SelLength = Len(.text3.Text)
    .text3.Visible = True
End If

' Cara "GUILLEMET"
On Error Resume Next
If Mid(.text1.Text, ii + 1, 1) = """" And concatf = 1 Then
    .text3.Text = "�" & vvGuillemet ' (Alt255 devant pour visibilit�)
    .text3.Width = 0.3 * .Width
    .text3.SelStart = 0
    .text3.SelLength = Len(.text3.Text)
    .text3.Visible = True
End If

' MAJUSCULES
If sonocara = 1 And indif = 0 Then
    If Asc(Mid(.text1.Text, ii + 1, 1)) >= 65 And Asc(Mid(.text1.Text, ii + 1, 1)) <= 90 Then
        .text3.Text = "�" & vvMajuscule ' (Alt255 devant pour visibilit�)
        .text3.Width = 0.35 * .Width
        .text3.SelStart = 0
        .text3.SelLength = Len(.text3.Text)
        .text3.Visible = True
    End If
End If
End With
End Sub


' **************  Ajuster width, left de la TEXTBOX et size de la POLICE (font) *************
' Recadr� un peu plus � gauche 12/2011, modifi� pour zoom
Public Sub AdjustWidthAndSize(le�on_courante, t2v As Byte)
With le�on_courante
    Select Case Len(currentline)
    Case 1
        .text1.Left = 0.35 * .Width
        If t2v = 1 Then .text2.Left = 0.35 * .Width
        .text1.Width = 0.12 * .Width * zoomvalue
        If t2v = 1 Then .text2.Width = 0.12 * .Width * zoomvalue
        .text1.Font.Size = 1.6 * le�onfontsize * zoomvalue
        If t2v = 1 Then .text2.Font.Size = 1.6 * le�onfontsize * zoomvalue
    Case 2 To 3
        .text1.Left = 0.35 * .Width
        If t2v = 1 Then .text2.Left = 0.35 * .Width
        .text1.Width = 0.21 * .Width * zoomvalue
        If t2v = 1 Then .text2.Width = 0.21 * .Width * zoomvalue
        .text1.Font.Size = 1.5 * le�onfontsize * zoomvalue
        If t2v = 1 Then .text2.Font.Size = 1.5 * le�onfontsize * zoomvalue
    Case 4 To 6
        .text1.Left = 0.2 * .Width
        If t2v = 1 Then .text2.Left = 0.2 * .Width
        .text1.Width = 0.4 * .Width * zoomvalue
        If t2v = 1 Then .text2.Width = 0.4 * .Width * zoomvalue
        .text1.Font.Size = 1.4 * le�onfontsize * zoomvalue
        If t2v = 1 Then .text2.Font.Size = 1.4 * le�onfontsize * zoomvalue
    Case 7 To 11
        .text1.Left = 0.12 * .Width
        If t2v = 1 Then .text2.Left = 0.12 * .Width
        .text1.Width = 0.5 * .Width * zoomvalue
        If t2v = 1 Then .text2.Width = 0.5 * .Width * zoomvalue
        .text1.Font.Size = 1.3 * le�onfontsize * zoomvalue
        If t2v = 1 Then .text2.Font.Size = 1.3 * le�onfontsize * zoomvalue
    Case 12 To 18
        .text1.Left = 0.08 * .Width
        If t2v = 1 Then .text2.Left = 0.08 * .Width
        .text1.Width = 0.72 * .Width * zoomvalue
        If t2v = 1 Then .text2.Width = 0.72 * .Width * zoomvalue
        .text1.Font.Size = 1.2 * le�onfontsize * zoomvalue
        If t2v = 1 Then .text2.Font.Size = 1.2 * le�onfontsize * zoomvalue
    Case Else
        .text1.Left = 0.01 * .Width
        If t2v = 1 Then .text2.Left = 0.01 * .Width
        .text1.Width = 0.85 * .Width * zoomvalue
        If t2v = 1 Then .text2.Width = 0.85 * .Width * zoomvalue
        .text1.Font.Size = le�onfontsize * zoomvalue
        If t2v = 1 Then .text2.Font.Size = le�onfontsize * zoomvalue
    End Select
' Pour la sono, il faut placer le label1 au-dessus de text1
    .label1.Left = .text1.Left
' Il faut placer le text3 � droite du label1
    .text3.Left = .text1.Left + .label1.Width
End With
End Sub


' *************************  DETECTER la FIN de la PHRASE  *****************************
Public Sub DetectPhraseEnd()
        If pasdepoint = 1 Then
            iistop = Len(currentline)
            Exit Sub
        End If
        'D�tecter ".. ou "..." de pr�f�rence � "."
        iistop1 = InStr(iistop + 1, currentline, ".")
        If iistop1 = iistop + 1 Then iistop = iistop1
        iistop1 = InStr(iistop + 1, currentline, ".")
        If iistop1 = iistop + 1 Then iistop = iistop1
        
        'Annuler ce "." s'il est suivi d'un chiffre juste apr�s, exemple "3.12"
        iistop0 = InStr(iistop + 1, currentline, "0")
        If iistop0 = iistop + 1 Then iistop = InStr(iistop + 1, currentline, ".")
        iistop0 = InStr(iistop + 1, currentline, "1")
        If iistop0 = iistop + 1 Then iistop = InStr(iistop + 1, currentline, ".")
        iistop0 = InStr(iistop + 1, currentline, "2")
        If iistop0 = iistop + 1 Then iistop = InStr(iistop + 1, currentline, ".")
        iistop0 = InStr(iistop + 1, currentline, "3")
        If iistop0 = iistop + 1 Then iistop = InStr(iistop + 1, currentline, ".")
        iistop0 = InStr(iistop + 1, currentline, "4")
        If iistop0 = iistop + 1 Then iistop = InStr(iistop + 1, currentline, ".")
        iistop0 = InStr(iistop + 1, currentline, "5")
        If iistop0 = iistop + 1 Then iistop = InStr(iistop + 1, currentline, ".")
        iistop0 = InStr(iistop + 1, currentline, "6")
        If iistop0 = iistop + 1 Then iistop = InStr(iistop + 1, currentline, ".")
        iistop0 = InStr(iistop + 1, currentline, "7")
        If iistop0 = iistop + 1 Then iistop = InStr(iistop + 1, currentline, ".")
        iistop0 = InStr(iistop + 1, currentline, "8")
        If iistop0 = iistop + 1 Then iistop = InStr(iistop + 1, currentline, ".")
        iistop0 = InStr(iistop + 1, currentline, "9")
        If iistop0 = iistop + 1 Then iistop = InStr(iistop + 1, currentline, ".")
            
        'D�tecter le premier de "." ou "!" ou "?" (suite)
        If (iistop2 <= iistop) And (iistop2 <> 0) Then iistop = iistop2
        If (iistop3 <= iistop) And (iistop3 <> 0) Then iistop = iistop3
        If iistop <= iistart Then iistop = iistart + Len(currentline)
        
        'Inclure l'�ventuel espace qui suit la fin de la partie suivante
        iistop9 = InStr(iistop + 1, currentline, " ")
        If iistop9 = iistop + 1 Then iistop = iistop9
End Sub


' ********************************* PASDETAB *******************************************
' ***  La touche Tab peut �tre ainsi moins g�nante dans certaines le�ons
Public Sub pasdetab()
keyforce = 9
keyinhibit = 1
t2inhibit = 1
f2link = 0 'cas du Tab envoy� apr�s un appel F2, septembre 2007
'SendKeys "{ESC}"
'Sendkeys est remplac� par des actions keybd_event pour Windows Vista juin 2007
keybd_event VK_ESCAPE, 0, 0, 0
keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
echapbis = 0
End Sub


' ****************** RESULTS : FICHIER de R�SULTATS ****************************************
Public Sub results()
' Ouverture
If nivo = msgStandard Then Open vfile & "\R�sultat-Standard.doc" For Output As #1
If nivo = msgPersonnalis� Then Open vfile & "\R�sultat-Personnalis�.doc" For Output As #1

' Pour �viter les close#1 lors du load implicite des menus qui appellent OpenAndSuffix
consult = 1

' Remplissage du fichier de r�sultats
Print #1, msgR�sultats & nom & "."
If nivo = msgStandard Then Print #1, msgNiveauStandard
If nivo = msgPersonnalis� Then Print #1, msgNiveauPersonnalis�
    If InStr(10, Menu_principal.list1.List(0), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(0)
    If InStr(10, Menu_principal.list1.List(1), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(1)
    If InStr(10, Menu_principal.list1.List(2), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(2)
    If InStr(10, Menu_principal.list1.List(3), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(3)
    For kk = 0 To Menu_le�on1.list1.ListCount - 1
        If InStr(10, Menu_le�on1.list1.List(kk), "%", 1) > 0 Then Print #1, "   1" & Menu_le�on1.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(4), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(4)
    For kk = 0 To Menu_le�on2.list1.ListCount - 1
        If InStr(10, Menu_le�on2.list1.List(kk), "%", 1) > 0 Then Print #1, "   2" & Menu_le�on2.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(5), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(5)
    For kk = 0 To Menu_le�on3.list1.ListCount - 1
        If InStr(10, Menu_le�on3.list1.List(kk), "%", 1) > 0 Then Print #1, "   3" & Menu_le�on3.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(6), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(6)
    For kk = 0 To Menu_le�on4.list1.ListCount - 1
        If InStr(10, Menu_le�on4.list1.List(kk), "%", 1) > 0 Then Print #1, "   4" & Menu_le�on4.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(7), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(7)
    For kk = 0 To Menu_le�on5.list1.ListCount - 1
        If InStr(10, Menu_le�on5.list1.List(kk), "%", 1) > 0 Then Print #1, "   5" & Menu_le�on5.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(8), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(8)
    For kk = 0 To Menu_le�on6.list1.ListCount - 1
        If InStr(10, Menu_le�on6.list1.List(kk), "%", 1) > 0 Then Print #1, "   6" & Menu_le�on6.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(9), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(9)
    For kk = 0 To Menu_le�on7.list1.ListCount - 1
        If InStr(10, Menu_le�on7.list1.List(kk), "%", 1) > 0 Then Print #1, "   7" & Menu_le�on7.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(10), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(10)
    For kk = 0 To Menu_le�on8.list1.ListCount - 1
        If InStr(10, Menu_le�on8.list1.List(kk), "%", 1) > 0 Then Print #1, "   8" & Menu_le�on8.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(11), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(11)
    For kk = 0 To Menu_le�on9.list1.ListCount - 1
        If InStr(10, Menu_le�on9.list1.List(kk), "%", 1) > 0 Then Print #1, "   9" & Menu_le�on9.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(12), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(12)
    For kk = 0 To Menu_le�on10.list1.ListCount - 1
        If InStr(10, Menu_le�on10.list1.List(kk), "%", 1) > 0 Then Print #1, "  10" & Menu_le�on10.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(13), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(13)
    For kk = 0 To Menu_le�on11.list1.ListCount - 1
        If InStr(10, Menu_le�on11.list1.List(kk), "%", 1) > 0 Then Print #1, "  11" & Menu_le�on11.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(14), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(14)
    For kk = 0 To Menu_le�on12.list1.ListCount - 1
        If InStr(10, Menu_le�on12.list1.List(kk), "%", 1) > 0 Then Print #1, "  12" & Menu_le�on12.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(15), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(15)
    For kk = 0 To Menu_le�on13.list1.ListCount - 1
        If InStr(10, Menu_le�on13.list1.List(kk), "%", 1) > 0 Then Print #1, "  13" & Menu_le�on13.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(16), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(16)
    For kk = 0 To Menu_le�on14.list1.ListCount - 1
        If InStr(10, Menu_le�on14.list1.List(kk), "%", 1) > 0 Then Print #1, "  14" & Menu_le�on14.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(17), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(17)
    For kk = 0 To Menu_le�on15.list1.ListCount - 1
        If InStr(10, Menu_le�on15.list1.List(kk), "%", 1) > 0 Then Print #1, "  15" & Menu_le�on15.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(18), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(18)
    For kk = 0 To Menu_le�on16.list1.ListCount - 1
        If InStr(10, Menu_le�on16.list1.List(kk), "%", 1) > 0 Then Print #1, "  16" & Menu_le�on16.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(19), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(19)
    For kk = 0 To Menu_le�on17.list1.ListCount - 1
        If InStr(10, Menu_le�on17.list1.List(kk), "%", 1) > 0 Then Print #1, "  17" & Menu_le�on17.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(20), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(20)
    For kk = 0 To Menu_le�on18.list1.ListCount - 1
        If InStr(10, Menu_le�on18.list1.List(kk), "%", 1) > 0 Then Print #1, "  18" & Menu_le�on18.list1.List(kk)
    Next kk
    If InStr(10, Menu_principal.list1.List(21), "%", 1) > 0 Then Print #1, Menu_principal.list1.List(21)
    For kk = 0 To Menu_le�on19.list1.ListCount - 1
        If InStr(10, Menu_le�on19.list1.List(kk), "%", 1) > 0 Then Print #1, "  19" & Menu_le�on19.list1.List(kk)
    Next kk
    Close #1
    
' Restaurer les possibilit�s de close#x avec OpenAndSuffix
consult = 0
Unload Menu_le�on1
Unload Menu_le�on2
Unload Menu_le�on3
Unload Menu_le�on4
Unload Menu_le�on5
Unload Menu_le�on6
Unload Menu_le�on7
Unload Menu_le�on8
Unload Menu_le�on9
Unload Menu_le�on10
Unload Menu_le�on11
Unload Menu_le�on12
Unload Menu_le�on13
Unload Menu_le�on14
Unload Menu_le�on15
Unload Menu_le�on16
Unload Menu_le�on17
Unload Menu_le�on18
Unload Menu_le�on19
End Sub


' *****************  REGISTER : ENREGISTRER les R�SULTATS  *********************************
' modifi� 12/2011 pour zoom et couleurs
Public Sub register()
If vfile = "" Then Exit Sub
' Taux de r�ussite (moyenne, exo par exo, numle�on en derni�re col)
Open vfile & "\pctok.txt" For Output As #4
For jj = 0 To 24
    Write #4, pctok(jj, 0), pctok(jj, 1), pctok(jj, 2), pctok(jj, 3), pctok(jj, 4), pctok(jj, 5), pctok(jj, 6), pctok(jj, 7), pctok(jj, 8), jj - 2
Next jj
For jj = 25 To 49
    Write #4, pctok(jj, 0), pctok(jj, 1), pctok(jj, 2), pctok(jj, 3), pctok(jj, 4), pctok(jj, 5), pctok(jj, 6), pctok(jj, 7), pctok(jj, 8), jj - 27
Next jj
Close #4

' Vitesse (moyenne, exo par exo, nume�on en derni�re col)
Open vfile & "\vitok.txt" For Output As #9
For jj = 0 To 24
    Write #9, vitok(jj, 0), vitok(jj, 1), vitok(jj, 2), vitok(jj, 3), vitok(jj, 4), vitok(jj, 5), vitok(jj, 6), vitok(jj, 7), vitok(jj, 8), jj - 2
Next jj
For jj = 25 To 49
    Write #9, vitok(jj, 0), vitok(jj, 1), vitok(jj, 2), vitok(jj, 3), vitok(jj, 4), vitok(jj, 5), vitok(jj, 6), vitok(jj, 7), vitok(jj, 8), jj - 27
Next jj
Close #9

' Fichier INI (conserve l'�tat d'avancement de l'utilisateur)
Open vfile & "\" & nom & ".ini" For Output As #16
Write #16, numle�on, numexo, nivo, debexplilevel, biplevel, debgenlevel, zoomlevel, colorslevel
Close #16
End Sub


' **************  OPENANDCOUNT lines and max line length of FIC  *****************************
' Utilis� seulement dans les dict�es, sinon on pr�f�re OpenAndSuffix
' qui ajoute Alt255 pour meilleure sonorisation par Jaws du dernier caract�re de la ligne envoy�e
Public Sub OpenAndCount(fictxt)
Close #1
Open vpath & "Le�ons\" & nivoRep & "\" & fictxt For Input As #1
jj = 0
mm = 0
Do While Not EOF(1)
    Line Input #1, currentline
    If Len(currentline) > mm Then mm = Len(currentline)
    jj = jj + 1
Loop
nblines = jj
If nblines = 1 Then derligne = 1
Close #1
Open vpath & "Le�ons\" & nivoRep & "\" & fictxt For Input As #1
End Sub


' ***********  OPENANDSUFFIX with Alt255 at end of line, and COUNT  **************************
Public Sub OpenAndSuffix(fictxt As String, suffix As Byte)
' Suffix=1 pour ouverture et ajout du suffixe en bout de ligne
' Suffix=0 pour restauration/suppression

'Test
Close #1
If fictxt = "fautes.txt" Then
    If Dir(vfile & "\fautes.txt") = "" Then Exit Sub
End If
If fictxt <> "fautes.txt" Then
    If Dir(vpath & "Le�ons\" & nivoRep & "\" & fictxt) = "" Then Exit Sub
End If

' Compter d'abord le nb de lignes, si ouverture et mise en place du suffixe
If suffix = 1 Then
    If fictxt = "fautes.txt" Then
        Open vfile & "\fautes.txt" For Input As #1
    End If
    If fictxt <> "fautes.txt" Then
        Open vpath & "Le�ons\" & nivoRep & "\" & fictxt For Input As #1
    End If
    jj = 0
    mm = 0
    Do While Not EOF(1)
        Line Input #1, currentline
        If Len(currentline) > mm Then mm = Len(currentline)
        jj = jj + 1
    Loop
    nblines = jj
    If nblines = 1 Then derligne = 1
    Close #1
End If
                
' M�nage pour pr�parer ajout/suppression du suffixe
On Error Resume Next
Kill vpath & "Le�ons\" & nivoRep & "\tmp_tmp.txt"

' Ouvrir la le�on d'input et le tmp_tmp d'output
Open vpath & "Le�ons\" & nivoRep & "\tmp_tmp.txt" For Output As #8
If fictxt = "fautes.txt" Then
    Open vfile & "\fautes.txt" For Input As #7
End If
If fictxt <> "fautes.txt" Then
    Open vpath & "Le�ons\" & nivoRep & "\" & fictxt For Input As #7
End If
Do While Not EOF(7)
    Line Input #7, currentline
 
    ' Suffix=1, Modifier en AJOUTANT le suffixe blanc dur Alt255 s'il n'existe pas d�j�
    If suffix = 1 Then
        If Right(currentline, 1) <> "�" Then currentline = currentline & "�"
    End If
    
    ' Suffix=0, Modifier en SUPPRIMANT le suffixe blanc dur Alt255 s'il n'existe pas d�j�
    If suffix = 0 Then
        If Right(currentline, 1) = "�" Then currentline = Left(currentline, Len(currentline) - 1)
    End If
    
    ' Ecrire dans le fichier tmp_tmp
    Print #8, currentline
Loop
Close #7
Close #8

' Valider le texte de le�on suffix�e (pas 0 octets)
If Dir(vpath & "Le�ons\" & nivoRep & "\tmp_tmp.txt") = "" Then
ElseIf FileLen(vpath & "Le�ons\" & nivoRep & "\tmp_tmp.txt") < 4 Then
Else
    If fictxt = "fautes.txt" Then
        FileCopy vpath & "Le�ons\" & nivoRep & "\tmp_tmp.txt", vfile & "\fautes.txt"
    End If
    If fictxt <> "fautes.txt" Then
        FileCopy vpath & "Le�ons\" & nivoRep & "\tmp_tmp.txt", vpath & "Le�ons\" & nivoRep & "\" & fictxt
    End If
End If

' Ouvrir enfin la le�on suffix�e, si ouverture et mise en place du suffixe
Close #1
If suffix = 1 Then
    If fictxt = "fautes.txt" Then
        Open vfile & "\fautes.txt" For Input As #1
    End If
    If fictxt <> "fautes.txt" Then
        Open vpath & "Le�ons\" & nivoRep & "\" & fictxt For Input As #1
    End If
End If
' M�nage
On Error Resume Next
Kill vpath & "Le�ons\" & nivoRep & "\tmp_tmp.txt"
End Sub


' **************  PlaceInMsgAide  place file lines in msgtext0  ******************************
' Met successivement toutes les lignes d'un fichier dans une variable puis "msgbox" de cette variable
Public Sub placeinmsgaide(fictxt)
If Dir(vpath & fictxt) = "" Then
    MsgBox msgNofic, 0, ""
    Exit Sub
End If
Close #1
Open vpath & fictxt For Input As #1
msgtext0 = ""
Do While Not EOF(1)
    Line Input #1, currentline
    msgtext0 = msgtext0 & currentline & CRLF
Loop
Close #1
' Attention, texte de msgbox limit� � 1024 caract�res
MsgBox msgtext0, 0, ""
End Sub


' **************  PLACEINMSGTEXT0  place file lines in msgtext0  ******************************
' Met successivement toutes les lignes d'un fichier dans une variable (msgbox hors routine)
Public Sub placeinmsgtext0(fictxt)
Close #1
Open vpath & "Le�ons\" & nivoRep & "\" & fictxt For Input As #1
msgtext0 = ""
Do While Not EOF(1)
    Line Input #1, currentline
    msgtext0 = msgtext0 & currentline & CRLF
Loop
msgtext0 = msgtext0 + pressez_entr�e
Close #1
End Sub


' ******  MENU_REFRESH : R�AFFICHER un MENU, r�sultats utilisateur en bout de ligne  *********
Public Sub menu_refresh(ficmnu, mnu)
If nivo = msgStandard Then kk = 0
If nivo = msgPersonnalis� Then kk = 25

' Rep�rer la longueur mm de la plus grande ligne et le nombre nbexo de lignes(d'exercices)
Open vpath & ficmnu For Input As #2
jj = 0
mm = 0
Do While Not EOF(2)
    Line Input #2, currentmenuline
    If Len(currentmenuline) > mm Then mm = Len(currentmenuline)
    jj = jj + 1
Loop
nbexo = jj
Close #2

' Inclure les r�sultats
Open vpath & ficmnu For Input As #2
incomplet = 0
jj = 0
Do While Not EOF(2)
    Line Input #2, currentmenuline
    If pctok(numle�on + kk, jj + 1) = 0 Then
        mnu.list1.List(jj) = currentmenuline
        incomplet = 1
    Else
        'nb nn d'espaces d'alignement � la fin de la ligne menu, avant le r�sultat pctok
        nn = mm - Len(currentmenuline) + 1
        pp = 0: tempo = " "
        Do While pp < nn
            tempo = tempo & " "
            pp = pp + 1
        Loop
        If Not pctok(numle�on + kk, jj + 1) = 100 Then tempo = tempo & " "
        If vitok(numle�on + kk, jj + 1) = 0 Then mnu.list1.List(jj) = currentmenuline & tempo & pctok(numle�on + kk, jj + 1) & "%"
        If Not vitok(numle�on + kk, jj + 1) = 0 Then mnu.list1.List(jj) = currentmenuline & tempo & pctok(numle�on + kk, jj + 1) & "% " & vitok(numle�on + kk, jj + 1) & msgMotsMinute
    End If
    jj = jj + 1
Loop
Close #2

' Cr�er la moyenne pctok (et la vitesse moyenne vitok) pour une le�on compl�t�e
If incomplet = 0 Then
    pp = 0
    For jj = 0 To nbexo - 1
        pp = pp + pctok(numle�on + kk, jj + 1)
    Next jj
    pctok(numle�on + kk, 0) = pp / nbexo
    pp = 0
    For jj = 0 To nbexo - 1
        pp = pp + vitok(numle�on + kk, jj + 1)
    Next jj
    vitok(numle�on + kk, 0) = pp / nbexo
End If

' Positionner la surbrillance
If numexo < mnu.list1.ListCount Then
    mnu.list1.Selected(numexo) = True
Else
    mnu.list1.Selected(0) = True
End If
End Sub


' *********************  MENU_REPEAT R�p�te la ligne menu en cours  ***********************
Public Sub menu_repeat()
keyinhibit = 0
numindex = menu_courant.list1.ListIndex
Unload menu_courant
menu_courant.list1.ListIndex = numindex
menu_courant.Show 1
End Sub


' **********************  PASS Passer � la le�on suivante  ********************************
Public Sub pass()

    ' Pour le cas o� on serait pass� auparavant par l'option C Consulter les r�sultats (module results)
    If numle�on = 3 Then Set menu_courant = Menu_le�on1
    If numle�on = 4 Then Set menu_courant = Menu_le�on2
    If numle�on = 5 Then Set menu_courant = Menu_le�on3
    If numle�on = 6 Then Set menu_courant = Menu_le�on4
    If numle�on = 7 Then Set menu_courant = Menu_le�on5
    If numle�on = 8 Then Set menu_courant = Menu_le�on6
    If numle�on = 9 Then Set menu_courant = Menu_le�on7
    If numle�on = 10 Then Set menu_courant = Menu_le�on8
    If numle�on = 11 Then Set menu_courant = Menu_le�on9
    If numle�on = 12 Then Set menu_courant = Menu_le�on10
    If numle�on = 13 Then Set menu_courant = Menu_le�on11
    If numle�on = 14 Then Set menu_courant = Menu_le�on12
    If numle�on = 15 Then Set menu_courant = Menu_le�on13
    If numle�on = 16 Then Set menu_courant = Menu_le�on14
    If numle�on = 17 Then Set menu_courant = Menu_le�on15
    If numle�on = 18 Then Set menu_courant = Menu_le�on16
    If numle�on = 19 Then Set menu_courant = Menu_le�on17
    If numle�on = 20 Then Set menu_courant = Menu_le�on18
    If numle�on = 21 Then Set menu_courant = Menu_le�on19
    
    If numle�on = 3 Then Set menu_suivant = Menu_le�on2
    If numle�on = 4 Then Set menu_suivant = Menu_le�on3
    If numle�on = 5 Then Set menu_suivant = Menu_le�on4
    If numle�on = 6 Then Set menu_suivant = Menu_le�on5
    If numle�on = 7 Then Set menu_suivant = Menu_le�on6
    If numle�on = 8 Then Set menu_suivant = Menu_le�on7
    If numle�on = 9 Then Set menu_suivant = Menu_le�on8
    If numle�on = 10 Then Set menu_suivant = Menu_le�on9
    If numle�on = 11 Then Set menu_suivant = Menu_le�on10
    If numle�on = 12 Then Set menu_suivant = Menu_le�on11
    If numle�on = 13 Then Set menu_suivant = Menu_le�on12
    If numle�on = 14 Then Set menu_suivant = Menu_le�on13
    If numle�on = 15 Then Set menu_suivant = Menu_le�on14
    If numle�on = 16 Then Set menu_suivant = Menu_le�on15
    If numle�on = 17 Then Set menu_suivant = Menu_le�on16
    If numle�on = 18 Then Set menu_suivant = Menu_le�on17
    If numle�on = 19 Then Set menu_suivant = Menu_le�on18
    If numle�on = 20 Then Set menu_suivant = Menu_le�on19
    If numle�on = 21 Then Set menu_suivant = Menu_le�on19
    
' Ne PAS PASSER au MENU de la LE�ON SUIVANTE
If nextle�on = 0 Then
    
    ' Chargement
    Unload menu_courant
    menu_courant.Show 1

' PASSER au MENU de la LE�ON SUIVANTE
Else
    ' Pour refresh de la moyenne de tous les pctok de la le�on
    Unload menu_courant
    menu_courant.Show  ' Pas de stop sur le show !
    Unload menu_courant
    ' Exo manquant ou Message donnant la moyenne
    If nivo = msgStandard Then kk = 0
    If nivo = msgPersonnalis� Then kk = 25
    If pctok(numle�on + kk, 0) = 0 Then
        Unload menu_courant
        menu_courant.Show 1
    Else
PA10:
        pagenum = 0
        msgtext0 = CRLF + msgLaLe�on + Str(numle�on - 2) + msgEstTermin�e + CRLF + msgSes + Str(menucount) + msgR�ussite + Str(pctok(numle�on + kk, 0)) + msgPourcent + "." + pressez_Le�onSuivante
        Msgform.Show 1
        If msgf = 2 Then GoTo PA10
        If msgf = 1 Then
            numexo = 0
            numle�on = numle�on + 1
            Unload menu_courant
            Unload Menu_principal
            Menu_principal.Show 1
        End If
        If msgf = 0 Then
            numexo = numexo - 1
            Unload menu_courant
            menu_courant.Show 1
        End If
    End If
    nextle�on = 0
End If
End Sub


' *************************  EPELLATION  ***************************************************
Public Sub epellation(le�on_courante)
With le�on_courante
    qq = Len(.text2.Text)
    .text1.SelStart = qq
    .text1.SelLength = 0
    Call Sleep(260)  ' N�cessaire en Win98, bizarre et critique
    
    'recherche l'espace suivant ou fin de ligne, utilise qq car ii faux quand on a d�pass� les iwrongbismax fautes
    pp = InStr(1, .text1.Text, " ", 1)
    If pp > 0 Then
        jj = 0: kk = 0
        Do
            jj = jj + 1
            If Mid(.text1.Text, qq + jj, 1) = " " Or qq + jj = ll + 1 Then
                kk = jj
            End If
            If jj >= Len(.text1.Text) Then kk = Len(.text1.Text)
        Loop While kk = 0
        .text1.SelLength = kk - 1
        If sonocara = 1 And erepeat = 0 Then Call Sleep(cadencemot)
    Else
        kk = Len(.text1.Text) + 1  ' Pour traitement par Control+Espace
    End If
    
    ' �pellation (limit�e � 16 caras)
    For jj = 0 To kk - 2
        If jj < 16 Then
            .text1.SelLength = 1
            Call Sleep(260)
            .text1.SelStart = .text1.SelStart + 1
        End If
    Next jj
    .text1.SelStart = qq
    .text1.SelLength = 0
    Call Sleep(260)  ' N�cessaire en Win98
End With
End Sub


' *************************  NIVEAUX  *****************************************************
' modifi� 12/2011 pour zoom et couleurs
Public Sub niveaux()
If nivo = msgStandard Then
    menu_courant.Standard.Checked = True
End If
If nivo = msgPersonnalis� Then
    menu_courant.Personnalis�.Checked = True
End If
If debexplilevel = msgNormal Then
    menu_courant.DebExpliNormal.Checked = True
    debexplivalue = ""
End If
If debexplilevel = msgRapide Then
    menu_courant.DebExpliRapide.Checked = True
    debexplivalue = "   "
End If
If biplevel = msgClassique Then
    menu_courant.BipClassique.Checked = True
End If
If biplevel = msgDiff�rent Then
    menu_courant.BipDiff�rent.Checked = True
End If
If debgenlevel = msgLent Then
    menu_courant.DebGenLent.Checked = True
    debgenvalue = " "
End If
If debgenlevel = msgMoyen Then
    menu_courant.DebGenMoyen.Checked = True
    debgenvalue = "  "
End If
If debgenlevel = msgVite Then
    menu_courant.DebGenVite.Checked = True
    debgenvalue = "   "
End If
If zoomlevel = msgNoZoom Then
     menu_courant.NoZoom.Checked = True
End If
If zoomlevel = msgWithZoom Then
    menu_courant.WithZoom.Checked = True
End If
If colorslevel = msgBasicColors Then
    menu_courant.BasicColors.Checked = True
End If
If colorslevel = msgOtherColors Then
    menu_courant.OtherColors.Checked = True
End If
End Sub


' *******************  BIP  **************************************************************
Public Sub bip(le�on) 'modifi� avril 2008
On Error Resume Next
le�on.Picture1.Visible = True
'If Dir(vpath & "sonbeep.exe") <> "" And biplevel = msgClassique Then
    'Shell vpath & "sonbeep.exe", 0
If Dir(vpath & "pop.wav") <> "" And biplevel = msgClassique Then
    sndPlaySound vpath & "pop.wav", SND_NOWAIT
'ElseIf Dir(vpath & "sonbip.exe") <> "" And biplevel = msgDiff�rent Then
    'Shell vpath & "sonbip.exe", 0
ElseIf biplevel = msgDiff�rent Then
    Emettre (660)
Else
    Beep
End If
End Sub


' ***********************  DEBIT d'explications NORMAL  *************************************
' Le script Jaws est bas� sur les titres des fen�tres mont�es � l'utilisateur,
' lesdits titres incluent des espaces en-t�te
Public Sub DebExpliNormal()
debexplilevel = msgNormal
debexplivalue = ""
numindex = menu_courant.list1.ListIndex
Unload menu_courant
MsgBox msgExpli & msgNormal & ". ", 0, "" & debexplilevel & "."
menu_courant.list1.ListIndex = numindex
menu_courant.Show 1
End Sub


' ******  DEBIT d'explications RAPIDE  par changement de la variable OffsetExpli dans le script Jaws *******
' Le script Jaws est bas� sur les titres des fen�tres mont�es � l'utilisateur,
' lesdits titres incluent des espaces en-t�te
Public Sub DebExpliRapide()
debexplilevel = msgRapide
debexplivalue = "   "
numindex = menu_courant.list1.ListIndex
Unload menu_courant
MsgBox msgExpli & msgRapide & ". ", 0, "   " & debexplilevel & "."
menu_courant.list1.ListIndex = numindex
menu_courant.Show 1
End Sub


' ******  DEBIT g�n�ral LENT par changement de la variable OffsetGen dans le script Jaws  *******
' Le script Jaws est bas� sur les titres des fen�tres mont�es � l'utilisateur,
' lesdits titres incluent des espaces apr�s le premier mot
Public Sub DebGenLent()
debgenlevel = msgLent
debgenvalue = " "
numindex = menu_courant.list1.ListIndex
Unload menu_courant
MsgBox msgD�bit & msgLent & ".", 0, msgLent & "."
menu_courant.list1.ListIndex = numindex
menu_courant.Show 1
End Sub


' ******  DEBIT g�n�ral PLUS VITE par changement de la variable Offset dans le scipt  *******
' Le script Jaws est bas� sur les titres des fen�tres mont�es � l'utilisateur,
' lesdits titres incluent des espaces apr�s le premier mot
Public Sub DebGenMoyen()
debgenlevel = msgMoyen
debgenvalue = "  "
numindex = menu_courant.list1.ListIndex
Unload menu_courant
MsgBox msgD�bit & msgMoyen & ".", 0, msgMoyen & "."
menu_courant.list1.ListIndex = numindex
menu_courant.Show 1
End Sub


' ******  DEBIT g�n�ral PLUS VITE par changement de la variable Offset dans le scipt  *******
' Le script Jaws est bas� sur les titres des fen�tres mont�es � l'utilisateur,
' lesdits titres incluent des espaces apr�s le premier mot
Public Sub DebGenVite()
debgenlevel = msgVite
debgenvalue = "   "
numindex = menu_courant.list1.ListIndex
Unload menu_courant
MsgBox msgD�bit & msgVite & ".", 0, msgVite & "."
menu_courant.list1.ListIndex = numindex
menu_courant.Show 1
End Sub


' ***********************  BIPClassique  **************************************************
Public Sub BipClassique()  'modifi� avril 2008
biplevel = msgClassique
numindex = menu_courant.list1.ListIndex
If Dir(vpath & "pop.wav") <> "" Then
    sndPlaySound vpath & "pop.wav", SND_NOWAIT
Else
    Beep
End If
Unload menu_courant
MsgBox msgBipsAre & msgClassique & ". ", 0, debexplivalue & biplevel & "."
menu_courant.list1.ListIndex = numindex
menu_courant.Show 1
End Sub


' **************************  BIPDiff�rent  **************************************************
Public Sub BipDiff�rent()  'modifi� avril 2008
biplevel = msgDiff�rent
numindex = menu_courant.list1.ListIndex
Emettre (660)
Unload menu_courant
MsgBox msgBipsAre & msgDiff�rent & ". ", 0, debexplivalue & biplevel & "."
menu_courant.list1.ListIndex = numindex
menu_courant.Show 1
End Sub


' **************************  sonbip2tons  **************************************************
Public Sub sonbip2tons()  'cr�� en avril 2008
If Dir(vpath & "PianoUp4-F-A.wav") <> "" And biplevel = msgClassique Then
    sndPlaySound vpath & "PianoUp4-F-A.wav", SND_NOWAIT
Else
    txtTps = 0.25
    Emettre (350)
    Emettre (440)
    txtTps = 0.05
End If
End Sub


' **************************  BasicColors  **************************************************
' cr�� 12/2011
Public Sub BasicColors()
colorslevel = msgBasicColors
numindex = menu_courant.list1.ListIndex
Unload menu_courant
MsgBox "L'affichage sera en " & msgBasicColors & ".", 0, msgBasicColors & "."
menu_courant.list1.ListIndex = numindex
menu_courant.Show 1
End Sub

' **************************  OtherColors  **************************************************
' cr�� 12/2011
Public Sub OtherColors()
colorslevel = msgOtherColors
numindex = menu_courant.list1.ListIndex
Unload menu_courant
MsgBox "L'affichage sera en " & msgOtherColors & ".", 0, msgOtherColors & "."
menu_courant.list1.ListIndex = numindex
menu_courant.Show 1
End Sub

' **************************  NoZoom  **************************************************
' cr�� 12/2011
Public Sub NoZoom()
zoomlevel = msgNoZoom
numindex = menu_courant.list1.ListIndex
zoomfactor = 1
Unload menu_courant
MsgBox "L'affichage sera " & msgNoZoom & ".", 0, msgNoZoom & "."
menu_courant.list1.ListIndex = numindex
menu_courant.Show 1
End Sub

' **************************  WithZoom  **************************************************
' cr�� 12/2011
Public Sub WithZoom()
zoomlevel = msgWithZoom
numindex = menu_courant.list1.ListIndex
zoomfactor = zoomvalue
Unload menu_courant
MsgBox "L'affichage sera " & msgWithZoom & ".", 0, msgWithZoom & "."
menu_courant.list1.ListIndex = numindex
menu_courant.Show 1
End Sub


' ****************  CLEAN temporary files  **************************************************
Public Sub clean()
If Not Dir(vpath & "menu_courant.txt") = "" Then
    On Error Resume Next
    Kill vpath & "menu_courant.txt"
End If
End Sub


' ****************  RESETMSG  ***********************************************************
Public Sub resetmsg()
inexo = 1
For jj = 0 To 149
    msgtext1(jj) = ""
    msgtext2(jj) = ""
Next jj
End Sub


' **************** CANCELWIN *********************************************************
' Annule l'effet de la touche Windows par l'appel bref de Control+Alt+Suppr fen�tre prioritaire Windows de gestion des t�ches
' suivi d'un d�lai calibr� par l'exp�rience, et d'un Echap qui d�truit ladite fen�tre
Public Sub cancelwin(nobip, bipobject, menucase)
If menucase = 0 Then msgpb.Show 1 ' indispensable
If menucase = 1 Then msgpbmenu.Show 1 ' indispensable
'If nobip = 0 Then keyinhibit = 1   ' pour ne pas bipper sur Control+Alt+Suppr, inutile juin 2007
If nobip = 1 Then keyinhibit = 4   ' pour ne pas bipper sur Control+Alt+Suppr
'SendKeys "^%{DEL}", True   ' combinaison Control+Alt+Suppr
'Sendkeys est remplac� par des actions keybd_event pour Windows Vista juin 2007
keybd_event VK_CONTROL, 0, 0, 0
keybd_event VK_MENU, 0, 0, 0
keybd_event VK_DELETE, 0, 0, 0
keybd_event VK_DELETE, 0, KEYEVENTF_KEYUP, 0
keybd_event VK_MENU, 0, KEYEVENTF_KEYUP, 0
keybd_event VK_CONTROL, 0, KEYEVENTF_KEYUP, 0
'Call Sleep(400)  ' vraiment indispensable avril 2004, supprim� en juin 2007 ?
echapbis = -1
'SendKeys "{ESC}", True
'Sendkeys est remplac� par des actions keybd_event pour Windows Vista juin 2007
keybd_event VK_ESCAPE, 0, 0, 0
keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
echapbis = 0
If menucase = 0 Then msgpb.Show 1 ' indispensable
If menucase = 1 Then msgpbmenu.Show 1 ' indispensable
If nobip = 0 Then Module_routines.bip bipobject
keyinhibit = 0 ' septembre 2007
End Sub


' ********************* ITEM 1 du MENU PRINCIPAL : PRESENTATION *****************************
' Tous les messages de cette introduction sont des variables pg-- regroup�es dans le global.bas
Public Sub presentation()
Unload Menu_principal
numle�on = 0
L10:
pagenum = 1
msgtext0 = CRLF + pgia1 + pressez
Msgform.Show 1
If msgf = 33 Then Beep
If msgf = 2 Or msgf = 33 Then GoTo L10
If msgf = 1 Or msgf = 34 Then
L11:
    pagenum = 2
    msgtext0 = CRLF + pgia2 + pressez
    Msgform.Show 1
    If msgf = 33 Then GoTo L10
    If msgf = 2 Then GoTo L11
    If msgf = 1 Or msgf = 34 Then
L12:
        pagenum = 3
        msgtext0 = CRLF + pgia3 + pressez
        Msgform.Show 1
        If msgf = 33 Then GoTo L11
        If msgf = 2 Then GoTo L12
        If msgf = 1 Or msgf = 34 Then
L13:
            pagenum = 4: pagemax = 1
            msgtext0 = CRLF + pgia4 + pressez
            If nivo = msgStandard Then pctok(0, 0) = 100
            If nivo = msgPersonnalis� Then pctok(25, 0) = 100
            Msgform.Show 1
            If msgf = 33 Then GoTo L12
            If msgf = 34 Then Beep
            If msgf = 2 Or msgf = 34 Then GoTo L13
            If msgf = 1 Then numle�on = numle�on + 1
        End If
    End If
End If

' Retour
Menu_principal.Show 1
End Sub


' ******************** ITEM 2 du MENU PRINCIPAL : POURQUI ***********************************
' Tous les messages de cette introduction sont des variables pg-- regroup�es dans le global.bas
Public Sub pourqui()
Unload Menu_principal
numle�on = 1
L20:
pagenum = 1
msgtext0 = CRLF + pgib1 + pressez
Msgform.Show 1
If msgf = 33 Then Beep
If msgf = 2 Or msgf = 33 Then GoTo L20
If msgf = 1 Or msgf = 34 Then
L21:
        pagenum = 2
        msgtext0 = CRLF + pgib2 + pressez
        Msgform.Show 1
        If msgf = 33 Then GoTo L20
        If msgf = 2 Then GoTo L21
        If msgf = 1 Or msgf = 34 Then
L22:
        pagenum = 3
        msgtext0 = CRLF + pgib3 + pressez
        Msgform.Show 1
        If msgf = 33 Then GoTo L21
        If msgf = 2 Then GoTo L22
        If msgf = 1 Or msgf = 34 Then
L23:
            pagenum = 4: pagemax = 1
            msgtext0 = CRLF + pgib4 + pressez
            If nivo = msgStandard Then pctok(1, 0) = 100
            If nivo = msgPersonnalis� Then pctok(26, 0) = 100
            Msgform.Show 1
            If msgf = 33 Then GoTo L22
            If msgf = 34 Then Beep
            If msgf = 2 Or msgf = 34 Then GoTo L23
            If msgf = 1 Then numle�on = numle�on + 1
        End If
    End If
End If

' Retour
Menu_principal.Show 1
End Sub


' ******************** ITEM 3 du MENU PRINCIPAL : CONSEILS **********************************
' Tous les messages de cette introduction sont des variables pg-- regroup�es dans le global.bas
Public Sub conseils()
Unload Menu_principal
numle�on = 2
L30:
pagenum = 1
msgtext0 = CRLF + pgic1 + pressez
Msgform.Show 1
If msgf = 33 Then Beep
If msgf = 2 Or msgf = 33 Then GoTo L30
If msgf = 1 Or msgf = 34 Then
L31:
    pagenum = 2
    msgtext0 = CRLF + pgic2 + pressez
    Msgform.Show 1
    If msgf = 33 Then GoTo L30
    If msgf = 2 Then GoTo L31
    If msgf = 1 Or msgf = 34 Then
L32:
        pagenum = 3
        msgtext0 = CRLF + pgic3 + pressez
        Msgform.Show 1
        If msgf = 33 Then GoTo L31
        If msgf = 2 Then GoTo L32
        If msgf = 1 Or msgf = 34 Then
L33:
            pagenum = 4
            msgtext0 = CRLF + pgic4 + pressez
            Msgform.Show 1
            If msgf = 33 Then GoTo L32
            If msgf = 2 Then GoTo L33
            If msgf = 1 Or msgf = 34 Then
L34:
                pagenum = 5: pagemax = 1
                msgtext0 = CRLF + pgic5 + pressez
                If nivo = msgStandard Then pctok(2, 0) = 100
                If nivo = msgPersonnalis� Then pctok(27, 0) = 100
                Msgform.Show 1
                If msgf = 33 Then GoTo L33
                If msgf = 34 Then Beep
                If msgf = 2 Or msgf = 34 Then GoTo L34
                If msgf = 1 Then numle�on = numle�on + 1
            End If
        End If
    End If
End If

' Retour
Menu_principal.Show 1
End Sub


' ********************* IsCapsLockOn  *****************************************************
' D�tecte l'�tat de la touche Verrouilage-Majuscule
Public Function IsCapsLockOn() As Boolean
Dim o As OSVERSIONINFO
o.dwOSVersionInfoSize = Len(o)
GetVersionEx o
Dim keys(0 To 255) As Byte
GetKeyboardState keys(0)
IsCapsLockOn = keys(VK_CAPITAL)
End Function
' ********************* IsNumLockOn  *****************************************************
' D�tecte l'�tat du Verrouilage-Num�rique
Public Function IsNumLockOn() As Boolean
Dim o As OSVERSIONINFO
o.dwOSVersionInfoSize = Len(o)
GetVersionEx o
Dim keys(0 To 255) As Byte
GetKeyboardState keys(0)
IsNumLockOn = keys(VK_NUMLOCK)
End Function
' ********************* IsScrollLockOn  ****************************************************
' D�tecte l'�tat du Arr�tD�fil
Public Function IsScrollLockOn()
Dim o As OSVERSIONINFO
o.dwOSVersionInfoSize = Len(o)
GetVersionEx o
Dim keys(0 To 255) As Byte
GetKeyboardState keys(0)
IsScrollLockOn = keys(VK_SCROLL)
End Function


'************  SETKEYS Switches CAPSLOCK, NUMLOCK, SCROLL ON or OFF  **********************
'**** Un simple sendkeys ne fonctionnerait pas sous Win XP, d'o� cette longue routine *****
'**** supposant les API d�clar�es keybd_event, getkeyboardstate, setkeyboardstate...  *****
Public Function SetKeys(Optional capslock As Variant)
Dim keys(0 To 255) As Byte
Dim o As OSVERSIONINFO
Dim CapsLockState As Boolean
Dim CapsLockStateBis As Boolean
Dim NumLockState As Boolean
Dim ScrollLockState As Boolean
o.dwOSVersionInfoSize = Len(o)
GetVersionEx o
GetKeyboardState keys(0)
CapsLockState = keys(VK_CAPITAL)
CapsLockStateBis = keys(VK_CAPITAL_BIS)
NumLockState = keys(VK_NUMLOCK)
ScrollLockState = keys(VK_SCROLL)
Select Case capslock

Case "CAPSLOCK_ON"
If CapsLockState = False Then
    'If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    If o.dwPlatformId = 1 Then  ' Mieux reconnu par Win95/Win98
        ' For Win 95/98
        keys(VK_CAPITAL) = 1
        SetKeyboardState keys(0)
    ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        ' For Win XP : Press and release either key itself, either key "MAJ", as per user adjustment !
        keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        keybd_event VK_CAPITAL_BIS, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        keybd_event VK_CAPITAL_BIS, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If
End If

Case "CAPSLOCK_OFF"
If CapsLockState = True Then
    'If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    If o.dwPlatformId = 1 Then  ' Mieux reconnu par Win95/Win98
        ' For Win 95/98
        keys(VK_CAPITAL) = 0
        SetKeyboardState keys(0)
    ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        ' For Win XP : Press and release either key itself, either key "MAJ", as per user adjustment !
        keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        keybd_event VK_CAPITAL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
        keybd_event VK_CAPITAL_BIS, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        keybd_event VK_CAPITAL_BIS, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If
End If

Case "NUMLOCK_ON"
If NumLockState = False Then
    'If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    If o.dwPlatformId = 1 Then  ' Mieux reconnu par Win95/Win98
        If winstop = 2 Then winstop = 1 ' Astuce n�cessaire en Win95/Win98 pour Le�on16D
        keys(VK_NUMLOCK) = 1
        SetKeyboardState keys(0)
    ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If
End If

Case "NUMLOCK_OFF"
If NumLockState = True Then
    'If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    If o.dwPlatformId = 1 Then  ' Mieux reconnu par Win95/Win98
        keys(VK_NUMLOCK) = 0
        SetKeyboardState keys(0)
    ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        keybd_event VK_NUMLOCK, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If
End If

Case "SCROLLLOCK_ON"
If ScrollLockState = False Then
    'If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    If o.dwPlatformId = 1 Then  ' Mieux reconnu par Win95/Win98
        keys(VK_SCROLL) = 1
        SetKeyboardState keys(0)
    ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If
End If

Case "SCROLLLOCK_OFF"
If ScrollLockState = True Then
    'If o.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    If o.dwPlatformId = 1 Then  ' Mieux reconnu par Win95/Win98
        keys(VK_SCROLL) = 0
        SetKeyboardState keys(0)
    ElseIf o.dwPlatformId = VER_PLATFORM_WIN32_NT Then
        keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or 0, 0
        keybd_event VK_SCROLL, &H45, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    End If
End If

Case Else
MsgBox " SetKeys : �CHEC sur la touche Verrouillage ", 0, ""
End Select
End Function


' ******************** Emettre Joue un son ************************************************************
Public Sub Emettre(Frequence As String) ' ajout avril 2008
If txtTps = "" Then 'Si le txttps est vide
    'MsgBox "Vous devez inscrire la dur�e du son", vbCritical, "Attention!"
    txtTps = "" 'Vide txttps
    Exit Sub
ElseIf Val(txtTps) > 20 Then
    'MsgBox "La dur�e du son, en secondes, ne doit pas d�passer 20 secondes", vbCritical, "Attention!"
    txtTps = "" 'Vide txttps
    Exit Sub
Else 'Si txttps n'est ni vide, ni supp�rieur � 20
    APIBeep Frequence, txtTps * 1000
    'Joue le son d'une fr�quence donn�e, avec la dur�e voulue
    'La dur�e est d�finie par txttps (*1000 car le PC comprend la dur�e en millisecondes
    lblFreq = Frequence 'Indique la fr�quence dans lblFreq
End If
End Sub



' ***************  restore_locks  Restore les touches de Verrouillage CAPSLOCK...  *********
Public Sub restore_locks()
' CapsLock (pb : �viter select multiples incongrus)
'If bCapsLockState = "False" Then Module_routines.SetKeys "CAPSLOCK_OFF"
'If bCapsLockState = "True" Then Module_routines.SetKeys "CAPSLOCK_ON"
Module_routines.SetKeys "CAPSLOCK_OFF"

' Numlock
If bNumLockState = "False" Then Module_routines.SetKeys "NUMLOCK_OFF"
If bNumLockState = "True" Then Module_routines.SetKeys "NUMLOCK_ON"

' ScrollLock
If bScrollLockState = "False" Then Module_routines.SetKeys "SCROLLLOCK_OFF"
If bScrollLockState = "True" Then Module_routines.SetKeys "SCROLLLOCK_ON"
End Sub


' ***************  SCROLLRESULTS  D�filement page par page par MSGFORM  **************
Public Sub scrollresults(start, qty)
    msgtext0 = ""
    Close #3
    Open vfileresults For Input As #3
    For jj = 1 To start
        If Not EOF(3) Then
            Line Input #3, currentline
        End If
    Next jj
    For jj = 1 To qty
        If Not EOF(3) Then
            Line Input #3, currentline
            msgtext0 = msgtext0 & CRLF & currentline
        Else
            stopscroll = 1
        End If
    Next jj
    msgtext0 = msgtext0 + pressez
SCR1:
    ' Pas de pagenum = 0 !
    Msgform.Show 1
    If msgf = 33 Then stopscroll = 0
    If msgf = 2 Then GoTo SCR1
    If msgf = 0 Then
        Close #3
        Menu_principal.Show 1
    End If
    Close #3
End Sub


' ******************** ZOOMFORM : ZFACTOR quantifi� selon r�solution �cran  ***************
' modifi� mars 2008, ajout de d�tection de la hauteur de la premi�re fen�tre "Bienvenue" pour tenir sur les �crans 16/9� aussi
' IMPORTANT : le zoom zfactor est d�fini par la largeur de la premi�re fen�tre "Bienvenue" !
Public Sub zoomform(forme_courante)
scrw = Screen.Width
frmw = forme_courante.Width
scrh = Screen.Height
frmh = forme_courante.Height
' zoom par rapport � la largeur et hauteur de la fen�tre montr�e
zfactor = (scrw + scrh) / (frmw + frmh)
' zfactor risque d'�tre inf�rieur � 1 dans le cas d'une d�finition d'�cran de 640x480,
' ceci est r�solu en construisant des fen�tres Visual Basic pas trop larges.
' En g�n�ral il vaut mieux cr�er un zoom un peu faible pour ne pas sortir de l'�cran
If zfactor > 1 Then zfactor = 0.97 * zfactor
'MsgBox scrw & "   " & scrh & "   " & frmw & "   " & frmh & "   " & scrw / scrh & "   " & zfactor
End Sub


' ***************  DIMOBJECT Restore les dimensions des objets selon �cran  *************
' modifi� 12/2011
Public Sub dimobject(object)
On Error Resume Next
object.Height = zfactor * zoomfactor * object.Height
On Error Resume Next
object.Width = zfactor * zoomfactor * object.Width
On Error Resume Next
object.Top = zfactor * zoomfactor * object.Top
On Error Resume Next
object.Left = zfactor * zoomfactor * object.Left
On Error Resume Next
object.Font.Size = zfactor * zoomfactor * object.Font.Size
On Error Resume Next
object.X1 = zfactor * zoomfactor * object.X1
On Error Resume Next
object.Y1 = zfactor * zoomfactor * object.Y1
On Error Resume Next
object.X2 = zfactor * zoomfactor * object.X2
On Error Resume Next
object.Y2 = zfactor * zoomfactor * object.Y2
End Sub


' ***************  DIMENSION Restore toutes les dimensions selon �cran  *****************
' modifi� 12/2011
Public Sub Dimension(forme_courante)
' Feuille
forme_courante.Height = zfactor * zoomfactor * forme_courante.Height
forme_courante.Width = zfactor * zoomfactor * forme_courante.Width
forme_courante.Top = Screen.Height / 2 - forme_courante.Height / 2
forme_courante.Left = Screen.Width / 2 - forme_courante.Width / 2

' Objets dans la feuille
On Error Resume Next
Module_routines.dimobject forme_courante.Text0
On Error Resume Next
Module_routines.dimobject forme_courante.text1
On Error Resume Next
Module_routines.dimobject forme_courante.text2
On Error Resume Next
Module_routines.dimobject forme_courante.text3
On Error Resume Next
Module_routines.dimobject forme_courante.text4
On Error Resume Next
Module_routines.dimobject forme_courante.text5
On Error Resume Next
Module_routines.dimobject forme_courante.Label0
On Error Resume Next
Module_routines.dimobject forme_courante.label1
On Error Resume Next
Module_routines.dimobject forme_courante.Label2
On Error Resume Next
Module_routines.dimobject forme_courante.Label3
On Error Resume Next
Module_routines.dimobject forme_courante.Label4
On Error Resume Next
Module_routines.dimobject forme_courante.Label5
On Error Resume Next
Module_routines.dimobject forme_courante.list1
On Error Resume Next
Module_routines.dimobject forme_courante.Quitter
On Error Resume Next
Module_routines.dimobject forme_courante.Continuer
On Error Resume Next
Module_routines.dimobject forme_courante.Pr�c�dent
On Error Resume Next
Module_routines.dimobject forme_courante.Suivant
End Sub

' *********** COLORS 12/2011 ****************************************************************
Public Sub Colors(forme_courante)

' Basic colors
If colorslevel = msgBasicColors Then
    
    forme_courante.BackColor = f_grisfonc�
    
    On Error Resume Next
    forme_courante.list1.BackColor = f_vertp�le
    On Error Resume Next
    forme_courante.list1.ForeColor = f_noir
    
    On Error Resume Next
    forme_courante.Text0.BackColor = f_vertp�le
    On Error Resume Next
    forme_courante.Text0.ForeColor = f_noir

    On Error Resume Next
    forme_courante.text1.BackColor = f_bleuclair
    On Error Resume Next
    forme_courante.text1.ForeColor = f_noir

    On Error Resume Next
    forme_courante.text2.BackColor = f_bleuclair
    On Error Resume Next
    forme_courante.text2.ForeColor = f_noir
    
    On Error Resume Next
    forme_courante.text3.BackColor = f_vertp�le
    On Error Resume Next
    forme_courante.text3.ForeColor = f_noir

    On Error Resume Next
    forme_courante.text4.BackColor = f_vertp�le
    On Error Resume Next
    forme_courante.text4.ForeColor = f_noir

    On Error Resume Next
    forme_courante.text5.BackColor = f_vertp�le
    On Error Resume Next
    forme_courante.text5.ForeColor = f_noir
    
    On Error Resume Next
    forme_courante.Text6.BackColor = f_vertp�le
    On Error Resume Next
    forme_courante.Text6.ForeColor = f_noir
    
    On Error Resume Next
    forme_courante.label1.BackColor = f_gris
    On Error Resume Next
    forme_courante.label1.ForeColor = f_noir

    On Error Resume Next
    forme_courante.Label2.BackColor = f_grisfonc�
    On Error Resume Next
    forme_courante.Label2.ForeColor = f_noir

    On Error Resume Next
    forme_courante.Label3.BackColor = f_grisp�le
    On Error Resume Next
    forme_courante.Label3.ForeColor = f_vert

    On Error Resume Next
    forme_courante.Label4.BackColor = f_gris
    On Error Resume Next
    forme_courante.Label4.ForeColor = f_noir

    On Error Resume Next
    forme_courante.Label5.BackColor = f_gris
    On Error Resume Next
    forme_courante.Label5.ForeColor = f_noir
End If

' Other colors
If colorslevel = msgOtherColors Then
    
    forme_courante.BackColor = f_vertsombre
    
    On Error Resume Next
    forme_courante.list1.BackColor = f_marronsombre
    On Error Resume Next
    forme_courante.list1.ForeColor = f_vertp�le
    
    On Error Resume Next
    forme_courante.Text0.BackColor = f_marronsombre
    On Error Resume Next
    forme_courante.Text0.ForeColor = f_blanc

    On Error Resume Next
    forme_courante.text1.BackColor = f_presquenoir
    On Error Resume Next
    forme_courante.text1.ForeColor = f_blanc

    On Error Resume Next
    forme_courante.text2.BackColor = f_presquenoir
    On Error Resume Next
    forme_courante.text2.ForeColor = f_blanc
    
    On Error Resume Next
    forme_courante.text3.BackColor = f_marronsombre
    On Error Resume Next
    forme_courante.text3.ForeColor = f_blanc

    On Error Resume Next
    forme_courante.text4.BackColor = f_marronsombre
    On Error Resume Next
    forme_courante.text4.ForeColor = f_blanc

    On Error Resume Next
    forme_courante.text5.BackColor = f_marronsombre
    On Error Resume Next
    forme_courante.text5.ForeColor = f_blanc
    
    On Error Resume Next
    forme_courante.Text6.BackColor = f_presquenoir
    On Error Resume Next
    forme_courante.Text6.ForeColor = f_blanc
    
    On Error Resume Next
    forme_courante.label1.BackColor = f_vert
    On Error Resume Next
    forme_courante.label1.ForeColor = f_blanc

    On Error Resume Next
    forme_courante.Label2.BackColor = f_vertsombre
    On Error Resume Next
    forme_courante.Label2.ForeColor = f_bleuclair

    On Error Resume Next
    forme_courante.Label3.BackColor = f_vertp�le
    On Error Resume Next
    forme_courante.Label3.ForeColor = f_vert

    On Error Resume Next
    forme_courante.Label4.BackColor = f_vertsombre
    On Error Resume Next
    forme_courante.Label4.ForeColor = f_blanc

    On Error Resume Next
    forme_courante.Label5.BackColor = f_vertsombre
    On Error Resume Next
    forme_courante.Label5.ForeColor = f_blanc
End If

End Sub


' *******************  MSHOW  **************************************************************
Public Sub mshow(menu) 'avril 2008 puis 12/2011
'If repjawsnames = "" Then menu.Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "." & CRLF & msgSonori & msgNoSono & CRLF & msgKeyboard & clavierType & ", " & country
'If repjawsnames <> "" Then menu.Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "." & CRLF & msgSonori & repjawsnames & CRLF & msgSpeedExp & debexplilevel & msgSpeedGen & debgenlevel & "." & CRLF & msgKeyboard & clavierType & ", " & country
' *** NVDA
repNVDA = ""
If repNVDA = Dir("c:\Program Files (x86)\NVDA", vbDirectory) Then
repNVDA = Dir("c:\Program Files (x86)\NVDA", vbDirectory)
On Error Resume Next
repNVDA = Dir("c:\Program Files\NVDA", vbDirectory)
Else
repNVDA = Dir("c:\Program Files\NVDA", vbDirectory)
On Error Resume Next
repNVDA = Dir("c:\Program Files (x86)\NVDA", vbDirectory)
End If
If repNVDA <> "" Then
repNVDA = "NVDA  "
End If
svnames = repNVDA & repjawsnames

menu.Label2.Font.Size = 7 * zfactor * zoomfactor
If svnames = "" Then menu.Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "." & CRLF & msgSonori & msgNoSono & CRLF & msgBip & biplevel & ". " & msgBipComment & CRLF & msgDisplay & zoomlevel & ", " & colorslevel & ". " & msgBipComment & CRLF & msgKeyboard & clavierType & ", " & country
If svnames <> "" Then menu.Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "." & CRLF & msgSonori & svnames & CRLF & msgSpeedExp & debexplilevel & msgSpeedGen & debgenlevel & "." & CRLF & msgBip & biplevel & ". " & msgBipComment & CRLF & msgDisplay & zoomlevel & ", " & colorslevel & ". " & msgBipComment & CRLF & msgKeyboard & clavierType & ", " & country
End Sub


' *******************  MODIFY the DEFAULT.JKM  **********************************
Public Sub modify_jkm(keytype As String)
' Non utilis� ! : ancienne routine avant l'ajout du Timer9 des le�ons 1 et 13
' Permettait de modifier le fichier default.jkm de Jaws et ainsi de d�sactiver certains codes touches parasites envoy�s par Jaws
' Ces modifications de default.jkm ne sont pas vues par les versions Jaws6, donc elles sont inutilisables

' keytype : Basics �tait mis dans inits
' keytype : Others �tait mis dans le load menu_le�on1 menu_le�on8 menu_le�on12 menu_le�on13 menu_le�on16 menu_le�on17
' BOUCLE sur les r�ps JAWS trouv�s
ii = 0
Do While repj(ii) <> ""
    If repj(ii) <> "." And repj(ii) <> ".." Then
        repjawsfra = ujaws(ii) & repj(ii) & repjawscountry
                
        ' Tester
        If Dir(repjawsfra & "default.jkm") = "" Then GoTo JKMM
        If Dir(repjawsfra & "apprenticlavier_default.jkm") <> "" Then
            On Error Resume Next  ' N�cessaire si commandes � cadence rapide
            Kill repjawsfra & "apprenticlavier_default.jkm"
        End If

        ' Ouvrir le default d'input et le apprenticlavier_default.jkm d'output
        Open repjawsfra & "apprenticlavier_default.jkm" For Output As #6
        Open repjawsfra & "default.jkm" For Input As #5
        Do While Not EOF(5)
            Line Input #5, currentline
    
            ' Mises en commentaires par ;ZZ
            ' Pour tout � cause de AltGr avec/sans Espace (Jaws401!) : RightAlt+Control+Ctrl devraient �tre annul�s auparavant par le setup, pas seulement ici ?
            ' Pour "le�on1" et "le�on13" (Jaws401!) : LeftAlt, RightAlt devraient �tre annul�s auparavant par le setup, pas seulement ici ?
            ' Tous menus et le�ons : touches LeftWindows et RightWindows
            ' Pour que les commandes Jaws Ins-F4... marchent apr�s avoir quitt� : RightAlt
            ' Pour que Alt+F4 fonctionne et qu'on retrouve les commandes propres : Control + Ctrl
            ' Pour "le�on1" et "le�on13" : touches Echap, LeftAlt, RightAlt
            ' Pour le�on 8G sur � qui fait beep : LeftShift, RightShift
            ' Pour le�on 13B : RetourArri�re/BackSpace pour Jaws451 et +
            ' Pour le�on 16D : NumInsert pour Jaws451 et +, et Insert pour Jaws500
            ' Pour le�on 17A : NumInsert+ExtInsert en 17A si on utilise la touche AltGr !
            ' Pour le�ons 17B 17C 17D : la plupart de ces touches sauf Win
            If keytype = "Basics" Then
                ' Permanent, d� � un bug Jaws en Version 401
                If UCase(Left(currentline, 16)) = "RIGHTALT=13|3|2|" Then currentline = "RightAlt=13|3|1|" & Right(currentline, Len(currentline) - 16)
                ' Basics
                If UCase(Left(currentline, 8)) = "LEFTALT=" Then currentline = ";ZZLeftAlt=" & Right(currentline, Len(currentline) - 8)
                If UCase(Left(currentline, 9)) = "RIGHTALT=" Then currentline = ";ZZRightAlt=" & Right(currentline, Len(currentline) - 9)
                If UCase(Left(currentline, 8)) = "CONTROL=" Then currentline = ";ZZControl=" & Right(currentline, Len(currentline) - 8)
                If UCase(Left(currentline, 5)) = "CTRL=" Then currentline = ";ZZCtrl=" & Right(currentline, Len(currentline) - 5)
                
                If UCase(Left(currentline, 12)) = "LEFTWINDOWS=" Then currentline = ";ZZLeftWindows=" & Right(currentline, Len(currentline) - 12)
                If UCase(Left(currentline, 13)) = "RIGHTWINDOWS=" Then currentline = ";ZZRightWindows=" & Right(currentline, Len(currentline) - 13)
            End If
            If keytype = "Others" Then
                If UCase(Left(currentline, 9)) = "ESCAPE=UP" Then currentline = ";ZZEscape=Up" & Right(currentline, Len(currentline) - 9)
                If UCase(Left(currentline, 10)) = "LEFTSHIFT=" Then currentline = ";ZZLeftShift=" & Right(currentline, Len(currentline) - 10)
                If UCase(Left(currentline, 11)) = "RIGHTSHIFT=" Then currentline = ";ZZRightShift=" & Right(currentline, Len(currentline) - 11)
                If UCase(Left(currentline, 14)) = "BACKSPACE=JAWS" Then currentline = ";ZZBackSpace=Jaws" & Right(currentline, Len(currentline) - 14)
                If UCase(Left(currentline, 10)) = "EXTINSERT=" Then currentline = ";ZZExtInsert=" & Right(currentline, Len(currentline) - 10)
                If UCase(Left(currentline, 10)) = "NUMINSERT=" Then currentline = ";ZZNumInsert=" & Right(currentline, Len(currentline) - 10)
                If UCase(Left(currentline, 7)) = "INSERT=" Then currentline = ";ZZInsert=" & Right(currentline, Len(currentline) - 7)
            End If
            Print #6, currentline
        Loop
        Close #5
        Close #6

        ' Valider le nouveau default.jkm (pas moins de 32 octets, en r�alit� beaucoup plus)
        If Dir(repjawsfra & "apprenticlavier_default.jkm") = "" Then Exit Sub
        If FileLen(repjawsfra & "apprenticlavier_default.jkm") < 32 Then Exit Sub
        
        ' Effacer l'ancien default.jkm, indispensable en Jaws 6, sinon le nouveau default.jkm n'est pas r�ellement appel� par Jaws
        On Error Resume Next
        Kill repjawsfra & "default.jkm"
        
        ' Mettre en place le nouveau default.jkm
        On Error Resume Next
        FileCopy repjawsfra & "apprenticlavier_default.jkm", repjawsfra & "default.jkm"

    End If

JKMM:
    ' Num�ro de version JAWS suivante
    ii = ii + 1
Loop
End Sub


' *******************  RESTORE the DEFAULT.JKM  **********************************
Public Sub restore_jkm(keytype As String)
' Non utilis� : ancienne subroutine avant l'ajout de Timer9, avant Jaws6, pour restauration du fichier default.jkm de Jaws permet de d�sactiver certains codes touches parasites envoy�s par Jaws

' keytype : All �tait mis dans AuRevoir
' keytype : Others �tait mis dans le quitter_click de menu_le�on1 menu_le�on8 menu_le�on12 menu_le�on13 menu_le�on16 menu_le�on17

' BOUCLE sur les r�ps JAWS trouv�s
ii = 0
Do While repj(ii) <> ""
    If repj(ii) <> "." And repj(ii) <> ".." Then
        repjawsfra = ujaws(ii) & repj(ii) & repjawscountry

        ' Tester
        If Dir(repjawsfra & "default.jkm") = "" Then GoTo JKMR
        If Dir(repjawsfra & "apprenticlavier_default.jkm") <> "" Then
            On Error Resume Next  ' N�cessaire si le�on1 + F3 + Alt+F4 � cadence rapide
            Kill repjawsfra & "apprenticlavier_default.jkm"
        End If

        ' Ouvrir le default d'input et le apprenticlavier_default.jkm d'output
        Open repjawsfra & "apprenticlavier_default.jkm" For Output As #11
        Open repjawsfra & "default.jkm" For Input As #10
        Do While Not EOF(10)
            Line Input #10, currentline
    
            ' Restaurations
            If keytype = "All" Then
                ' Des blancs pourraient appara�tre � droite du ";" d� � un automatisme d'�dition
                If UCase(Left(currentline, 3)) = ";ZZ" Then currentline = Right(currentline, Len(currentline) - 3)
                If UCase(Left(currentline, 4)) = "; ZZ" Then currentline = Right(currentline, Len(currentline) - 4)
                If UCase(Left(currentline, 5)) = ";  ZZ" Then currentline = Right(currentline, Len(currentline) - 5)
                If UCase(Left(currentline, 6)) = ";   ZZ" Then currentline = Right(currentline, Len(currentline) - 6)
                If UCase(Left(currentline, 7)) = ";    ZZ" Then currentline = Right(currentline, Len(currentline) - 7)
                
                ' Pour annuler les blancs dans les lignes ZY permanentes si elles existent
                If UCase(Left(currentline, 4)) = "; ZY" Then currentline = ";ZY" & Right(currentline, Len(currentline) - 4)
                If UCase(Left(currentline, 5)) = ";  ZY" Then currentline = ";ZY" & Right(currentline, Len(currentline) - 5)
                If UCase(Left(currentline, 6)) = ";   ZY" Then currentline = ";ZY" & Right(currentline, Len(currentline) - 6)
                If UCase(Left(currentline, 7)) = ";    ZY" Then currentline = ";ZY" & Right(currentline, Len(currentline) - 7)
            End If
            
            If keytype = "Others" Then
                If UCase(Left(currentline, 12)) = ";ZZESCAPE=UP" Then currentline = "Escape=Up" & Right(currentline, Len(currentline) - 12)
                If UCase(Left(currentline, 13)) = "; ZZESCAPE=UP" Then currentline = "Escape=Up" & Right(currentline, Len(currentline) - 13)
                If UCase(Left(currentline, 14)) = ";  ZZESCAPE=UP" Then currentline = "Escape=Up" & Right(currentline, Len(currentline) - 14)
                
                If UCase(Left(currentline, 13)) = ";ZZLEFTSHIFT=" Then currentline = "LeftShift=" & Right(currentline, Len(currentline) - 13)
                If UCase(Left(currentline, 14)) = "; ZZLEFTSHIFT=" Then currentline = "LeftShift=" & Right(currentline, Len(currentline) - 14)
                If UCase(Left(currentline, 15)) = ";  ZZLEFTSHIFT=" Then currentline = "LeftShift=" & Right(currentline, Len(currentline) - 15)
                
                If UCase(Left(currentline, 14)) = ";ZZRIGHTSHIFT=" Then currentline = "RightShift=" & Right(currentline, Len(currentline) - 14)
                If UCase(Left(currentline, 15)) = "; ZZRIGHTSHIFT=" Then currentline = "RightShift=" & Right(currentline, Len(currentline) - 15)
                If UCase(Left(currentline, 16)) = ";  ZZRIGHTSHIFT=" Then currentline = "RightShift=" & Right(currentline, Len(currentline) - 16)
                
                If UCase(Left(currentline, 17)) = ";ZZBACKSPACE=JAWS" Then currentline = "BackSpace=Jaws" & Right(currentline, Len(currentline) - 17)
                If UCase(Left(currentline, 18)) = "; ZZBACKSPACE=JAWS" Then currentline = "BackSpace=Jaws" & Right(currentline, Len(currentline) - 18)
                If UCase(Left(currentline, 19)) = ";  ZZBACKSPACE=JAWS" Then currentline = "BackSpace=Jaws" & Right(currentline, Len(currentline) - 19)
                                
                If UCase(Left(currentline, 13)) = ";ZZNUMINSERT=" Then currentline = "NumInsert=" & Right(currentline, Len(currentline) - 13)
                If UCase(Left(currentline, 14)) = "; ZZNUMINSERT=" Then currentline = "NumInsert=" & Right(currentline, Len(currentline) - 14)
                If UCase(Left(currentline, 15)) = ";  ZZNUMINSERT=" Then currentline = "NumInsert=" & Right(currentline, Len(currentline) - 15)
                
                If UCase(Left(currentline, 13)) = ";ZZEXTINSERT=" Then currentline = "ExtInsert=" & Right(currentline, Len(currentline) - 13)
                If UCase(Left(currentline, 14)) = "; ZZEXTINSERT=" Then currentline = "ExtInsert=" & Right(currentline, Len(currentline) - 14)
                If UCase(Left(currentline, 15)) = ";  ZZEXTINSERT=" Then currentline = "ExtInsert=" & Right(currentline, Len(currentline) - 15)
                
                If UCase(Left(currentline, 10)) = ";ZZINSERT=" Then currentline = "Insert=" & Right(currentline, Len(currentline) - 10)
                If UCase(Left(currentline, 11)) = "; ZZINSERT=" Then currentline = "Insert=" & Right(currentline, Len(currentline) - 11)
                If UCase(Left(currentline, 12)) = ";  ZZINSERT=" Then currentline = "Insert=" & Right(currentline, Len(currentline) - 12)
                End If
            
            Print #11, currentline
        Loop
        Close #10
        Close #11

        ' Valider le nouveau default.jkm (pas moins de 32 octets, en r�alit� beaucoup plus)
        If Dir(repjawsfra & "apprenticlavier_default.jkm") = "" Then Exit Sub
        If FileLen(repjawsfra & "apprenticlavier_default.jkm") < 32 Then Exit Sub
        On Error Resume Next
        FileCopy repjawsfra & "apprenticlavier_default.jkm", repjawsfra & "default.jkm"

        ' M�nage
        On Error Resume Next
        Kill repjawsfra & "apprenticlavier_default.jkm"

    End If
    
JKMR:
    ' Num�ro de version JAWS suivante
    ii = ii + 1
Loop
End Sub


' *******************  MODIFY the SYMBOLS.INI  **********************************
' Pour les versions Jaws avant Jaws5, ceci permet d'�viter des prononciations parasites des touches par Jaws
' Fl�che dit Vide, Retour-Arri�re dit Espace
Public Sub Modify_symbols(keytype As String)
' keytype : Arrows
' BOUCLE sur les r�ps JAWS trouv�s
ii = 0
Do While repj(ii) <> ""
    If repj(ii) <> "." And repj(ii) <> ".." Then
        repjawsfra = ujaws(ii) & repj(ii) & repjawscountry
                
        ' Tester
        If Dir(repjawsfra & "symbols.ini") = "" Then GoTo SYMM
        If Dir(repjawsfra & "symbols_default.ini") <> "" Then
            On Error Resume Next  ' N�cessaire si commandes � cadence rapide
            Kill repjawsfra & "symbols_default.ini"
        End If

        ' Ouvrir le symbols.ini d'input et le symbols_default.ini d'output
        Open repjawsfra & "symbols_default.ini" For Output As #13
        Open repjawsfra & "symbols.ini" For Input As #12
        Do While Not EOF(12)
            Line Input #12, currentline
    
            ' 1 modification
            ' Pour "le�on1" et "le�on13" : touches Fl�ches
            If keytype = "Arrows" Then
                If UCase(currentline) = "BLANK=VIDE" Then currentline = "Blank=�"   ' Alt255 final
            End If
            Print #13, currentline
        Loop
        Close #12
        Close #13

        ' Valider le nouveau symbols.ini (pas moins de 32 octets, en r�alit� beaucoup plus)
        If Dir(repjawsfra & "symbols_default.ini") = "" Then Exit Sub
        If FileLen(repjawsfra & "symbols_default.ini") < 32 Then Exit Sub
        On Error Resume Next
        FileCopy repjawsfra & "symbols_default.ini", repjawsfra & "symbols.ini"

    End If

SYMM:
    ' Num�ro de version JAWS suivante
    ii = ii + 1
Loop
End Sub


' *******************  RESTORE the SYMBOLS.INI  **********************************
Public Sub restore_symbols(keytype As String)
' keytype : All
' BOUCLE sur les r�ps JAWS trouv�s
ii = 0
Do While repj(ii) <> ""
    If repj(ii) <> "." And repj(ii) <> ".." Then
        repjawsfra = ujaws(ii) & repj(ii) & repjawscountry

        ' Tester
        If Dir(repjawsfra & "symbols.ini") = "" Then GoTo SYMR
        If Dir(repjawsfra & "symbols_default.ini") <> "" Then
            On Error Resume Next  ' N�cessaire si commandes � cadence rapide
            Kill repjawsfra & "symbols_default.ini"
        End If

        ' Ouvrir le default d'input et le default d'output
        Open repjawsfra & "symbols_default.ini" For Output As #15
        Open repjawsfra & "symbols.ini" For Input As #14
        Do While Not EOF(14)
            Line Input #14, currentline
    
            ' 1 ligne � restaurer
            ' Pour le�ons 1 et le�ons 13 : touches Fl�ches
            If keytype = "All" Then
                If UCase(currentline) = "BLANK=�" Then currentline = "Blank=Vide"  ' Avec Alt 255 final
                If UCase(Right(currentline, 6)) = "BLANK=" Then currentline = "Blank=Vide"  ' Sans Alt255 final
            End If
            
            Print #15, currentline
        Loop
        Close #14
        Close #15

        ' Valider le nouveau symbols.ini (pas moins de 32 octets, en r�alit� beaucoup plus)
        If Dir(repjawsfra & "symbols_default.ini") = "" Then Exit Sub
        If FileLen(repjawsfra & "symbols_default.ini") < 32 Then Exit Sub
        On Error Resume Next
        FileCopy repjawsfra & "symbols_default.ini", repjawsfra & "symbols.ini"

        ' M�nage
        On Error Resume Next
        Kill repjawsfra & "symbols_default.ini"

    End If
    
SYMR:
    ' Num�ro de version JAWS suivante
    ii = ii + 1
Loop
End Sub


' ****************  HELP_F1  *****************************************************************
' Aide g�n�rale
Public Sub help_f1(le�on)
If f1msgform = 1 Then Exit Sub
On Error Resume Next
If le�on.text1.Text = "F1" Then Exit Sub
HF1:
msgtext0 = msgAide
If scorecourant <> "" Then msgtext0 = msgtext0 & CRLF & msgScore & scorecourant & CRLF
msgtext0 = msgtext0 & CRLF & msgCommandesDispo
msgtext0 = msgtext0 & CRLF & msgF1Aide & CRLF & msgF2Loc & CRLF & msgF3AM
If typele�on >= 2 Then msgtext0 = msgtext0 & CRLF & msgEspace
If typele�on >= 3 Then msgtext0 = msgtext0 & CRLF & msgCtrlEspace & CRLF & msgAltEspace & CRLF & msgMajEspace
If echapbismax = 0 Then msgtext0 = msgtext0 & CRLF & msgSortir & CRLF & msgAltF4 & pressez
If echapbismax = 1 Then msgtext0 = msgtext0 & CRLF & msgSortir2 & CRLF & msgAltF4 & pressez
If echapbismax = 2 Then msgtext0 = msgtext0 & CRLF & msgSortir3 & CRLF & msgAltF4 & pressez
fbc = fbc_f1
ffc = ffc_f1

' Voici la fen�tre d'aide sans titre page
pagenum = 0
fullscreeninhibit = 1
Msgform.Show 1
fullscreeninhibit = 0
fbc = fbc_default
ffc = ffc_default
If msgf = 2 Then GoTo HF1
On Error Resume Next
le�on.text2.SetFocus
noechapF1 = 0
End Sub


' ****************  HELP_F1M  *****************************************************************
' Aide quand on se trouve dans un menu, modifi� 12/2011 zoom et couleurs
Public Sub help_f1m()
keyinhibit = 2

svnames = repNVDA & repjawsnames

HF1M:
If nivo = "" Then
    If svnames = "" Then msgtext0 = msgAide & CRLF & msgNoficSono & CRLF2 & msgCommandesDispo & CRLF & msgF3AM & CRLF & msgEntr�eContinuer
    'If svnames <> "" Then msgtext0 = msgAide & CRLF & svnames & msgD�tect� & CRLF2 & msgCommandesDispo & CRLF & msgF3AM & CRLF & msgEntr�eContinuer
    If svnames <> "" Then msgtext0 = msgAide & CRLF & "Vocalisation : " & svnames & CRLF2 & msgCommandesDispo & CRLF & msgF3AM & CRLF & msgEntr�eContinuer
Else
    If svnames = "" Then msgtext0 = msgAide & CRLF & msgNoficSono & CRLF & msgUserIs & nom & "." & CRLF & msgLevelIs & UCase(nivo) & "." & CRLF & msgSpeedExpIs & UCase(debexplilevel) & "." & CRLF & msgSpeedGenIs & UCase(debgenlevel) & "." & CRLF & msgBipsAre & UCase(biplevel) & "." & CRLF & zoomlevel & ". " & colorslevel & "." & CRLF2 & msgCommandesDispo & CRLF & msgChoisir & CRLF & msgOptions & CRLF & msgF1Aide & CRLF & msgF3AM
    'If svnames <> "" Then msgtext0 = msgAide & CRLF & svnames & msgD�tect� & CRLF & msgUserIs & nom & "." & CRLF & msgLevelIs & UCase(nivo) & "." & CRLF & msgSpeedExpIs & UCase(debexplilevel) & "." & CRLF & msgSpeedGenIs & UCase(debgenlevel) & "." & CRLF & msgBipsAre & UCase(biplevel) & "." & CRLF & zoomlevel & ". " & colorslevel & "." & CRLF2 & msgCommandesDispo & CRLF & msgChoisir & CRLF & msgOptions & CRLF & msgF1Aide & CRLF & msgF3AM
    If svnames <> "" Then msgtext0 = msgAide & CRLF & "Vocalisation : " & svnames & CRLF & msgUserIs & nom & "." & CRLF & msgLevelIs & UCase(nivo) & "." & CRLF & msgSpeedExpIs & UCase(debexplilevel) & "." & CRLF & msgSpeedGenIs & UCase(debgenlevel) & "." & CRLF & msgBipsAre & UCase(biplevel) & "." & CRLF & zoomlevel & ". " & colorslevel & "." & CRLF2 & msgCommandesDispo & CRLF & msgChoisir & CRLF & msgOptions & CRLF & msgF1Aide & CRLF & msgF3AM
End If
msgtext0 = msgtext0 & CRLF & msgSortir & CRLF & msgAltF4 & pressez
fbc = fbc_f1
ffc = ffc_f1

' Voici la fen�tre d'aide sans titre page
pagenum = 0
fullscreeninhibit = 1
Msgform.Show 1
fullscreeninhibit = 0
fbc = fbc_default
ffc = ffc_default
If msgf = 2 Then GoTo HF1M
End Sub


' ****************  HELP_F3  ****************************************************************
' Aide-m�moire sauf dans les menus
Public Sub help_f3(le�on)
If Left(le�on.text1.Text, lt1) = "F3" Then Exit Sub
Set le�on_courante = le�on
Aidef3.Show 1
End Sub


' ****************  HELP_F3M  ****************************************************************
' Aide-M�moire quand on se trouve dans un menu
Public Sub help_f3m()
Aidef3.Show 1
End Sub


' ********************  QUIT_L  Quitter une le�on  *****************************************
Public Sub quit_l() 'Public n�cessaire pour score
On Error Resume Next
Unload le�on_courante
Close #1
bascule = 0
bipinhibit = 0
derligne = 0
elapsed = 0
erepeat = 0
espacevalid = 0
f2link = 0
fullscreeninhibit = 0
iiante = 0
iiprec = 0
iistartp = 0
inexo = 0
iter = 0
iwrong = 0
iwrongbis = 0
iwrongbismax = 5
iwrongCRmax = 0
iwrongl = 0
KeyFirst = 0: KeySecond = 0: KeyThird = 0
keyinhibit = 0
le�onfontsize5 = 18 * zoomvalue '12/2011
llold = -2
noalt = 1
nodoublesono = 0
noechapF1 = 0
noF1 = 0
notab = 1 'septembre 2007
numpad = 0
passb = 0
pasdepoint = 0
scorecourant = "100 %"
timevalid = 0
typele�on = 0
zz = 0
fautecourante = ""
fauteprec = ""
For jj = 0 To 149
    fautesur(jj) = ""
Next jj
For jj = 0 To 149
    nboccur(jj) = 0
Next jj
Module_routines.register
Module_routines.SetKeys "CAPSLOCK_OFF"
Module_routines.SetKeys "NUMLOCK_OFF"
Module_routines.SetKeys "SCROLLLOCK_OFF"

' Faut-il PASSER au MENU de la LE�ON SUIVANTE ?
Module_routines.pass
End Sub


' *****************  QuitQuit (quitter d�finitivement) ***********************************
Public Sub QuitQuit()
quitactive = 1
Module_routines.register
Module_routines.clean
Module_routines.AuRevoir
End Sub


' ***************  Au REVOIR  Ecran de sortie pour quitter  *******************************
Public Sub AuRevoir()
quitactive = 1  ' Replacer ici, pour le cas Bienvenue avec Alt+F4 r�p�t�

' Restaurer les fichiers Jaws
'Module_routines.restore_jkm "All"
Module_routines.restore_symbols "All"

pagenum = 0  ' Placer ici, sinon le "Au revoir" par Echap + Alt+F4 prend le titre : "Page x"
Unload Msgform
If altf4 = 1 Then
    msgtext0 = CRLF + " Alt+F4.  " & msgAurevoir & " "
Else
    msgtext0 = CRLF + "     " & msgAurevoir & "    "
End If
fsize = 3 * fsizedefault * zfactor
fbc = fbc_quit
ffc = ffc_quit
Msgform.Quitter.Caption = msgQuitter & msg�chap
Msgform.Quitter.Visible = "true"
Msgform.Continuer.Visible = "false"
timeout = 1  ' Pour quitter l'�cran de sortie automatiquement
keyinhibit = 2
Msgform.Show 1

' Restaurer les touches de verrouillage
Module_routines.restore_locks
End
End Sub

