Attribute VB_Name = "Module_SetUpGlobal"
'Ce logiciel libre est disponible sous licence GNU/GPL,
'dont une copie se trouvera dans le fichier gpl.txt,
'avec une traduction française non officielle gpl-fr.txt.

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Global CRLF, CRLF2, CRLF3, CRLF4 As String 'carrierreturn-linefeed (13-10) une, deux, trois, quatre, cinq, 10 fois
Global bannerversion, bannercopyright, clavierType As String  'bannière_app_version, bannière_copyright, type_clavier
Global repj(9), repjexe(9), vpath, windir As String 'rep_config_jaws_si_plusieurs, rep_jaws_exe, path_de_setup, path_windows
Global repjaws, repUsers, repjawscountry, country, repjawsfra, ujaws(9), repjawsconfig, repjawsnames As String 'rep_jaws_base, pour version Windows, rep_jaws_settings, pays, rép_jaws_pour_jcf, version_récente_jaws, unité_rép_jaws, sous-rep_des_config, all_jaws_reps_names
Global repsys, repsys1, repsys2, repsys3, repsys4 As String 'répertoires_pour_installer
Global tempo, tempo1, tempo2, tempo3, tempo4, texte_bienvenue As String 'variables
Global info, info_lance, annul, msgtext0, pressez As String 'info, info_lance, info_suppressio, info_annulation, texte_bienvenue, message_pressez
Global msgAurevoir, currentline As String 'message_au_revoir, current_line
Global jsspath, lperso As String 'jss_path, noms_joker_des_leçons_avec_chemin
Global ii, kk, inst As Integer 'variables_de_boucle, inst=-1_init_ou_=0_désinst_ou_=1_en_cours_d'inst
Global echapbis, vmsg As Integer 'echap_bissée
Global fbackcolor, fbackcolor_default, fbackcolor_special, fcolor, fcolor_default, fcolor_special, fsize As Integer 'font_back_color, font_color, font_size
Global keycode, shift, keyascii As Integer 'keycode_réponse_user, shift_réponse_user, keyascii_user
Global rrs(9), rrt(9) As Integer 'variables_selon_le_type_de_version_Jaws
Global keyinhibit, timein, timebienv, timeout, msgf, stopscroll, f1expli As Byte 'inhibit_keyup_after_msgbox, timer_msgform, timer_bienvenue, timeout=1_pour_quitter, réponse_msgform, stopscroll, f1expli
Global FullScreenSwitch As Byte 'FullScreenSwitch = 1 for full screen show, FullScreenSwitch = 1 for application
Global scrw, scrh, frmw, frmh, zfactor As Variant 'screen_width, height, form_width, height, zoom_factor
Global msgInstall, msgDésinst, msgDésinstall, msgErreur1, msgNoSono, msgInfo, msgPatientez, msgNoFic, msgDétecté As String
Global msgSupprimé, msgASupprimer, msgPage, msgVousEtiez, msgRecommencez, msgEchap, msgAttention As String
Global msgInstaller, msgDésinstaller, msgAide, msgAnnuler, msgContinuer, msgQuitter, msgKeyboard As String


' **************** MAIN *******************************************************************
Public Sub main()
' Touches à pb :
' Alt Droit 17, Alt Gauche 18, Win Gauche 91, Win Droit 92, Menu contextuel 93, Renvoyé par Sendkeys 145

' Paramètres de la ligne de commande
' /I  Installer
' /D  Désinstaller
' /DQ Désinstallation silencieuse (Quiet)
' /P  Simple Message "Patientez"

' Variables
FullScreenSwitch = 1 ' FullScreenSwitch = 0 for debug, FullScreenSwitch = 1 for application
CRLF = Chr(13) + Chr(10)
CRLF2 = Chr(13) + Chr(10) + CRLF
CRLF3 = Chr(13) + Chr(10) + CRLF2
CRLF4 = Chr(13) + Chr(10) + CRLF3
tempo = "": tempo1 = "": tempo2 = "": tempo3 = ""
echapbis = 0
keyinhibit = 0
fcolor_default = &H80000008
fbackcolor_default = &HFFFF00 '12/2011
fcolor_special = &H4040
fbackcolor_special = &H80C0FF
fcolor = fcolor_default
fbackcolor = fbackcolor_default

fsize = 14
f1expli = 0
inst = -1
timein = 0: timebienv = 0: timeout = 0

' VARIABLES à TRADUIRE
Module_SetUpVar.SetUpVariables

' Répertoires
repsys = "": repsys1 = "": repsys2 = "": repsys3 = "": repsys4 = ""
repjaws = "": repjawsfra = "": repjawsconfig = "": repjawsnames = ""

' Chemins du fichier exe, du user, de all_users, de windows
vpath = App.Path
If Right(vpath, 1) <> "\" Then vpath = vpath & "\"
If Left(LCase(vpath), 19) = "c:\apprenticlavier\" Or Left(LCase(vpath), 12) = "c:\appren~1\" Or Left(LCase(vpath), 12) = "c:\appren~2\" Then
    MsgBox msgErreur1, 0, ""
    End
End If

'Le raccourci bureau sera déterminé par l'installateur NSIS
'userprofile = Environ("USERPROFILE")
'allusersprofile = Environ("ALLUSERSPROFILE")
'windir = Environ("WINDIR")

' Préliminaire
If Command = "/p" Or Command = "/P" Then Module_SetUpGlobal.Patience
Module_SetUpGlobal.SonoCheck
End Sub


'*********  SonoCheck : Préparer la Copie des fichiers de configuration jcf, jsb...  *******
Public Sub SonoCheck()

' Tester la présence de ApprentiClavier.jcf dans le rép où se trouve Setup.exe
If Dir(vpath & "ApprentiClavier.jcf") = "" Then

    ' Initialisation
    If inst = -1 Then
        If Command = "/i" Or Command = "/I" Then
            MsgBox msgInstall & bannerversion, 0, ""
            Module_SetUpGlobal.install
            Exit Sub
        End If
        If Command = "/d" Or Command = "/D" Then
            vmsg = MsgBox(msgDésinstall, 3, "")
            If vmsg = 6 Then
                Module_SetUpGlobal.desinstall
                Exit Sub
            Else
                End
            End If
        End If
        If Command = "/dq" Or Command = "/DQ" Or Command = "/dQ" Or Command = "/Dq" Then
            Module_SetUpGlobal.desinstall
            Exit Sub
        End If
        Module_SetUpGlobal.bienv
    End If
    
    ' En cours d'installation
    If inst = 1 Then
        keyinhibit = 2
        MsgBox msgNoFic & "ApprentiClavier.jcf", 48, ""
        Module_SetUpGlobal.msgfinal
    End If
Else
    
    ' Cas Normal, copier les config vers V-Jaws toujours, puis localiser les réps Jaws
    Module_SetUpGlobal.SonoVocal
    Module_SetUpGlobal.SonoLocate
End If
End Sub


'**************  SonoVocal : Copier les fichiers jcf, jdf, jkf, jss vers V-Jaws  ***********
Public Sub SonoVocal()
' Skipper, si on est en initialisation (inst = -1) , ou en désinstallation (inst = 0)
If inst < 1 Then Exit Sub

' JCF : Copier ApprentiClavier.jcf vers le rép V-Jaws
If Dir("c:\ApprentiClavier\V-Jaws", vbDirectory) = "" Then MkDir "c:\ApprentiClavier\V-Jaws"
If Dir(vpath & "ApprentiClavier.jcf") <> "" Then FileCopy vpath & "ApprentiClavier.jcf", "c:\ApprentiClavier\V-Jaws\ApprentiClavier.jcf"

' JDF : Copier ApprentiClavier.jdf vers le rép V-Jaws (dictionnaire)
If Dir(vpath & "ApprentiClavier.jdf") <> "" Then FileCopy vpath & "ApprentiClavier.jdf", "c:\ApprentiClavier\V-Jaws\ApprentiClavier.jdf"

' JSS : Copier ApprentiClavier.jss vers le rép V-Jaws
If Dir(vpath & "ApprentiClavier.jss") <> "" Then FileCopy vpath & "ApprentiClavier.jss", "c:\ApprentiClavier\V-Jaws\ApprentiClavier.jss"

' JSB : copier ApprentiClavier.jsb (script compilé Version 5 ou Suivantes)
If Dir(vpath & "ApprentiClavier.jsb") <> "" Then FileCopy vpath & "ApprentiClavier.jsb", "c:\ApprentiClavier\V-Jaws\ApprentiClavier.jsb"

' JSB : copier ApprentiClavier-Jaws401.jsb (script compilé version 4.01 pas ultérieur!)
If Dir(vpath & "ApprentiClavier-Jaws401.jsb") <> "" Then FileCopy vpath & "ApprentiClavier-Jaws401.jsb", "c:\ApprentiClavier\V-Jaws\ApprentiClavier-Jaws401.jsb"

End Sub


'*******  SonoLocate : Localiser le Jaws pour les fichiers jcf, jsb... et BIENVENUE  *******
Public Sub SonoLocate()
If inst <> 1 Then repjawsnames = ""
ii = 0

' ****** CODE IDENTIQUE à CELUI du SONOLOCATE de ROUTINE.BAS dans ApprentiClavier.vbp ******
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

' BOUCLE en ujaws = C:\ pour REPÉRER les réps CONFIG de VERSIONS JAWS, si plusieurs réps c:\jaws??? !!
If repjaws <> "" Then
    repj(ii) = repjaws  '1er résultat
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

' BOUCLE en ujaws = D:\ pour REPÉRER les réps CONFIG de VERSIONS JAWS, si plusieurs réps c:\jaws???
If repjaws <> "" Then
    repj(ii) = repjaws  '1er résultat
    
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

' BOUCLE en ujaws = C:\Documents... pour REPÉRER les réps CONFIG de VERSIONS JAWS, si plusieurs réps c:\Documents and Settings\...jaws\?.??
If repjaws <> "" Then
    If repjaws <> "." And repjaws <> ".." Then
        repj(ii) = repjaws  '1er résultat
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

' BOUCLE en ujaws = d:\Documents... pour REPÉRER les réps CONFIG de VERSIONS JAWS, si plusieurs réps d:\Documents and Settings\...jaws\?.??
If repjaws <> "" Then
    If repjaws <> "." And repjaws <> ".." Then
        repj(ii) = repjaws  '1er résultat
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

' Éviter deux affichages même version de jaws
If repUsers = "" Then

' *** VERSIONS JAWS6 et Suivantes, pour Windows XP, en C:\Documents and Settings\All Users\Application Data\Freedom Scientific\Jaws\??.??\Settings\Fra
repjaws = ""
ujaws(ii) = "c:\"
On Error Resume Next
repjaws = Dir("c:\Documents and Settings\All Users\Application Data\Freedom Scientific\Jaws\??.??", vbDirectory) 'septembre 2007
' BOUCLE en ujaws = C:\Documents... pour REPÉRER les réps CONFIG de VERSIONS JAWS, si plusieurs réps c:\Documents and Settings\...jaws\?.??
If repjaws <> "" Then
    If repjaws <> "." And repjaws <> ".." Then
        repj(ii) = repjaws  '1er résultat
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

' BOUCLE en ujaws = D:\Documents... pour REPÉRER les réps CONFIG de VERSIONS JAWS, si plusieurs réps c:\Documents and Settings\...jaws\??.??
If repjaws <> "" Then
    If repjaws <> "." And repjaws <> ".." Then
        repj(ii) = repjaws  '1er résultat
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
'MsgBox "0 " & ujaws(0) & repj(0) & "   1 " & ujaws(1) & repj(1) & "   2 " & ujaws(2) & repj(2) & "   3 " & ujaws(3) & repj(3) & "   4 " & ujaws(4) & repj(4) & "   5 " & ujaws(5) & repj(5) & "   6 " & ujaws(6) & repj(6) & "   7 " & ujaws(7) & repj(7)
' ******* FIN du CODE IDENTIQUE à CELUI du SONOLOCATE de ROUTINE.BAS dans ApprentiClavier.vbp ***************

' Jaws INTROUVABLE
If repj(0) = "" Then
    ujaws(ii) = ""
    ' Si initialisation
    If inst = -1 Then Module_SetUpGlobal.bienv
    ' Si en cours d'installation
    If inst = 1 Then Module_SetUpGlobal.infojaws
    Exit Sub
End If

' BOUCLE D'INSTALLATION/DESINSTALLATION sur les réps JAWS trouvés
ii = 0
Do While repj(ii) <> ""
    If repj(ii) <> "." And repj(ii) <> ".." Then
        repjawsfra = ujaws(ii) & repj(ii) & repjawscountry
        
        ' Tester que le sous-rep jaws des configurations est trouvé
        repjawsconfig = ""
        repjawsconfig = Dir(repjawsfra, vbDirectory)
        
        If repjawsconfig <> "" Then
            ' Type de Version Jaws
            rrs(ii) = InStr(1, repjawsfra, "Documents and Settings\All Users\Application Data\Freedom Scientific\Jaws", 1)
            rrt(ii) = InStr(1, repjawsfra, "ProgramData\Freedom Scientific\Jaws", 1) ' ajout mars 2008
            
            ' Copies (ou effacement, ou restauration) des fichiers de configuration ApprentiClavier
            ' Si annulation de l'installation
            If inst <= -2 Then
                Module_SetUpGlobal.Restore_Jkm "ZZ"
                Module_SetUpGlobal.Restore_Symbols "All"
                Module_SetUpGlobal.KillSetupJcf
            End If
            
            ' Si Initialisation, modifié septembre 2007
            If inst = -1 Then
                ' Versions JAWS 3.7, 4 et 5
                If rrs(ii) = 0 And rrt(ii) = 0 Then  'modifié mars 2008
                    repjawsnames = repjawsnames & LCase(ujaws(ii)) & LCase(repj(ii)) & "  "
                Else
                'Versions JAWS 6 et suivantes
                    If InStr(1, Right(LCase(repj(ii)), 5), "\") = 0 Then repjawsnames = repjawsnames & LCase(ujaws(ii)) & "jaws" & Right(LCase(repj(ii)), 5) & "  "
                    If InStr(1, Right(LCase(repj(ii)), 5), "\") = 1 Then repjawsnames = repjawsnames & LCase(ujaws(ii)) & "jaws" & Right(LCase(repj(ii)), 4) & "  "
                    If InStr(1, Right(LCase(repj(ii)), 5), "\") = 2 Then repjawsnames = repjawsnames & LCase(ujaws(ii)) & "jaws" & Right(LCase(repj(ii)), 3) & "  "
                End If
                Module_SetUpGlobal.SonoDinstall
            End If
                        
            ' Si désinstallation
            If inst = 0 Then Module_SetUpGlobal.deljcf
            
            ' Si en cours d'installation, Sonocopy et Modification des fichiers default.jkm (erreur des anciennes versions Jaws401...)
            If inst = 1 Then
                Module_SetUpGlobal.SonoCopy
                Module_SetUpGlobal.Modify_Jkm "AltCtrl"
            End If
        End If
    End If
    ' Numéro de version JAWS suivante
    ii = ii + 1
Loop

' Boucle terminée, lancer BIENVENUE ou POURSUIVRE
' INIT selon les paramètres de la LIGNE DE COMMANDE
If inst = -1 Then
    If Command = "/i" Or Command = "/I" Then
        MsgBox msgInstall & bannerversion, 0, ""
        Module_SetUpGlobal.install
        Exit Sub
    End If
    If Command = "/d" Or Command = "/D" Then
        vmsg = MsgBox(msgDésinstall, 3, "")
        If vmsg = 6 Then
            Module_SetUpGlobal.desinstall
            Exit Sub
        Else
            End
        End If
    End If
    If Command = "/dq" Or Command = "/DQ" Or Command = "/Dq" Or Command = "/dQ" Then
        Module_SetUpGlobal.desinstall
        Exit Sub
    End If
    Module_SetUpGlobal.bienv
End If

' SUITE de l'INSTALLATION
If inst = 1 Then Module_SetUpGlobal.infojaws
End Sub


'**************  sonocopy : Copier les fichiers jcf, jsb...  ****************************
Public Sub SonoCopy()
' BOUCLE sur les réps JAWS trouvés localisée dans sonolocate qui définit repjawsfra !
' Tester que le sous-rep jaws des configurations est trouvé (voir aussi subroutine infojaws)
repjawsconfig = Dir(repjawsfra, vbDirectory)

' Si PAS DE RÉPERTOIRE CONFIG JAWS, on ne copie pas les config, mais on poursuit vers infojaws et msgfinal
If repjawsconfig = "" Then Exit Sub

' Si on est en init (inst = -1) , ou en désinstallation (inst = 0), on sort
If inst < 1 Then Exit Sub

' JCF : Copier ApprentiClavier.jcf vers le répertoire des configurations JAWS
If Dir(vpath & "ApprentiClavier.jcf") <> "" Then FileCopy vpath & "ApprentiClavier.jcf", repjawsfra & "ApprentiClavier.jcf"

' JDF : Copier ApprentiClavier.jdf vers le rép settings fra de jaws (dictionnaire)
If Dir(vpath & "ApprentiClavier.jdf") <> "" Then FileCopy vpath & "ApprentiClavier.jdf", repjawsfra & "ApprentiClavier.jdf"

' JSS : Copier ApprentiClavier.jss vers le rép settings\fra de jaws (script en clair)
If Dir(vpath & "ApprentiClavier.jss") <> "" Then
    FileCopy vpath & "ApprentiClavier.jss", repjawsfra & "ApprentiClavier.jss"

    ' JSB : Il faut toujours UTILISER LE COMPILATEUR de la version JAWS approprié, pb de compatibilité
    ' Il faudra donner le chemin de l'exécutable scompile : on définit repjexe
    ' Versions JAWS 3.7, 4 et 5
    If rrs(ii) = 0 And rrt(ii) = 0 Then  'modifié mars 2008
        repjexe(ii) = repj(ii)
    Else
    ' Versions JAWS 6 et suivantes, modifié septembre 2007
        If InStr(1, Right(LCase(repj(ii)), 5), "\") = 0 Then repjexe(ii) = "Program Files\Freedom Scientific\Jaws\" & Right(LCase(repj(ii)), 5)
        If InStr(1, Right(LCase(repj(ii)), 5), "\") = 1 Then repjexe(ii) = "Program Files\Freedom Scientific\Jaws\" & Right(LCase(repj(ii)), 4)
        If InStr(1, Right(LCase(repj(ii)), 5), "\") = 2 Then repjexe(ii) = "Program Files\Freedom Scientific\Jaws\" & Right(LCase(repj(ii)), 3)
    End If
    
    ' Localiser le SCOMPILE.EXE
    tempo = ""
    tempo = Dir(ujaws(ii) & repjexe(ii) & "\scompile.exe")
    If tempo <> "" Then
    
        ' Se placer dans le chemin des CONFIG, ou, dans le chemin du compilateur, selon le type de Version JAWS
        ' Y copier le compilateur, ou, le script jss, selon le type de Version JAWS
        ChDrive (ujaws(ii))
        
        ' Versions JAWS 3.7, 4 et 5
        If rrs(ii) = 0 And rrt(ii) = 0 Then ' modifié mars 2008
            ChDir (repjawsfra)
            FileCopy ujaws(ii) & repjexe(ii) & "\scompile.exe", repjawsfra & "scompile.exe"
            
        ' Versions JAWS 6 et Suivantes (copie du jss car selon les versions Jaws, l'emplacement requis pour le jss varie?)
        Else
            ChDir (ujaws(ii) & repjexe(ii))
            FileCopy "c:\ApprentiClavier\V-Jaws\ApprentiClavier.jss", ujaws(ii) & repjexe(ii) & "\ApprentiClavier.jss"
        End If
        
        ' COMPILER le jss
        On Error Resume Next
        Module_exec.ExecAndWait "scompile.exe ApprentiClavier.jss"
        
        ' Versions JAWS 3.7, 4 et 5
        If rrs(ii) = 0 And rrt(ii) = 0 Then 'modifié mars 2008
            ' Ménage
            On Error Resume Next
            Kill repjawsfra & "scompile.exe"
        End If
        
        ' Versions JAWS 6 et Suivantes (rrs pour Windows XP, rrt pour Windows VISTA) modifié mars 2008
        If rrs(ii) > 0 Or rrt(ii) > 0 Then
            If ujaws(ii) & repjexe(ii) & "\" <> repjawsfra Then
                On Error Resume Next
                FileCopy ujaws(ii) & repjexe(ii) & "\ApprentiClavier.jsb", repjawsfra & "ApprentiClavier.jsb"
                ' Ménage
                On Error Resume Next
                Kill ujaws(ii) & repjexe(ii) & "\ApprentiClavier.jss"
                On Error Resume Next
                Kill ujaws(ii) & repjexe(ii) & "\ApprentiClavier.jsb"
            End If
        End If
    End If
End If

' Si ECHEC de SCOMPILE, c'est normalement dû à un Jaws trop ancien
' on copie alors le jsb (obtenu spécifiquement en Version 4.01 max) de V-Jaws
If Dir(repjawsfra & "ApprentiClavier.jsb") = "" Then
        If LCase(repj(ii)) = "jfw37" Or LCase(repj(ii)) = "jaws401" Then
            If Dir("c:\ApprentiClavier\V-Jaws\ApprentiClavier-Jaws401.jsb") <> "" Then FileCopy "c:\ApprentiClavier\V-Jaws\ApprentiClavier-Jaws401.jsb", repjawsfra & "ApprentiClavier.jsb"
        Else
            If Dir("c:\ApprentiClavier\V-Jaws\ApprentiClavier.jsb") <> "" Then FileCopy "c:\ApprentiClavier\V-Jaws\ApprentiClavier.jsb", repjawsfra & "ApprentiClavier.jsb"
        End If
End If

End Sub


'************  Sonorisation Jaws pour l'install ApprentiClavier_Setup  *********************
Public Sub SonoDinstall()

' Copier ApprentiClavier.jcf vers le rép settings\fra de jaws (config)
FileCopy vpath & "ApprentiClavier.jcf", repjawsfra & "ApprentiClavier_Setup.jcf"

' Il ne semble pas utile de copier le jss et le jsb !

End Sub



'**************  Fenêtre de BIENVENUE  **************************************************
Public Sub bienv()
    SetUpBienvenue.Show 1
End Sub



'**************  INSTALL  les fichiers  *************************************************
Public Sub install()

inst = 1
Module_SetUpGlobal.check    'tests de présence pour les fichiers indispensables
Module_SetUpGlobal.dllcopy  'copie la dll pour vb4
Module_SetUpGlobal.dirapp   'rép c:\ApprentiClavier, Leçons\Personnalisé\leçonsxx.txt, lance infojaws, annexes, crée le msgfinal
End Sub



'****************  TESTs de PRÉSENCE des FICHIERS INDISPENSABLES à INSTALLER  ***************
Public Sub check()

If Dir(vpath & "Vb40032.dll") = "" Then
    MsgBox msgNoFic & "VB40032.DLL" + CRLF2 + annul, 48, ""
    inst = -2
    Module_SetUpGlobal.SonoLocate
    End
End If

If Dir(vpath & "apprenticlavier.exe") = "" Then
    MsgBox msgNoFic & "APPRENTICLAVIER.EXE" + CRLF2 + annul, 48, ""
    inst = -2
    Module_SetUpGlobal.SonoLocate
    End
End If

End Sub



'**************  COPIER la DLL de VISUAL BASIC  *******************************************
Public Sub dllcopy()
' Créer le rép c:\ApprentiClavier
tempo = Dir("c:\ApprentiClavier", vbDirectory)
If tempo = "" Then MkDir "c:\ApprentiClavier"

' Copier Vb40032.dll du rép où se trouve Setup.exe vers le rép c:\ApprentiClavier
If UCase(vpath) = "C:\APPRENTICLAVIER\" Or UCase(vpath) = "C:\APPREN~1\" Or UCase(vpath) = "C:\APPREN~2\" Then
    MsgBox msgErreur1
    inst = -2
    Module_SetUpGlobal.SonoLocate
    End
End If
FileCopy vpath & "Vb40032.dll", "c:\ApprentiClavier\Vb40032.dll"
End Sub



'******  DIRapp : COPIER vers C:\APPRENTICLAVIER  ****
Public Sub dirapp()
' Créer le rép c:\ApprentiClavier
If Dir("c:\ApprentiClavier", vbDirectory) = "" Then MkDir "c:\ApprentiClavier"

' Présence de leçons personnalisées dans le rép d'installation ?
' ATTENTION, le rép Perso sera construit par l'installateur NSIS qui fera ensuite appel au Setup
If Dir(vpath & "Perso\*.txt") = "" Then GoTo INST2

' Créer alors le rép c:\ApprentiClavier\Leçons\Personnalisé
If Dir("c:\ApprentiClavier\Leçons", vbDirectory) = "" Then MkDir "c:\ApprentiClavier\Leçons"
If Dir("c:\ApprentiClavier\Leçons\Personnalisé", vbDirectory) = "" Then MkDir "c:\ApprentiClavier\Leçons\Personnalisé"

lperso = vpath & "perso\*.txt"
tempo = Dir(lperso)

' BOUCLE de COPIE sur le contenu du rép Perso (leçons???.txt)
Do While tempo <> ""
    ' Recherche
    If tempo <> "." And tempo <> ".." Then
        FileCopy vpath & "Perso\" & tempo, "c:\ApprentiClavier\Leçons\Personnalisé\" & tempo
    End If
    tempo = Dir
Loop

INST2:
FileCopy vpath & "ApprentiClavier.exe", "c:\ApprentiClavier\ApprentiClavier.exe"

' Lance infojaws et crée le msgfinal
Module_SetUpGlobal.SonoCheck
End Sub



'***************  INFOJAWS : MESSAGE sur la sonorisation  ***********************************
Public Sub infojaws()
' **** Détection de NVDA
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
info = "NVDA " & repjawsnames 'info & repNVDA '"NVDA  " & msgDétecté
End If

If repjawsnames <> "" Or repNVDA <> "" Then
    info = CRLF & "Détection de : " & info & CRLF & info_lance
Else
    info = CRLF & msgNoSono & CRLF & info_lance
End If
Module_SetUpGlobal.msgfinal
End Sub


'*********************  MESSAGE FINAL *****************************************************
Public Sub msgfinal()
' Auparavant, copie les fichiers annexes, Leçons\Personnalisé\info.txt, alire.txt, sonorisation.txt
Module_SetUpGlobal.annexes

RAC1:
' MESSAGE FINAL SANS PARMS de ligne de commande
If Command = "" Then
    msgtext0 = info & pressez
    SetUpMsgform.Caption = msgInfo
    SetUpMsgform.Show 1
    If msgf = 2 Then GoTo RAC1

' MESSAGE FINAL AVEC PARMS de ligne de commande
Else
    MsgBox info & msgAurevoir, 0, ""
End If

' Ménage du ApprentiClavier_Setup
inst = -2
Module_SetUpGlobal.SonoLocate
End Sub



'********  ANNEXES : Copier le fichier qui bippe les fautes, et fichiers info  ************
Public Sub annexes() 'modifié avril 2008
' Tester la présence de Pop.wav
If Dir(vpath & "Pop.wav") = "" Then
    keyinhibit = 2
    MsgBox msgNoFic & "Pop.wav", 64, ""
Else
    ' Copier vers c:\ApprentiClavier\SonBip.exe
    FileCopy vpath & "Pop.wav", "c:\ApprentiClavier\Pop.wav"
End If

' Tester la présence de PianoUp4-F-A.wav.exe
If Dir(vpath & "PianoUp4-F-A.wav") = "" Then
    keyinhibit = 2
    MsgBox msgNoFic & "PianoUp4-F-A.wav", 64, ""
Else
    ' Copier vers c:\ApprentiClavier\SonBip2.exe
    FileCopy vpath & "PianoUp4-F-A.wav", "c:\ApprentiClavier\PianoUp4-F-A.wav"
End If

' COPIE des fichiers TXT du répertoire où se trouve l'exécutable de setup
If Dir(vpath & "*.txt") = "" Then Exit Sub
lperso = vpath & "*.txt"
tempo = Dir(lperso)
' BOUCLE de COPIE sur les fichiers *.txt (alire.txt, sonorisation.txt, etc.)
Do While tempo <> ""
    ' Recherche
    If tempo <> "." And tempo <> ".." Then
        FileCopy vpath & tempo, "c:\ApprentiClavier\" & tempo
    End If
    tempo = Dir
Loop

' COPIE des fichiers WAV du répertoire où se trouve l'exécutable de setup
If Dir(vpath & "*.wav") = "" Then Exit Sub
lperso = vpath & "*.wav"
tempo = Dir(lperso)
' BOUCLE de COPIE sur les fichiers *.txt (alire.txt, sonorisation.txt, etc.)
Do While tempo <> ""
    ' Recherche
    If tempo <> "." And tempo <> ".." Then
        FileCopy vpath & tempo, "c:\ApprentiClavier\" & tempo
    End If
    tempo = Dir
Loop

End Sub



'**************  DÉSINSTALL  les fichiers  ***********************************************
Public Sub desinstall()
inst = 0
Module_SetUpGlobal.delrep
Module_SetUpGlobal.SonoLocate
Unload SetUpBienvenue
If Command <> "/dq" And Command <> "/DQ" And Command <> "/Dq" And Command <> "/dQ" Then MsgBox msgSupprimé, 0, ""
inst = -2
Module_SetUpGlobal.SonoLocate
End
End Sub



'**************  EFFACE le REP C:\APPRENTICLAVIER  **************************************
Public Sub delrep()
repsys = Dir("c:\ApprentiClavier\", vbDirectory)

' BOUCLE0 sur le contenu du rép ApprentiClavier
Do While repsys <> ""

    ' Recherche
    If repsys <> "." And repsys <> ".." Then
        tempo1 = "c:\ApprentiClavier\" & repsys
        If (GetAttr(tempo1) And vbDirectory) = vbDirectory Then
            
            ' Cas d'un répertoire "c:\ApprentiClavier\toto"
            repsys1 = Dir(tempo1 & "\", vbDirectory)
            
            ' BOUCLE1 sur le contenu du "c:\ApprentiClavier\toto"
            Do While repsys1 <> ""

                ' Recherche
                If repsys1 <> "." And repsys1 <> ".." Then
                    tempo2 = tempo1 & "\" & repsys1
                    If (GetAttr(tempo2) And vbDirectory) = vbDirectory Then
            
                        ' Cas d'un sous-répertoire "c:\ApprentiClavier\toto\titi"
                        repsys2 = Dir(tempo2 & "\", vbDirectory)
            
                        ' BOUCLE2 sur le contenu du sous-rép sous ApprentiClavier
                        Do While repsys2 <> ""

                            ' Recherche
                            If repsys2 <> "." And repsys2 <> ".." Then
                                tempo3 = tempo2 & "\" & repsys2
                                If (GetAttr(tempo3) And vbDirectory) = vbDirectory Then
            
                                    ' Cas d'un sous-sous-répertoire "c:\ApprentiClavier\toto\titi\tutu"
                                    repsys3 = Dir(tempo3 & "\", vbDirectory)
            
                                    ' BOUCLE3 sur le contenu du sous-sous-rép "c:\ApprentiClavier\toto\titi\tutu"
                                    Do While repsys3 <> ""

                                        ' Recherche
                                        If repsys3 <> "." And repsys3 <> ".." Then
                                            tempo4 = tempo3 & "\" & repsys3
                                            If (GetAttr(tempo4) And vbDirectory) = vbDirectory Then
            
                                                ' Cas d'un sous-sous-sous-répertoire "c:\ApprentiClavier\toto\titi\tutu\tyty"
                                                repsys4 = Dir(tempo4 & "\", vbDirectory)
                                                If repsys4 <> "" Then
                                                    MsgBox msgASupprimer, 48, ""
                                                    Exit Sub
                                                End If
            
                                            Else
                                            ' Cas d'un fichier du sous-sous-rép "c:\ApprentiClavier\toto\titi\tutu"
                                            Kill tempo4
                                            End If
                                        End If
                                        repsys3 = Dir    ' Get next entry.
                                    Loop     ' FIN DE BOUCLE

                                    On Error Resume Next
                                    RmDir tempo3
                                    repsys2 = Dir(tempo2 & "\", vbDirectory)
            
                                Else
                                ' Cas d'un fichier du sous-rép "c:\ApprentiClavier\toto\titi"
                                Kill tempo3
                                End If
                            End If
                            repsys2 = Dir    ' Get next entry.
                        Loop     ' FIN DE BOUCLE

                        On Error Resume Next
                        RmDir tempo2
                        repsys1 = Dir(tempo1 & "\", vbDirectory)
                
                    Else
                    ' Cas d'un fichier du rép "c:\ApprentiClavier\toto"
                    Kill tempo2
                    End If
                End If
                repsys1 = Dir    ' Get next entry.
            Loop     ' FIN DE BOUCLE

            On Error Resume Next
            RmDir tempo1
            repsys = Dir("c:\ApprentiClavier\", vbDirectory)
    
            Else
            ' Cas d'un fichier sous "c:\ApprentiClavier"
            Kill tempo1
        End If
    End If
    repsys = Dir    ' Get next entry.
Loop     ' FIN DE BOUCLE

On Error Resume Next
RmDir "c:\ApprentiClavier"
End Sub



'***********  EFFACE les FICHIERs de CONFIG JAWS APPRENTICLAVIER.JCF, jsb...  ************
Public Sub deljcf()
' BOUCLE sur les réps JAWS trouvés localisée dans sonolocate qui définit repjawsfra !

' Effacer les fichiers jcf, jsb,... France
If repjawsfra <> "" Then
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier.jbs"
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier.jcf"
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier.jdf"
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier.jgf"
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier.jkm"
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier.jka"
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier.jsb"
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier.jsd"
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier.jsh"
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier.jsm"
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier.jss"
End If

' Ajout effacement configuration Italie, mars 2008
If Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" <> "" Then
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier.jbs"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier.jcf"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier.jdf"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier.jgf"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier.jkm"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier.jka"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier.jsb"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier.jsd"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier.jsh"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier.jsm"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier.jss"
End If

' Ajout effacement configuration Canada, mars 2008
If Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" <> "" Then
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier.jbs"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier.jcf"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier.jdf"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier.jgf"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier.jkm"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier.jka"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier.jsb"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier.jsd"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier.jsh"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier.jsm"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier.jss"
End If

' Ajout effacement configuration USA, mars 2008
If Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" <> "" Then
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier.jbs"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier.jcf"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier.jdf"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier.jgf"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier.jkm"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier.jka"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier.jsb"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier.jsd"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier.jsh"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier.jsm"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier.jss"
End If

End Sub


' *******************  MODIFY the DEFAULT.JKM  **********************************
Public Sub Modify_Jkm(keytype As String)
' keytype : AltCtrl
' BOUCLE sur les réps JAWS trouvés localisée dans sonolocate qui définit repjawsfra !
                
' Tester
If Dir(repjawsfra & "default.jkm") = "" Then Exit Sub
If Dir(repjawsfra & "apprenticlavier_default.jkm") <> "" Then
    On Error Resume Next  ' Nécessaire si commandes à cadence rapide
    Kill repjawsfra & "apprenticlavier_default.jkm"
End If

' Ouvrir le default d'input et le apprenticlavier_default.jkm d'output
Open repjawsfra & "apprenticlavier_default.jkm" For Output As #8
Open repjawsfra & "default.jkm" For Input As #7
Do While Not EOF(7)
    Line Input #7, currentline
    
    ' Pour réparer le Bug Jaws401 sur RightAlt !
    If keytype = "AltCtrl" Then
        If UCase(Left(currentline, 9)) = "RIGHTALT=13|3|2|" Then currentline = "RightAlt=13|3|1|" & Right(currentline, Len(currentline) - 16)
    End If
    Print #8, currentline
Loop
Close #7
Close #8

' Valider le nouveau default.jkm (pas moins de 32 octets, en réalité beaucoup plus)
If Dir(repjawsfra & "apprenticlavier_default.jkm") = "" Then Exit Sub
If FileLen(repjawsfra & "apprenticlavier_default.jkm") < 32 Then Exit Sub
On Error Resume Next
FileCopy repjawsfra & "apprenticlavier_default.jkm", repjawsfra & "default.jkm"
End Sub


' *******************  RESTORE the DEFAULT.JKM  ****************************************
Public Sub Restore_Jkm(keytype As String)
' keytype: ZZ
' BOUCLE sur les réps JAWS trouvés localisée dans sonolocate qui définit repjawsfra !

' Tester
If Dir(repjawsfra & "default.jkm") = "" Then Exit Sub
If Dir(repjawsfra & "apprenticlavier_default.jkm") <> "" Then Kill repjawsfra & "apprenticlavier_default.jkm"

' Ouvrir le default d'input et le apprenticlavier_default.jkm d'output
Open repjawsfra & "apprenticlavier_default.jkm" For Output As #6
Open repjawsfra & "default.jkm" For Input As #5
Do While Not EOF(5)
    Line Input #5, currentline
    
    ' Lignes à restaurer = toutes celles commençant par ";ZZ"
    ' Les premières versions de ApprentiClavier.exe (jusqu'àversion 1.04) pouvait modifier le default.jkm
    ' 2 blancs pourraient apparaître à droite du ";", dû à un automatisme d'édition
    If keytype = "ZZ" Then
        If UCase(Left(currentline, 3)) = ";ZZ" Then currentline = Right(currentline, Len(currentline) - 3)
        If UCase(Left(currentline, 4)) = "; ZZ" Then currentline = Right(currentline, Len(currentline) - 4)
        If UCase(Left(currentline, 5)) = ";  ZZ" Then currentline = Right(currentline, Len(currentline) - 5)
        If UCase(Left(currentline, 6)) = ";   ZZ" Then currentline = Right(currentline, Len(currentline) - 6)
        If UCase(Left(currentline, 7)) = ";    ZZ" Then currentline = Right(currentline, Len(currentline) - 7)
    End If
    Print #6, currentline
Loop
Close #5
Close #6

' Valider le nouveau default.jkm (pas moins de 32 octets, en réalité beaucoup plus)
If Dir(repjawsfra & "apprenticlavier_default.jkm") = "" Then Exit Sub
If FileLen(repjawsfra & "apprenticlavier_default.jkm") < 32 Then Exit Sub
FileCopy repjawsfra & "apprenticlavier_default.jkm", repjawsfra & "default.jkm"
 
' Ménage
On Error Resume Next
Kill repjawsfra & "apprenticlavier_default.jkm"
End Sub



' *******************  RESTORE the SYMBOLS.INI  **********************************
Public Sub Restore_Symbols(keytype As String)
' keytype : All
' BOUCLE sur les réps JAWS trouvés localisée dans sonolocate qui définit repjawsfra !

' Tester
If Dir(repjawsfra & "symbols.ini") = "" Then Exit Sub
If Dir(repjawsfra & "symbols_default.ini") <> "" Then Kill repjawsfra & "symbols_default.ini"

' Ouvrir le default d'input et le default d'output
Open repjawsfra & "symbols_default.ini" For Output As #15
Open repjawsfra & "symbols.ini" For Input As #14
Do While Not EOF(14)
    Line Input #14, currentline
    
    ' 1 ligne à restaurer
    ' Pour leçons 1 et leçons 13 : touches Flèches
    If keytype = "All" Then
        If UCase(currentline) = "BLANK= " Then currentline = "Blank=Vide"  ' Avec Alt 255 final
        If UCase(Right(currentline, 6)) = "BLANK=" Then currentline = "Blank=Vide"  ' Sans Alt255 final
    End If
            
    Print #15, currentline
Loop
Close #14
Close #15

' Valider le nouveau symbols.ini (pas moins de 32 octets, en réalité beaucoup plus)
If Dir(repjawsfra & "symbols_default.ini") = "" Then Exit Sub
If FileLen(repjawsfra & "symbols_default.ini") < 32 Then Exit Sub
FileCopy repjawsfra & "symbols_default.ini", repjawsfra & "symbols.ini"

' Ménage
On Error Resume Next
Kill repjawsfra & "symbols_default.ini"
End Sub


' **************** CANCELWIN : annule effet touche Windows  *******************************
Public Sub cancelwin(nobip, menucase)
' Annule l'effet de la touche Windows par Control-Alt-Suppr, délai, Echap
If menucase = 0 Then msgpb.Show 1 ' indispensable
If menucase = 1 Then msgpbmenu.Show 1 ' indispensable
If nobip = 0 Then keyinhibit = 0   ' pour ne pas bipper sur Control-Alt-Suppr
If nobip = 1 Then keyinhibit = 4   ' pour ne pas bipper sur Control-Alt-Suppr
SendKeys "^%{DEL}", True
Call Sleep(400)  ' indispensable
echapbis = -1
SendKeys "{ESC}", True
echapbis = 0
If menucase = 0 Then msgpb.Show 1 ' indispensable
If menucase = 1 Then msgpbmenu.Show 1 ' indispensable
If nobip = 0 Then Beep
End Sub



' ******************** ZOOMFORM : ZFACTOR quantifié selon résolution écran  ***************
Sub zoomform(forme_courante) 'modifiée mars 2008
scrw = Screen.Width
scrh = Screen.Height
frmw = forme_courante.Width
frmh = forme_courante.Height
' zoom par rapport à la largeur
zfactor = 0.9 * scrw / frmw
End Sub



' ***************  DIMOBJECT Restore les dimensions des objets selon écran  *************
Sub dimobject(object)
On Error Resume Next
object.Height = zfactor * object.Height
On Error Resume Next
object.Width = zfactor * object.Width
On Error Resume Next
object.Top = zfactor * object.Top
On Error Resume Next
object.Left = zfactor * object.Left
On Error Resume Next
object.Font.Size = zfactor * object.Font.Size
End Sub



' ***************  DIMENSION Restore toutes les dimensions selon écran  *****************
Sub Dimension(forme_courante)
' Feuille
forme_courante.Height = zfactor * forme_courante.Height
forme_courante.Width = zfactor * forme_courante.Width
forme_courante.Top = Screen.Height / 2 - forme_courante.Height / 2
forme_courante.Left = Screen.Width / 2 - forme_courante.Width / 2

' Objets dans la feuille
On Error Resume Next
Module_SetUpGlobal.dimobject forme_courante.Label0
On Error Resume Next
Module_SetUpGlobal.dimobject forme_courante.Text0
On Error Resume Next
Module_SetUpGlobal.dimobject forme_courante.Label1
On Error Resume Next
Module_SetUpGlobal.dimobject forme_courante.text1
On Error Resume Next
Module_SetUpGlobal.dimobject forme_courante.List1
On Error Resume Next
Module_SetUpGlobal.dimobject forme_courante.Installer
On Error Resume Next
Module_SetUpGlobal.dimobject forme_courante.Désinstaller
On Error Resume Next
Module_SetUpGlobal.dimobject forme_courante.aide
On Error Resume Next
Module_SetUpGlobal.dimobject forme_courante.Annuler
On Error Resume Next
Module_SetUpGlobal.dimobject forme_courante.Continuer
On Error Resume Next
Module_SetUpGlobal.dimobject forme_courante.Quitter
End Sub



' ***************  PATIENCE  Ecran d'entrée éventuel  **************************************
Sub Patience()
msgtext0 = msgPatientez
fsize = 28
SetUpMsgform.Continuer.Visible = False
timeout = 1  ' Pour quitter l'écran de sortie automatiquement
SetUpMsgform.Caption = ""
SetUpMsgform.Show 1
'Ici c'est SetUpMsgform et son timeout qui prend la main
End Sub



' ***************  Au REVOIR  Ecran de sortie pour quitter  *******************************
Sub AuRevoir()
Unload SetUpMsgform
msgtext0 = msgAurevoir
fsize = 28
SetUpMsgform.Continuer.Visible = False
timeout = 1  ' Pour quitter l'écran de sortie automatiquement
Unload SetUpBienvenue
SetUpMsgform.Caption = ""
SetUpMsgform.Show 1
'Ici c'est SetUpMsgform et son timeout qui prend la main
End Sub



' *******************  Efface le fichier JCF du SETUP  ************************************
Public Sub KillSetupJcf()
' BOUCLE sur les réps JAWS trouvés localisée dans sonolocate qui définit repjawsfra !
If repjawsfra <> "" Then
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier_Setup.jcf"
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier_Setup.jsb"
    On Error Resume Next
    Kill repjawsfra & "ApprentiClavier_Setup.jss"
End If

' Ajout effacement configuration Italie, mars 2008
If Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" <> "" Then
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier_Setup.jcf"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier_Setup.jsb"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "ita\" & "ApprentiClavier_Setup.jss"
End If

' Ajout effacement configuration Canada, mars 2008
If Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" <> "" Then
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier_Setup.jcf"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier_Setup.jsb"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "frc\" & "ApprentiClavier_Setup.jss"
End If

' Ajout effacement configuration USA, mars 2008
If Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" <> "" Then
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier_Setup.jcf"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier_Setup.jsb"
    On Error Resume Next
    Kill Left(repjawsfra, Len(repjawsfra) - 4) & "enu\" & "ApprentiClavier_Setup.jss"
End If

End Sub



' ***************  SCROLLRESULTS  Défilement page par page par SetUpMsgform  **************
Sub scrollresults(start, qty, page)
    msgtext0 = ""
    Close #3
    Open vpath & "alire.txt" For Input As #3
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
    msgtext0 = msgtext0 & pressez
SCR1:
    SetUpMsgform.Caption = msgPage & page & "."
    SetUpMsgform.Show 1
    If msgf = 2 Then GoTo SCR1
    If msgf = 0 Then
        Close #3
        Unload SetUpMsgform
    End If
    Close #3
End Sub

