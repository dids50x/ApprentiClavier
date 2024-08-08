Name "Apprenticlavier"

; Nom du fichier compilé qui sera exécuté
OutFile "ApC110_AZERTY-FR_FRA_inst.exe"

; Evite la fenêtre d'ouverture de nsis (Silent Uninstall nécessaire si Silent Install utilisé !)
SilentInstall silent
SilentUnInstall silent


; ****************** DECHARGEMENT ************************************************
Section "" 

; Décharger tous les fichiers de base dans un répertoire temporaire
setoutpath "c:\temp\apcz"
file ApprentiClavier.exe
file ApprentiClavier.jcf
file ApprentiClavier.jdf
file ApprentiClavier.jsb
file ApprentiClavier-Jaws401.jsb
file ApprentiClavier.jss

; Décharger aussi les fichiers complémentaires
file *.txt
file *.wav

; Décharger le fichier d'installation/désinstallation et la dll pour Visual Basic 4, dans le répertoire temporaire et dans le répertoire windows
file ApprentiClavier_Setup.exe
file VB40032.dll
setoutpath "$WINDIR"
file ApprentiClavier_Setup.exe
file VB40032.dll

; Copier les leçons personnalisées jointes à l'exécutable d'installation
call CopyTxt

SectionEnd


; ************ SI SUCCES du DECHARGEMENT ***********************************
Function .onInstSuccess

; Lancer ApprentiClavier_Setup.exe
ExecWait "c:\temp\apcz\ApprentiClavier_Setup.exe"

; Selon le choix utilisateur ou le succès de l'installation !
IfFileExists "c:\ApprentiClavier\ApprentiClavier.exe" Y1 Y2

Y1:
; Prévoir les clés de désinstallation pour Windows (NE PAS PRECISER LA VERSION)
WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ApprentiClavier" "DisplayName" "ApprentiClavier"
WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ApprentiClavier" "UninstallString" "$WINDIR\ApprentiClavier_uninstall.exe"
WriteRegDword HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ApprentiClavier" "NoModify" "1"

; Ecrire le fichier de désinstallation
WriteUninstaller "$WINDIR\ApprentiClavier_uninstall.exe"

; Inscrire dans le menu Démarrer
CreateDirectory "$SMPROGRAMS\ApprentiClavier"
CreateShortCut "$SMPROGRAMS\ApprentiClavier\ApprentiClavier.lnk" "c:\ApprentiClavier\ApprentiClavier.exe" "" "" 0
CreateShortCut "$SMPROGRAMS\ApprentiClavier\Désinstaller ApprentiClavier.lnk" "$WINDIR\ApprentiClavier_uninstall.exe" "" "" 0
CreateShortCut "$SMPROGRAMS\ApprentiClavier\Informations générales.lnk" "c:\ApprentiClavier\alire.txt" "" "" 0
CreateShortCut "$SMPROGRAMS\ApprentiClavier\Informations sur la sonorisation.lnk" "c:\ApprentiClavier\sonorisation.txt" "" "" 0
CreateShortCut "$SMPROGRAMS\ApprentiClavier\Informations pour les enseignants.lnk" "c:\ApprentiClavier\leçons\Personnalisé\info.txt" "" "" 0

; Inscrire une icône dans le bureau (Mettre Espace entre Apprenti et Clavier)
Delete "$DESKTOP\ApprentiClavier.lnk"
CreateShortCut "$DESKTOP\Apprenti Clavier.lnk" "c:\ApprentiClavier\ApprentiClavier.exe" "" "$WINDIR\notepad.exe" 0
goto Y3

Y2:
; Installation ratée ou désinstallation demandée, relancer le même exé avec option /D Quiet
ExecWait '"c:\temp\apcz\ApprentiClavier_setup.exe" /DQ'

; Supprimer les clés du registre
DeleteRegKey HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ApprentiClavier"

; Supprimer les shortcuts
Delete "$SMPROGRAMS\ApprentiClavier\*.*"
rmdir "$SMPROGRAMS\ApprentiClavier"

; Supprimer l'icône du bureau
Delete "$DESKTOP\ApprentiClavier.lnk"
Delete "$DESKTOP\Apprenti Clavier.lnk"

; Supprimer les fichiers résiduels
rmdir /r "c:\ApprentiClavier"

; Supprimer les exécutables d'installation/désinstallation
Delete "$WINDIR\ApprentiClavier_setup.exe"
Delete "$WINDIR\ApprentiClavier_uninstall.exe"
Delete "$WINDIR\VB40032.dll"

Y3:
; Message final
;messagebox MB_OK|MB_TOPMOST "FIN DE L'INSTALLATION. AU REVOIR"

; Supprimer le déchargement temporaire
rmdir /r "c:\temp\apcz"
FunctionEnd


; *************************************************************************
Function CopyTxt

; Décharger/Copier les leçons personnalisées jointes à l'exécutable d'installation
createdirectory "c:\temp\apcz\perso"
IfFileExists "$EXEDIR\leçon1A.txt" A1 VB1
A1:
copyfiles "$EXEDIR\leçon1A.txt" "c:\temp\apcz\perso\leçon1A.txt"
VB1:
IfFileExists "$EXEDIR\leçon1B.txt" B1 VC1
B1:
copyfiles "$EXEDIR\leçon1B.txt" "c:\temp\apcz\perso\leçon1B.txt"
VC1:
IfFileExists "$EXEDIR\leçon1C.txt" C1 VA2
C1:
copyfiles "$EXEDIR\leçon1C.txt" "c:\temp\apcz\perso\leçon1C.txt"
VA2:
IfFileExists "$EXEDIR\leçon2A.txt" A2 VB2
A2:
copyfiles "$EXEDIR\leçon2A.txt" "c:\temp\apcz\perso\leçon2A.txt"
VB2:
IfFileExists "$EXEDIR\leçon2B.txt" B2 VC2
B2:
copyfiles "$EXEDIR\leçon2B.txt" "c:\temp\apcz\perso\leçon2B.txt"
VC2:
IfFileExists "$EXEDIR\leçon2C.txt" C2 VD2
C2:
copyfiles "$EXEDIR\leçon2C.txt" "c:\temp\apcz\perso\leçon2C.txt"
VD2:
IfFileExists "$EXEDIR\leçon2D.txt" D2 VE2
D2:
copyfiles "$EXEDIR\leçon2D.txt" "c:\temp\apcz\perso\leçon2D.txt"
VE2:
IfFileExists "$EXEDIR\leçon2E.txt" E2 VF2
E2:
copyfiles "$EXEDIR\leçon2E.txt" "c:\temp\apcz\perso\leçon2E.txt"
VF2:
IfFileExists "$EXEDIR\leçon2F.txt" F2 VG2
F2:
copyfiles "$EXEDIR\leçon2F.txt" "c:\temp\apcz\perso\leçon2F.txt"
VG2:
IfFileExists "$EXEDIR\leçon2G.txt" G2 VH2
G2:
copyfiles "$EXEDIR\leçon2G.txt" "c:\temp\apcz\perso\leçon2G.txt"
VH2:
IfFileExists "$EXEDIR\leçon2H.txt" H2 VA3
H2:
copyfiles "$EXEDIR\leçon2H.txt" "c:\temp\apcz\perso\leçon2H.txt"
VA3:
IfFileExists "$EXEDIR\leçon3A.txt" A3 VB3
A3:
copyfiles "$EXEDIR\leçon3A.txt" "c:\temp\apcz\perso\leçon3A.txt"
VB3:
IfFileExists "$EXEDIR\leçon3B.txt" B3 VC3
B3:
copyfiles "$EXEDIR\leçon3B.txt" "c:\temp\apcz\perso\leçon3B.txt"
VC3:
IfFileExists "$EXEDIR\leçon3C.txt" C3 VA4
C3:
copyfiles "$EXEDIR\leçon3C.txt" "c:\temp\apcz\perso\leçon3C.txt"
VA4:
IfFileExists "$EXEDIR\leçon4A.txt" A4 VB4
A4:
copyfiles "$EXEDIR\leçon4A.txt" "c:\temp\apcz\perso\leçon4A.txt"
VB4:
IfFileExists "$EXEDIR\leçon4B.txt" B4 VC4
B4:
copyfiles "$EXEDIR\leçon4B.txt" "c:\temp\apcz\perso\leçon4B.txt"
VC4:
IfFileExists "$EXEDIR\leçon4C.txt" C4 VD4
C4:
copyfiles "$EXEDIR\leçon4C.txt" "c:\temp\apcz\perso\leçon4C.txt"
VD4:
IfFileExists "$EXEDIR\leçon4D.txt" D4 VE4
D4:
copyfiles "$EXEDIR\leçon4D.txt" "c:\temp\apcz\perso\leçon4D.txt"
VE4:
IfFileExists "$EXEDIR\leçon4E.txt" E4 VF4
E4:
copyfiles "$EXEDIR\leçon4E.txt" "c:\temp\apcz\perso\leçon4E.txt"
VF4:
IfFileExists "$EXEDIR\leçon4F.txt" F4 VG4
F4:
copyfiles "$EXEDIR\leçon4F.txt" "c:\temp\apcz\perso\leçon4F.txt"
VG4:
IfFileExists "$EXEDIR\leçon4G.txt" G4 VH4
G4:
copyfiles "$EXEDIR\leçon4G.txt" "c:\temp\apcz\perso\leçon4G.txt"
VH4:
IfFileExists "$EXEDIR\leçon4H.txt" H4 VA5
H4:
copyfiles "$EXEDIR\leçon4H.txt" "c:\temp\apcz\perso\leçon4H.txt"
VA5:
IfFileExists "$EXEDIR\leçon5A.txt" A5 VB5
A5:
copyfiles "$EXEDIR\leçon5A.txt" "c:\temp\apcz\perso\leçon5A.txt"
VB5:
IfFileExists "$EXEDIR\leçon5B.txt" B5 VC5
B5:
copyfiles "$EXEDIR\leçon5B.txt" "c:\temp\apcz\perso\leçon5B.txt"
VC5:
IfFileExists "$EXEDIR\leçon5C.txt" C5 VA6
C5:
copyfiles "$EXEDIR\leçon5C.txt" "c:\temp\apcz\perso\leçon5C.txt"
VA6:
IfFileExists "$EXEDIR\leçon6A.txt" A6 VB6
A6:
copyfiles "$EXEDIR\leçon6A.txt" "c:\temp\apcz\perso\leçon6A.txt"
VB6:
IfFileExists "$EXEDIR\leçon6B.txt" B6 VC6
B6:
copyfiles "$EXEDIR\leçon6B.txt" "c:\temp\apcz\perso\leçon6B.txt"
VC6:
IfFileExists "$EXEDIR\leçon6C.txt" C6 VA7
C6:
copyfiles "$EXEDIR\leçon6C.txt" "c:\temp\apcz\perso\leçon6C.txt"
VA7:
IfFileExists "$EXEDIR\leçon7A.txt" A7 VB7
A7:
copyfiles "$EXEDIR\leçon7A.txt" "c:\temp\apcz\perso\leçon7A.txt"
VB7:
IfFileExists "$EXEDIR\leçon7B.txt" B7 VC7
B7:
copyfiles "$EXEDIR\leçon7B.txt" "c:\temp\apcz\perso\leçon7B.txt"
VC7:
IfFileExists "$EXEDIR\leçon7C.txt" C7 VA8
C7:
copyfiles "$EXEDIR\leçon7C.txt" "c:\temp\apcz\perso\leçon7C.txt"
VA8:
IfFileExists "$EXEDIR\leçon8A.txt" A8 VB8
A8:
copyfiles "$EXEDIR\leçon8A.txt" "c:\temp\apcz\perso\leçon8A.txt"
VB8:
IfFileExists "$EXEDIR\leçon8B.txt" B8 VC8
B8:
copyfiles "$EXEDIR\leçon8B.txt" "c:\temp\apcz\perso\leçon8B.txt"
VC8:
IfFileExists "$EXEDIR\leçon8C.txt" C8 VD8
C8:
copyfiles "$EXEDIR\leçon8C.txt" "c:\temp\apcz\perso\leçon8C.txt"
VD8:
IfFileExists "$EXEDIR\leçon8D.txt" D8 VE8
D8:
copyfiles "$EXEDIR\leçon8D.txt" "c:\temp\apcz\perso\leçon8D.txt"
VE8:
IfFileExists "$EXEDIR\leçon8E.txt" E8 VF8
E8:
copyfiles "$EXEDIR\leçon8E.txt" "c:\temp\apcz\perso\leçon8E.txt"
VF8:
IfFileExists "$EXEDIR\leçon8F.txt" F8 VG8
F8:
copyfiles "$EXEDIR\leçon8F.txt" "c:\temp\apcz\perso\leçon8F.txt"
VG8:
IfFileExists "$EXEDIR\leçon8G.txt" G8 VH8
G8:
copyfiles "$EXEDIR\leçon8G.txt" "c:\temp\apcz\perso\leçon8G.txt"
VH8:
IfFileExists "$EXEDIR\leçon8H.txt" H8 VA9
H8:
copyfiles "$EXEDIR\leçon8H.txt" "c:\temp\apcz\perso\leçon8H.txt"
VA9:
IfFileExists "$EXEDIR\leçon9A.txt" A9 VB9
A9:
copyfiles "$EXEDIR\leçon9A.txt" "c:\temp\apcz\perso\leçon9A.txt"
VB9:
IfFileExists "$EXEDIR\leçon9B.txt" B9 VC9
B9:
copyfiles "$EXEDIR\leçon9B.txt" "c:\temp\apcz\perso\leçon9B.txt"
VC9:
IfFileExists "$EXEDIR\leçon9C.txt" C9 VD9
C9:
copyfiles "$EXEDIR\leçon9C.txt" "c:\temp\apcz\perso\leçon9C.txt"
VD9:
IfFileExists "$EXEDIR\leçon9D.txt" D9 VA10
D9:
copyfiles "$EXEDIR\leçon9D.txt" "c:\temp\apcz\perso\leçon9D.txt"
VA10:
IfFileExists "$EXEDIR\leçon10A.txt" A10 VB10
A10:
copyfiles "$EXEDIR\leçon10A.txt" "c:\temp\apcz\perso\leçon10A.txt"
VB10:
IfFileExists "$EXEDIR\leçon10B.txt" B10 VC10
B10:
copyfiles "$EXEDIR\leçon10B.txt" "c:\temp\apcz\perso\leçon10B.txt"
VC10:
IfFileExists "$EXEDIR\leçon10C.txt" C10 VA11
C10:
copyfiles "$EXEDIR\leçon10C.txt" "c:\temp\apcz\perso\leçon10C.txt"
VA11:
IfFileExists "$EXEDIR\leçon11A.txt" A11 VB11
A11:
copyfiles "$EXEDIR\leçon11A.txt" "c:\temp\apcz\perso\leçon11A.txt"
VB11:
IfFileExists "$EXEDIR\leçon11B.txt" B11 VC11
B11:
copyfiles "$EXEDIR\leçon11B.txt" "c:\temp\apcz\perso\leçon11B.txt"
VC11:
IfFileExists "$EXEDIR\leçon11C.txt" C11 VA12
C11:
copyfiles "$EXEDIR\leçon11C.txt" "c:\temp\apcz\perso\leçon11C.txt"
VA12:
IfFileExists "$EXEDIR\leçon12A.txt" A12 VB12
A12:
copyfiles "$EXEDIR\leçon12A.txt" "c:\temp\apcz\perso\leçon12A.txt"
VB12:
IfFileExists "$EXEDIR\leçon12B.txt" B12 VC12
B12:
copyfiles "$EXEDIR\leçon12B.txt" "c:\temp\apcz\perso\leçon12B.txt"
VC12:
IfFileExists "$EXEDIR\leçon12C.txt" C12 VD12
C12:
copyfiles "$EXEDIR\leçon12C.txt" "c:\temp\apcz\perso\leçon12C.txt"
VD12:
IfFileExists "$EXEDIR\leçon12D.txt" D12 VA13
D12:
copyfiles "$EXEDIR\leçon12D.txt" "c:\temp\apcz\perso\leçon12D.txt"
VA13:
IfFileExists "$EXEDIR\leçon13A.txt" A13 VB13
A13:
copyfiles "$EXEDIR\leçon13A.txt" "c:\temp\apcz\perso\leçon13A.txt"
VB13:
IfFileExists "$EXEDIR\leçon13B.txt" B13 VC13
B13:
copyfiles "$EXEDIR\leçon13B.txt" "c:\temp\apcz\perso\leçon13B.txt"
VC13:
IfFileExists "$EXEDIR\leçon13C.txt" C13 VD13
C13:
copyfiles "$EXEDIR\leçon13C.txt" "c:\temp\apcz\perso\leçon13C.txt"
VD13:
IfFileExists "$EXEDIR\leçon13D.txt" D13 VE13
D13:
copyfiles "$EXEDIR\leçon13D.txt" "c:\temp\apcz\perso\leçon13D.txt"
VE13:
IfFileExists "$EXEDIR\leçon13E.txt" E13 VF13
E13:
copyfiles "$EXEDIR\leçon13E.txt" "c:\temp\apcz\perso\leçon13E.txt"
VF13:
IfFileExists "$EXEDIR\leçon13F.txt" F13 VA14
F13:
copyfiles "$EXEDIR\leçon13F.txt" "c:\temp\apcz\perso\leçon13F.txt"
VA14:
IfFileExists "$EXEDIR\leçon14A.txt" A14 VB14
A14:
copyfiles "$EXEDIR\leçon14A.txt" "c:\temp\apcz\perso\leçon14A.txt"
VB14:
IfFileExists "$EXEDIR\leçon14B.txt" B14 VC14
B14:
copyfiles "$EXEDIR\leçon14B.txt" "c:\temp\apcz\perso\leçon14B.txt"
VC14:
IfFileExists "$EXEDIR\leçon14C.txt" C14 VA15
C14:
copyfiles "$EXEDIR\leçon14C.txt" "c:\temp\apcz\perso\leçon14C.txt"
VA15:
IfFileExists "$EXEDIR\leçon15A.txt" A15 VB15
A15:
copyfiles "$EXEDIR\leçon15A.txt" "c:\temp\apcz\perso\leçon15A.txt"
VB15:
IfFileExists "$EXEDIR\leçon15B.txt" B15 VC15
B15:
copyfiles "$EXEDIR\leçon15B.txt" "c:\temp\apcz\perso\leçon15B.txt"
VC15:
IfFileExists "$EXEDIR\leçon15C.txt" C15 VA16
C15:
copyfiles "$EXEDIR\leçon15C.txt" "c:\temp\apcz\perso\leçon15C.txt"
VA16:
IfFileExists "$EXEDIR\leçon16A.txt" A16 VB16
A16:
copyfiles "$EXEDIR\leçon16A.txt" "c:\temp\apcz\perso\leçon16A.txt"
VB16:
IfFileExists "$EXEDIR\leçon16B.txt" B16 VC16
B16:
copyfiles "$EXEDIR\leçon16B.txt" "c:\temp\apcz\perso\leçon16B.txt"
VC16:
IfFileExists "$EXEDIR\leçon16C.txt" C16 VD16
C16:
copyfiles "$EXEDIR\leçon16C.txt" "c:\temp\apcz\perso\leçon16C.txt"
VD16:
IfFileExists "$EXEDIR\leçon16D.txt" D16 VA17
D16:
copyfiles "$EXEDIR\leçon16D.txt" "c:\temp\apcz\perso\leçon16D.txt"
VA17:
IfFileExists "$EXEDIR\leçon17A.txt" A17 VB17
A17:
copyfiles "$EXEDIR\leçon17A.txt" "c:\temp\apcz\perso\leçon17A.txt"
VB17:
IfFileExists "$EXEDIR\leçon17B.txt" B17 VC17
B17:
copyfiles "$EXEDIR\leçon17B.txt" "c:\temp\apcz\perso\leçon17B.txt"
VC17:
IfFileExists "$EXEDIR\leçon17C.txt" C17 VD17
C17:
copyfiles "$EXEDIR\leçon17C.txt" "c:\temp\apcz\perso\leçon17C.txt"
VD17:
IfFileExists "$EXEDIR\leçon17D.txt" D17 VA18
D17:
copyfiles "$EXEDIR\leçon17D.txt" "c:\temp\apcz\perso\leçon17D.txt"
VA18:
IfFileExists "$EXEDIR\leçon18A.txt" A18 VB18
A18:
copyfiles "$EXEDIR\leçon18A.txt" "c:\temp\apcz\perso\leçon18A.txt"
VB18:
IfFileExists "$EXEDIR\leçon18B.txt" B18 VC18
B18:
copyfiles "$EXEDIR\leçon18B.txt" "c:\temp\apcz\perso\leçon18B.txt"
VC18:
IfFileExists "$EXEDIR\leçon18C.txt" C18 VD18
C18:
copyfiles "$EXEDIR\leçon18C.txt" "c:\temp\apcz\perso\leçon18C.txt"
VD18:
IfFileExists "$EXEDIR\leçon18D.txt" D18 VE18
D18:
copyfiles "$EXEDIR\leçon18D.txt" "c:\temp\apcz\perso\leçon18D.txt"
VE18:
IfFileExists "$EXEDIR\leçon18E.txt" E18 VA19
E18:
copyfiles "$EXEDIR\leçon18E.txt" "c:\temp\apcz\perso\leçon18E.txt"
VA19:
IfFileExists "$EXEDIR\leçon19A.txt" A19 VB19
A19:
copyfiles "$EXEDIR\leçon19A.txt" "c:\temp\apcz\perso\leçon19A.txt"
VB19:
IfFileExists "$EXEDIR\leçon19B.txt" B19 VC19
B19:
copyfiles "$EXEDIR\leçon19B.txt" "c:\temp\apcz\perso\leçon19B.txt"
VC19:
IfFileExists "$EXEDIR\leçon19C.txt" C19 VD19
C19:
copyfiles "$EXEDIR\leçon19C.txt" "c:\temp\apcz\perso\leçon19C.txt"
VD19:
IfFileExists "$EXEDIR\leçon19D.txt" D19 VA20
D19:
copyfiles "$EXEDIR\leçon19D.txt" "c:\temp\apcz\perso\leçon19D.txt"
VA20:

FunctionEnd


; *************************************************************************
Section "Uninstall"

; Lancer l'exécutable de désinstallation
ExecWait '"$WINDIR\ApprentiClavier_setup.exe" /D'

; Selon le choix utilisateur, ou le succès de l'installation !
IfFileExists "c:\ApprentiClavier\ApprentiClavier.exe" Z2 Z1

Z1:
; Désinstallation complète
; Supprimer les clés du registre
;messagebox MB_OK|MB_TOPMOST "Effacement des clés de la base de registres"
DeleteRegKey HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ApprentiClavier"

; Supprimer les shortcuts
Delete "$SMPROGRAMS\ApprentiClavier\*.*"
rmdir "$SMPROGRAMS\ApprentiClavier"

; Supprimer l'icône du bureau
Delete "$DESKTOP\ApprentiClavier.lnk"
Delete "$DESKTOP\Apprenti Clavier.lnk"

; Supprimer les fichiers résiduels
rmdir /r "c:\ApprentiClavier"

; Supprimer les exécutables d'installation/désinstallation
Delete "$WINDIR\ApprentiClavier_setup.exe"
Delete "$WINDIR\ApprentiClavier_uninstall.exe"
Delete "$WINDIR\VB40032.dll"

Z2:
; Supprimer le déchargement temporaire
; rmdir /r "c:\temp\apcz"
SectionEnd