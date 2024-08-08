Name "Apprenticlavier"

; Nom du fichier compil� qui sera ex�cut�
OutFile "ApC110_AZERTY-FR_FRA_inst.exe"

; Evite la fen�tre d'ouverture de nsis (Silent Uninstall n�cessaire si Silent Install utilis� !)
SilentInstall silent
SilentUnInstall silent


; ****************** DECHARGEMENT ************************************************
Section "" 

; D�charger tous les fichiers de base dans un r�pertoire temporaire
setoutpath "c:\temp\apcz"
file ApprentiClavier.exe
file ApprentiClavier.jcf
file ApprentiClavier.jdf
file ApprentiClavier.jsb
file ApprentiClavier-Jaws401.jsb
file ApprentiClavier.jss

; D�charger aussi les fichiers compl�mentaires
file *.txt
file *.wav

; D�charger le fichier d'installation/d�sinstallation et la dll pour Visual Basic 4, dans le r�pertoire temporaire et dans le r�pertoire windows
file ApprentiClavier_Setup.exe
file VB40032.dll
setoutpath "$WINDIR"
file ApprentiClavier_Setup.exe
file VB40032.dll

; Copier les le�ons personnalis�es jointes � l'ex�cutable d'installation
call CopyTxt

SectionEnd


; ************ SI SUCCES du DECHARGEMENT ***********************************
Function .onInstSuccess

; Lancer ApprentiClavier_Setup.exe
ExecWait "c:\temp\apcz\ApprentiClavier_Setup.exe"

; Selon le choix utilisateur ou le succ�s de l'installation !
IfFileExists "c:\ApprentiClavier\ApprentiClavier.exe" Y1 Y2

Y1:
; Pr�voir les cl�s de d�sinstallation pour Windows (NE PAS PRECISER LA VERSION)
WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ApprentiClavier" "DisplayName" "ApprentiClavier"
WriteRegStr HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ApprentiClavier" "UninstallString" "$WINDIR\ApprentiClavier_uninstall.exe"
WriteRegDword HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ApprentiClavier" "NoModify" "1"

; Ecrire le fichier de d�sinstallation
WriteUninstaller "$WINDIR\ApprentiClavier_uninstall.exe"

; Inscrire dans le menu D�marrer
CreateDirectory "$SMPROGRAMS\ApprentiClavier"
CreateShortCut "$SMPROGRAMS\ApprentiClavier\ApprentiClavier.lnk" "c:\ApprentiClavier\ApprentiClavier.exe" "" "" 0
CreateShortCut "$SMPROGRAMS\ApprentiClavier\D�sinstaller ApprentiClavier.lnk" "$WINDIR\ApprentiClavier_uninstall.exe" "" "" 0
CreateShortCut "$SMPROGRAMS\ApprentiClavier\Informations g�n�rales.lnk" "c:\ApprentiClavier\alire.txt" "" "" 0
CreateShortCut "$SMPROGRAMS\ApprentiClavier\Informations sur la sonorisation.lnk" "c:\ApprentiClavier\sonorisation.txt" "" "" 0
CreateShortCut "$SMPROGRAMS\ApprentiClavier\Informations pour les enseignants.lnk" "c:\ApprentiClavier\le�ons\Personnalis�\info.txt" "" "" 0

; Inscrire une ic�ne dans le bureau (Mettre Espace entre Apprenti et Clavier)
Delete "$DESKTOP\ApprentiClavier.lnk"
CreateShortCut "$DESKTOP\Apprenti Clavier.lnk" "c:\ApprentiClavier\ApprentiClavier.exe" "" "$WINDIR\notepad.exe" 0
goto Y3

Y2:
; Installation rat�e ou d�sinstallation demand�e, relancer le m�me ex� avec option /D Quiet
ExecWait '"c:\temp\apcz\ApprentiClavier_setup.exe" /DQ'

; Supprimer les cl�s du registre
DeleteRegKey HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ApprentiClavier"

; Supprimer les shortcuts
Delete "$SMPROGRAMS\ApprentiClavier\*.*"
rmdir "$SMPROGRAMS\ApprentiClavier"

; Supprimer l'ic�ne du bureau
Delete "$DESKTOP\ApprentiClavier.lnk"
Delete "$DESKTOP\Apprenti Clavier.lnk"

; Supprimer les fichiers r�siduels
rmdir /r "c:\ApprentiClavier"

; Supprimer les ex�cutables d'installation/d�sinstallation
Delete "$WINDIR\ApprentiClavier_setup.exe"
Delete "$WINDIR\ApprentiClavier_uninstall.exe"
Delete "$WINDIR\VB40032.dll"

Y3:
; Message final
;messagebox MB_OK|MB_TOPMOST "FIN DE L'INSTALLATION. AU REVOIR"

; Supprimer le d�chargement temporaire
rmdir /r "c:\temp\apcz"
FunctionEnd


; *************************************************************************
Function CopyTxt

; D�charger/Copier les le�ons personnalis�es jointes � l'ex�cutable d'installation
createdirectory "c:\temp\apcz\perso"
IfFileExists "$EXEDIR\le�on1A.txt" A1 VB1
A1:
copyfiles "$EXEDIR\le�on1A.txt" "c:\temp\apcz\perso\le�on1A.txt"
VB1:
IfFileExists "$EXEDIR\le�on1B.txt" B1 VC1
B1:
copyfiles "$EXEDIR\le�on1B.txt" "c:\temp\apcz\perso\le�on1B.txt"
VC1:
IfFileExists "$EXEDIR\le�on1C.txt" C1 VA2
C1:
copyfiles "$EXEDIR\le�on1C.txt" "c:\temp\apcz\perso\le�on1C.txt"
VA2:
IfFileExists "$EXEDIR\le�on2A.txt" A2 VB2
A2:
copyfiles "$EXEDIR\le�on2A.txt" "c:\temp\apcz\perso\le�on2A.txt"
VB2:
IfFileExists "$EXEDIR\le�on2B.txt" B2 VC2
B2:
copyfiles "$EXEDIR\le�on2B.txt" "c:\temp\apcz\perso\le�on2B.txt"
VC2:
IfFileExists "$EXEDIR\le�on2C.txt" C2 VD2
C2:
copyfiles "$EXEDIR\le�on2C.txt" "c:\temp\apcz\perso\le�on2C.txt"
VD2:
IfFileExists "$EXEDIR\le�on2D.txt" D2 VE2
D2:
copyfiles "$EXEDIR\le�on2D.txt" "c:\temp\apcz\perso\le�on2D.txt"
VE2:
IfFileExists "$EXEDIR\le�on2E.txt" E2 VF2
E2:
copyfiles "$EXEDIR\le�on2E.txt" "c:\temp\apcz\perso\le�on2E.txt"
VF2:
IfFileExists "$EXEDIR\le�on2F.txt" F2 VG2
F2:
copyfiles "$EXEDIR\le�on2F.txt" "c:\temp\apcz\perso\le�on2F.txt"
VG2:
IfFileExists "$EXEDIR\le�on2G.txt" G2 VH2
G2:
copyfiles "$EXEDIR\le�on2G.txt" "c:\temp\apcz\perso\le�on2G.txt"
VH2:
IfFileExists "$EXEDIR\le�on2H.txt" H2 VA3
H2:
copyfiles "$EXEDIR\le�on2H.txt" "c:\temp\apcz\perso\le�on2H.txt"
VA3:
IfFileExists "$EXEDIR\le�on3A.txt" A3 VB3
A3:
copyfiles "$EXEDIR\le�on3A.txt" "c:\temp\apcz\perso\le�on3A.txt"
VB3:
IfFileExists "$EXEDIR\le�on3B.txt" B3 VC3
B3:
copyfiles "$EXEDIR\le�on3B.txt" "c:\temp\apcz\perso\le�on3B.txt"
VC3:
IfFileExists "$EXEDIR\le�on3C.txt" C3 VA4
C3:
copyfiles "$EXEDIR\le�on3C.txt" "c:\temp\apcz\perso\le�on3C.txt"
VA4:
IfFileExists "$EXEDIR\le�on4A.txt" A4 VB4
A4:
copyfiles "$EXEDIR\le�on4A.txt" "c:\temp\apcz\perso\le�on4A.txt"
VB4:
IfFileExists "$EXEDIR\le�on4B.txt" B4 VC4
B4:
copyfiles "$EXEDIR\le�on4B.txt" "c:\temp\apcz\perso\le�on4B.txt"
VC4:
IfFileExists "$EXEDIR\le�on4C.txt" C4 VD4
C4:
copyfiles "$EXEDIR\le�on4C.txt" "c:\temp\apcz\perso\le�on4C.txt"
VD4:
IfFileExists "$EXEDIR\le�on4D.txt" D4 VE4
D4:
copyfiles "$EXEDIR\le�on4D.txt" "c:\temp\apcz\perso\le�on4D.txt"
VE4:
IfFileExists "$EXEDIR\le�on4E.txt" E4 VF4
E4:
copyfiles "$EXEDIR\le�on4E.txt" "c:\temp\apcz\perso\le�on4E.txt"
VF4:
IfFileExists "$EXEDIR\le�on4F.txt" F4 VG4
F4:
copyfiles "$EXEDIR\le�on4F.txt" "c:\temp\apcz\perso\le�on4F.txt"
VG4:
IfFileExists "$EXEDIR\le�on4G.txt" G4 VH4
G4:
copyfiles "$EXEDIR\le�on4G.txt" "c:\temp\apcz\perso\le�on4G.txt"
VH4:
IfFileExists "$EXEDIR\le�on4H.txt" H4 VA5
H4:
copyfiles "$EXEDIR\le�on4H.txt" "c:\temp\apcz\perso\le�on4H.txt"
VA5:
IfFileExists "$EXEDIR\le�on5A.txt" A5 VB5
A5:
copyfiles "$EXEDIR\le�on5A.txt" "c:\temp\apcz\perso\le�on5A.txt"
VB5:
IfFileExists "$EXEDIR\le�on5B.txt" B5 VC5
B5:
copyfiles "$EXEDIR\le�on5B.txt" "c:\temp\apcz\perso\le�on5B.txt"
VC5:
IfFileExists "$EXEDIR\le�on5C.txt" C5 VA6
C5:
copyfiles "$EXEDIR\le�on5C.txt" "c:\temp\apcz\perso\le�on5C.txt"
VA6:
IfFileExists "$EXEDIR\le�on6A.txt" A6 VB6
A6:
copyfiles "$EXEDIR\le�on6A.txt" "c:\temp\apcz\perso\le�on6A.txt"
VB6:
IfFileExists "$EXEDIR\le�on6B.txt" B6 VC6
B6:
copyfiles "$EXEDIR\le�on6B.txt" "c:\temp\apcz\perso\le�on6B.txt"
VC6:
IfFileExists "$EXEDIR\le�on6C.txt" C6 VA7
C6:
copyfiles "$EXEDIR\le�on6C.txt" "c:\temp\apcz\perso\le�on6C.txt"
VA7:
IfFileExists "$EXEDIR\le�on7A.txt" A7 VB7
A7:
copyfiles "$EXEDIR\le�on7A.txt" "c:\temp\apcz\perso\le�on7A.txt"
VB7:
IfFileExists "$EXEDIR\le�on7B.txt" B7 VC7
B7:
copyfiles "$EXEDIR\le�on7B.txt" "c:\temp\apcz\perso\le�on7B.txt"
VC7:
IfFileExists "$EXEDIR\le�on7C.txt" C7 VA8
C7:
copyfiles "$EXEDIR\le�on7C.txt" "c:\temp\apcz\perso\le�on7C.txt"
VA8:
IfFileExists "$EXEDIR\le�on8A.txt" A8 VB8
A8:
copyfiles "$EXEDIR\le�on8A.txt" "c:\temp\apcz\perso\le�on8A.txt"
VB8:
IfFileExists "$EXEDIR\le�on8B.txt" B8 VC8
B8:
copyfiles "$EXEDIR\le�on8B.txt" "c:\temp\apcz\perso\le�on8B.txt"
VC8:
IfFileExists "$EXEDIR\le�on8C.txt" C8 VD8
C8:
copyfiles "$EXEDIR\le�on8C.txt" "c:\temp\apcz\perso\le�on8C.txt"
VD8:
IfFileExists "$EXEDIR\le�on8D.txt" D8 VE8
D8:
copyfiles "$EXEDIR\le�on8D.txt" "c:\temp\apcz\perso\le�on8D.txt"
VE8:
IfFileExists "$EXEDIR\le�on8E.txt" E8 VF8
E8:
copyfiles "$EXEDIR\le�on8E.txt" "c:\temp\apcz\perso\le�on8E.txt"
VF8:
IfFileExists "$EXEDIR\le�on8F.txt" F8 VG8
F8:
copyfiles "$EXEDIR\le�on8F.txt" "c:\temp\apcz\perso\le�on8F.txt"
VG8:
IfFileExists "$EXEDIR\le�on8G.txt" G8 VH8
G8:
copyfiles "$EXEDIR\le�on8G.txt" "c:\temp\apcz\perso\le�on8G.txt"
VH8:
IfFileExists "$EXEDIR\le�on8H.txt" H8 VA9
H8:
copyfiles "$EXEDIR\le�on8H.txt" "c:\temp\apcz\perso\le�on8H.txt"
VA9:
IfFileExists "$EXEDIR\le�on9A.txt" A9 VB9
A9:
copyfiles "$EXEDIR\le�on9A.txt" "c:\temp\apcz\perso\le�on9A.txt"
VB9:
IfFileExists "$EXEDIR\le�on9B.txt" B9 VC9
B9:
copyfiles "$EXEDIR\le�on9B.txt" "c:\temp\apcz\perso\le�on9B.txt"
VC9:
IfFileExists "$EXEDIR\le�on9C.txt" C9 VD9
C9:
copyfiles "$EXEDIR\le�on9C.txt" "c:\temp\apcz\perso\le�on9C.txt"
VD9:
IfFileExists "$EXEDIR\le�on9D.txt" D9 VA10
D9:
copyfiles "$EXEDIR\le�on9D.txt" "c:\temp\apcz\perso\le�on9D.txt"
VA10:
IfFileExists "$EXEDIR\le�on10A.txt" A10 VB10
A10:
copyfiles "$EXEDIR\le�on10A.txt" "c:\temp\apcz\perso\le�on10A.txt"
VB10:
IfFileExists "$EXEDIR\le�on10B.txt" B10 VC10
B10:
copyfiles "$EXEDIR\le�on10B.txt" "c:\temp\apcz\perso\le�on10B.txt"
VC10:
IfFileExists "$EXEDIR\le�on10C.txt" C10 VA11
C10:
copyfiles "$EXEDIR\le�on10C.txt" "c:\temp\apcz\perso\le�on10C.txt"
VA11:
IfFileExists "$EXEDIR\le�on11A.txt" A11 VB11
A11:
copyfiles "$EXEDIR\le�on11A.txt" "c:\temp\apcz\perso\le�on11A.txt"
VB11:
IfFileExists "$EXEDIR\le�on11B.txt" B11 VC11
B11:
copyfiles "$EXEDIR\le�on11B.txt" "c:\temp\apcz\perso\le�on11B.txt"
VC11:
IfFileExists "$EXEDIR\le�on11C.txt" C11 VA12
C11:
copyfiles "$EXEDIR\le�on11C.txt" "c:\temp\apcz\perso\le�on11C.txt"
VA12:
IfFileExists "$EXEDIR\le�on12A.txt" A12 VB12
A12:
copyfiles "$EXEDIR\le�on12A.txt" "c:\temp\apcz\perso\le�on12A.txt"
VB12:
IfFileExists "$EXEDIR\le�on12B.txt" B12 VC12
B12:
copyfiles "$EXEDIR\le�on12B.txt" "c:\temp\apcz\perso\le�on12B.txt"
VC12:
IfFileExists "$EXEDIR\le�on12C.txt" C12 VD12
C12:
copyfiles "$EXEDIR\le�on12C.txt" "c:\temp\apcz\perso\le�on12C.txt"
VD12:
IfFileExists "$EXEDIR\le�on12D.txt" D12 VA13
D12:
copyfiles "$EXEDIR\le�on12D.txt" "c:\temp\apcz\perso\le�on12D.txt"
VA13:
IfFileExists "$EXEDIR\le�on13A.txt" A13 VB13
A13:
copyfiles "$EXEDIR\le�on13A.txt" "c:\temp\apcz\perso\le�on13A.txt"
VB13:
IfFileExists "$EXEDIR\le�on13B.txt" B13 VC13
B13:
copyfiles "$EXEDIR\le�on13B.txt" "c:\temp\apcz\perso\le�on13B.txt"
VC13:
IfFileExists "$EXEDIR\le�on13C.txt" C13 VD13
C13:
copyfiles "$EXEDIR\le�on13C.txt" "c:\temp\apcz\perso\le�on13C.txt"
VD13:
IfFileExists "$EXEDIR\le�on13D.txt" D13 VE13
D13:
copyfiles "$EXEDIR\le�on13D.txt" "c:\temp\apcz\perso\le�on13D.txt"
VE13:
IfFileExists "$EXEDIR\le�on13E.txt" E13 VF13
E13:
copyfiles "$EXEDIR\le�on13E.txt" "c:\temp\apcz\perso\le�on13E.txt"
VF13:
IfFileExists "$EXEDIR\le�on13F.txt" F13 VA14
F13:
copyfiles "$EXEDIR\le�on13F.txt" "c:\temp\apcz\perso\le�on13F.txt"
VA14:
IfFileExists "$EXEDIR\le�on14A.txt" A14 VB14
A14:
copyfiles "$EXEDIR\le�on14A.txt" "c:\temp\apcz\perso\le�on14A.txt"
VB14:
IfFileExists "$EXEDIR\le�on14B.txt" B14 VC14
B14:
copyfiles "$EXEDIR\le�on14B.txt" "c:\temp\apcz\perso\le�on14B.txt"
VC14:
IfFileExists "$EXEDIR\le�on14C.txt" C14 VA15
C14:
copyfiles "$EXEDIR\le�on14C.txt" "c:\temp\apcz\perso\le�on14C.txt"
VA15:
IfFileExists "$EXEDIR\le�on15A.txt" A15 VB15
A15:
copyfiles "$EXEDIR\le�on15A.txt" "c:\temp\apcz\perso\le�on15A.txt"
VB15:
IfFileExists "$EXEDIR\le�on15B.txt" B15 VC15
B15:
copyfiles "$EXEDIR\le�on15B.txt" "c:\temp\apcz\perso\le�on15B.txt"
VC15:
IfFileExists "$EXEDIR\le�on15C.txt" C15 VA16
C15:
copyfiles "$EXEDIR\le�on15C.txt" "c:\temp\apcz\perso\le�on15C.txt"
VA16:
IfFileExists "$EXEDIR\le�on16A.txt" A16 VB16
A16:
copyfiles "$EXEDIR\le�on16A.txt" "c:\temp\apcz\perso\le�on16A.txt"
VB16:
IfFileExists "$EXEDIR\le�on16B.txt" B16 VC16
B16:
copyfiles "$EXEDIR\le�on16B.txt" "c:\temp\apcz\perso\le�on16B.txt"
VC16:
IfFileExists "$EXEDIR\le�on16C.txt" C16 VD16
C16:
copyfiles "$EXEDIR\le�on16C.txt" "c:\temp\apcz\perso\le�on16C.txt"
VD16:
IfFileExists "$EXEDIR\le�on16D.txt" D16 VA17
D16:
copyfiles "$EXEDIR\le�on16D.txt" "c:\temp\apcz\perso\le�on16D.txt"
VA17:
IfFileExists "$EXEDIR\le�on17A.txt" A17 VB17
A17:
copyfiles "$EXEDIR\le�on17A.txt" "c:\temp\apcz\perso\le�on17A.txt"
VB17:
IfFileExists "$EXEDIR\le�on17B.txt" B17 VC17
B17:
copyfiles "$EXEDIR\le�on17B.txt" "c:\temp\apcz\perso\le�on17B.txt"
VC17:
IfFileExists "$EXEDIR\le�on17C.txt" C17 VD17
C17:
copyfiles "$EXEDIR\le�on17C.txt" "c:\temp\apcz\perso\le�on17C.txt"
VD17:
IfFileExists "$EXEDIR\le�on17D.txt" D17 VA18
D17:
copyfiles "$EXEDIR\le�on17D.txt" "c:\temp\apcz\perso\le�on17D.txt"
VA18:
IfFileExists "$EXEDIR\le�on18A.txt" A18 VB18
A18:
copyfiles "$EXEDIR\le�on18A.txt" "c:\temp\apcz\perso\le�on18A.txt"
VB18:
IfFileExists "$EXEDIR\le�on18B.txt" B18 VC18
B18:
copyfiles "$EXEDIR\le�on18B.txt" "c:\temp\apcz\perso\le�on18B.txt"
VC18:
IfFileExists "$EXEDIR\le�on18C.txt" C18 VD18
C18:
copyfiles "$EXEDIR\le�on18C.txt" "c:\temp\apcz\perso\le�on18C.txt"
VD18:
IfFileExists "$EXEDIR\le�on18D.txt" D18 VE18
D18:
copyfiles "$EXEDIR\le�on18D.txt" "c:\temp\apcz\perso\le�on18D.txt"
VE18:
IfFileExists "$EXEDIR\le�on18E.txt" E18 VA19
E18:
copyfiles "$EXEDIR\le�on18E.txt" "c:\temp\apcz\perso\le�on18E.txt"
VA19:
IfFileExists "$EXEDIR\le�on19A.txt" A19 VB19
A19:
copyfiles "$EXEDIR\le�on19A.txt" "c:\temp\apcz\perso\le�on19A.txt"
VB19:
IfFileExists "$EXEDIR\le�on19B.txt" B19 VC19
B19:
copyfiles "$EXEDIR\le�on19B.txt" "c:\temp\apcz\perso\le�on19B.txt"
VC19:
IfFileExists "$EXEDIR\le�on19C.txt" C19 VD19
C19:
copyfiles "$EXEDIR\le�on19C.txt" "c:\temp\apcz\perso\le�on19C.txt"
VD19:
IfFileExists "$EXEDIR\le�on19D.txt" D19 VA20
D19:
copyfiles "$EXEDIR\le�on19D.txt" "c:\temp\apcz\perso\le�on19D.txt"
VA20:

FunctionEnd


; *************************************************************************
Section "Uninstall"

; Lancer l'ex�cutable de d�sinstallation
ExecWait '"$WINDIR\ApprentiClavier_setup.exe" /D'

; Selon le choix utilisateur, ou le succ�s de l'installation !
IfFileExists "c:\ApprentiClavier\ApprentiClavier.exe" Z2 Z1

Z1:
; D�sinstallation compl�te
; Supprimer les cl�s du registre
;messagebox MB_OK|MB_TOPMOST "Effacement des cl�s de la base de registres"
DeleteRegKey HKLM "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ApprentiClavier"

; Supprimer les shortcuts
Delete "$SMPROGRAMS\ApprentiClavier\*.*"
rmdir "$SMPROGRAMS\ApprentiClavier"

; Supprimer l'ic�ne du bureau
Delete "$DESKTOP\ApprentiClavier.lnk"
Delete "$DESKTOP\Apprenti Clavier.lnk"

; Supprimer les fichiers r�siduels
rmdir /r "c:\ApprentiClavier"

; Supprimer les ex�cutables d'installation/d�sinstallation
Delete "$WINDIR\ApprentiClavier_setup.exe"
Delete "$WINDIR\ApprentiClavier_uninstall.exe"
Delete "$WINDIR\VB40032.dll"

Z2:
; Supprimer le d�chargement temporaire
; rmdir /r "c:\temp\apcz"
SectionEnd