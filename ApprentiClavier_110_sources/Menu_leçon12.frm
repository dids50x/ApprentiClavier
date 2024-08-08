VERSION 4.00
Begin VB.Form Menu_leçon12 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Menu leçon 12"
   ClientHeight    =   6120
   ClientLeft      =   585
   ClientTop       =   1905
   ClientWidth     =   9870
   ControlBox      =   0   'False
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   6930
   Left            =   525
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   9870
   Top             =   1155
   Width           =   9990
   Begin VB.ListBox List1 
      BackColor       =   &H0000C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      ItemData        =   "Menu_leçon12.frx":0000
      Left            =   120
      List            =   "Menu_leçon12.frx":0002
      TabIndex        =   0
      Top             =   840
      Width           =   9135
   End
   Begin VB.CommandButton Quitter 
      Caption         =   " &Quitter vers Menu Principal  (Échap)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   7560
      TabIndex        =   2
      Top             =   4320
      Width           =   2265
   End
   Begin VB.Label Label3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Label3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Width           =   7335
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   7215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Choisissez."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.Menu Fichier 
      Caption         =   "&Fichier"
      Begin VB.Menu Quitter_bm 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu Options 
      Caption         =   "&Options"
      Begin VB.Menu Standard 
         Caption         =   "Niveau &Standard"
      End
      Begin VB.Menu Personnalisé 
         Caption         =   "Niveau &Personnalisé"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu DebExpliNormal 
         Caption         =   "Débit des explications &Normal"
      End
      Begin VB.Menu DebExpliRapide 
         Caption         =   "Débit des explications &Rapide"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu DebGenLent 
         Caption         =   "Débit général &Lent"
      End
      Begin VB.Menu DebGenMoyen 
         Caption         =   "Débit général &Moyen"
      End
      Begin VB.Menu DebGenVite 
         Caption         =   "Débit général &Vite"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu BipClassique 
         Caption         =   "Bip &Classique"
      End
      Begin VB.Menu BipDifférent 
         Caption         =   "Bip &Différent"
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu BasicColors 
         Caption         =   "Couleurs &basiques"
      End
      Begin VB.Menu OtherColors 
         Caption         =   "A&utres couleurs"
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu NoZoom 
         Caption         =   "Sans z&oom"
      End
      Begin VB.Menu WithZoom 
         Caption         =   "Avec &zoom"
      End
   End
   Begin VB.Menu Aide 
      Caption         =   "&Aide"
      Begin VB.Menu AideGénérale 
         Caption         =   "&Aide générale"
         Shortcut        =   {F1}
      End
      Begin VB.Menu AideMémoire 
         Caption         =   "Aide-Mémoire"
         Shortcut        =   {F3}
      End
      Begin VB.Menu Separator0 
         Caption         =   "-"
      End
      Begin VB.Menu Enseignant 
         Caption         =   "Aide pour l'&Enseignant"
      End
      Begin VB.Menu Sonorisation 
         Caption         =   "Aide sur la &Sonorisation"
      End
      Begin VB.Menu Separator 
         Caption         =   "-"
      End
      Begin VB.Menu Aproposde 
         Caption         =   "A &propos de"
      End
   End
End
Attribute VB_Name = "Menu_leçon12"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'*******************************  LOAD  ************************************************
' MENU : Insertion, Suppression, Déplacements rapides
Private Sub Form_Load()
If FullScreenSwitch = 1 Then WindowState = 2
Menu_leçon12.Quitter.Caption = msgQuitterMP & msgÉchap

'Paramètres
If numleçon <> 14 Then numexo = 0
numleçon = 14   ' Toujours leçon n + 2
Set menu_courant = Menu_leçon12
Set menu_suivant = Menu_leçon13
Module_routines.Colors Menu_leçon12  '12/2011
Module_routines.Dimension Menu_leçon12
Module_menus.menu_reset "menu_leçon12.txt"
Module_routines.menu_refresh "menu_courant.txt", Menu_leçon12
Module_routines.mshow Menu_leçon12
Label3.Caption = bannerVersion & ", " & bannerCopyright
Module_routines.niveaux
Module_routines.MenuEditorTrans Menu_leçon12
menucount = menu_courant.list1.ListCount
echapbismax = 0   ' echapbismax + 1 coups Échap pour sortir
indif = 0: sonocara = 1

' Attention : le script Jaws jss détecte les blancs avant et au milieu du titre (caption)
Menu_leçon12.Caption = debexplivalue & bannerMenu & debgenvalue & bannerLeçon & " 12"
Menu_leçon12.label1.Caption = msgChoisissez

' Ici, pas dans quit_l, sinon sono transitoire du bureau
If consult = 0 Then Module_routines.OpenAndSuffix exo_courant, 0

' Pour se déplacer dans le menu par les initiales lettres
Module_routines.SetKeys "NUMLOCK_OFF"
End Sub


'**************************** LIST1_DBLCLICK  ******************************************
Private Sub List1_DblClick()
'****************  EXERCICE 12A ********************************************
' MENU : Insertion/Remplacement, Suppression
' Attention ,pour que Jaws ne perturbe pas le fonctionnement de Insertion, il faut décocher
' la case à cocher "Insertion du clavier étendu" dans le configurateur Jaws/Options-clavier
If list1.ListIndex = 0 Then
    numexo = 0  ' 0 pour A
    Unload Menu_leçon12
    tempo = Dir(vpath & "Leçons\" & nivoRep & "\leçon12A.txt")
    If tempo = "" Then
ML10:
        pagenum = 0
        msgtext0 = CRLF + msgNofic & vpath & "Leçons\" & nivoRep & "\leçon12A.txt. " + perso_methode + pressez
        Msgform.Show 1
        If msgf = 2 Then GoTo ML10
        Menu_leçon12.Show 1
        Exit Sub
    End If
ML11:
    pagenum = 1
    msgtext0 = CRLF + pg12a1 + pressez
    Msgform.Show 1
    If msgf = 33 Then Beep
    If msgf = 2 Or msgf = 33 Then GoTo ML11
    If msgf = 1 Or msgf = 34 Then
ML12:
        pagenum = 2
        msgtext0 = CRLF + pg12a2 + pressez
        Msgform.Show 1
        If msgf = 33 Then GoTo ML11
        If msgf = 2 Then GoTo ML12
        If msgf = 1 Or msgf = 34 Then
ML13:
            pagenum = 3
            msgtext0 = CRLF + pg12a3 + pressez
            Msgform.Show 1
            If msgf = 33 Then GoTo ML12
            If msgf = 2 Then GoTo ML13
            If msgf = 1 Or msgf = 34 Then
ML14:
                pagenum = 4: pagemax = 1
                msgtext0 = CRLF + pg12a4 + pressez
                Msgform.Show 1
                If msgf = 33 Then GoTo ML13
                If msgf = 34 Then Beep
                If msgf = 2 Or msgf = 34 Then GoTo ML14
                If msgf = 1 Then
                    exo_courant = "leçon12A.txt"
                    
                    ' Msg d'encouragements et d'explications
                    Module_routines.resetmsg
                    msgtext2(1) = CRLF + pg12am1
                    msgtext2(4) = CRLF + pg12am2
                    
                    ' Go
                    espacevalid = 1 'Pour accepter ESPACE pour RÉPÉTER
                    leçon1.Caption = bannerLeçon & " 12 A."
                    leçon1.Show 1
                End If
            End If
        End If
    End If
    If msgf = 0 Then Menu_leçon12.Show 1
End If

'****************  EXERCICE 12B ********************************************
' MENU : Touches début et Fin
If list1.ListIndex = 1 Then
    numexo = 1
    Unload Menu_leçon12
    tempo = Dir(vpath & "Leçons\" & nivoRep & "\leçon12B.txt")
    If tempo = "" Then
ML20:
        pagenum = 0
        msgtext0 = CRLF + msgNofic & vpath & "Leçons\" & nivoRep & "\leçon12B.txt. " + perso_methode + pressez
        Msgform.Show 1
        If msgf = 2 Then GoTo ML20
        Menu_leçon12.Show 1
        Exit Sub
    End If
ML21:
    pagenum = 1
    msgtext0 = CRLF + pg12b1 + pressez
    Msgform.Show 1
    If msgf = 33 Then Beep
    If msgf = 2 Or msgf = 33 Then GoTo ML21
    If msgf = 1 Or msgf = 34 Then
ML22:
        pagenum = 2: pagemax = 1
        msgtext0 = CRLF + pg12b2 + pressez
        Msgform.Show 1
        If msgf = 33 Then GoTo ML21
        If msgf = 34 Then Beep
        If msgf = 2 Or msgf = 34 Then GoTo ML22
        If msgf = 1 Then
            exo_courant = "leçon12B.txt"
            
            ' Msg d'encouragements et d'explications
            Module_routines.resetmsg
            msgtext1(16) = pg12bm1
            
            ' Go
            espacevalid = 1 'Pour accepter ESPACE pour RÉPÉTER
            leçon1.Caption = bannerLeçon & " 12 B."
            leçon1.Show 1
        End If
    End If
    If msgf = 0 Then Menu_leçon12.Show 1
End If

'****************  EXERCICE 12C ********************************************
' MENU : Page-Précédente et Page-Suivante
If list1.ListIndex = 2 Then
    numexo = 2
    Unload Menu_leçon12
    tempo = Dir(vpath & "Leçons\" & nivoRep & "\leçon12C.txt")
    If tempo = "" Then
ML30:
        pagenum = 0
        msgtext0 = CRLF + msgNofic & vpath & "Leçons\" & nivoRep & "\leçon12C.txt. " + perso_methode + pressez
        Msgform.Show 1
        If msgf = 2 Then GoTo ML30
        Menu_leçon12.Show 1
        Exit Sub
    End If
ML31:
    pagenum = 1
    msgtext0 = CRLF + pg12c1 + pressez
    Msgform.Show 1
    If msgf = 33 Then Beep
    If msgf = 2 Or msgf = 33 Then GoTo ML31
    If msgf = 1 Or msgf = 34 Then
ML32:
        pagenum = 2: pagemax = 1
        msgtext0 = CRLF + pg12c2 + pressez
        Msgform.Show 1
        If msgf = 33 Then GoTo ML31
        If msgf = 34 Then Beep
        If msgf = 2 Or msgf = 34 Then GoTo ML32
        If msgf = 1 Then
            exo_courant = "leçon12C.txt"
            
            ' Msg d'encouragements et d'explications
            Module_routines.resetmsg
            
            ' Go
            espacevalid = 1 'Pour accepter ESPACE pour RÉPÉTER
            leçon1.Caption = bannerLeçon & " 12 C."
            leçon1.Show 1
        End If
    End If
    If msgf = 0 Then Menu_leçon12.Show 1
End If

'****************  EXERCICE 12D ********************************************
' MENU : Impression, ArrêtDéfil, Pause
If list1.ListIndex = 3 Then
    numexo = 3
    Unload Menu_leçon12
    tempo = Dir(vpath & "Leçons\" & nivoRep & "\leçon12D.txt")
    If tempo = "" Then
ML40:
        pagenum = 0
        msgtext0 = CRLF + msgNofic & vpath & "Leçons\" & nivoRep & "\leçon12D.txt. " + perso_methode + pressez
        Msgform.Show 1
        If msgf = 2 Then GoTo ML40
        Menu_leçon12.Show 1
        Exit Sub
    End If
ML41:
    pagenum = 1
    msgtext0 = CRLF + pg12d1 + pressez
    Msgform.Show 1
    If msgf = 33 Then Beep
    If msgf = 2 Or msgf = 33 Then GoTo ML41
    If msgf = 1 Or msgf = 34 Then
ML42:
        pagenum = 2: pagemax = 1
        msgtext0 = CRLF + pg12d2 + pressez
        Msgform.Show 1
        If msgf = 33 Then GoTo ML41
        If msgf = 34 Then Beep
        If msgf = 2 Or msgf = 34 Then GoTo ML42
        If msgf = 1 Then
            exo_courant = "leçon12D.txt"
            
            ' Msg d'encouragements et d'explications
            Module_routines.resetmsg
            
            ' Go
            espacevalid = 1 'Pour accepter ESPACE pour RÉPÉTER
            leçon1.Caption = bannerLeçon & " 12 D."
            leçon1.Show 1
        End If
    End If
    If msgf = 0 Then Menu_leçon12.Show 1
End If

'************************ Fin de Dbl_click **********************************
End Sub


' ******************** LIST1_KEYDOWN **********************************************
Private Sub list1_KeyDown(Keycode As Integer, shift As Integer)
' Win 91 et Win 92 (voir en plus Text2_KeyUp)
If Keycode = 91 Or Keycode = 92 Then Module_routines.cancelwin 0, Menu_leçon12, 1
End Sub


' ******************** LIST1_KEYUP **********************************************
Private Sub List1_KeyUp(Keycode As Integer, shift As Integer)
' Win 91 et Win 92 (voir en plus Text2_KeyDown)
If Keycode = 91 Or Keycode = 92 Then Module_routines.cancelwin 0, Menu_leçon12, 1

' Echappement
If Keycode = 27 Then
    If echapbis >= 0 Then
        Quitter_Click
    Else
        echapbis = echapbis + 1
    End If
End If

' Entrée
If Keycode = 13 Then
    If keyinhibit <> 0 Then
        keyinhibit = 0
    Else
        List1_DblClick
    End If
End If

' Touche F2
If Keycode = 113 Then
    quitF2 = 1
    msgtext0 = pressez_F2 + pressez_touche
    Msgform.Show 1
End If

'Touche Alt+F4 pour quitter
If Keycode = 115 And shift = 4 Then
    altf4 = 1
    If quitactive = 0 Then Module_routines.QuitQuit
End If
End Sub


'****************************  LIST1_KEYPRESS  *****************************************
Private Sub List1_KeyPress(KeyAscii As Integer)
Module_routines.SetKeys "CAPSLOCK_OFF"
Module_routines.SetKeys "NUMLOCK_OFF"
Module_routines.SetKeys "SCROLLLOCK_OFF"
echapbis = 0  'Reset après appel menu Options

' Pour sonoriser en répétant la ligne menu en cours
If KeyAscii = 32 Then Module_routines.menu_repeat
End Sub


'*******************************  QUITTER  *********************************************
Private Sub Quitter_Click()
' Décharger/recharger
Unload Menu_leçon12
Unload Menu_principal  'reset label2
Menu_principal.Show 1
End Sub

Private Sub Quitter_KeyPress(KeyAscii As Integer)
If KeyAscii = 81 Or KeyAscii = 113 Then Quitter_Click
End Sub


' **************  COMMANDES de la BARRE de MENU  *******************
Private Sub Fichier_Click()
echapbis = -1
End Sub

Private Sub Options_Click()
echapbis = -1
End Sub

Private Sub Aide_Click()
echapbis = -1
End Sub

Private Sub Quitter_bm_Click()
If quitactive = 0 Then Module_routines.QuitQuit
End Sub

Private Sub Personnalisé_Click()
nivo = msgPersonnalisé
nivoRep = "Personnalisé" 'immuable, ne pas traduire
numexo = list1.ListIndex
Unload Menu_leçon12
MsgBox msgLevelIs & nivo & ".", 0, debexplivalue & nivo
keyinhibit = 1
Menu_leçon12.Show 1
End Sub

Private Sub Standard_Click()
nivo = msgStandard
nivoRep = "Standard" 'immuable, ne pas traduire
numexo = list1.ListIndex
Unload Menu_leçon12
MsgBox msgLevelIs & nivo & ".", 0, debexplivalue & nivo
keyinhibit = 1
Menu_leçon12.Show 1
End Sub

Private Sub DebExpliNormal_Click()
keyinhibit = 1
Module_routines.DebExpliNormal
End Sub

Private Sub DebExpliRapide_Click()
keyinhibit = 1
Module_routines.DebExpliRapide
End Sub

Private Sub DebGenLent_Click()
keyinhibit = 1
Module_routines.DebGenLent
End Sub

Private Sub DebGenMoyen_Click()
keyinhibit = 1
Module_routines.DebGenMoyen
End Sub

Private Sub DebGenVite_Click()
keyinhibit = 1
Module_routines.DebGenVite
End Sub

Private Sub BipClassique_Click()
keyinhibit = 1
Module_routines.BipClassique
End Sub

Private Sub BipDifférent_Click()
keyinhibit = 1
Module_routines.BipDifférent
End Sub

'12/2011
Private Sub BasicColors_Click()
keyinhibit = 1
Module_routines.BasicColors
End Sub

'12/2011
Private Sub OtherColors_Click()
keyinhibit = 1
Module_routines.OtherColors
End Sub

'12/2011
Private Sub NoZoom_Click()
keyinhibit = 1
Module_routines.NoZoom
End Sub

'12/2011
Private Sub WithZoom_Click()
keyinhibit = 1
Module_routines.WithZoom
End Sub

Private Sub AideGénérale_Click()
Module_routines.help_f1m
End Sub

Private Sub AideMémoire_Click()
Module_routines.help_f3m
End Sub

Public Sub Enseignant_Click()
Module_routines.placeinmsgaide "\Leçons\Personnalisé\info.txt"
keyinhibit = 1
End Sub

Public Sub Sonorisation_Click()
Module_routines.placeinmsgaide "sonorisation.txt"
keyinhibit = 1
End Sub

Private Sub Aproposde_Click()
Menu_principal.Aproposde_Click
End Sub

