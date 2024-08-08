VERSION 4.00
Begin VB.Form Menu_principal 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Menu principal"
   ClientHeight    =   6480
   ClientLeft      =   765
   ClientTop       =   1965
   ClientWidth     =   9855
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
   Height          =   7290
   Left            =   705
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   9855
   Top             =   1215
   Width           =   9975
   Begin VB.Timer Timer1 
      Interval        =   400
      Left            =   8400
      Top             =   0
   End
   Begin VB.CommandButton Quitter 
      Caption         =   "&Quitter  (Échap)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   7800
      TabIndex        =   1
      Top             =   4920
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   4110
      ItemData        =   "Menu_principal.frx":0000
      Left            =   120
      List            =   "Menu_principal.frx":0028
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   9255
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
      Top             =   6000
      Width           =   7485
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
      Height          =   1335
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   7605
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
      Height          =   315
      Left            =   3360
      TabIndex        =   2
      Top             =   0
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
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu DebExpliNormal 
         Caption         =   "Débit des explications &Normal"
      End
      Begin VB.Menu DebExpliRapide 
         Caption         =   "Débit des explications &Rapide"
      End
      Begin VB.Menu Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu DebGenLent 
         Caption         =   "Débit général  &Lent"
      End
      Begin VB.Menu DebGenMoyen 
         Caption         =   "Débit général &Moyen"
      End
      Begin VB.Menu DebGenVite 
         Caption         =   "Débit général  &Vite"
      End
      Begin VB.Menu Sep3 
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
      Begin VB.Menu Sep6 
         Caption         =   "-"
      End
      Begin VB.Menu Restart 
         Caption         =   "Redémarrer à la prem&ière leçon"
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
      Begin VB.Menu Seperator0 
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
         Caption         =   "A  &propos de"
      End
   End
End
Attribute VB_Name = "Menu_principal"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' ********************  LOAD  *******************************
Private Sub Form_Load()
' La fenêtre remplit l'écran sauf en mode debug
If FullScreenSwitch = 1 Then WindowState = 2

'Paramètres
ii = 0 'Reset après la routine de recherche du rép Jaws
winstop = 0
Module_menus.menu_reset "menu_principal.txt"
Set menu_courant = Menu_principal

'Taille des objets de la fenêtre selon la définition de l'écran
Module_routines.Colors Menu_principal  '12/2011
Module_routines.Dimension Menu_principal
Module_routines.niveaux
Module_routines.MenuEditorTrans Menu_principal
If nivo = msgStandard Then kk = 0
If nivo = msgPersonnalisé Then kk = 25

' Attention : le script Jaws jss détecte les blancs avant et au milieu du titre (caption)
Menu_principal.Caption = debexplivalue & bannerMenu & debgenvalue & bannerPrincipal
Menu_principal.label1.Caption = msgChoisissez
Menu_principal.Quitter.Caption = msgQuitter & msgÉchap

' Repérer la longueur mm de la plus grande ligne
Open vpath & "menu_courant.txt" For Input As #2
jj = 0
mm = 0
Do While Not EOF(2)
    Line Input #2, currentline
    If Len(currentline) > mm Then mm = Len(currentline)
    jj = jj + 1
Loop
Close #2

' Inclure les résultats
Open vpath & "menu_courant.txt" For Input As #2
jj = 0
Do While Not EOF(2)
    Line Input #2, currentline
    If pctok(jj + kk, 0) = 0 Then
        Menu_principal.list1.List(jj) = currentline
    Else
        'nb nn d'espaces d'alignement à la fin de la ligne menu, avant le résultat pctok
        nn = mm - Len(currentline) + 1
        pp = 0: tempo = ""
        Do While pp < nn
            tempo = tempo & " "
            pp = pp + 1
        Loop
        If Not pctok(jj + kk, 0) = 100 Then tempo = tempo & " "
        If vitok(jj + kk, 0) = 0 Then Menu_principal.list1.List(jj) = currentline & tempo & pctok(jj + kk, 0) & "%"
        If Not vitok(jj + kk, 0) = 0 Then Menu_principal.list1.List(jj) = currentline & tempo & pctok(jj + kk, 0) & "% " & vitok(jj + kk, 0) & msgMotsMinute
    End If
    jj = jj + 1
Loop
Close #2
list1.Selected(numleçon) = True
Module_routines.mshow Menu_principal
Label3.Caption = bannerVersion & ", " & bannerCopyright
numpad = 0  ' Default, nécessaire pour retour après menu_leçon16

' Pour pouvoir se déplacer dan le menu principal par les initiales chiffres
Module_routines.SetKeys "NUMLOCK_ON"
End Sub

Private Sub Timer1_Timer()
list1.Visible = True
Quitter.Visible = True
Timer1.Enabled = False
End Sub


' ******************  DBLCLICK  ******************************
Private Sub List1_DblClick()
If list1.ListIndex = 0 Then
    Module_routines.presentation
    Exit Sub
End If
If list1.ListIndex = 1 Then
    Module_routines.pourqui
    Exit Sub
End If
If list1.ListIndex = 2 Then
    Module_routines.conseils
    Exit Sub
End If
If list1.ListIndex = 3 Then
    Unload Menu_principal
    Menu_leçon1.Show 1
End If
If list1.ListIndex = 4 Then
    Unload Menu_principal
    Menu_leçon2.Show 1
End If
If list1.ListIndex = 5 Then
    Unload Menu_principal
    Menu_leçon3.Show 1
End If
If list1.ListIndex = 6 Then
    Unload Menu_principal
    Menu_leçon4.Show 1
End If
If list1.ListIndex = 7 Then
    Unload Menu_principal
    Menu_leçon5.Show 1
End If
If list1.ListIndex = 8 Then
    Unload Menu_principal
    Menu_leçon6.Show 1
End If
If list1.ListIndex = 9 Then
    Unload Menu_principal
    Menu_leçon7.Show 1
End If
If list1.ListIndex = 10 Then
    Unload Menu_principal
    Menu_leçon8.Show 1
End If
If list1.ListIndex = 11 Then
    Unload Menu_principal
    Menu_leçon9.Show 1
End If
If list1.ListIndex = 12 Then
    Unload Menu_principal
    Menu_leçon10.Show 1
End If
If list1.ListIndex = 13 Then
    Unload Menu_principal
    Menu_leçon11.Show 1
End If
If list1.ListIndex = 14 Then
    Unload Menu_principal
    Menu_leçon12.Show 1
End If
If list1.ListIndex = 15 Then
    Unload Menu_principal
    Menu_leçon13.Show 1
End If
If list1.ListIndex = 16 Then
    Unload Menu_principal
    Menu_leçon14.Show 1
End If
If list1.ListIndex = 17 Then
    Unload Menu_principal
    Menu_leçon15.Show 1
End If
If list1.ListIndex = 18 Then
    Unload Menu_principal
    Menu_leçon16.Show 1
End If
If list1.ListIndex = 19 Then
    Unload Menu_principal
    Menu_leçon17.Show 1
End If
If list1.ListIndex = 20 Then
    Unload Menu_principal
    Menu_leçon18.Show 1
End If
If list1.ListIndex = 21 Then
    Unload Menu_principal
    Menu_leçon19.Show 1
End If

' Voir les résultats (attention à numleçon perdu par la routine results)
If list1.ListIndex = 22 Then
    Unload Menu_principal
    tempnum = numleçon
    Module_routines.results
    numleçon = tempnum
    vfileresults = ""
    If nivo = msgStandard And Dir(vfile & "\Résultat-Standard.doc") <> "" Then vfileresults = vfile & "\Résultat-Standard.doc"
    If nivo = msgPersonnalisé And Dir(vfile & "\Résultat-Personnalisé.doc") <> "" Then vfileresults = vfile & "\Résultat-Personnalisé.doc"
    If vfileresults = "" Then
        MsgBox msgNofic & vfile & "\Résultat-" & nivoRep & ".doc" & msgPrincPour & nom & msgPrincDansniveau & nivo & ".", 0, ""
        keyinhibit = 1
        Menu_principal.Show 1
        Exit Sub
    End If
    
    ' Information
L22A:
    pagenum = 1
    msgtext0 = CRLF + msgPrincContenu & vfile & "\Résultat-" & nivoRep & ".doc" + pressez
    Msgform.Show 1
    If msgf = 33 Then Beep
    If msgf = 2 Or msgf = 33 Then GoTo L22A
    If msgf = 0 Then
        Menu_principal.Show 1
        Exit Sub
    End If
    
    ' Mise des résultats dans msgtext0, par pages visibles de 17 lignes
L22B:
    stopscroll = 0
    pagenum = 2
    Module_routines.scrollresults 0, 15
    If msgf = 33 Then GoTo L22A
    If stopscroll = 1 Then GoTo L22I
L22C:
    pagenum = 3
    Module_routines.scrollresults 15, 15
    If msgf = 33 Then GoTo L22B
    If stopscroll = 1 Then GoTo L22I
L22D:
    pagenum = 4
    Module_routines.scrollresults 30, 15
    If msgf = 33 Then GoTo L22C
    If stopscroll = 1 Then GoTo L22I
L22E:
    pagenum = 5
    Module_routines.scrollresults 45, 15
    If msgf = 33 Then GoTo L22D
    If stopscroll = 1 Then GoTo L22I
L22F:
    pagenum = 6
    Module_routines.scrollresults 60, 15
    If msgf = 33 Then GoTo L22E
    If stopscroll = 1 Then GoTo L22I
L22G:
    pagenum = 7
    Module_routines.scrollresults 75, 15
    If msgf = 33 Then GoTo L22F
    If stopscroll = 1 Then GoTo L22I
L22H:
    pagenum = 8
    Module_routines.scrollresults 90, 15
    If msgf = 33 Then GoTo L22G
L22I:
    pagenum = pagenum + 1
L22J:
    pagemax = 1
    msgtext0 = CRLF + msgPrincTerminé
    Msgform.Text0.Font.Size = 2 * Msgform.Text0.Font.Size
    Msgform.Show 1
    If msgf = 33 Then
        If pagenum = 3 Then GoTo L22B
        If pagenum = 4 Then GoTo L22C
        If pagenum = 5 Then GoTo L22D
        If pagenum = 6 Then GoTo L22E
        If pagenum = 7 Then GoTo L22F
        If pagenum = 8 Then GoTo L22G
        If pagenum = 9 Then GoTo L22H
    End If
    If msgf = 34 Then GoTo L22J
    Menu_principal.Show 1
End If

If list1.ListIndex = 23 Then End
End Sub


' ******************** LIST1_KEYDOWN **********************************************
Private Sub list1_KeyDown(Keycode As Integer, shift As Integer)
' Win 91 et Win 92 (voir en plus Text2_KeyUp)
If Keycode = 91 Or Keycode = 92 Then
    winstop = winstop + 1
    Module_routines.cancelwin 0, Menu_principal, 1
End If
End Sub


' ******************** LIST1_KEYUP **********************************************
Private Sub List1_KeyUp(Keycode As Integer, shift As Integer)
' Win 91 et Win 92 (voir en plus Text2_KeyDown)
If Keycode = 91 Or Keycode = 92 Then Module_routines.cancelwin 0, Menu_principal, 1

' Echappement
If Keycode = 27 Then
' Winstop stoppe Win ou Win+E, Win+F, Win+L, Win+R, Win+U
    If winstop > 0 Then
        winstop = winstop - 1
        Exit Sub
    End If
    If echapbis = -1 Then
        echapbis = 0
    Else
BV11:
        Unload Menu_principal
        msgtext0 = pressez_quit
        fsize = 1.5 * fsizedefault * zfactor
        fbc = fbc_quit
        ffc = ffc_quit
        pagenum = 0
        Msgform.Quitter.Caption = msgQuitter & msgÉchap
        Msgform.Show 1
        ffc = ffc_default
        fbc = fbc_default
        fsize = fsizedefault * zfactor
        If msgf = 2 Then GoTo BV11
        If msgf = 0 Then Quitter_Click
        Menu_principal.Show 1
        keyinhibit = 1
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


' *******************  KEYPRRESS  *********************************
Private Sub List1_KeyPress(KeyAscii As Integer)
Module_routines.SetKeys "CAPSLOCK_OFF"
Module_routines.SetKeys "SCROLLLOCK_OFF"
echapbis = 0

' Pour sonoriser en répétant la ligne menu en cours
If KeyAscii = 32 Then Module_routines.menu_repeat
End Sub


' *******************  QUITTER  ***********************************
Private Sub Quitter_Click()
If quitactive = 0 Then Module_routines.QuitQuit
End Sub

Private Sub Quitter_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    On Error Resume Next
    list1.SetFocus
End If
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
Quitter_Click
End Sub

Private Sub Personnalisé_Click()
nivo = msgPersonnalisé
nivoRep = "Personnalisé" 'immuable, ne pas traduire
numleçon = list1.ListIndex
Unload Menu_principal
MsgBox msgLevelIs & nivo & ".", 0, debexplivalue & nivo
keyinhibit = 0
Menu_principal.Show 1
End Sub

Private Sub Standard_Click()
nivo = msgStandard
nivoRep = "Standard" 'immuable, ne pas traduire
numleçon = list1.ListIndex
Unload Menu_principal
MsgBox msgLevelIs & nivo & ".", 0, debexplivalue & nivo
keyinhibit = 0
Menu_principal.Show 1
End Sub

Private Sub DebExpliNormal_Click()
keyinhibit = 0
Module_routines.DebExpliNormal
End Sub

Private Sub DebExpliRapide_Click()
keyinhibit = 0
Module_routines.DebExpliRapide
End Sub

Private Sub DebGenLent_Click()
keyinhibit = 0
Module_routines.DebGenLent
End Sub

Private Sub DebGenMoyen_Click()
keyinhibit = 0
Module_routines.DebGenMoyen
End Sub

Private Sub DebGenVite_Click()
keyinhibit = 0
Module_routines.DebGenVite
End Sub

Private Sub BipClassique_Click()
keyinhibit = 0
Module_routines.BipClassique
End Sub

Private Sub BipDifférent_Click()
keyinhibit = 0
Module_routines.BipDifférent
End Sub

'12/2011
Private Sub BasicColors_Click()
keyinhibit = 0
Module_routines.BasicColors
End Sub

'12/2011
Private Sub OtherColors_Click()
keyinhibit = 0
Module_routines.OtherColors
End Sub

'12/2011
Private Sub NoZoom_Click()
keyinhibit = 0
Module_routines.NoZoom
End Sub

'12/2011
Private Sub WithZoom_Click()
keyinhibit = 0
Module_routines.WithZoom
End Sub

Private Sub Restart_Click() 'Effacer les résultats de l'utilisateur et reset à la première leçon
vmsgbox = MsgBox(msgRestart & nom & " ?" & msgRestartCmd, 1, msgRestartTitle)
If vmsgbox = 2 Then Exit Sub

' Reset the user ini file to lesson 1
numleçon = 0
numexo = 0

' Reset the pctok table (last column is for lesson number)
For jj = 0 To 49
    For kk = 0 To 8
    pctok(jj, kk) = 0
    Next kk
    'Visualise numéro de leçon en dernière col
    If jj < 25 Then pctok(jj, kk) = jj - 2
    If jj >= 25 Then pctok(jj, kk) = jj - 27
Next jj

' Reset the vitok table (last column is for lesson number)
For jj = 0 To 49
    For kk = 0 To 8
    vitok(jj, kk) = 0
    Next kk
    'Visualise numéro de leçon en dernière col
    If jj < 25 Then vitok(jj, kk) = jj - 2
    If jj >= 25 Then vitok(jj, kk) = jj - 27
Next jj

'Reload menu
Unload Menu_principal
keyinhibit = 0
Menu_principal.Show 1
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

Public Sub Aproposde_Click()
keyinhibit = 1
'MsgBox bannerVersion & ", " & bannerFunction & ", " & CRLF & bannerCopyright & ", Hervé Béranger, " & bannerAuthorAddress & "." & CRLF2 & bannerNosell & CRLF2 & bannerThanks & CRLF2 & msgTypeClavier & CRLF2 & msgTranslator, 0, debexplivalue
MsgBox bannerVersion & ", " & bannerFunction & ", " & CRLF & bannerCopyright & ", Hervé Béranger, " & bannerAuthorAddress & "." & CRLF2 & bannerNosell & CRLF2 & bannerThanks & CRLF2 & msgTypeClavier, 0, debexplivalue
End Sub

