VERSION 4.00
Begin VB.Form Menu_le�on15 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Menu le�on 15"
   ClientHeight    =   5535
   ClientLeft      =   600
   ClientTop       =   1935
   ClientWidth     =   9810
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
   Height          =   6345
   Left            =   540
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   9810
   Top             =   1185
   Width           =   9930
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
      Height          =   1785
      ItemData        =   "Menu_le�on15.frx":0000
      Left            =   120
      List            =   "Menu_le�on15.frx":000D
      TabIndex        =   0
      Top             =   960
      Width           =   8085
   End
   Begin VB.CommandButton Quitter 
      Caption         =   " &Quitter vers Menu Principal  (�chap)"
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
      Left            =   7680
      TabIndex        =   2
      Top             =   4080
      Width           =   2115
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
      Top             =   5040
      Width           =   7455
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
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   3240
      Width           =   7335
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
      Height          =   375
      Left            =   2640
      TabIndex        =   1
      Top             =   480
      Width           =   2295
   End
   Begin VB.Menu Fichier 
      Caption         =   "&Fichier"
      Begin VB.Menu Quitter_bm 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu OPtions 
      Caption         =   "&Options"
      Begin VB.Menu Standard 
         Caption         =   "Niveau &Standard"
      End
      Begin VB.Menu Personnalis� 
         Caption         =   "Niveau &Personnalis�"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu DebExpliNormal 
         Caption         =   "D�bit des explications &Normal"
      End
      Begin VB.Menu DebExpliRapide 
         Caption         =   "D�bit des explications &Rapide"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu DebGenLent 
         Caption         =   "D�bit g�n�ral &Lent"
      End
      Begin VB.Menu DebGenMoyen 
         Caption         =   "D�bit g�n�ral &Moyen"
      End
      Begin VB.Menu DebGenVite 
         Caption         =   "D�bit g�n�ral &Vite"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu BipClassique 
         Caption         =   "Bip &Classique"
      End
      Begin VB.Menu BipDiff�rent 
         Caption         =   "Bip &Diff�rent"
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
      Begin VB.Menu AideG�n�rale 
         Caption         =   "&Aide g�n�rale"
         Shortcut        =   {F1}
      End
      Begin VB.Menu AideM�moire 
         Caption         =   "Aide-M�moire"
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
Attribute VB_Name = "Menu_le�on15"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
'*****************  LOAD  ************************************************
' MENU : Encore des dict�es, plus rapides
Private Sub Form_Load()
If FullScreenSwitch = 1 Then WindowState = 2
Menu_le�on15.Quitter.Caption = msgQuitterMP & msg�chap

'Param�tres
If numle�on <> 17 Then numexo = 0
numle�on = 17   ' Toujours le�on n + 2
Set menu_courant = Menu_le�on15
Set menu_suivant = Menu_le�on16
Module_routines.Colors Menu_le�on15  '12/2011
Module_routines.Dimension Menu_le�on15
Module_menus.menu_reset "menu_le�on15.txt"
Module_routines.menu_refresh "menu_courant.txt", Menu_le�on15
Module_routines.mshow Menu_le�on15
Label3.Caption = bannerVersion & ", " & bannerCopyright
Module_routines.niveaux
Module_routines.MenuEditorTrans Menu_le�on15
menucount = menu_courant.list1.ListCount
indif = 0: sonocara = 0
echapbismax = 0   ' echapbismax + 1 coups �chap pour sortir

' Attention : le script Jaws jss d�tecte les blancs avant et au milieu du titre (caption)
Menu_le�on15.Caption = debexplivalue & bannerMenu & debgenvalue & bannerLe�on & " 15"
Menu_le�on15.label1.Caption = msgChoisissez

' Ici, pas dans quit_l, sinon sono transitoire du bureau
If consult = 0 Then Module_routines.OpenAndSuffix exo_courant, 0

' Pour se d�placer dans le menu par les initiales lettres
Module_routines.SetKeys "NUMLOCK_OFF"
End Sub


'******************** DBL_CLICK  ******************************************
Private Sub List1_DblClick()
'****************  EXERCICE 15A ********************************************
If list1.ListIndex = 0 Then
    numexo = 0
    Unload Menu_le�on15
    tempo = Dir(vpath & "Le�ons\" & nivoRep & "\le�on15A.txt")
    If tempo = "" Then
ML10:
        pagenum = 0
        msgtext0 = CRLF + msgNofic & vpath & "Le�ons\" & nivoRep & "\le�on15A.txt. " + perso_methode + pressez
        Msgform.Show 1
        If msgf = 2 Then GoTo ML10
        Menu_le�on15.Show 1
        Exit Sub
    End If
ML11:
    pagenum = 0
    msgtext0 = CRLF + pg15a0 + pressez
    Msgform.Show 1
    If msgf = 2 Then GoTo ML11
    If msgf = 1 Then
        Module_routines.placeinmsgtext0 "le�on15A.txt"
        vmsgbox = MsgBox(msgtext0, 1, msgDict�e & "15 A.")
        If vmsgbox = 2 Then
            keyinhibit = 1
            Menu_le�on15.Show 1
            Exit Sub
        End If
        exo_courant = "le�on15A.txt"
            
        ' Msg d'encouragements et d'explications
        Module_routines.resetmsg
         
        ' Go
        le�on14.Caption = bannerLe�on & " 15 A"  'Pas de point dans le titre, car sono avec ponctuation
        le�on14.Show 1
    End If
    If msgf = 0 Then Menu_le�on15.Show 1
End If

'****************  EXERCICE 15B ********************************************
If list1.ListIndex = 1 Then
    numexo = 1
    Unload Menu_le�on15
    tempo = Dir(vpath & "Le�ons\" & nivoRep & "\le�on15B.txt")
    If tempo = "" Then
ML20:
        pagenum = 0
        msgtext0 = CRLF + msgNofic & vpath & "Le�ons\" & nivoRep & "\le�on15B.txt. " + perso_methode + pressez
        Msgform.Show 1
        If msgf = 2 Then GoTo ML20
        Menu_le�on15.Show 1
        Exit Sub
    End If
ML21:
    pagenum = 0
    msgtext0 = CRLF + pg15b0 + pressez
    Msgform.Show 1
    If msgf = 2 Then GoTo ML21
    If msgf = 1 Then
        Module_routines.placeinmsgtext0 "le�on15B.txt"
        vmsgbox = MsgBox(msgtext0, 1, msgDict�e & "15 B.")
        If vmsgbox = 2 Then
            keyinhibit = 1
            Menu_le�on15.Show 1
            Exit Sub
        End If
        exo_courant = "le�on15B.txt"
                        
        ' Msg d'encouragements et d'explications
        Module_routines.resetmsg
            
        ' Go
        le�on14.Caption = bannerLe�on & " 15 B"  'Pas de point dans le titre, car sono avec ponctuation
        le�on14.Show 1
    End If
    If msgf = 0 Then Menu_le�on15.Show 1
End If

'****************  EXERCICE 15C ********************************************
If list1.ListIndex = 2 Then
    numexo = 2
    Unload Menu_le�on15
    tempo = Dir(vpath & "Le�ons\" & nivoRep & "\le�on15C.txt")
    If tempo = "" Then
ML30:
        pagenum = 0
        msgtext0 = CRLF + msgNofic & vpath & "Le�ons\" & nivoRep & "\le�on15C.txt. " + perso_methode + pressez
        Msgform.Show 1
        If msgf = 2 Then GoTo ML30
        Menu_le�on15.Show 1
        Exit Sub
    End If
ML31:
    pagenum = 0
    msgtext0 = CRLF + pg15c0 + pressez
    Msgform.Show 1
    If msgf = 2 Then GoTo ML31
    If msgf = 1 Then
        Module_routines.placeinmsgtext0 "le�on15C.txt"
        vmsgbox = MsgBox(msgtext0, 1, msgDict�e & "15 C.")
        If vmsgbox = 2 Then
            keyinhibit = 1
            Menu_le�on15.Show 1
            Exit Sub
        End If
        exo_courant = "le�on15C.txt"
                        
        ' Msg d'encouragements et d'explications
        Module_routines.resetmsg
            
        ' Go
        le�on14.Caption = bannerLe�on & " 15 C"  'Pas de point dans le titre, car sono avec ponctuation
        le�on14.Show 1
    End If
    If msgf = 0 Then Menu_le�on15.Show 1
End If

' ********************  end of sub  ******************************************
End Sub


' ******************** LIST1_KEYDOWN **********************************************
Private Sub list1_KeyDown(Keycode As Integer, shift As Integer)
' Win 91 et Win 92 (voir en plus Text2_KeyUp)
If Keycode = 91 Or Keycode = 92 Then Module_routines.cancelwin 0, Menu_le�on15, 1
End Sub


' ******************** LIST1_KEYUP **********************************************
Private Sub List1_KeyUp(Keycode As Integer, shift As Integer)
' Win 91 et Win 92 (voir en plus Text2_KeyDown)
If Keycode = 91 Or Keycode = 92 Then Module_routines.cancelwin 0, Menu_le�on15, 1
' Echappement
If Keycode = 27 Then
    If echapbis >= 0 Then
        If keyinhibit = 0 Then Quitter_Click
    Else
        echapbis = echapbis + 1
    End If
End If

' Entr�e
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

' Les msgbox procurent des key-ups ind�sirables avec les 3 commandes Entr�e Non Oui
If keyinhibit = 1 Then keyinhibit = 0
End Sub


'*****************  LIST1_KEYPRESS  ***********************************************
Private Sub List1_KeyPress(KeyAscii As Integer)
Module_routines.SetKeys "CAPSLOCK_OFF"
Module_routines.SetKeys "NUMLOCK_OFF"
Module_routines.SetKeys "SCROLLLOCK_OFF"
echapbis = 0  'Reset apr�s appel menu Options

' Pour sonoriser en r�p�tant la ligne menu en cours
If KeyAscii = 32 Then Module_routines.menu_repeat
End Sub


'********************  QUITTER  *********************************************
Private Sub Quitter_Click()
Unload Menu_le�on15
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

Private Sub Personnalis�_Click()
nivo = msgPersonnalis�
nivoRep = "Personnalis�" 'immuable, ne pas traduire
numexo = list1.ListIndex
Unload Menu_le�on15
MsgBox msgLevelIs & nivo & ".", 0, debexplivalue & nivo
keyinhibit = 1
Menu_le�on15.Show 1
End Sub

Private Sub Standard_Click()
nivo = msgStandard
nivoRep = "Standard" 'immuable, ne pas traduire
numexo = list1.ListIndex
Unload Menu_le�on15
MsgBox msgLevelIs & nivo & ".", 0, debexplivalue & nivo
keyinhibit = 1
Menu_le�on15.Show 1
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

Private Sub BipDiff�rent_Click()
keyinhibit = 1
Module_routines.BipDiff�rent
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

Private Sub AideG�n�rale_Click()
Module_routines.help_f1m
End Sub

Private Sub AideM�moire_Click()
Module_routines.help_f3m
End Sub

Public Sub Enseignant_Click()
Module_routines.placeinmsgaide "\Le�ons\Personnalis�\info.txt"
keyinhibit = 1
End Sub

Public Sub Sonorisation_Click()
Module_routines.placeinmsgaide "sonorisation.txt"
keyinhibit = 1
End Sub

Private Sub Aproposde_Click()
Menu_principal.Aproposde_Click
End Sub

