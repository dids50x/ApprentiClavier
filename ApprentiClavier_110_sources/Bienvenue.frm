VERSION 4.00
Begin VB.Form Bienvenue 
   BackColor       =   &H00808080&
   Caption         =   "Bienvenue dans ApprentiClavier"
   ClientHeight    =   6090
   ClientLeft      =   585
   ClientTop       =   1920
   ClientWidth     =   10425
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   Height          =   6930
   KeyPreview      =   -1  'True
   Left            =   465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   10425
   Top             =   1200
   Width           =   10665
   Begin VB.Timer Timer2 
      Interval        =   12000
      Left            =   8760
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   120
      Picture         =   "Bienvenue.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      HideSelection   =   0   'False
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3720
      Width           =   6615
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8160
      Top             =   0
   End
   Begin VB.TextBox Text0 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      HideSelection   =   0   'False
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   8265
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
      Height          =   1335
      Left            =   7920
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      HideSelection   =   0   'False
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   2700
      Width           =   6615
   End
   Begin VB.Line Line13 
      BorderWidth     =   2
      X1              =   7920
      X2              =   7680
      Y1              =   3960
      Y2              =   3720
   End
   Begin VB.Line Line12 
      BorderWidth     =   2
      X1              =   7920
      X2              =   8040
      Y1              =   3960
      Y2              =   3600
   End
   Begin VB.Line Line11 
      X1              =   7920
      X2              =   7680
      Y1              =   3960
      Y2              =   3840
   End
   Begin VB.Line Line10 
      X1              =   7920
      X2              =   8040
      Y1              =   3960
      Y2              =   3720
   End
   Begin VB.Line Line9 
      BorderWidth     =   2
      X1              =   7680
      X2              =   7920
      Y1              =   3360
      Y2              =   3960
   End
   Begin VB.Line Line8 
      X1              =   7680
      X2              =   7560
      Y1              =   3360
      Y2              =   3480
   End
   Begin VB.Line Line7 
      BorderWidth     =   2
      X1              =   7560
      X2              =   7560
      Y1              =   3360
      Y2              =   3480
   End
   Begin VB.Line Line6 
      X1              =   7560
      X2              =   7440
      Y1              =   3240
      Y2              =   3120
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   7560
      X2              =   7440
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Line Line4 
      BorderWidth     =   2
      X1              =   7800
      X2              =   7920
      Y1              =   3360
      Y2              =   3240
   End
   Begin VB.Line Line3 
      X1              =   7800
      X2              =   7920
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   7800
      X2              =   7800
      Y1              =   3120
      Y2              =   3000
   End
   Begin VB.Line Line1 
      X1              =   7680
      X2              =   7800
      Y1              =   3120
      Y2              =   3000
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   3
      FillColor       =   &H00FF80FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   7560
      Shape           =   3  'Circle
      Top             =   3120
      Width           =   255
   End
   Begin VB.Label Label4 
      BackColor       =   &H00808080&
      Caption         =   "Label4"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   9
      Top             =   4440
      Width           =   7455
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
      ForeColor       =   &H00008000&
      Height          =   345
      Left            =   240
      TabIndex        =   7
      Top             =   5400
      Width           =   7455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Par sécurité, une deuxième fois, TAPEZ votre NOM :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   960
      TabIndex        =   4
      Top             =   3240
      Width           =   6855
   End
   Begin VB.Label Label0 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Il faudra souvent appuyer sur la touche Entrée :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   960
      TabIndex        =   0
      Top             =   150
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tapez votre NOM, ou tapez simplement sur Entrée :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   960
      TabIndex        =   2
      Top             =   2250
      Width           =   6855
   End
   Begin VB.Menu Fichier 
      Caption         =   "&Fichier"
      Begin VB.Menu Quitter_bm 
         Caption         =   "&Quitter"
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
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu Reset 
         Caption         =   "Redémarrer à la prem&ière leçon"
      End
      Begin VB.Menu Sep4 
         Caption         =   "-"
      End
      Begin VB.Menu Aproposde 
         Caption         =   "A &Propos de"
      End
   End
End
Attribute VB_Name = "Bienvenue"
Attribute VB_Creatable = False
Attribute VB_Exposed = False

' ********************  LOAD  *******************************
Private Sub Form_Load()
If App.PrevInstance = True Then End
If FullScreenSwitch = 1 Then WindowState = 2
Set menu_courant = Bienvenue
Bienvenue.Caption = msgBienvenue & bannerVersion
Bienvenue.Label0.Caption = msgBienvFaudra
Bienvenue.Quitter.Caption = msgQuitter & msgÉchap
Bienvenue.Label4.Font.Size = 9 'avril 2008

' Détecter la résolution d'écran et redimensionner les fenêtres,
' le zoom zfactor est défini par la largeur de la première fenêtre bienvenue !
' commentaire 12/ 2011 : le zoomfactor et les options user ne seront définis que lorsque le user sera connu, donc à l'affichage du menu principal
Module_routines.zoomform Bienvenue
Module_routines.Dimension Bienvenue

' Font.Size des caras de msgform selon définition écran
fsize = fsizedefault * zfactor

' Dessin de la fleur à la bonne échelle
Module_routines.dimobject Bienvenue.Shape1
Module_routines.dimobject Bienvenue.Line1
Module_routines.dimobject Bienvenue.Line2
Module_routines.dimobject Bienvenue.Line3
Module_routines.dimobject Bienvenue.Line4
Module_routines.dimobject Bienvenue.Line5
Module_routines.dimobject Bienvenue.Line6
Module_routines.dimobject Bienvenue.Line7
Module_routines.dimobject Bienvenue.Line8
Module_routines.dimobject Bienvenue.Line9
Module_routines.dimobject Bienvenue.Line10
Module_routines.dimobject Bienvenue.Line11
Module_routines.dimobject Bienvenue.Line12
Module_routines.dimobject Bienvenue.Line13

' Préparer l'apparition des infos dans la fenêtre
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
text1.Visible = False
label1.Caption = ""
text2.Visible = False
Label2.Caption = ""
Label3.Caption = bannerVersion & ", " & bannerCopyright
If svnames = "" Then Label4.Caption = msgSonori & msgNoSono & CRLF2 & msgKeyboard & clavierType & ". " & country
If svnames <> "" Then Label4.Caption = msgSonori & svnames & CRLF2 & msgKeyboard & clavierType & ". " & country
'If repjawsnames = "" Then Label4.Caption = msgSonori & msgNoSono & CRLF2 & msgKeyboard & clavierType & ". " & country
'If repjawsnames <> "" Then Label4.Caption = msgSonori & repjawsnames & CRLF2 & msgKeyboard & clavierType & ". " & country
keyinhibit = 1
timeover = 0
Module_routines.MenuEditorTrans Bienvenue
End Sub


Private Sub Timer1_Timer()
Quitter.TabStop = True
On Error Resume Next
Text0.SetFocus
keyinhibit = 0

' Faire apparaître ce message après des délais timer1 pour meilleure sono
If timein = 4 Then
    If timeover = 0 Then Text0.Text = CRLF + msgPressEnter
    If timeover = 1 Then Text0.Text = msgEnter + CRLF + msgPressEnter
    Text0.SelStart = 0
    Text0.SelLength = Len(Text0.Text)
    Timer1.Enabled = False
    Timer2.Enabled = True
End If
timein = timein + 1
End Sub


Private Sub Timer2_Timer()
If timein >= 4 Then timeover = 1 'signifie que l'utilisateur est perplexe devant la demande d'appuyer sur Entrée
timein = 0
Label0.Caption = ""
Label0.BackColor = Bienvenue.BackColor
Text0.SelStart = 0
Text0.SelLength = 0
Timer1.Enabled = True
Timer2.Enabled = False
End Sub


' *******************  TEXT0_KEYDOWN  *********************************
Private Sub text0_keydown(Keycode As Integer, shift As Integer)
'Debug.Print "keydown1=" & Keycode & "  winstop=" & winstop & "  keyinh=" & keyinhibit
' Win 91 et Win 92 (voir aussi text0_KeyUp)
If Keycode = 91 Or Keycode = 92 Then
    ' Winstop = 1 stoppe Win ou Win+E, Win+F, Win+L, Win+R, Win+U juin 2007
    winstop = 1
    'Module_routines.cancelwin 0, Bienvenue, 1   remplacé par un Escape juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    Exit Sub
End If

' Touche Tab incorrecte (non reconnue par KeyUp)
If Keycode = 9 Then Module_routines.pasdetab
End Sub


' *******************  TEXT0_KEYUP  *********************************
Private Sub text0_keyup(Keycode As Integer, shift As Integer)
' Touche Tab refusée (placer avant keyinhibit)
If keyforce = 9 Then
    Module_routines.bip Bienvenue
    Text0.Text = msgEnter + CRLF + msgPressEnter
    Text0.SelStart = 0
    Text0.SelLength = Len(Text0.Text)
    keyforce = 0: keyinhibit = 0
    Exit Sub
End If

' Si on vient d'un msgbox
If keyinhibit = 1 Then
    keyinhibit = 0
    Exit Sub
End If

' Winstop stoppe Win ou Win+E, Win+F, Win+L, Win+R, Win+U
If winstop > 0 Then
    winstop = winstop - 1
    Exit Sub
End If

' Win 91 et Win 92 (voir aussi text0_KeyDown)
If Keycode = 91 Or Keycode = 92 Then
    Module_routines.cancelwin 0, Bienvenue, 1
    Exit Sub
End If

' Menu-Contextuel 93
If Keycode = 93 Then
    keyinhibit = 1
    echapbis = 0
    'SendKeys "{ESC}"
    'Sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    Module_routines.bip Bienvenue
    Exit Sub
End If

'Touche Alt+F4 pour quitter
If Keycode = 115 And shift = 4 Then
    altf4 = 1
    If quitactive = 0 Then Module_routines.AuRevoir
End If

'Touche F10 mène à la barre menu, il faut échapper
If Keycode = 121 Then
    keyinhibit = 1
    'SendKeys "{ESC}", True
    'Sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
End If

' Échappement (5 fois Échap font quitter même sans confirm)
If Keycode = 27 And keyinhibit = 0 Then
    echapbis = echapbis + 1
    If echapbis > 5 Then End
    Unload Bienvenue
BV10:
    ' Changer les couleurs pour avertir qu'on quitte
    msgtext0 = pressez_quit
    fsize = 1.5 * fsizedefault * zfactor
    fbc = fbc_quit
    ffc = ffc_quit
    Msgform.Quitter.Caption = msgQuitter & msgÉchap
    pagenum = 0
    Msgform.Show 1
    ffc = ffc_default
    fbc = fbc_default
    fsize = fsizedefault * zfactor
    If msgf = 2 Then GoTo BV10
    If msgf = 0 Then Quitter_Click
    Bienvenue.Show 1
End If

' Bonne réponse vvEntrée
If Keycode = 13 Then
    Timer1.Enabled = False
    Timer2.Enabled = False
    
    ' Dessin de la fleur supprimé
    Shape1.Visible = False
    Line1.Visible = False
    Line2.Visible = False
    Line3.Visible = False
    Line4.Visible = False
    Line5.Visible = False
    Line6.Visible = False
    Line7.Visible = False
    Line8.Visible = False
    Line9.Visible = False
    Line10.Visible = False
    Line11.Visible = False
    Line12.Visible = False
    Line13.Visible = False
    
    ' Suite
    Label0.Caption = ""
    Label0.BackColor = Bienvenue.BackColor
    Picture1.Visible = False
    Text0.Visible = False
    text1.Visible = True
    label1.Caption = msgBienvUsername
    echapbis = 0
    old = vvSansNom
    text1.Text = ""
    text1.SelStart = 0
    text1.SelLength = 0
    On Error Resume Next
    text1.SetFocus
    Exit Sub
End If

' Mauvaises réponses au lieu de vvEntrée
If Keycode <> 27 Then
    ' Mauvaises réponses réitérées 7 fois font Quitter
    echapbis = echapbis + 1
    If echapbis > 8 Then End
    Text0.SelStart = 0
    Text0.SelLength = 0  '1 si pas d'écho clavier
    On Error Resume Next
    Text0.SetFocus
    Module_routines.bip Bienvenue
    Call Sleep(cadenceligne)
    Bienvenue.Cls
    If echapbis > 7 Then
        Text0.Text = msgRelaunch + CRLF + msgAurevoir
    Else
        Timer1.Enabled = False
        Timer2.Enabled = False
        Text0.Text = msgEnter + CRLF + msgPressEnter
    End If
    Text0.SelStart = 0
    Text0.SelLength = Len(Text0.Text)
    On Error Resume Next
    Text0.SetFocus
    Call Sleep(3 * cadenceligne)
    Text0.SelStart = 0
    Text0.SelLength = 0
    Exit Sub
End If
End Sub


' *******************  TEXT1_KEYDOWN  *********************************
Private Sub text1_KeyDown(Keycode As Integer, shift As Integer)
' Win 91 et Win 92 (voir aussi text1_KeyUp)
If Keycode = 91 Or Keycode = 92 Then
    ' Winstop = 1 stoppe Win ou Win+E, Win+F, Win+L, Win+R, Win+U juin 2007
    winstop = 1
    'Module_routines.cancelwin 0, Bienvenue, 1    remplacé par un Escape juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    Exit Sub
End If

' Touche Tab incorrecte (non reconnue par KeyUp)
If Keycode = 9 Then Module_routines.pasdetab
End Sub

' *********************  TEXT1_KEYUP  *********************************
Private Sub text1_KeyUp(Keycode As Integer, shift As Integer)
' Touche Tab refusée (placer avant keyinhibit)
If keyforce = 9 Then
    Module_routines.bip Bienvenue
    text1.Text = old: text1.SelStart = Len(text1.Text)
    keyforce = 0: keyinhibit = 0
    Exit Sub
End If

' Wintop stoppe Win ou Win+E
If winstop > 0 Then
    winstop = winstop - 1
    Exit Sub
End If

' Si on vient d'un msgbox
If keyinhibit = 1 Then
    keyinhibit = 0
    Exit Sub
End If

' Échappement
If Keycode = 27 Then
    If echapbis >= 0 Then
        echapbis = echapbis + 1
BV11:
        msgtext0 = pressez_quit
        fsize = 1.5 * fsizedefault * zfactor
        fbc = fbc_quit
        ffc = ffc_quit
        Msgform.Quitter.Caption = msgQuitter & msgÉchap
        pagenum = 0
        Msgform.Show 1
        ffc = ffc_default
        fbc = fbc_default
        fsize = fsizedefault * zfactor
        If msgf = 2 Then GoTo BV11
        If msgf = 0 Then Quitter_Click
        keyinhibit = 0
        old = vvSansNom
        text1.Text = ""
        text1.SelStart = 0
        text1.SelLength = 0
        On Error Resume Next
        text1.SetFocus
    Else
        echapbis = echapbis + 1
    End If
    Exit Sub
End If

' Win 91 et Win 92
If Keycode = 91 Or Keycode = 92 Then
    Module_routines.cancelwin 0, Bienvenue, 1
    Exit Sub
End If

' Menu-Contextuel 93
If Keycode = 93 Then
    keyinhibit = 1
    echapbis = 0
    'SendKeys "{ESC}"
    'Sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    Module_routines.bip Bienvenue
    Exit Sub
End If

'Touche Alt+F4 pour quitter
If Keycode = 115 And shift = 4 Then
    altf4 = 1
    Module_routines.AuRevoir
End If

'Touche F10 mène à la barre menu
If Keycode = 121 Then
    echapbis = echapbis - 1
    'SendKeys "{ESC}", True
    'Sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
End If

' Touche RetArr
If Keycode = 8 Then
    text1.Text = Left(text1.Text, Len(text1.Text))
    On Error Resume Next
    old = Left(old, Len(old) - 1)
    text1.SelStart = Len(text1.Text)
    Exit Sub
End If

' 1er cara ESPACE invalide, fait RÉPÉTER
If Mid(text1.Text, 1, 1) = " " Then
    Module_routines.bip Bienvenue
    text1.Text = msgBienvRetape
    text1.SelStart = 0
    text1.SelLength = Len(text1.Text)
    On Error Resume Next
    text1.SetFocus
    Call Sleep(2000)
    old = vvSansNom
    text1.Text = ""
    text1.SelStart = 0
    text1.SelLength = 0
    On Error Resume Next
    text1.SetFocus
    Exit Sub
End If

' 1ère VALIDATION du nom utilisateur
If Keycode = 13 Then
    If UCase(old) = UCase(vvSansNom) Or old = "" Then
        nom = vvSansNom
        Module_routines.data_user
        Exit Sub
    End If
    If text1.Text = "" Then Exit Sub
    
    'Check for valid path name
    If InStr(old, "<") Or InStr(old, ">") Or InStr(old, "?") Or InStr(old, ",") Or InStr(old, ".") Or InStr(old, ";") Or InStr(old, "/") Or InStr(old, ":") Or InStr(old, "§") Or InStr(old, "!") Or InStr(old, "*") Or InStr(old, "µ") Or InStr(old, "%") Or InStr(old, "&") Or InStr(old, """") Or InStr(old, "'") Or InStr(old, "(") Or InStr(old, ")") Or InStr(old, "°") Or InStr(old, "=") Or InStr(old, "+") Or InStr(old, "²") Then
        ' ERREUR sur la saisie du NOM
        text1.Text = msgBienvRedo
        text1.SelStart = 0
        text1.SelLength = Len(text1.Text)
        On Error Resume Next
        text1.SetFocus
        Call Sleep(2000)
        Unload Bienvenue
        Bienvenue.Show 1
    End If
    
    ' Si le nom existe déjà dans les utilisateurs, on ne le redemande pas
    If Dir(vpath & "Utilisateurs\" & old & "\pctok.txt") <> "" Then
        nom = old
        Module_routines.data_user
        Exit Sub
    End If
    
    ' Si le nom n'existe pas encore, on le redemande pour assurer l'orthographe
    nom_temp = old
    old = ""
    text1.Text = ""
    text1.Visible = False
    text2.Visible = True
    label1.Caption = ""
    Label2.Caption = msgBienvRepeat
    On Error Resume Next
    text2.SetFocus
End If

' Touches refusées Ctrl, Maj, Alt, VerrMaj... lancement du logiciel par un raccourci
If shift > 1 Then
    Module_routines.bip Bienvenue
    On Error Resume Next
    text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    text1.SelStart = Len(text1.Text)
    Exit Sub
End If

' Touches de faible code
If Keycode < 27 Then Exit Sub

' Touches refusées telles que le & 49, quote 52, ( 53, souligné 56, égal 187, ? et virgule 188, . et ; 190, slash 191, ° 219, * 220, ! 223, inférieur 226
If Keycode = 49 Or Keycode = 51 Or Keycode = 52 Or Keycode = 53 Or Keycode = 56 Or Keycode = 187 Or Keycode = 188 Or Keycode = 190 Or Keycode = 191 Or Keycode = 192 Or Keycode = 219 Or Keycode = 220 Or Keycode > 221 Then
    Module_routines.bip Bienvenue
    On Error Resume Next
    text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    text1.SelStart = Len(text1.Text)
    Exit Sub
End If

' Touches refusées avant la touche espace
' Touches refusées (Flèches) avant le a grave 48, le é 50, le tiret 54, le è 55
If (Keycode > 27 And Keycode < 32) Or (Keycode > 32 And Keycode < 48) Then Exit Sub

' Touches refusées après le ç 57, avant le A ou a 65
' Touches refusées au-delà de z 90, avant le u grave 192
' Touches refusées au-delà de u grave 192, avant le tréma ou circonflexe 221
' Touche refusée %
' Touches refusées au-delà de tréma ou circonflexe 221
' Touches refusées chiffres 0 à 9
If (Keycode > 27 And Keycode < 32) Or (Keycode > 32 And Keycode < 48) Or (shift = 1 And Keycode > 47 And Keycode < 61) Or (Keycode > 57 And Keycode < 65) Or (Keycode > 90 And Keycode < 192) Or (Keycode = 192 And shift = 1) Or (Keycode > 192 And Keycode < 221) Then
    Module_routines.bip Bienvenue
    On Error Resume Next
    text1.Text = Left(text1.Text, Len(text1.Text) - 1)
    text1.SelStart = Len(text1.Text)
    Exit Sub
End If

' Autres touches : Echo clavier [script setjcfoption(OPT_TYPING_ECHO,1)incompatible Win 98]
tempo = Right(text1.Text, 1)
'If UCase(old) = vvSansNom Then text1.Text = tempo
'If UCase(old) <> vvSansNom Then text1.Text = old & tempo

' Troncature du nom à 26 caras
If Len(old) > 26 Then
    Beep
    text1.Text = old
End If
old = text1.Text
text1.SelStart = Len(text1.Text)
text1.SelLength = 0
On Error Resume Next
text1.SetFocus
End Sub


' *******************  TEXT2_KEYDOWN  *********************************
Private Sub Text2_KeyDown(Keycode As Integer, shift As Integer)
' Win 91 et Win 92 (voir aussi text2_KeyUp)
If Keycode = 91 Or Keycode = 92 Then
    ' Winstop = 1 stoppe Win ou Win+E, Win+F, Win+L, Win+R, Win+U juin 2007
    winstop = 1
    'Module_routines.cancelwin 0, Bienvenue, 1    remplacé par un Escape juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    Exit Sub
End If

' Touche Tab incorrecte (non reconnue par KeyUp)
If Keycode = 9 Then Module_routines.pasdetab
End Sub


' *********************  TEXT2_KEYUP  *********************************
Private Sub Text2_KeyUp(Keycode As Integer, shift As Integer)
' Touche Tab refusée
If keyforce = 9 Then
    Module_routines.bip Bienvenue
    text2.Text = old: text2.SelStart = Len(text2.Text)
    keyforce = 0: keyinhibit = 0
    Exit Sub
End If

' Wintop stoppe Win ou Win+E
If winstop > 0 Then
    winstop = winstop - 1
    Exit Sub
End If

' Si on vient d'un msgbox
If keyinhibit = 1 Then
    keyinhibit = 0
    Exit Sub
End If

' Win 91 et Win 92
If Keycode = 91 Or Keycode = 92 Then
    Module_routines.cancelwin 0, Bienvenue, 1
    Exit Sub
End If

' Menu-Contextuel 93
If Keycode = 93 Then
    keyinhibit = 1
    echapbis = 0
    'SendKeys "{ESC}"
    'Sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    Module_routines.bip Bienvenue
    Exit Sub
End If

'Touche Alt+F4 pour quitter
If Keycode = 115 And shift = 4 Then
    altf4 = 1
    Module_routines.AuRevoir
End If

'Touche F10 mène à la barre menu
If Keycode = 121 Then
    echapbis = echapbis - 1
    'SendKeys "{ESC}", True
    'Sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
End If

' Échappement
If Keycode = 27 And keyinhibit = 0 Then
    If echapbis >= 0 Then
        echapbis = echapbis + 1
        If echapbis > 8 Then End
BV12:
        msgtext0 = pressez_quit
        fsize = 1.5 * fsizedefault * zfactor
        fbc = fbc_quit
        ffc = ffc_quit
        Msgform.Quitter.Caption = msgQuitter & msgÉchap
        pagenum = 0
        Msgform.Show 1
        ffc = ffc_default
        fbc = fbc_default
        fsize = fsizedefault * zfactor
        If msgf = 2 Then GoTo BV12
        If msgf = 0 Then Quitter_Click
        keyinhibit = 0
        old = vvSansNom
        text2.Text = ""
        text2.SelStart = 0
        text2.SelLength = 0
        On Error Resume Next
        text2.SetFocus
    Else
        echapbis = echapbis + 1
    End If
    Exit Sub
End If

' Touche RetArr
If Keycode = 8 Then
    text2.Text = Left(text2.Text, Len(text2.Text))
    On Error Resume Next
    old = Left(old, Len(old) - 1)
    text2.SelStart = Len(text2.Text)
    Exit Sub
End If

' 1er cara ESPACE invalide, fait RÉPÉTER
If Mid(text2.Text, 1, 1) = " " Then
    Module_routines.bip Bienvenue
    text2.Text = msgBienvRep
    text2.SelStart = 0
    text2.SelLength = Len(text2.Text)
    On Error Resume Next
    text2.SetFocus
    Call Sleep(2000)
    old = ""
    text2.Text = ""
    text2.SelStart = 0
    text2.SelLength = 0
    On Error Resume Next
    text2.SetFocus
    Exit Sub
End If

' 2ème VALIDATION du nom utilisateur
If Keycode = 13 Then
    If UCase(old) = vvSansNom Then
        nom = vvSansNom
        Module_routines.data_user
        Exit Sub
    End If
    If text2.Text = "" Then Exit Sub
    If UCase(old) = UCase(nom_temp) Then
        nom = nom_temp
        Module_routines.data_user
    Else
        ' ERREUR sur la saisie du NOM
        text2.Text = msgBienvRedo
        text2.SelStart = 0
        text2.SelLength = Len(text2.Text)
        On Error Resume Next
        text2.SetFocus
        Call Sleep(2000)
        Unload Bienvenue
        Bienvenue.Show 1
    End If
    Exit Sub
End If

' Touches refusées Ctrl, Maj, Alt, VerrMaj... lancement du logiciel par un raccourci
If shift > 1 Then
    Module_routines.bip Bienvenue
    On Error Resume Next
    text2.Text = Left(text2.Text, Len(text2.Text) - 1)
    text2.SelStart = Len(text2.Text)
    Exit Sub
End If

' Touches de faible code
If Keycode < 27 Then Exit Sub

' Touches refusées telles que le & 49, quote 52, ( 53, souligné 56, égal 187, ? et virgule 188, . et ; 190, slash 191, ° 219, * 220, ! 223, inférieur 226
If Keycode = 49 Or Keycode = 51 Or Keycode = 52 Or Keycode = 53 Or Keycode = 56 Or Keycode = 187 Or Keycode = 188 Or Keycode = 190 Or Keycode = 191 Or Keycode = 192 Or Keycode = 219 Or Keycode = 220 Or Keycode > 221 Then
    Module_routines.bip Bienvenue
    On Error Resume Next
    text2.Text = Left(text2.Text, Len(text2.Text) - 1)
    text2.SelStart = Len(text2.Text)
    Exit Sub
End If

' Touches refusées avant la touche espace
' Touches refusées (Flèches) avant le a grave 48, le é 50, le tiret 54, le è 55
If (Keycode > 27 And Keycode < 32) Or (Keycode > 32 And Keycode < 48) Then Exit Sub

' Touches refusées après le ç 57, avant le A ou a 65
' Touches refusées au-delà de z 90, avant le u grave 192
' Touches refusées au-delà de u grave 192, avant le tréma ou circonflexe 221
' Touche refusée %
' Touches refusées au-delà de tréma ou circonflexe 221
' Touches refusées chiffres 0 à 9
If (Keycode > 27 And Keycode < 32) Or (Keycode > 32 And Keycode < 48) Or (shift = 1 And Keycode > 47 And Keycode < 61) Or (Keycode > 57 And Keycode < 65) Or (Keycode > 90 And Keycode < 192) Or (Keycode = 192 And shift = 1) Or (Keycode > 192 And Keycode < 221) Then
    Module_routines.bip Bienvenue
    On Error Resume Next
    text2.Text = Left(text2.Text, Len(text2.Text) - 1)
    text2.SelStart = Len(text2.Text)
    Exit Sub
End If


' Autres touches : Echo clavier [script setjcfoption(OPT_TYPING_ECHO,1)incompatible Win 98]
tempo = Right(text2.Text, 1)
'If UCase(old) = vvSansNom Then text2.Text = tempo
'If UCase(old) <> vvSansNom Then text2.Text = old & tempo

' Troncature du nom à 26 caras
If Len(old) > 26 Then
    Beep
    text2.Text = old
End If
old = text2.Text
text2.SelStart = Len(text2.Text)
text2.SelLength = 0
On Error Resume Next
text2.SetFocus
End Sub


' *******************  QUITTER  ***********************************
Private Sub Quitter_Click()
'Module_routines.restore_locks
Module_routines.AuRevoir
End Sub

Private Sub Quitter_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Quitter_Click
If KeyAscii = 81 Or KeyAscii = 113 Then Quitter_Click
End Sub


' **************  COMMANDES de la BARRE de MENU  *******************
Private Sub Quitter_bm_Click()
'Module_routines.restore_locks
Module_routines.AuRevoir
End Sub

Private Sub Fichier_Click()
keyinhibit = 1
End Sub

Private Sub Aide_Click()
keyinhibit = 1
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

Private Sub Reset_Click()
vmsgbox = MsgBox(msgReset, 0, msgResetTitle)
End Sub

Private Sub Aproposde_Click()
Menu_principal.Aproposde_Click
End Sub

