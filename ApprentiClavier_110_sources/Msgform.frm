VERSION 4.00
Begin VB.Form Msgform 
   BackColor       =   &H00E0E0E0&
   Caption         =   " "
   ClientHeight    =   6495
   ClientLeft      =   630
   ClientTop       =   1470
   ClientWidth     =   9975
   ControlBox      =   0   'False
   Height          =   7005
   KeyPreview      =   -1  'True
   Left            =   570
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   9975
   Top             =   1020
   Width           =   10095
   Begin VB.CommandButton Suivant 
      Caption         =   "&Suivant"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1800
      TabIndex        =   4
      Top             =   5640
      Width           =   1365
   End
   Begin VB.CommandButton Précédent 
      Caption         =   "&Précédent"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   150
      TabIndex        =   3
      Top             =   5640
      Width           =   1515
   End
   Begin VB.CommandButton Continuer 
      Caption         =   "&Continuer (Entrée)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   3900
      TabIndex        =   1
      Top             =   5640
      Width           =   2565
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   6600
      Top             =   6000
   End
   Begin VB.TextBox Text0 
      BackColor       =   &H00C00000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   5535
      HideSelection   =   0   'False
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   9375
   End
   Begin VB.CommandButton quitter 
      Caption         =   "&Quitter (Échap)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   7680
      TabIndex        =   2
      Top             =   5640
      Width           =   2115
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   6600
      Top             =   5520
   End
End
Attribute VB_Name = "Msgform"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' ********************  LOAD  *******************************
Private Sub Form_Load()
If FullScreenSwitch = 1 And fullscreeninhibit = 0 Then WindowState = 2
Msgform.Quitter.Caption = msgQuitter & msgÉchap
Msgform.Continuer.Caption = msgContinuer
Msgform.Précédent.Caption = msgPrécédent
Msgform.Suivant.Caption = msgSuivant

'Paramètres
msgf = 9
If keyinhibit <> 2 Then keyinhibit = 1
echapbis = 0
Module_routines.Colors Msgform  '12/2011
Module_routines.Dimension Msgform
If pagenum > 0 Then
    Msgform.Caption = debexplivalue & msgPage & pagenum & "."   'pour jaws jss
    If pagenum = 1 Then Précédent.Enabled = False
    If pagemax = 1 Then Suivant.Enabled = False
End If
If pagenum = 0 Then
    Msgform.Caption = debexplivalue     'pour jaws jss
    Précédent.Visible = False
    Suivant.Visible = False
End If
'12/2011
'Text0.BackColor = fbc
'Text0.ForeColor = ffc
firstmove = 0
End Sub

Private Sub Timer1_Timer()
Text0.Text = msgtext0
Text0.Font.Size = fsize * zoomfactor  '12/2011
Text0.SelStart = 0
Text0.SelLength = Len(Text0.Text)
On Error Resume Next
Text0.SetFocus
Timer1.Enabled = False
If keyinhibit = 1 Then keyinhibit = 0
End Sub

Private Sub Timer2_Timer()
Text0.SelStart = 0
Text0.SelLength = 0
Timer2.Enabled = False
If timeout = 1 Then Unload Msgform
End Sub


' *******************  TEXT0_KEYdown  *********************************
Private Sub text0_keydown(Keycode As Integer, shift As Integer)
' Win 91 et Win 92
If Keycode = 91 Or Keycode = 92 Then
    msgpbmenu.Show 1
    Module_routines.cancelwin 0, Msgform, 0
    msgpbmenu.Show 1
    Exit Sub
End If

' F2 inutile (voir text0_keyup)
If quitF2 = 1 Then Exit Sub

' Échappement (complément de KeyUp)
If Keycode = 27 Then Exit Sub

' Entrée ou OUI pour continuer
If Keycode = 13 Then
    Exit Sub
End If

' Touche F1 Aide pour le score avec fautes
If Keycode = 112 And f1msgform = 1 Then
    msgf = 3
    Unload Msgform
    Exit Sub
End If

' Eviter message après Quitquit final
If quitactive = 1 Then Exit Sub

' Touche F1 Aide par dessus l'info F1
If Keycode = 112 And f1msgform = 0 Then
    Unload Msgform
    If typeleçon = 0 Then Module_routines.help_f1m
    Exit Sub
End If

' Touche F2 (voir text0_keyup)
If Keycode = 113 And typeleçon > 0 Then
    Unload Msgform
    Exit Sub
End If
If Keycode = 113 And typeleçon = 0 Then
    Text0.Text = ""
    Text0.SelStart = 0
    Text0.SelLength = 0
    Msgform.Cls
    Call Sleep(cadencecara)
    echapbis = echapbis + 1
    Text0.Font.Size = 1.2 * fsizedefault * zfactor
    noechapF1 = 1
    Text0.Text = pressez_F2 + msgFormPressez
    Text0.SelStart = 0
    Text0.SelLength = Len(Text0.Text)
    Call Sleep(cadenceligne)
    Text0.SelStart = 0
    Text0.SelLength = 0
    Exit Sub
End If

' Touches Maj, Control, Alt, VerrNum
If (Keycode > 15) And (Keycode < 24) Then
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
    Exit Sub
End If

' Quitte le pavé numérique avec NUMLOCK_OFF
If Keycode = 144 Then Exit Sub

' ESPACE pour répéter
If Keycode = 32 Then
    msgf = 2
End If

' PagePrécSuiv 33 et 34, DebFin 36 et 35, Flèches 37 à 40
If (Keycode > 32) And (Keycode < 41) Then
    Timer2.Enabled = False
    If firstmove = 0 Then Text0.SelStart = 0
    firstmove = firstmove + 1
    Exit Sub
End If

' F3 (voir keyup)
If Keycode = 114 Then Exit Sub

'Touche Alt+F4 pour quitter (voir text0_keyup)
If Keycode = 115 And shift = 4 Then Exit Sub

' AUTRES touches
If Keycode > 40 Or Keycode = 8 Then
    Text0.Text = ""
    Text0.SelStart = 0
    Text0.SelLength = 0
    Msgform.Cls
    Call Sleep(cadencecara)
    echapbis = echapbis + 1
    ' TROP d'ERREURS RÉPÉTÉES
    If echapbis < 15 Then
        Text0.Font.Size = 1.2 * fsizedefault * zfactor
        noechapF1 = 1
        Text0.Text = CRLF + msgFormVousétiez + CRLF2 + msgFormPressez
        Text0.SelStart = 0
        Text0.SelLength = Len(Text0.Text)
        Call Sleep(cadenceligne)
        Text0.SelStart = 0
        Text0.SelLength = 0
        Exit Sub
    Else
        echapbis = 0
        msgf = 0
        Text0.Text = msgFormRecommencer
        Text0.SelStart = 0
        Text0.SelLength = Len(Text0.Text)
        Call Sleep(2000)
        If nivo = msgStandard Then pctok(0, 0) = 0
        If nivo = msgPersonnalisé Then pctok(25, 0) = 0
    End If
End If
    
' Suite
    Unload Msgform
End Sub


' *******************  TEXT0_KEYup  *********************************
Private Sub text0_keyup(Keycode As Integer, shift As Integer)
' F2 inutile (voir text0_keyup)
If quitF2 = 1 Then
    quitF2 = 0
    Unload Msgform
    Exit Sub
End If

' Echappement
If Keycode = 27 And forcepause = 2 Then
    forcepause = 0
    Exit Sub
End If
If Keycode = 27 Then
    If echapbis < 0 Then
        echapbis = echapbis + 1
        Exit Sub
    Else
        msgf = 0
        pagemax = 0
        pagenum = 0
        Unload Msgform
        If typeleçon = 1 Then
            If noechapF1 = 0 Then
                Module_routines.quit_l
            End If
        End If
    End If
End If

' Entrée ou OUI pour continuer sauf si keyinhibit=2 venant de la barre des menus
If Keycode = 13 And keyinhibit < 2 Then
    msgf = 1
    pagemax = 0
    Unload Msgform
End If
If (Keycode = 13 Or Keycode = 112) And keyinhibit = 2 Then keyinhibit = 0

' Touche F2
If Keycode = 113 And f1msgform = 1 Then
    help_f2 leçon_courante
    Exit Sub
End If

' F3
If Keycode = 114 Then
    Unload Msgform
    Module_routines.help_f3m
    ' Pas de pagenum = 0 !
    Msgform.Show 1
End If

' PagePrécSuiv 33 et 34
If pagenum > 0 Then
    ' Page Précédente
    If Keycode = 33 Then
        msgf = 33
        pagemax = 0
        Unload Msgform
    End If
    ' Page Suivante
If Keycode = 34 Then
        msgf = 34
        Unload Msgform
    End If
End If

'Touche Alt+F4 pour quitter
If Keycode = 115 And shift = 4 Then
    altf4 = 1
    If quitactive = 0 Then Module_routines.QuitQuit
End If
End Sub

' *******************  Boutons CONTINUER, QUITTER  ***********************************
Private Sub Quitter_Click()
    msgf = 0
    pagemax = 0
    Unload Msgform
End Sub

Private Sub Quitter_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    On Error Resume Next
    Text0.SetFocus
End If
If KeyAscii = 13 Then Quitter_Click
End Sub

Private Sub continuer_click()
    msgf = 1
    pagemax = 0
    Unload Msgform
End Sub

Private Sub continuer_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    On Error Resume Next
    Text0.SetFocus
End If
If KeyAscii = 13 Then continuer_click
End Sub

Private Sub Précédent_Click()
    msgf = 33
    pagemax = 0
    Unload Msgform
End Sub

Private Sub Précédent_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    On Error Resume Next
    Text0.SetFocus
End If
If KeyAscii = 13 Then Précédent_Click
End Sub

Private Sub Suivant_Click()
    msgf = 34
    pagemax = 0
    Unload Msgform
End Sub

Private Sub Suivant_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    On Error Resume Next
    Text0.SetFocus
End If
If KeyAscii = 13 Then Suivant_Click
End Sub

