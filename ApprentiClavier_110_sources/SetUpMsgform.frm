VERSION 4.00
Begin VB.Form SetUpMsgform 
   Caption         =   " "
   ClientHeight    =   5940
   ClientLeft      =   1305
   ClientTop       =   1605
   ClientWidth     =   9270
   ControlBox      =   0   'False
   Height          =   6450
   KeyPreview      =   -1  'True
   Left            =   1245
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   9270
   Top             =   1155
   Width           =   9390
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Interval        =   2500
      Left            =   8640
      Top             =   5400
   End
   Begin VB.CommandButton Continuer 
      Caption         =   "&CONTINUER (Entrée)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   960
      TabIndex        =   1
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Timer Timer2 
      Interval        =   3000
      Left            =   4560
      Top             =   5400
   End
   Begin VB.TextBox Text0 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   5205
      HideSelection   =   0   'False
      Left            =   150
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
   Begin VB.CommandButton Quitter 
      Caption         =   "&QUITTER (Échap)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5520
      TabIndex        =   2
      Top             =   5160
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   240
      Top             =   5400
   End
End
Attribute VB_Name = "SetUpMsgform"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' ********************   LOAD  ****************************************
Private Sub Form_Load()
If FullScreenSwitch = 1 Then WindowState = 2
Text0.Text = msgtext0
Text0.Font.Size = fsize
Text0.BackColor = fbackcolor
Text0.ForeColor = fcolor
Continuer.Caption = msgContinuer
Quitter.Caption = msgQuitter
Module_SetUpGlobal.Dimension SetUpMsgform
If timein > 0 Then
    Text0.SelStart = 0
    Text0.SelLength = Len(Text0.Text)
Else
    Text0.SelStart = 0
    Text0.SelLength = 0
End If
msgf = 9
If keyinhibit <> 2 Then keyinhibit = 1
If f1expli = 0 Then SetUpMsgform.Quitter.Visible = False
If f1expli = 1 Then SetUpMsgform.Quitter.Visible = True
End Sub

' *******************  TIMER1  ***************************************
Private Sub Timer1_Timer()
If timein = 0 Then
    Text0.SelStart = 0
    Text0.SelLength = Len(Text0.Text)
Else
    Text0.SelStart = 0
    Text0.SelLength = 0
End If
On Error Resume Next
Text0.SetFocus
Timer1.Enabled = False
Timer2.Enabled = True
If keyinhibit <> 1 Then keyinhibit = 0
End Sub


' *******************  TIMER3  ***************************************
Private Sub Timer3_Timer()
Text0.SelStart = 0
Text0.SelLength = 0
On Error Resume Next
Text0.SetFocus
Timer3.Enabled = False
End Sub


' *******************  TIMER2  ***************************************
Private Sub Timer2_Timer()
timein = timein + 1
If timeout > 0 Then timeout = timeout + 1

' Cadence de relance de la sonorisation du message apparaissant dans la fenêtre
If timein = 7 Or timein = 14 Then
    If f1expli = 0 Then
        Text0.SelStart = 0
        Text0.SelLength = Len(Text0.Text)
    End If
End If
Timer1.Enabled = True
Timer2.Enabled = False

' Quitter au bout du timeout
If timeout >= 2 Then
    Unload SetUpMsgform
    inst = -2
    Module_SetUpGlobal.SonoLocate
    End
End If
End Sub


' *******************  TEXT0_KEYdown  *********************************
Private Sub text0_keydown(keycode As Integer, shift As Integer)
' Win 91 et Win 92
If keycode = 91 Or keycode = 92 Then
    msgpbmenu.Show 1
    Module_SetUpGlobal.cancelwin 0, 0
    msgpbmenu.Show 1
    Exit Sub
End If

' Échappement (complément de KeyUp)
If keycode = 27 Then Exit Sub

' Touche F1 Aide
If keycode = 112 Then
    msgf = 3
    Exit Sub
End If

'Touche Alt+F4 pour quitter
If keycode = 115 And shift = 4 Then
    Text0.Text = ""
    Text0.SelStart = 0
    Text0.SelLength = 0
    SetUpMsgform.Cls
    Call Sleep(50)
    Text0.Text = CRLF2 + "   Alt+F4.  " + msgAurevoir
    Text0.SelStart = 0
    Text0.SelLength = Len(Text0.Text)
    Call Sleep(1500)
    Unload SetUpMsgform
    inst = -2
    Module_SetUpGlobal.SonoLocate
    End
End If

' Touches Maj, Control, Alt, VerrNum
If (keycode > 15) And (keycode < 24) Then
    Exit Sub
End If

' Menu Contextuel 93
If keycode = 93 Then
    keyinhibit = 1
    echapbis = 0
    SendKeys "{ESC}"
    Exit Sub
End If

' Quitte le pavé numérique avec NUMLOCK_OFF
If keycode = 144 Then Exit Sub

' Entrée ou OUI pour CONTINUER
If keycode = 13 Then Exit Sub

' ESPACE pour répéter
If keycode = 32 Then
    timein = 0
    msgf = 2
End If

' PagePréc pour revenir
If keycode = 33 Then
    timein = 0
    msgf = 33
End If

' PageSuiv pour continuer
If keycode = 34 Then
    timein = 0
    msgf = 34
    If keyinhibit = 2 Then Unload SetUpMsgform
End If

' PagePrécSuiv 33 et 34, DebFin 36 et 35, Flèches 37 à 40
If (keycode > 32) And (keycode < 41) Then
    Timer1.Enabled = False
    Timer1.Enabled = True
    Exit Sub
End If

' En provenance de bienvenue I Installer ou D Désinstaller
If keycode = 73 Or keycode = 105 Then Exit Sub
If keycode = 68 Or keycode = 100 Then Exit Sub

' AUTRES touches
If keycode > 40 Then
    Text0.Text = Chr(keycode)
    Text0.SelStart = 0
    Text0.SelLength = 1
    Call Sleep(300)
    Text0.Text = ""
    Text0.SelStart = 0
    Text0.SelLength = 0
    SetUpMsgform.Cls
    Call Sleep(50)
    echapbis = echapbis + 1
    
    ' TROP d'ERREURS RÉPÉTÉES
    If echapbis < 6 Then
        Text0.Text = CRLF + msgVousEtiez
        Text0.SelStart = 0
        Text0.SelLength = Len(Text0.Text)
        Call Sleep(300)
        Text0.SelStart = 0
        Text0.SelLength = 0
        Exit Sub
    Else
        echapbis = 0
        msgf = 0
        Text0.Text = msgRecommencez
        Text0.SelStart = 0
        Text0.SelLength = Len(Text0.Text)
    End If
End If
    
' Suite
    Unload SetUpMsgform
End Sub


' *******************  TEXT0_KEYup  *********************************
Private Sub text0_keyup(keycode As Integer, shift As Integer)
' Echappement
If keycode = 27 Then
    msgf = 0
    
    ' Aurevoir, pour éviter que la sono cite la tâche de background
    On Error Resume Next
    Text0.SetFocus
    If Text0.Text <> msgAurevoir And f1expli = 0 Then
        Text0.Text = ""
        Text0.SelStart = 0
        Text0.SelLength = 0
        SetUpMsgform.Cls
        Call Sleep(50)
        Text0.Text = CRLF2 + msgEchap + msgAurevoir
        Text0.SelStart = 0
        Text0.SelLength = Len(Text0.Text)
        Call Sleep(1500)
    End If
    Unload SetUpMsgform
    If f1expli = 0 Then
        inst = -2
        Module_SetUpGlobal.SonoLocate
        End
    End If
    If f1expli = 1 Then keyinhibit = 0
End If

' Touche F1 Aide
'If keycode = 112 Then
'    msgf = 3
'    Unload SetUpMsgform
'    Exit Sub
'End If

' Entrée ou OUI pour CONTINUER sauf si keyinhibit=2 venant d'un message
If (keycode = 13 Or keycode = 33 Or keycode = 34) And keyinhibit < 2 Then
    If timeout = 0 Then continuer_click
End If
If keycode = 13 And keyinhibit = 2 Then keyinhibit = 0
End Sub

' *******************  Boutons CONTINUER, QUITTER  ***********************************
Private Sub Quitter_Click()
    msgf = 0
    stopscroll = 1
    Unload SetUpMsgform
End Sub

Private Sub Quitter_KeyPress(keyascii As Integer)
If keyascii = 27 Then
    On Error Resume Next
    Text0.SetFocus
End If
If keyascii = 13 Then Quitter_Click
End Sub

Private Sub continuer_click()
    If msgf <> 33 And msgf <> 34 Then msgf = 1
    On Error Resume Next
    Text0.SetFocus
    timein = 0
    Text0.Text = ""
    SetUpMsgform.Cls
    Call Sleep(100)
    If f1expli = 0 Then Module_SetUpGlobal.AuRevoir
    If f1expli = 1 Then Unload SetUpMsgform
End Sub

Private Sub continuer_KeyPress(keyascii As Integer)
If keyascii = 27 Then
    On Error Resume Next
    Text0.SetFocus
End If
If keyascii = 13 Then continuer_click
End Sub

