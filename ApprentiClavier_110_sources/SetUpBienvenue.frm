VERSION 4.00
Begin VB.Form SetUpBienvenue 
   Caption         =   "Installation de ApprentiClavier"
   ClientHeight    =   6510
   ClientLeft      =   1050
   ClientTop       =   1695
   ClientWidth     =   8760
   Height          =   7020
   KeyPreview      =   -1  'True
   Left            =   990
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   8760
   Top             =   1245
   Width           =   8880
   Begin VB.CommandButton Aide 
      Caption         =   " Aide    (F1)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   4500
      TabIndex        =   4
      Top             =   3900
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton Désinstaller 
      Caption         =   "&Désinstaller"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   2400
      TabIndex        =   3
      Top             =   3900
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.CommandButton Installer 
      Caption         =   "&Installer"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   300
      TabIndex        =   2
      Top             =   3900
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   6900
      Top             =   150
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6150
      Top             =   150
   End
   Begin VB.TextBox Text0 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      HideSelection   =   0   'False
      Left            =   450
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   900
      Width           =   7365
   End
   Begin VB.CommandButton Annuler 
      Caption         =   "Annuler  (Échap)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   6450
      TabIndex        =   5
      Top             =   3900
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   5160
      Width           =   7815
   End
   Begin VB.Label Label0 
      Caption         =   "Attention"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1350
      TabIndex        =   0
      Top             =   450
      Width           =   1665
   End
End
Attribute VB_Name = "SetUpBienvenue"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' ********************  LOAD  *********************************************
Private Sub Form_Load()
If App.PrevInstance = True Then End
If FullScreenSwitch = 1 Then WindowState = 2
Caption = msgInstall & bannerversion
Label0.Caption = msgAttention
Label1.Caption = msgKeyboard & clavierType & ". " & country + CRLF + bannerversion & ", " & bannercopyright
Installer.Caption = msgInstaller
Désinstaller.Caption = msgDésinstaller
aide.Caption = msgAide
Annuler.Caption = msgAnnuler
Module_SetUpGlobal.zoomform SetUpBienvenue
Module_SetUpGlobal.Dimension SetUpBienvenue
timein = 0    ' Nécessaire après passage par F1
timebienv = 0
End Sub


' ***********************  TIMER1  ****************************************
Private Sub Timer1_Timer()
Text0.Text = texte_bienvenue
Text0.SelStart = 0
Text0.SelLength = 0
On Error Resume Next
Text0.SetFocus
Installer.Visible = True
Désinstaller.Visible = True
aide.Visible = True
Annuler.Visible = True

' Surbrillance du texte après des délais timer1
If timebienv > 4 Then
    Timer1.Enabled = False
    Timer2.Enabled = True
    Text0.SelStart = 0
    Text0.SelLength = Len(Text0.Text)
    Call Sleep(800)
    Text0.SelStart = 0
    Text0.SelLength = 0
End If

' Compteur de temps
timebienv = timebienv + 1
End Sub


' ***********************  TIMER2  *****************************************
Private Sub Timer2_Timer()
On Error Resume Next
Text0.SetFocus
Timer2.Enabled = False
Timer1.Enabled = True
End Sub


' *******************  TEXT0_KEYDOWN  *************************************
Private Sub text0_keydown(keycode As Integer, shift As Integer)
' Win 91 et Win 92 (voir aussi text0_KeyUp)
If keycode = 91 Or keycode = 92 Then
    keyinhibit = 1
    Module_SetUpGlobal.cancelwin 0, 1
    On Error Resume Next
    Text0.SetFocus
    Exit Sub
End If

' F1 Explications
If keycode = 112 Then Aide_click

' BONNE RÉPONSE "I INSTALLER"
If keycode = 73 Or keycode = 105 Then Installer_Click

' BONNE RÉPONSE "D DÉSINSTALLER" 68 ou 100 = D ou d
If keycode = 68 Or keycode = 100 Then Désinstaller_Click
End Sub


' *******************  TEXT0_KEYUP  *********************************
Private Sub text0_keyup(keycode As Integer, shift As Integer)
' Si on vient d'un msgbox
If keyinhibit = 1 Then
    keyinhibit = 0
    Exit Sub
End If

' Win 91 et Win 92 (voir aussi text0_KeyDown)
If keycode = 91 Or keycode = 92 Then
    Unload SetUpBienvenue
    Module_SetUpGlobal.cancelwin 0, 1
    SetUpBienvenue.Show   ' Non-modal sans 1 pour annuler correctement la touche windows
    Exit Sub
End If

' Menu Contextuel 93
If keycode = 93 Then
    keyinhibit = 1
    echapbis = 0
    SendKeys "{ESC}"
    Beep
    Exit Sub
End If

' Echappement
If keycode = 27 Then
    Unload SetUpBienvenue
    keyinhibit = 1
    MsgBox annul, 0, ""
    inst = -2
    Module_SetUpGlobal.SonoLocate
    End
End If

'F1 Explications
If keycode = 112 Then Exit Sub

' BONNE RÉPONSE "I INSTALLER"
If keycode = 73 Or keycode = 105 Then Exit Sub

' BONNE RÉPONSE "D DÉSINSTALLER"
If keycode = 68 Or keycode = 100 Then Exit Sub

' Touches Maj, Control, Alt, VerrNum inertes
If (keycode > 15) And (keycode < 24) Then Exit Sub

' Touche Échap (complément)
If keycode = 27 Then

' PagePrécSuiv 33 et 34, DebFin 36 et 35, Flèches 37 à 40
ElseIf (keycode > 32) And (keycode < 41) Then
    Timer1.Enabled = False
    Timer2.Enabled = False
    Exit Sub

' Mauvaises réponses au lieu de "I" ou "D"
Else
    Text0.SelStart = 0
    Text0.SelLength = 1
    On Error Resume Next
    Text0.SetFocus
    Beep
    Call Sleep(300)
    SetUpBienvenue.Cls
    Text0.Text = texte_bienvenue
    Text0.SelStart = 0
    Text0.SelLength = Len(Text0.Text)
    On Error Resume Next
    Text0.SetFocus
    Call Sleep(800)
    Text0.SelStart = 0
    Text0.SelLength = 0
    Exit Sub
End If
End Sub


' ****************  BOUTONS Installer, Désinstaller, Aide, Annuler  *******************
Private Sub Installer_Keyup(keycode As Integer, shift As Integer)
' Echappement
If keycode = 27 Then
    Unload SetUpBienvenue
    keyinhibit = 1
    MsgBox annul, 0, ""
    inst = -2
    Module_SetUpGlobal.SonoLocate
    End
End If
End Sub

Private Sub Installer_Click()
' Nécéssaire pour quand on valide sur le bouton !
keyinhibit = 2

' Installer
Text0.Text = ""
SetUpBienvenue.Cls
Call Sleep(50)
Text0.Text = CRLF + msgPatientez
Text0.SelStart = 0
Text0.SelLength = Len(Text0.Text)
Call Sleep(1000)
Module_SetUpGlobal.install
End Sub

Private Sub désinstaller_keyup(keycode As Integer, shift As Integer)
' Echappement
If keycode = 27 Then
    Unload SetUpBienvenue
    keyinhibit = 1
    MsgBox annul, 0, ""
    inst = -2
    Module_SetUpGlobal.SonoLocate
    End
End If
End Sub

Private Sub Désinstaller_Click()
' Désinstaller
Text0.Text = ""
SetUpBienvenue.Cls
Call Sleep(50)
Text0.Text = CRLF + msgDésinst + CRLF + msgPatientez
Text0.SelStart = 0
Text0.SelLength = Len(Text0.Text)
Call Sleep(2000)
Module_SetUpGlobal.desinstall
End Sub

Private Sub aide_keyup(keycode As Integer, shift As Integer)
' Echappement
If keycode = 27 Then
    Unload SetUpBienvenue
    keyinhibit = 1
    MsgBox annul, 0, ""
    inst = -2
    Module_SetUpGlobal.SonoLocate
    End
End If
End Sub

Private Sub Aide_click()
' Nécéssaire pour quand on valide sur le bouton !
keyinhibit = 2

' Aide
    tempo = ""
    tempo = Dir(vpath & "alire.txt")
    If tempo = "" Then Exit Sub
    Unload SetUpBienvenue
    f1expli = 1
    fsize = 9
    stopscroll = 0
SCR0:
' Scrollresults x, y, z signifie : à partir de la ligne x, montrer les y lignes suivantes, numéro-page
    Module_SetUpGlobal.scrollresults 0, 19, 1
    If msgf = 0 Then stopscroll = 1
    If msgf = 33 Then GoTo SCR0
SCR1:
' Scrollresults x, y, z signifie : à partir de la ligne x, montrer les y lignes suivantes, numéro-page
    If stopscroll = 0 Then Module_SetUpGlobal.scrollresults 19, 17, 2
    If msgf = 0 Then stopscroll = 1
    If msgf = 33 Then GoTo SCR0
    If msgf = 34 Then stopscroll = 0
SCR2:
' Scrollresults x, y, z signifie : à partir de la ligne x, montrer les y lignes suivantes, numéro-page
    If stopscroll = 0 Then Module_SetUpGlobal.scrollresults 36, 18, 3
    If msgf = 0 Then stopscroll = 1
    If msgf = 33 Then
        stopscroll = 0
        GoTo SCR1
    End If
    If msgf = 34 Then
        stopscroll = 0
        GoTo SCR2
    End If
    f1expli = 0
    fsize = 14
    SetUpBienvenue.Show 1
End Sub

Private Sub Annuler_Click()
inst = -2
Module_SetUpGlobal.SonoLocate
End
End Sub

Private Sub Annuler_KeyPress(keyascii As Integer)
If keyascii = 27 And f1expli = 0 Then Annuler_Click
End Sub

