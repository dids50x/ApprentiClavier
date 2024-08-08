VERSION 4.00
Begin VB.Form leçon4 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   6585
   ClientLeft      =   525
   ClientTop       =   1380
   ClientWidth     =   9825
   ControlBox      =   0   'False
   Height          =   7095
   Left            =   465
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9825
   Top             =   930
   Width           =   9945
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7320
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   450
      Width           =   1365
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   750
      Picture         =   "Leçon4.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   450
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text3 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      HideSelection   =   0   'False
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1500
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.TextBox Text4 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      HideSelection   =   0   'False
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1050
      Visible         =   0   'False
      Width           =   9765
   End
   Begin VB.Timer Timer1 
      Interval        =   1600
      Left            =   7920
      Top             =   1650
   End
   Begin VB.CommandButton Quitter 
      Caption         =   " &Quitter vers   Menu   (Échap)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   7800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5160
      Width           =   1965
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3480
      Width           =   9015
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      HideSelection   =   0   'False
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2280
      Width           =   9015
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F1=Aide générale       F2=Description de la touche       F3=Aide-Mémoire"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   150
      TabIndex        =   11
      Top             =   0
      Width           =   7215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      Caption         =   "Score :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   6240
      TabIndex        =   10
      Top             =   600
      Width           =   975
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
      TabIndex        =   9
      Top             =   6000
      Width           =   7575
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
      TabIndex        =   4
      Top             =   4560
      Width           =   7575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tapez :"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2850
      TabIndex        =   2
      Top             =   1560
      Width           =   1065
   End
End
Attribute VB_Name = "leçon4"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' ********************  LOAD  *************************************************
Private Sub Form_Load()
If FullScreenSwitch = 1 Then WindowState = 2
Quitter.Caption = msgQuitterVers + CRLF + bannerMenu + CRLF + msgÉchap

'Paramètres
cadencecara = 50: cadenceligne = 260
typeleçon = 3
Set leçon_courante = leçon4
Module_routines.Colors leçon4  '12/2011
Module_routines.Dimension leçon4
'If repjawsnames = "" Then Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "."
'If repjawsnames <> "" Then Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "." & CRLF & msgSonori & repjawsnames & CRLF & msgSpeedExp & debexplilevel & msgSpeedGen & debgenlevel & "."
Module_routines.mshow leçon4 'avril 2008
label1.Caption = msgTapez2
Label3.Caption = bannerVersion & ", " & bannerCopyright
Label4.Caption = msgScore
Label5.Caption = msgF1F2F3
Module_routines.SetKeys "CAPSLOCK_OFF"
Module_routines.SetKeys "NUMLOCK_OFF"
Module_routines.SetKeys "SCROLLLOCK_OFF"
sonocara = 1

' Première currentline (après le bip 2 tons, sinon bip inaudible en Win98)
'If Dir(vpath & "sonbip2.exe") <> "" Then Module_exec.ExecAndWait vpath & "sonbip2.exe"
Module_routines.sonbip2tons 'avril 2008
Module_routines.OpenAndSuffix exo_courant, 1
On Error Resume Next
If EOF(1) Then
    Close #1
    Exit Sub
Else
    Line Input #1, currentline
    leçonfontsize = (36 * leçon4.Width) / 11746
    Module_routines.AdjustWidthAndSize leçon4, 1
End If
End Sub

' ********************  LOAD suite ********************************************
Private Sub Timer1_Timer()
leçon4.text1.Text = currentline
leçon4.text1.SelStart = 0
leçon4.text1.SelLength = Len(leçon4.text1.Text)
' Sleep ferait faire double sono en Win 98
'Call Sleep(cadenceligne)
nbcaras = Len(leçon4.text1.Text)
iwrong = 0: iwrongbis = 0: iter = 0
cara1 = "": cara2 = ""
Module_routines.text2text1 1, sonocara, 0, 0, 0, 0
Timer1.Enabled = False
End Sub


' **********  TEXT1_KEYUP Events rares sur text1 ***************************
Private Sub text1_KeyUp(Keycode As Integer, shift As Integer)
If Keycode = 27 Then
    Quitter_Click
    Exit Sub
End If
'Curseur sur text1 interdit, passer à text2
On Error Resume Next
leçon4.text2.SetFocus
End Sub


' ****************  TEXT2_KEYDOWN  ******************************************
Private Sub Text2_KeyDown(Keycode As Integer, shift As Integer)
' TOUCHES à PB, annule la commande réalisée simultanément par Windows
' Win 91 et Win 92 (voir en plus Text2_KeyUp)
If Keycode = 91 Or Keycode = 92 Then
    Module_routines.cancelwin 0, leçon4, 0
    Exit Sub
End If

' Échappement
If Keycode = 27 Then Exit Sub

'RetArr interdit
If Keycode = 8 Then
    Module_routines.bip leçon4
    t2inhibit = 1
    ii = ii + 1
    leçon4.text2.Text = Left(leçon4.text1.Text, ii)
    t2inhibit = 0
    leçon4.text2.SelStart = Len(leçon4.text2.Text)
End If

'Combinaison Maj+ESPACE (répétera la (FIN de la) ligne)
If Keycode = 32 And shift = 1 Then lrepeat = 1

'Combinaison Control+ESPACE (répétera le mot)
If Keycode = 32 And shift = 2 Then wrepeat = 1

'Alt+ESPACE épellera le mot
If Keycode = 32 And shift = 4 Then
    erepeat = 1
    Module_routines.epellation leçon_courante
End If

'Touche Alt+F4 pour quitter
If Keycode = 115 And shift = 4 Then
    altf4 = 1
    If quitactive = 0 Then Module_routines.QuitQuit
End If
End Sub


' *************** TEXT2_KEYUP  ****************************************************
Private Sub Text2_KeyUp(Keycode As Integer, shift As Integer)
' Les msgbox procurent des key-ups indésirables avec les 3 commandes Entrée Non Oui
If keyinhibit = 1 Then
    keyinhibit = 0
    If Keycode = 13 Or Keycode = 78 Or Keycode = 79 Then Exit Sub
End If

' Échappement réitéré
If Keycode = 27 Then
    
    ' Pour Échap de la touche F2
    If keyinhibit = 2 Then
        keyinhibit = 0
        leçon4.text4.Visible = False
        leçon4.text1.SelLength = 0
        Call Sleep(cadencecara)
        
        ' La touche AltGr lance Echap qui ne lance la sélection de la fin du text1 que si noalt=0
        If erepeat = 1 Then
            erepeat = 0
        ElseIf noalt = 1 Then
            leçon4.text1.SelLength = 1
        Else
            leçon4.text1.SelLength = Len(leçon4.text1.Text)
        End If
        Exit Sub
    End If
    
    ' Autres Echap
    echapbis = echapbis + 1
    If echapbis > echapbismax Then
        echapbis = 0
        If keyinhibit = 0 Then Quitter_Click  'Evite de quitter après un Alt sonorisé
        Exit Sub
    End If
End If

' Pour certains traitements
If keyinhibit = 2 Then keyinhibit = 1

' TOUCHES à PB, annule la commande réalisée simultanément par Windows
' Win 91 et Win 92 (voir en plus Text2_KeyDown)
If Keycode = 91 Or Keycode = 92 Then
    Module_routines.cancelwin 0, leçon4, 0
    Exit Sub
End If

' AltDroit 17 (qui envoie ensuite 18 et évent-t 27)
If Keycode = 17 Then
    echapoff = 0
    Exit Sub
End If

' AltGauche 18 et Menu-Contextuel 93
'If Keycode = 18 Or Keycode = 93 Then
If Keycode = 93 Then
    keyinhibit = 2
    echapbis = -1
    'SendKeys "{ESC}"
    'Sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    echapbis = 0 + echapoff
    echapoff = 0
    Exit Sub
End If

' Sendkeys termine par 145
If Keycode = 145 Then Exit Sub

'Touche F1
If Keycode = 112 Then Module_routines.help_f1 leçon4

'Touche F2
If Keycode = 113 Then
    keyinhibit = 2
    avecf2 = 1
    help_f2 leçon4
    avecf2 = 0
End If

'Touche F3
If Keycode = 114 Then Module_routines.help_f3 leçon4

'Touche F10 mène à la barre menu
If Keycode = 121 Then
    echapbis = echapbis - 1
    'SendKeys "{ESC}", True
    'Sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
End If

'Flèche gauche interdite
If Keycode = 37 Then
    Module_routines.bip leçon4
    leçon4.text2.SelStart = Len(leçon4.text2.Text)
End If
End Sub


' **************************  TEXT2_CHANGE  *************************************
Private Sub Text2_Change()
Module_routines.text2text1 1, sonocara, 0, 0, 0, 0
End Sub


' ******************************  QUITTER  ***************************************
Private Sub Quitter_Click()
Module_routines.quit_l
End Sub

' *************************  QUITTER par le BOUTON  *****************************
Private Sub Quitter_KeyPress(KeyAscii As Integer)
If KeyAscii = 81 Or KeyAscii = 113 Then Quitter_Click
End Sub

