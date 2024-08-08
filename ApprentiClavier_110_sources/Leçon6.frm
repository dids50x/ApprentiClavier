VERSION 4.00
Begin VB.Form leçon6 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   6405
   ClientLeft      =   600
   ClientTop       =   1635
   ClientWidth     =   9825
   ControlBox      =   0   'False
   Height          =   6915
   Left            =   540
   LinkTopic       =   "Form1"
   ScaleHeight     =   6405
   ScaleWidth      =   9825
   Top             =   1185
   Width           =   9945
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   8550
      Top             =   2400
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   480
      Picture         =   "Leçon6.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   750
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text6 
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   600
      Width           =   2115
   End
   Begin VB.TextBox Text5 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   600
      Width           =   4815
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1350
      Visible         =   0   'False
      Width           =   9765
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
      Height          =   615
      HideSelection   =   0   'False
      Left            =   4050
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1950
      Visible         =   0   'False
      Width           =   2865
   End
   Begin VB.Timer Timer1 
      Interval        =   1900
      Left            =   7800
      Top             =   2400
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
      Top             =   5040
      Width           =   1965
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1650
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3600
      Width           =   4965
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      HideSelection   =   0   'False
      Left            =   1650
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2700
      Width           =   4965
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F1=Aide générale       F2=Description de la touche       F3=Aide-Mémoire"
      Height          =   315
      Left            =   150
      TabIndex        =   12
      Top             =   0
      Width           =   7515
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
      Left            =   3480
      TabIndex        =   11
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
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   5880
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
      Top             =   4440
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
      Height          =   525
      Left            =   2850
      TabIndex        =   2
      Top             =   2040
      Width           =   1065
   End
End
Attribute VB_Name = "leçon6"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' ***************************  LOAD  *******************************************
Private Sub Form_Load()
If FullScreenSwitch = 1 Then WindowState = 2
Quitter.Caption = msgQuitterVers + CRLF + bannerMenu + CRLF + msgÉchap

'Paramètres
cadencecara = 40: cadenceligne = 200
typeleçon = 3  'mais alea = 1 dans text2text1
Set leçon_courante = leçon6
Module_routines.Colors leçon6  '12/2011
Module_routines.Dimension leçon6
'If repjawsnames = "" Then Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "."
'If repjawsnames <> "" Then Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "." & CRLF & msgSonori & repjawsnames & CRLF & msgSpeedExp & debexplilevel & msgSpeedGen & debgenlevel & "."
Module_routines.mshow leçon6 'avril 2008
label1.Caption = msgTapez2
Label3.Caption = bannerVersion & ", " & bannerCopyright
Label4.Caption = msgScore
Label5.Caption = msgF1F2F3
nbcaras = 0
ii = 0: iwrong = 0: iwrongbis = 0: iter = 0
cara1 = "": cara2 = "": nbonscaras = 0
elapsed = 0: elapsedtot = 0: nbmots = 0
noF1 = 1
sonocara = 0

'Définir le temps si l'utilisateur démarre trop vite
startline = Now
Module_routines.SetKeys "CAPSLOCK_OFF"
Module_routines.SetKeys "NUMLOCK_OFF"
Module_routines.SetKeys "SCROLLLOCK_OFF"
'If Dir(vpath & "sonbip2.exe") <> "" Then Module_exec.ExecAndWait vpath & "sonbip2.exe"
Module_routines.sonbip2tons 'avril 2008
End Sub


' ***************************  LOAD suite  *************************************
Private Sub Timer1_Timer()
'leçonfontsize = (20 * leçon6.Width) / 11746 '12/2011 Don't use AdjustLineAndWidth
leçon6.text1.Font.Size = (36 * zoomvalue * leçon6.Width) / 11746 '12/2011
leçon6.text2.Font.Size = (36 * zoomvalue * leçon6.Width) / 11746 '12/2011
Randomize Timer
leçon6.text1.Text = datatext1(Int((nbli * Rnd) + 1))
leçon6.text1.SelStart = 0
leçon6.text1.SelLength = Len(leçon6.text1.Text)
' Sleep ferait faire double sono en Win 98
'Call Sleep(cadenceligne)
Timer1.Enabled = False
End Sub


' ***************************  TIMER3  *****************************************
Private Sub Timer3_Timer()
' Attendre la première frappe
If iter = 0 And iwrong = 0 And ii = 0 Then
    startline = Now

' Puis passer à la ligne suivante au bout de emax secondes
Else
    currentdate = Now
    elapsed = DateDiff("s", startline, currentdate)
    If elapsed > 0 And leçon6.text1.Text = "" Then
        GoTo T1
    End If
    If elapsed > emax And (iter > 0 Or iwrong > 0 Or ii > 0) Then
T1:
        Module_routines.lignesuivante 1, 0, 0
        iter = iter + 1
        If derligne = 2 Then
            derligne = 0
            Exit Sub
        End If
        ll = Len(leçon6.text1.Text): ii = Len(leçon6.text2.Text)
    End If
End If
leçon6.Text6.Text = elapsed & msgSecondes
End Sub


' **********  TEXT1_KEYUP Events rares sur text1 ***************************
Private Sub text1_KeyUp(Keycode As Integer, shift As Integer)
If Keycode = 27 Then
    Quitter_Click
    Exit Sub
End If
'Curseur sur text1 interdit, passer à text2
On Error Resume Next
leçon6.text2.SetFocus
End Sub


' ***********************  TEXT2_KEYDOWN  **************************************
Private Sub Text2_KeyDown(Keycode As Integer, shift As Integer)
' TOUCHES à PB, annule la commande réalisée simultanément par Windows
' Win 91 et Win 92 (voir en plus Text2_KeyUp)
If Keycode = 91 Or Keycode = 92 Then
    Module_routines.cancelwin 0, leçon6, 0
    Exit Sub
End If

' Échappement
If Keycode = 27 Then Exit Sub

'RetArr interdit
If Keycode = 8 Then
    Module_routines.bip leçon6
    t2inhibit = 1
    ii = ii + 1
    leçon6.text2.Text = Left(leçon6.text1.Text, ii)
    t2inhibit = 0
    leçon6.text2.SelStart = Len(leçon6.text2.Text)
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


' **************************  TEXT2 KEY_UP  *************************************
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
        leçon6.text4.Visible = False
        leçon6.text1.SelLength = 0
        Call Sleep(cadencecara)
        
        ' La touche AltGr lance Echap qui ne lance la sélection de la fin du text1 que si noalt=0
        If erepeat = 1 Then
            erepeat = 0
        ElseIf noalt = 1 Then
            leçon6.text1.SelLength = 1
        Else
            leçon6.text1.SelLength = Len(leçon6.text1.Text)
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
    Module_routines.cancelwin 0, leçon6, 0
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
If Keycode = 112 Then Module_routines.help_f1 leçon6

'Touche F2
If Keycode = 113 Then
    keyinhibit = 2
    avecf2 = 1
    help_f2 leçon6
    avecf2 = 0
End If

'Touche F3
If Keycode = 114 Then Module_routines.help_f3 leçon6

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
    Module_routines.bip leçon6
    leçon6.text2.SelStart = Len(leçon6.text2.Text)
End If
End Sub


' **************************  TEXT2_CHANGE  *************************************
Private Sub Text2_Change()
Module_routines.text2text1 1, sonocara, 1, 0, 0, 1
End Sub


' ****************************  QUITTER  ****************************************
Private Sub Quitter_Click()
Module_routines.quit_l
End Sub


' ************************  QUITTER par le BOUTON   *****************************
Private Sub Quitter_KeyPress(KeyAscii As Integer)
If KeyAscii = 81 Or KeyAscii = 113 Then Quitter_Click
End Sub

