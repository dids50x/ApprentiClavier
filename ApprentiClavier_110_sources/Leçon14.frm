VERSION 4.00
Begin VB.Form le�on14 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   6630
   ClientLeft      =   1305
   ClientTop       =   1545
   ClientWidth     =   9825
   ControlBox      =   0   'False
   Height          =   7140
   Left            =   1245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   9825
   Top             =   1095
   Width           =   9945
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   900
      Picture         =   "Le�on14.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   615
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
      Height          =   375
      Left            =   4440
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   240
      Width           =   4665
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   9150
      Top             =   1080
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   630
      Visible         =   0   'False
      Width           =   9645
   End
   Begin VB.TextBox Text3 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      HideSelection   =   0   'False
      Left            =   4950
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Timer Timer1 
      Interval        =   2200
      Left            =   8550
      Top             =   1080
   End
   Begin VB.CommandButton Quitter 
      Caption         =   " &Quitter vers          Menu        (�chap)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
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
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3330
      Width           =   9495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      HideSelection   =   0   'False
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1530
      Width           =   9495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F1=Aide g�n�rale       F2=Description de la touche       F3=Aide-M�moire"
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
      Width           =   7815
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
      Height          =   255
      Left            =   3360
      TabIndex        =   10
      Top             =   360
      Width           =   1035
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
      Top             =   6240
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
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   5040
      Width           =   7605
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tapez "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3750
      TabIndex        =   2
      Top             =   1080
      Width           =   1065
   End
End
Attribute VB_Name = "le�on14"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' ***************************  LOAD  *******************************************
Private Sub Form_Load()
If FullScreenSwitch = 1 Then WindowState = 2
Quitter.Caption = msgQuitterVers + CRLF + bannerMenu + CRLF + msg�chap

' Param�tres
cadencecara = 40: cadenceligne = 260
typele�on = 14
noalt = 0  ' Alt r�p�te la fin de la phrase
Set le�on_courante = le�on14
Module_routines.Colors le�on14  '12/2011
Module_routines.Dimension le�on14
'If repjawsnames = "" Then Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "."
'If repjawsnames <> "" Then Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "." & CRLF & msgSonori & repjawsnames & CRLF & msgSpeedExp & debexplilevel & msgSpeedGen & debgenlevel & "."
Module_routines.mshow le�on14 'avril 2008
label1.Caption = msgTapez
Label3.Caption = bannerVersion & ", " & bannerCopyright
Label4.Caption = msgScore
Label5.Caption = msgF1F2F3

' Reset scorecourant
pctt = 100
scorecourant = " 100%.   " & "0" & msgMots & "   0" & msgSecondes

' Premi�re currentline (apr�s le bip 2 tons, sinon bip inaudible en Win98)
' Utilisera OpenAnd Count sans Alt255 en fin de chaque ligne
'If Dir(vpath & "sonbip2.exe") <> "" Then Module_exec.ExecAndWait vpath & "sonbip2.exe" 'avril 2008
Module_routines.sonbip2tons 'avril 2008
Module_routines.OpenAndCount exo_courant
On Error Resume Next
If EOF(1) Then
    Close #1
    Exit Sub
Else
    Line Input #1, currentline
    ' Modifi� 12/2011
    le�onfontsize = (16 * le�on14.Width) / 11746
    Module_routines.AdjustWidthAndSize le�on14, 1
    zz = Len(currentline)
    iwrong = 0: iwrongbis = 0: iter = 0
    iwrongbismax = 2
    cara1 = "": cara2 = ""
    elapsed = 0: elapsedtot = 0: nbmots = 0
    
    'D�tecter la fin de la 1�re phrase par un ".", un "...", un "!", un "?".
    iistop = InStr(currentline, ".")
    iistop2 = InStr(currentline, "!")
    iistop3 = InStr(currentline, "?")
    
    'D�tecter la fin de la 1�re phrase
    Module_routines.DetectPhraseEnd
End If

' D�finir le temps si l'utilisateur d�marre trop vite
starttop = Now
Module_routines.SetKeys "CAPSLOCK_OFF"
Module_routines.SetKeys "NUMLOCK_OFF"
Module_routines.SetKeys "SCROLLLOCK_OFF"
End Sub


' ***************************  LOAD suite  *************************************
Private Sub Timer1_Timer()
' Timer1 laisse � Jaws le temps de lire l'�tiquette "Tapez" et la premi�re ligne
'Pour une sono "multiline" correcte, on ne prend d'abord bri�vement qu'un cara, puis on limitera � la phrase.
If iistop > 1 Then
    iistopf = iistop
    iistop = 1
End If
le�on14.text1.Text = Left(currentline, iistop)
le�on14.text1.SelStart = 0

If iistopf > iistop Then
    le�on14.text1.Text = Left(currentline, iistopf)
    le�on14.text1.SelStart = 0
    le�on14.text1.SelLength = iistopf
    ' Sleep ferait faire double sono en Win 98 ?
    Call Sleep(1250) 'en avril 2008, au lieu de 1000
    iistop = iistopf
End If
nbcaras = Len(le�on14.text1.Text)
ll = Len(le�on14.text1.Text): ii = Len(le�on14.text2.Text)
Module_routines.nextspace le�on14
Timer1.Enabled = False
On Error Resume Next
le�on14.text2.SetFocus
End Sub


' ***************************  TIMER2  *****************************************
Private Sub Timer2_Timer()
' Attendre la premi�re frappe pour chaque ligne
If iwrong = 0 And ii = 0 Then
    starttop = Now

' Puis compter le temps elapsed pour la ligne
Else
    currentdate = Now
    elapsed = DateDiff("s", starttop, currentdate)
End If

' Scorecourant � calculer/afficher d�s qu'on a frapp� la premi�re lettre
If ii > 0 Then scorecourant = CInt(pctt) & " %.   " & nbmots & msgMots & elapsed & msgSecondes
le�on14.text5.Text = scorecourant
End Sub


' **********  TEXT1_KEYUP Events rares sur text1 ***************************
Private Sub text1_KeyUp(Keycode As Integer, shift As Integer)
If Keycode = 27 Then
    Quitter_Click
    Exit Sub
End If
'Curseur sur text1 interdit, passer � text2
On Error Resume Next
le�on14.text2.SetFocus
End Sub


' ***********************  TEXT2_KEYDOWN  **************************************
Private Sub Text2_KeyDown(Keycode As Integer, shift As Integer)
' TOUCHES � PB, annule la commande r�alis�e simultan�ment par Windows
' Win 91 et Win 92 (voir en plus Text2_KeyUp)
If Keycode = 91 Or Keycode = 92 Then
    Module_routines.cancelwin 0, le�on14, 0
    Exit Sub
End If

' �chappement
If Keycode = 27 Then Exit Sub

'RetArr interdit
If Keycode = 8 Then
    Module_routines.bip le�on14
    t2inhibit = 1
    ii = ii + 1
    le�on14.text2.Text = Left(le�on14.text1.Text, ii)
    t2inhibit = 0
    le�on14.text2.SelStart = Len(le�on14.text2.Text)
End If

'Combinaison Maj+ESPACE (r�p�tera la ligne)
If Keycode = 32 And shift = 1 Then lrepeat = 1

'Combinaison Control+ESPACE (r�p�tera le mot)
If Keycode = 32 And shift = 2 Then wrepeat = 1

'Alt+ESPACE �pellera le mot
If Keycode = 32 And shift = 4 Then
    erepeat = 1
    Module_routines.epellation le�on_courante
End If

'Touche Alt+F4 pour quitter
If Keycode = 115 And shift = 4 Then
    altf4 = 1
    If quitactive = 0 Then Module_routines.QuitQuit
End If
End Sub


' **************************  TEXT2 KEY_UP  *************************************
Private Sub Text2_KeyUp(Keycode As Integer, shift As Integer)
' Les msgbox procurent des key-ups ind�sirables avec les 3 commandes Entr�e Non Oui
If keyinhibit = 1 Then
    keyinhibit = 0
    If Keycode = 13 Or Keycode = 78 Or Keycode = 79 Then Exit Sub
End If

' �chappement r�it�r�
If Keycode = 27 Then
    
    ' Pour �chap de la touche F2
    If keyinhibit = 2 Then
        keyinhibit = 0
        le�on14.text4.Visible = False
        le�on14.text1.SelLength = 0
        Call Sleep(cadencecara)
        
        ' La touche AltGr lance Echap qui ne lance la s�lection de la fin du text1 que si noalt=0
        If erepeat = 1 Then
            erepeat = 0
        ElseIf noalt = 1 Then
            le�on14.text1.SelLength = 1
        Else
            le�on14.text1.SelStart = ii
            le�on14.text1.SelLength = Len(le�on14.text1.Text)
        End If
        Exit Sub
    End If
    
    ' Autres Echap
    echapbis = echapbis + 1
    If echapbis > echapbismax Then
        echapbis = 0
        If keyinhibit = 0 Then Quitter_Click  'Evite de quitter apr�s un Alt sonoris�
        Exit Sub
    End If
End If

' Pour certains traitements
If keyinhibit = 2 Then keyinhibit = 1

' TOUCHES � PB, annule la commande r�alis�e simultan�ment par Windows
' Win 91 et Win 92 (voir en plus Text2_KeyDown)
If Keycode = 91 Or Keycode = 92 Then
    Module_routines.cancelwin 0, le�on14, 0
    Exit Sub
End If

' AltDroit 17 (qui envoie ensuite 18 et �vent-t 27)
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
    'Sendkeys est remplac� par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    echapbis = 0 + echapoff
    echapoff = 0
    Exit Sub
End If

' Sendkeys termine par 145
If Keycode = 145 Then Exit Sub

'Touche F1
If Keycode = 112 Then Module_routines.help_f1 le�on14

'Touche F2
If Keycode = 113 Then
    keyinhibit = 2
    avecf2 = 1
    help_f2 le�on14
    avecf2 = 0
End If

'Touche F3
If Keycode = 114 Then Module_routines.help_f3 le�on14

'Touche F10 m�ne � la barre menu
If Keycode = 121 Then
    echapbis = echapbis - 1
    'SendKeys "{ESC}", True
    'Sendkeys est remplac� par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
End If

'Fl�che gauche interdite
If Keycode = 37 Then
    Module_routines.bip le�on14
    le�on14.text2.SelStart = Len(le�on14.text2.Text)
End If
End Sub


' **************************  TEXT2_CHANGE  *************************************
Private Sub Text2_Change()
Module_routines.text2text1 indif, sonocara, 0, 1, 1, 0
End Sub


' ****************************  QUITTER  ****************************************
Private Sub Quitter_Click()
Module_routines.quit_l
End Sub

' ************************  QUITTER par le BOUTON   *****************************
Private Sub Quitter_KeyPress(KeyAscii As Integer)
If KeyAscii = 81 Or KeyAscii = 113 Then Quitter_Click
End Sub

