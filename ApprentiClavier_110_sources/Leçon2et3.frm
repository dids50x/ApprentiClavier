VERSION 4.00
Begin VB.Form Le�on2et3 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   6585
   ClientLeft      =   615
   ClientTop       =   1605
   ClientWidth     =   9735
   ControlBox      =   0   'False
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   7095
   Left            =   555
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   9735
   Top             =   1155
   Visible         =   0   'False
   Width           =   9855
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
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   600
      Picture         =   "Le�on2et3.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   450
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1800
      Left            =   8280
      Top             =   1920
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
      Top             =   1050
      Visible         =   0   'False
      Width           =   9735
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
      Left            =   4560
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1500
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Quitter 
      Caption         =   " &Quitter vers   Menu   (�chap)"
      Height          =   1215
      Left            =   7800
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5400
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
      Height          =   1065
      Left            =   3450
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3600
      Width           =   1365
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
      Height          =   1050
      HideSelection   =   0   'False
      Left            =   3450
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2250
      Width           =   1365
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
      Left            =   240
      TabIndex        =   11
      Top             =   0
      Width           =   7605
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
      Left            =   6000
      TabIndex        =   10
      Top             =   480
      Width           =   1125
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
      TabIndex        =   9
      Top             =   6240
      Width           =   7605
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
      Left            =   150
      TabIndex        =   3
      Top             =   4800
      Width           =   7545
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
      Height          =   585
      Left            =   3480
      TabIndex        =   0
      Top             =   1560
      Width           =   1065
   End
End
Attribute VB_Name = "Le�on2et3"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' ********************  LOAD  *************************************************
Private Sub Form_Load()
If FullScreenSwitch = 1 Then WindowState = 2
Quitter.Caption = msgQuitterVers + CRLF + bannerMenu + CRLF + msg�chap

'Param�tres
cadencecara = 200: cadenceligne = 350
typele�on = 2
Set le�on_courante = Le�on2et3
Module_routines.Colors Le�on2et3  '12/2011
Module_routines.Dimension Le�on2et3
'If repjawsnames = "" Then Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "."
'If repjawsnames <> "" Then Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "." & CRLF & msgSonori & repjawsnames & CRLF & msgSpeedExp & debexplilevel & msgSpeedGen & debgenlevel & "."
Module_routines.mshow Le�on2et3 'avril 2008
label1.Caption = msgTapez2
Label3.Caption = bannerVersion & ", " & bannerCopyright
Label4.Caption = msgScore
Label5.Caption = msgF1F2F3
Module_routines.SetKeys "CAPSLOCK_OFF"
Module_routines.SetKeys "NUMLOCK_OFF"
Module_routines.SetKeys "SCROLLLOCK_OFF"
sonocara = 1

' Premi�re currentline (apr�s le bip 2 tons, sinon bip inaudible en Win98 ; bipinhibit pour �viter le bip quand on quitte l'exo sur les fautes)
If exo_courant = "fautes.txt" Then bipinhibit = bipinhibit + 1
'If bipinhibit < 2 And Dir(vpath & "sonbip2.exe") <> "" Then Module_exec.ExecAndWait vpath & "sonbip2.exe" 'avril 2008
If bipinhibit < 2 Then Module_routines.sonbip2tons 'avril 2008
Module_routines.OpenAndSuffix exo_courant, 1
On Error Resume Next
If EOF(1) Then
    Close #1
    Exit Sub
Else
    Line Input #1, currentline
    le�onfontsize = (36 * Le�on2et3.Width) / 11746
    Module_routines.AdjustWidthAndSize Le�on2et3, 1
End If
End Sub

' ********************  LOAD suite ********************************************
Private Sub Timer1_Timer()
f1msgform = 0 'Pour accepter F1 pendant les exercices sur les fautes qui appellent le�on2et3
Le�on2et3.text1.Text = currentline
nbcaras = Len(Le�on2et3.text1.Text)
iwrong = 0: iwrongbis = 0: iter = 0
cara1 = "": cara2 = ""
Le�on2et3.text1.SelStart = 0
Le�on2et3.text1.SelLength = Len(Le�on2et3.text1.Text)
' Sleep ferait faire double sono en Win 98
'Call Sleep(cadenceligne)
Module_routines.text2text1 1, sonocara, 0, 0, 0, 0
Timer1.Enabled = False
End Sub


' **********  TEXT1_KEYUP Events rares sur text1 ***************************
Private Sub text1_KeyUp(Keycode As Integer, shift As Integer)
If Keycode = 27 Then
    Quitter_Click
    Exit Sub
End If
'Curseur sur text1 interdit, passer � text2
On Error Resume Next
Le�on2et3.text2.SetFocus
End Sub


' ****************  TEXT2_KEYDOWN  ******************************************
Private Sub Text2_KeyDown(Keycode As Integer, shift As Integer)
' TOUCHES � PB, annule la commande r�alis�e simultan�ment par Windows
' Win 91 et Win 92 (voir en plus Text2_KeyUp)
If Keycode = 91 Or Keycode = 92 Then
    Module_routines.cancelwin 0, Le�on2et3, 0
    Exit Sub
End If

' �chappement
If Keycode = 27 Then Exit Sub

'RetArr interdit
If Keycode = 8 Then
    Module_routines.bip Le�on2et3
    t2inhibit = 1
    ii = ii + 1
    Le�on2et3.text2.Text = Left(Le�on2et3.text1.Text, ii)
    t2inhibit = 0
    Le�on2et3.text2.SelStart = Len(Le�on2et3.text2.Text)
End If

'Combinaison Maj+ESPACE (r�p�tera la (FIN de la) ligne)
If Keycode = 32 And shift = 1 Then lrepeat = 1

'Combinaison Control+ESPACE (r�p�tera le mot)
If Keycode = 32 And shift = 2 Then wrepeat = 1

'Touche Alt+F4 pour quitter
If Keycode = 115 And shift = 4 Then
    altf4 = 1
    If quitactive = 0 Then Module_routines.QuitQuit
End If
End Sub

' *************** TEXT2_KEYUP  ****************************************************
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
        Le�on2et3.text4.Visible = False
        Le�on2et3.text1.SelLength = 0
        Call Sleep(cadencecara)
        
        ' La touche Alt lance Echap qui ne lance la s�lection de la fin du text1 que si noalt=0
        If noalt = 1 Then
            Le�on2et3.text1.SelLength = 1
        Else
            Le�on2et3.text1.SelLength = Len(Le�on2et3.text1.Text)
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
    Module_routines.cancelwin 1, Le�on2et3, 0
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
If Keycode = 112 Then Module_routines.help_f1 Le�on2et3

'Touche F2
If Keycode = 113 Then
    keyinhibit = 2
    avecf2 = 1
    help_f2 Le�on2et3
    avecf2 = 0
End If

'Touche F3
If Keycode = 114 Then Module_routines.help_f3 Le�on2et3

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
    Module_routines.bip Le�on2et3
    Le�on2et3.text2.SelStart = Len(Le�on2et3.text2.Text)
End If
End Sub


' **************************  TEXT2_CHANGE  *************************************
Private Sub Text2_Change()
Module_routines.text2text1 1, sonocara, 0, 0, 0, 0
End Sub


' *************************  QUITTER  ********************************************
Private Sub Quitter_Click()
Module_routines.quit_l
End Sub

' ******************  QUITTER par le BOUTON  ************************************
Private Sub Quitter_KeyPress(KeyAscii As Integer)
If KeyAscii = 81 Or KeyAscii = 113 Then Quitter_Click
End Sub

