VERSION 4.00
Begin VB.Form Aidef3 
   BackColor       =   &H00004080&
   Caption         =   "Mode Aide-mémoire, Échap pour sortir"
   ClientHeight    =   5100
   ClientLeft      =   585
   ClientTop       =   1665
   ClientWidth     =   9900
   ControlBox      =   0   'False
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
      Name            =   "Arial Narrow"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Height          =   5610
   Left            =   525
   LinkTopic       =   "Form1"
   ScaleHeight     =   5100
   ScaleWidth      =   9900
   Top             =   1215
   Visible         =   0   'False
   Width           =   10020
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   8760
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   7200
      Top             =   1200
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      HideSelection   =   0   'False
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   450
      Visible         =   0   'False
      Width           =   9855
   End
   Begin VB.TextBox Text3 
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
      HideSelection   =   0   'False
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Quitter 
      BackColor       =   &H00C0C0FF&
      Caption         =   "&Quitter l'aide-mémoire (Échap)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Left            =   7560
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   30
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      HideSelection   =   0   'False
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1800
      Width           =   9435
   End
   Begin VB.Label Label2 
      BackColor       =   &H00808080&
      Caption         =   "Label2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   345
      Left            =   0
      TabIndex        =   2
      Top             =   4560
      Width           =   7455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tapez :"
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
      Left            =   360
      TabIndex        =   0
      Top             =   1200
      Width           =   1125
   End
End
Attribute VB_Name = "Aidef3"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' ********************  LOAD  *************************************************
Private Sub Form_Load()
Aidef3.Caption = msgAideF3
Aidef3.label1 = msgTapez2
Quitter.Caption = msgQuitterAM & msgÉchap

' Aide-mémoire sur chaque touche
cadencecara = 300
keyinhibit = 0
mcinhibit = 0
avecf3 = 1

'Ne pas Setter leçon_courante
Module_routines.Dimension Aidef3
Label2.Caption = bannerVersion & ", " & bannerCopyright
End Sub

' ********************  LOAD suite ********************************************
Private Sub Timer1_Timer()
On Error Resume Next
Aidef3.text1.SetFocus
Aidef3.text1.Width = 0.35 * Aidef3.Width * zoomvalue '12/2011
Timer1.Enabled = False
End Sub

' ******  LEVE l'INHIBIT après codes parasites LOCKS engendrés par menu-contextuel *********
Private Sub Timer2_Timer()
mcinhibit = 0
Timer2.Enabled = False
End Sub

' ****************  TEXT1_KEYDOWN  ******************************************
Private Sub text1_KeyDown(Keycode As Integer, shift As Integer)
'Debug.Print "keycodef3down=" & Keycode
' Reset
Aidef3.text1.Visible = "false"
Aidef3.text4.Visible = "false"
If mcinhibit = 0 Then Aidef3.text1.Text = "" ' Pas de reset dans le cas Menu-Contextuel
If mcinhibit = 0 Then Aidef3.text4.Text = ""
Aidef3.Cls
Call Sleep(20)
Aidef3.text1.Visible = "true"
Aidef3.text4.Visible = "true"

' Win 91 et Win 92 (voir en plus text1_KeyUp)
If Keycode = 91 Or Keycode = 92 Then
    If menucase = 0 Then msgpb.Show 1 ' indispensable
    keyinhibit = 1
    'Call Sleep(400)  ' indispensable en 2004, supprimé en juin 2007
    echapbis = -1
    'SendKeys "{ESC}", True
    'Sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    echapbis = 0
    msgpb.Show 1 ' indispensable
    If Keycode = 91 Then Aidef3.text1.Text = vvWindowsGauche & " " 'Alt255
    If Keycode = 92 Then Aidef3.text1.Text = vvWindowsDroit & " " 'Alt255
    keyinhibit = 0 'septembre 2007
    Exit Sub

'Tab
ElseIf Keycode = 9 And shift = 0 Then: Aidef3.text1.Text = vvTabulationAvant & " " 'Alt255
ElseIf Keycode = 9 And shift = 1 Then: Aidef3.text1.Text = vvTabulationArrière & " " 'Alt255

'Control et Alt
ElseIf Keycode = 17 Then: Aidef3.text1.Text = vvControl & " " 'Alt255
ElseIf Keycode = 18 Then: Aidef3.text1.Text = vvAltOuAltGr & " " 'Alt255

' Échappement
ElseIf Keycode = 27 Then
    Aidef3.text1.Text = tempo
    Exit Sub

'Autres touches
ElseIf Keycode = 13 Then: Aidef3.text1.Text = vvEntrée & " " 'Alt255
ElseIf Keycode = 16 Then: Aidef3.text1.Text = vvMaj & " " 'Alt255. Ne pas mettre en KeyUp sinon Maj+X montrerait MAJ au lieu de X
ElseIf Keycode = 33 Then: Aidef3.text1.Text = vvPagePrécédente & " " 'Alt255
ElseIf Keycode = 34 Then: Aidef3.text1.Text = vvPageSuivante & " " 'Alt255
ElseIf Keycode = 35 Then: Aidef3.text1.Text = vvFin & " " 'Alt255
ElseIf Keycode = 36 Then: Aidef3.text1.Text = vvDébut & " " 'Alt255
ElseIf Keycode = 45 Then: Aidef3.text1.Text = vvInsertion & " " 'Alt255
ElseIf Keycode = 46 Then: Aidef3.text1.Text = vvSuppression & " " 'Alt255
ElseIf Keycode = 93 Then: Aidef3.text1.Text = vvMenuContextuel & " " 'Alt255

' Touches LOCKS : attention, menu-contextuel renvoie aussi l'état des locks, d'où le mcinhibit
ElseIf Keycode = 20 And shift = 0 And mcinhibit = 0 Then: Aidef3.text1.Text = vvVerrouillageMajuscules & " " 'Alt255
ElseIf Keycode = 144 And mcinhibit = 0 Then: Aidef3.text1.Text = vvVerrouillageNumérique & " " 'Alt255
ElseIf Keycode = 145 And mcinhibit = 0 Then: Aidef3.text1.Text = vvArrêtDéfil & " " 'Alt255

'Touches F1, F2, pas F3!, F4 à F12
ElseIf Keycode = 112 Then: Aidef3.text1.Text = "F1 "
ElseIf Keycode = 113 Then: Aidef3.text1.Text = "F2 "
ElseIf Keycode = 114 Then: Aidef3.text1.Text = "F3 "
ElseIf Keycode = 115 Then: Aidef3.text1.Text = "F4 "
ElseIf Keycode = 116 Then: Aidef3.text1.Text = "F5 "
ElseIf Keycode = 117 Then: Aidef3.text1.Text = "F6 "
ElseIf Keycode = 118 Then: Aidef3.text1.Text = "F7 "
ElseIf Keycode = 119 Then: Aidef3.text1.Text = "F8 "
ElseIf Keycode = 120 Then: Aidef3.text1.Text = "F9 "
ElseIf Keycode = 121 Then: Aidef3.text1.Text = "F10 "
ElseIf Keycode = 122 Then: Aidef3.text1.Text = "F11 "
ElseIf Keycode = 123 Then: Aidef3.text1.Text = "F12 "

'Commun
Else
    If mcinhibit = 0 Then ' Pas de reset dans le cas Menu-Contextuel
        Aidef3.text1.Text = ""
        Aidef3.text4.Text = ""
    End If
End If
Aidef3.text1.SelStart = 0
Aidef3.text1.SelLength = Len(Aidef3.text1.Text)
End Sub


' *************** TEXT1_KEYUP  ****************************************************
Private Sub text1_KeyUp(Keycode As Integer, shift As Integer)
' Échappement
If Keycode = 27 Then
    If keyinhibit = 0 Then
        avecf3 = 0
        echapbis = 0
        Quitter_Click
    End If
End If

'12/2011 Affichage
If Len(Aidef3.text1.Text) > 16 Then
    Aidef3.text1.Width = 0.85 * Aidef3.Width * zoomvalue
ElseIf Len(Aidef3.text1.Text) > 10 Then
    Aidef3.text1.Width = 0.55 * Aidef3.Width * zoomvalue
Else
    Aidef3.text1.Width = 0.35 * Aidef3.Width * zoomvalue
End If

' Les msgbox procurent des key-ups indésirables avec les 3 commandes Entrée Oui Non
If keyinhibit > 0 Then
    keyinhibit = keyinhibit - 1
    If Keycode = 13 Or Keycode = 78 Or Keycode = 79 Then Exit Sub
End If

'Touche Alt+F4 pour quitter
If Keycode = 115 And shift = 4 Then
    altf4 = 1
    If quitactive = 0 Then Module_routines.QuitQuit
End If

' TOUCHES à PB, annule la commande réalisée simultanément par Windows
' Win 91 et Win 92 (voir en plus text1_KeyDown)
If Keycode = 91 Or Keycode = 92 Then
    keyinhibit = 1
    'SendKeys "{ESC}", True
    'sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    msgpb.Show 1
    On Error Resume Next
    Aidef3.text1.SetFocus
    echapbis = 0
    If Keycode = 91 Then Aidef3.text1.Text = vvWindowsGauche & " " 'Alt255
    If Keycode = 92 Then Aidef3.text1.Text = vvWindowsDroit & " " 'Alt255
    keyinhibit = 0 'septembre 2007
    Exit Sub

' AltGr 17 (qui envoie ensuite 18 et évent-t 27)
ElseIf Keycode = 17 Then
    echapoff = -1
    Exit Sub

' AltGauche 18 et Menu-Contextuel 93
ElseIf Keycode = 93 Or Keycode = 18 Then
    tempo = Aidef3.text1.Text
    keyinhibit = 1
    mcinhibit = 1
    Timer2.Enabled = True
    echapbis = -1
    'SendKeys "{ESC}" ' Pour évacuer le code 18 Jaws qui fait suite à Echap pour sortir, et pour évacuer le menu-contextuel
    'sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    echapbis = 0 + echapoff
    echapoff = 0
    Exit Sub

' Autres touches spéciales
ElseIf Keycode = 8 Then: Aidef3.text1.Text = vvRetourArrière & " " 'Alt255
ElseIf Keycode = 19 Then: Aidef3.text1.Text = vvPause & " " 'Alt255
ElseIf Keycode = 37 Then: Aidef3.text1.Text = vvFlecheGauche & " " 'Alt255
ElseIf Keycode = 38 Then: Aidef3.text1.Text = vvFlecheHaut & " " 'Alt255
ElseIf Keycode = 39 Then: Aidef3.text1.Text = vvFlecheDroite & " " 'Alt255
ElseIf Keycode = 40 Then: Aidef3.text1.Text = vvFlecheBas & " " 'Alt255
ElseIf Keycode = 44 Then: Aidef3.text1.Text = vvImpression & " " 'Alt255

'Commun
End If
Aidef3.text1.SelStart = 0
Aidef3.text1.SelLength = 0
End Sub


' **************************  TEXT1_CHANGE  *************************************
Private Sub Text1_Change()
Aidef3.text1.SelStart = 0
Aidef3.text1.SelLength = 0
iiold = ii
ii = 0
Module_global.help_f2 Aidef3
ii = iiold
End Sub


' *************************  QUITTER  ********************************************
Private Sub Quitter_Click()
Unload Aidef3
tempo = ""
keyinhibit = 1 'Pour annuler l'Échap keyup de la leçon 1
If inexo = 1 Then
    On Error Resume Next
    leçon_courante.text2.SetFocus
End If
End Sub

' ******************  QUITTER par le BOUTON  ************************************
Private Sub Quitter_KeyPress(KeyAscii As Integer)
If KeyAscii = 81 Or KeyAscii = 113 Then Quitter_Click
End Sub

