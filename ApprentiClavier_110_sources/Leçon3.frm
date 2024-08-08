VERSION 4.00
Begin VB.Form Leçon3 
   ClientHeight    =   5535
   ClientLeft      =   1185
   ClientTop       =   2100
   ClientWidth     =   9645
   Height          =   6405
   Left            =   1125
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   9645
   Top             =   1290
   Width           =   9765
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
      Height          =   480
      HideSelection   =   0   'False
      Left            =   3750
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.TextBox Text4 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      HideSelection   =   0   'False
      Left            =   150
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   600
      Visible         =   0   'False
      Width           =   9315
   End
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   8160
      Top             =   1200
   End
   Begin VB.CommandButton Quitter 
      Caption         =   "&Quitter"
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
      Left            =   7050
      TabIndex        =   3
      Top             =   4050
      Width           =   1515
   End
   Begin VB.TextBox Text2 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1950
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   2700
      Width           =   4815
   End
   Begin VB.TextBox Text1 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      HideSelection   =   0   'False
      Left            =   1950
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1800
      Width           =   4815
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   450
      TabIndex        =   4
      Top             =   4650
      Width           =   4365
   End
   Begin VB.Label Label1 
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
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Menu AideGénérale 
      Caption         =   "F1=Aide générale"
   End
   Begin VB.Menu Description 
      Caption         =   "F2=Description de la touche"
   End
   Begin VB.Menu AideMémoire 
      Caption         =   "F3=Aide-Mémoire"
   End
End
Attribute VB_Name = "Leçon3"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' ********************  LOAD  *************************************************
Private Sub Form_Load()
' Prototype d'un DERNIER exo appelant un texte de lettres seules, de mots ou de lignes
cadencecara = 300: cadencemot = 200: cadenceligne = 300
Label2.Caption = "Utilisateur : " & nom & CRLF & "Niveau : " & nivo
If Not EOF(1) Then
    Line Input #1, currentline
    Leçon3.text1.Text = currentline
    Leçon3.text1.SelStart = 0
    Leçon3.text1.SelLength = Len(Leçon3.text1.Text)
    nbcaras = Len(Leçon3.text1.Text): iwrong = 0: iwrongbis = 0
    iter = 0
Else
    Close #1
End If
End Sub

' ********************  LOAD suite ********************************************
Private Sub Timer1_Timer()
Module_global.text2text1 Leçon3, menu_courant, 1
Timer1.Enabled = False
End Sub


' **********  TEXT1_KEYDOWN Events rares sur text1 ***************************
Private Sub text1_keyDown(keycode As Integer, shift As Integer)
If keycode = 27 Then
    Quitter_Click
    Exit Sub
End If
'Curseur sur text1 interdit, passer à text2
Leçon3.Text2.SetFocus
End Sub


' ****************  TEXT2_KEYDOWN  ******************************************
Private Sub Text2_KeyDown(keycode As Integer, shift As Integer)
'Echappement
If keycode = 27 Then
    If keyinhibit = 0 Then Quitter_Click
End If

'RetArr interdit
If keycode = 8 Then
    Beep
    t2inhibit = 1
    ii = ii + 1
    Leçon3.Text2.Text = Left(Leçon3.text1.Text, ii)
    t2inhibit = 0
    Leçon3.Text2.SelStart = Len(Leçon3.Text2.Text)
End If

'Combinaison Maj-ESPACE (répétera la ligne)
If keycode = 32 And shift = 1 Then lrepeat = 1

'Combinaison Control-ESPACE (répétera le mot)
If keycode = 32 And shift = 2 Then wrepeat = 1

'Touche F1
If keycode = 112 Then help_f1
'Touche F2
If keycode = 113 Then help_f2 Leçon3
'Touche F3
If keycode = 114 Then help_f3 Leçon3
End Sub


' *************** TEXT2_KEYUP  ****************************************************
Private Sub Text2_KeyUp(keycode As Integer, shift As Integer)
' Les msgbox procurent des key-ups indésirables avec les 3 commandes Entrée Oui Non
If keyinhibit = 1 Then
    keyinhibit = 0
    If keycode = 13 Or keycode = 78 Or keycode = 79 Then Exit Sub
End If

' TOUCHES à PB, annule la commande réalisée simultanément par Windows
' AltDroit 17 (qui envoie ensuite 18)
If keycode = 17 Then Exit Sub
' AltGauche 18
If keycode = 18 Then
    keyinhibit = 1
    SendKeys "{ESC}"
    echapbis = echapbis - 1
    Exit Sub
End If
' Sendkeys termine par 145
If keycode = 145 Then Exit Sub
    
'Flèche gauche interdite
If keycode = 37 Then
    Beep
    Leçon3.Text2.SelStart = Len(Leçon3.Text2.Text)
End If
End Sub


' **************************  TEXT2_CHANGE  *************************************
Private Sub Text2_Change()
'Module_global.text2text1 Leçon3, Menu_leçon3, 1
Module_global.text2text1 Leçon3, menu_courant, 1
End Sub


' *************************  QUITTER  ********************************************
Public Sub Quitter_Click()
Unload Leçon3
Close #1
iter = 0

' Faut-il PASSER au MENU de la LEçON SUIVANTE ?
If nextleçon = 0 Then
    Unload menu_courant
    menu_courant.Show
End If
If nextleçon = 1 Then
    If nivo = "Standard" Then kk = 0
    If nivo = "Modifié" Then kk = 25
    If pctok(numleçon + kk, 0) = 0 Then
        Unload menu_courant
        menu_courant.Show
    Else
        vmsgbox = MsgBox("La leçon  " + Str(numleçon - 2) + "  est terminée, avec une réussite moyenne de " + Str(pctok(numleçon + kk, 0)) + " pourcent." + CRLF + CRLF + "Voulez-vous PASSER à la leçon SUIVANTE, OUI ou NON ?", 3, "")
        If vmsgbox = 7 Or vmsgbox = 2 Then
            numexo = numexo - 1
            Unload menu_courant
            menu_courant.Show
        Else
            numexo = 0
            Unload menu_suivant
            menu_suivant.Show
        End If
    End If
    nextleçon = 0
End If
End Sub

' ******************  QUITTER par le BOUTON  ************************************
Private Sub Quitter_KeyPress(KeyAscii As Integer)
If KeyAscii = 81 Or KeyAscii = 113 Then Quitter_Click
End Sub

' *****************  Commandes de la BARRE de MENUS   ***************************
Private Sub AideGénérale_Click()
Module_global.help_f1
End Sub

Private Sub Description_Click()
Module_global.help_f2 Leçon3
End Sub

Private Sub AideMémoire_Click()
Module_global.help_f3 Leçon3
End Sub

