VERSION 4.00
Begin VB.Form leçon1 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   6450
   ClientLeft      =   990
   ClientTop       =   1485
   ClientWidth     =   9810
   ControlBox      =   0   'False
   Height          =   6960
   Left            =   930
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9810
   Top             =   1035
   Visible         =   0   'False
   Width           =   9930
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9000
      Top             =   4560
   End
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
      Left            =   6960
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   480
      Width           =   1215
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
      Left            =   5400
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1650
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   600
      Picture         =   "leçon1.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   450
      Visible         =   0   'False
      Width           =   615
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
      Top             =   1200
      Visible         =   0   'False
      Width           =   9765
   End
   Begin VB.CommandButton Quitter 
      Caption         =   " &Quitter vers Menu   (Échap)"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   7800
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5040
      Width           =   1965
   End
   Begin VB.Timer Timer1 
      Interval        =   2400
      Left            =   9000
      Top             =   1650
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   480
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3450
      Width           =   8115
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      HideSelection   =   0   'False
      Left            =   480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2280
      Width           =   8115
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
      Width           =   8415
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
      Top             =   600
      Width           =   945
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
      Width           =   7605
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tapez la touche :"
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1650
      Width           =   2535
   End
End
Attribute VB_Name = "leçon1"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' **********************  LOAD  ****************************************
Private Sub Form_Load()
If FullScreenSwitch = 1 Then WindowState = 2

' Paramètres
cadencecara = 300: cadenceligne = 350
typeleçon = 1
Set leçon_courante = leçon1
Module_routines.Colors leçon1  '12/2011
Module_routines.Dimension leçon1
'If repjawsnames = "" Then Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "."
'If repjawsnames <> "" Then Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "." & CRLF & msgSonori & repjawsnames & CRLF & msgSpeedExp & debexplilevel & msgSpeedGen & debgenlevel & "."
Module_routines.mshow leçon1 'avril 2008
label1.Caption = msgTapezTouche
Label3.Caption = bannerVersion & ", " & bannerCopyright
Label4.Caption = msgScore
Label5.Caption = msgF1F2F3
If echapbismax = 0 Then Quitter.Caption = msgQuitterVers + CRLF + bannerMenu + CRLF + msgÉchap
If echapbismax = 1 Then Quitter.Caption = msgQuitterVers + CRLF + bannerMenu + CRLF + msgÉchap2
If echapbismax > 1 Then Quitter.Caption = msgQuitterVers + CRLF + bannerMenu + CRLF + msgÉchap3

' Touches Windows demandées (avec ou sans Alt255)
On Error Resume Next
If UCase(Left(leçon1.text1.Text, Len(vvWindowsGauche))) = vvWindowsGauche Or UCase(Left(leçon1.text1.Text, Len(vvWindowsDroit))) = vvWindowsDroit Then Exit Sub

' Les 3 touches
Module_routines.SetKeys "CAPSLOCK_OFF"
Module_routines.SetKeys "SCROLLLOCK_OFF"
If numpad = 0 Then Module_routines.SetKeys "NUMLOCK_OFF"
If numpad >= 1 Or numpad = -1 Then
    winstop = 2
    Module_routines.SetKeys "NUMLOCK_ON"
End If

' Première currentline (après le bip 2 tons, sinon bip inaudible en Win98)
'If Dir(vpath & "sonbip2.exe") <> "" Then Module_exec.ExecAndWait vpath & "sonbip2.exe" 'avril 2008
Module_routines.sonbip2tons 'avril 2008
Module_routines.OpenAndSuffix exo_courant, 1
On Error Resume Next
If EOF(1) Then
    Close #1
    Exit Sub
Else
    Line Input #1, currentline
    leçonfontsize = (28 * leçon1.Width) / 11746
    Module_routines.AdjustWidthAndSize leçon1, 1
End If
End Sub


' **********************  LOAD  suite  ****************************************
Private Sub Timer1_Timer()
leçon1.text1.Text = currentline
nbcaras = 1: iwrong = 0: iwrongbis = 0
leçon1.text1.SelStart = 0
leçon1.text1.SelLength = Len(leçon1.text1.Text)

' Comparer
Module_routines.cara2ligne1 leçon1
Timer1.Enabled = False
End Sub


' ***************  Traitement d'une SEQUENCE de TOUCHES définie au bout du Timer9  ***********
Private Sub Timer9_Timer()
Timer9.Enabled = False

' VALEUR du KeyCode RÉEL
' Control
If KeyFirst = 17 And KeySecond = 0 Then Keycode = 17 '(avec ou sans JAWS)

' Alt
If KeyFirst = 18 And KeySecond = 0 Then Keycode = 18 '(sans JAWS)
If KeyFirst = 18 And KeySecond = 18 And KeyThird = 0 Then Keycode = 18 '(avec JAWS)

' AltGr (avec ou sans JAWS)
If KeyFirst = 17 And KeySecond = 18 And KeyThird = 0 Then Keycode = 255 '(avec JAWS451 et plus, ou sans JAWS)
If KeyFirst = 17 And KeySecond = 17 And KeyThird = 17 Then Keycode = 255 '(avec JAWS401)

' Echap (avec ou sans JAWS)
If KeyFirst = 27 And KeySecond = 0 Then Keycode = 27 '(sans JAWS)
If KeyFirst = 18 And KeySecond = 27 And KeyThird = 0 Then Keycode = 27 '(avec JAWS)

' ESPACE (avec demande text1.text="ESPACE")
If KeyFirst = 32 And KeySecond = 0 Then Keycode = 32 '(avec JAWS451 et plus, ou sans JAWS)
If KeyFirst = 32 And KeySecond = 17 Then Keycode = 32 '(avec JAWS401)

' Reset
'Debug.Print "Timer9Keycode : "; KeyFirst & " + " & KeySecond & " + " & KeyThird & " = " & Keycode
KeyFirst = 0: KeySecond = 0: KeyThird = 0

' BONNE RÉPONSE (code semblable à la subroutine text2_keyup ci-dessous) à la fin du temps timer9
If Keycode = KeyExpect Then
    
    ' Pour affichage immédiat du score
    Module_routines.scoreaffich leçon_courante, 0
                    
    ' Reset
    leçon1.text4.Text = ""
    leçon1.text4.Visible = False
    f2link = 0
L10bis:
    ' Bonne réponse, traiter jusqu'à obtenir la ligne suivante
    follow
Else
    ' ECHAP réitéré (sauf si Échap est la bonne réponse)
    If Keycode = 27 Then
        If f2link = 1 Then 'septembre 2007, sortie de F2
            leçon1.text4.Text = ""
            leçon1.text4.Visible = False
            f2link = 0
            keyinhibit = 0
        Else
            echapbis = echapbis + 1
            If echapbis > echapbismax Then
                echapbis = 0
                Quitter_Click
                Exit Sub
            End If
        End If
    End If
        
    ' FAUTE de FRAPPE
    If Keycode <> 20 Or (Keycode = 20 And UCase(leçon1.text1.Text) = "TAB ") Then 'Déverr Maj toléré sauf si on demande TAB, septembre 2007
        
        ' Demande RÉPÉTER par ESPACE, ajout septembre 2007
        If Keycode = 32 And espacevalid = 1 Then
            leçon1.text2.Text = ""
            leçon1.text1.SelStart = 0
            leçon1.text1.SelLength = 0
            Call Sleep(cadencemot)
            leçon1.text1.SelStart = 0
            leçon1.text1.SelLength = Len(leçon1.text1.Text)
            Exit Sub
        End If
                
        'Les demandes d'aide ne sont pas des fautes
        If Keycode = 27 Or Keycode = 112 Or Keycode = 113 Or Keycode = 114 Then Exit Sub

        'Vraies fautes
        If winstop = 0 Or UCase(leçon1.text1.Text) <> vvWindowsGauche Or UCase(leçon1.text1.Text) <> vvWindowsDroit Then
            Module_routines.bip leçon1
            iwrong = iwrong + 1: iwrongbis = iwrongbis + 1
            
            ' Pour affichage immédiat du score
            Module_routines.scoreaffich leçon_courante, 0
                    
            ' Reset de iwrongbis si l'utilisateur progresse
            iiprec = iter
            If iiprec > iiante Then
                iiante = iiprec
                iwrongbis = 1
            End If
            
            ' Reset de text2
            leçon1.text1.SelLength = 0
            Call Sleep(150)
            leçon1.text1.SelLength = Len(leçon1.text1)
            leçon1.text2.Text = ""
            leçon1.text2.SelStart = 0
            leçon1.Cls
            
            ' On atteint le nb max de fautes sur le cara
            If iwrongbis >= iwrongbismax Then GoTo L10bis
        End If
    End If
End If

End Sub


' *************  TEXT1_KEYDOWN Events rares sur text1 ***************************
Private Sub text1_KeyDown(Keycode As Integer, shift As Integer)
If Keycode = 27 Then
    Quitter_Click
    Exit Sub
End If

'Curseur sur text1 interdit, passer à text2
On Error Resume Next
leçon1.text2.SetFocus
End Sub


' ***********************  TEXT2 KEY_DOWN  ***********************************
Private Sub Text2_KeyDown(Keycode As Integer, shift As Integer)
'Debug.Print "KeyDown=" & Keycode & "   KeyInhibit=" & keyinhibit & "   Winstop=" & winstop

' Win 91 et Win 92 (voir en plus Text2_KeyUp)
If Keycode = 91 Or Keycode = 92 Then
    winstop = 4
    If (Keycode = 91 And UCase(Left(leçon1.text1.Text, Len(vvWindowsGauche))) = vvWindowsGauche) Or (Keycode = 92 And UCase(Left(leçon1.text1.Text, Len(vvWindowsDroit))) = vvWindowsDroit) Then
        Module_routines.cancelwin 1, leçon1, 0
L8: '(septembre 2007)
        ' Reset
        leçon1.text4.Text = ""
        leçon1.text4.Visible = False
        f2link = 0
        follow
    Else
        Module_routines.cancelwin 0, leçon1, 0
        iwrong = iwrong + 1: iwrongbis = iwrongbis + 1
        ' Pour affichage immédiat du score
        Module_routines.scoreaffich leçon_courante, 0
        ' On atteint le nb max de fautes (septembre 2007)
        If iwrongbis >= iwrongbismax Then GoTo L8
                    
    End If
    Exit Sub
End If

' Pour annuler l'effet de la touche Tab qui n'est pas sensible à KeyUp
If Keycode = 9 Then Module_routines.pasdetab

'Reset picture à conserver!
Picture1.Visible = False

'Échappement n'est pas traité ici, car Échap est en apprentisssage
'RetArr inutile

'Touche Alt+F4 pour quitter
If Keycode = 115 And shift = 4 Then
    altf4 = 1
    If quitactive = 0 Then Module_routines.QuitQuit
End If
End Sub


' *************** TEXT2_KEYUP  ****************************************************
Private Sub Text2_KeyUp(Keycode As Integer, shift As Integer)
'Debug.Print "1KeyUp=" & Keycode & " KeyInhibit=" & keyinhibit & " avecf2=" & avecf2 & " f2link=" & f2link & " Winstop=" & winstop & " keyexpect=" & KeyExpect

' TRAITEMENT des "SÉQUENCES de TOUCHES" dues à JAWS pour Echap, Alt, Control
' Après appel de F2 (f2link = 1), il faut passer par ce traitement, sinon pb sur AltGr par exemple (septembre 2007)
If (Keycode = 17 Or Keycode = 18 Or Keycode = 27 Or Keycode = 32) And (keyinhibit = 0 Or (keyinhibit = 1 And f2link = 1)) Then 'septembre 2007
    If KeyFirst <> 0 And KeySecond <> 0 And KeyThird = 0 Then KeyThird = Keycode
    If KeyFirst <> 0 And KeySecond = 0 Then KeySecond = Keycode
    If KeyFirst = 0 Then
        KeyFirst = Keycode
        Timer9.Enabled = True
    End If
End If

' STOP ICI si on a une "SÉQUENCE DE TOUCHES" (alors, c'est timer9 qui traite la réponse)!!!
If Timer9.Enabled = True Then Exit Sub

' Attention, permet de basculer réellement la touche NumLock avec Win98 + Jaws !
If numpad >= 1 Then Module_routines.SetKeys "NUMLOCK_ON"

' Winstop
If winstop > 0 Then winstop = winstop - 1
If winstop = 2 Then keyinhibit = 0  ' Pour que Alt (juste après Win) n'emmène pas sur la barre des menus

' Touche Tab, suite à module_routines.pasdetab, interdite ou engendre éventuel code 144
If Keycode = 144 Then
    If numpad = 0 Then Exit Sub
    If numpad >= 1 And keyinhibit = 1 Then Exit Sub
End If
If keyforce = 9 Then
    leçon_courante.text2.Text = ""
    leçon_courante.text2.SelStart = 0
    keyforce = 0
    keyinhibit = 0
    If notab = 1 Then
        GoTo L13
    Else
        Keycode = 9
    End If
End If

'Touche F2 = Aide, sauf si la leçon demande F1 F2 ou F3 (avec ou sans Alt255) (placer avant le traitement msgbox/keyinhibit=1)
If Keycode = 113 Then
    If UCase(leçon1.text1.Text) = "F2" Or UCase(leçon1.text1.Text) = "F2 " Then GoTo L9
    If UCase(leçon1.text1.Text) <> "F1" And UCase(leçon1.text1.Text) <> "F1 " And UCase(leçon1.text1.Text) <> "F3" And UCase(leçon1.text1.Text) <> "F3 " Then
        keyinhibit = 2
        avecf2 = 1
        help_f2 leçon1
        avecf2 = 0
    Else
        GoTo L13
    End If
End If

' Les msgbox procurent des key-ups indésirables
If keyinhibit = 1 Then
    keyinhibit = 0
    If Keycode = 13 Or Keycode = 78 Or Keycode = 79 Then Exit Sub 'Entrée ou N=Non ou O=Oui
    leçon1.text4.Visible = False
    If f2link = 1 And Keycode = 27 Then Exit Sub
    If f2link = 0 Then Exit Sub
End If

' Pour Échap de la touche F2
If keyinhibit = 2 Then keyinhibit = 1

' TOUCHES à PB, annule la commande réalisée simultanément par Windows
' Win 91 et Win 92 (voir en plus Text2_KeyDown)
If Keycode = 91 Or Keycode = 92 Then
    winstop = 4
    Module_routines.cancelwin 1, leçon1, 0
    If (Keycode = 91 And UCase(Left(leçon1.text1.Text, Len(vvWindowsGauche))) = vvWindowsGauche) Or (Keycode = 92 And UCase(Left(leçon1.text1.Text, Len(vvWindowsDroit))) = vvWindowsDroit) Then
        GoTo L9
    End If
    Exit Sub
End If

' Menu-Contextuel 93
If Keycode = 93 Then
    winstop = 0 'évite interaction avec touche windows précédemment utilisée, septembre 2007
    echapbis = -1
    keyinhibit = 1
    forcepause = 2 ' Pour ne pas quitter msgform quand un message d'explications suit, et pour ne pas échapper du score dans msgform si la leçon se termine par Alt ou AltGr
    'SendKeys "{ESC}"
    'Sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    echapbis = -1
    
    ' Menu-Contextuel non demandé
    If UCase(Left(leçon1.text1.Text, Len(vvMenuContextuel))) <> vvMenuContextuel And Keycode = 93 Then GoTo L13
End If

' Sendkeys termine par 145
If Keycode = 145 And Left(leçon1.text1.Text, 10) <> vvArrêtDéfil Then
    If Left(leçon1.text1.Text, Len(vvArrêtDéfil)) <> vvArrêtDéfil And UCase(Left(leçon1.text1.Text, Len(vvPause))) <> vvPause And UCase(Left(leçon1.text1.Text, Len(vvImpression))) <> vvImpression Then
        iwrong = iwrong - 1
        If iwrong < 0 Then iwrong = 0
        iwrongbis = iwrongbis - 1
        If iwrongbis < 0 Then iwrongbis = 0
        Exit Sub
    End If
End If

'Touche F1
If Keycode = 112 Then
    If UCase(leçon1.text1.Text) = "F1" Or UCase(leçon1.text1.Text) = "F1 " Then GoTo L9
    If UCase(leçon1.text1.Text) <> "F2" And UCase(leçon1.text1.Text) <> "F2 " And UCase(leçon1.text1.Text) <> "F3" And UCase(leçon1.text1.Text) <> "F3 " Then
        noechapF1 = 1
        Module_routines.help_f1 leçon1
    Else
        GoTo L13
    End If
End If

'Touche F3
If Keycode = 114 Then
    If UCase(leçon1.text1.Text) = "F3" Or UCase(leçon1.text1.Text) = "F3 " Then GoTo L9
    If UCase(leçon1.text1.Text) <> "F1" And UCase(leçon1.text1.Text) <> "F1 " And UCase(leçon1.text1.Text) <> "F2" And UCase(leçon1.text1.Text) <> "F2 " Then
        Module_routines.help_f3 leçon1
    Else
        GoTo L13
    End If
End If

'Touche F10 mène à la barre menu
If Keycode = 121 Then
    If UCase(leçon1.text1.Text) = "F10" Or UCase(leçon1.text1.Text) = "F10 " Then GoTo L9 'réparation bug, juin 2007
    echapbis = echapbis - 1
    'SendKeys "{ESC}", True
    'Sendkeys est remplacé par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
End If

' BONNE RÉPONSE
'Debug.Print "6KeyUp=" & Keycode & "  KeyInhibit=" & keyinhibit & "  f2link=" & f2link & "  Winstop=" & winstop & "  keyexpect=" & KeyExpect
If Keycode = KeyExpect Then
L9:
    ' Pour affichage immédiat du score
    Module_routines.scoreaffich leçon_courante, 0
                    
    ' Reset
    leçon1.text4.Text = ""
    leçon1.text4.Visible = False
    f2link = 0
L10:
    ' Bonne réponse, traiter jusqu'à obtenir la ligne suivante
    follow
Else
    ' ECHAP réitéré (sauf si Échap est la bonne réponse)
    If Keycode = 27 Then
        echapbis = echapbis + 1
        If echapbis > echapbismax Then
            echapbis = 0
            Quitter_Click
            Exit Sub
        End If
    End If
    
    ' FAUTE de FRAPPE
    If Keycode <> 20 Or (Keycode = 20 And UCase(leçon1.text1.Text) = "TAB ") Then 'Déverr Maj toléré sauf si on demande TAB, septembre 2007
        
        ' Demande RÉPÉTER par ESPACE
        If Keycode = 32 And espacevalid = 1 Then
            leçon1.text2.Text = ""
            leçon1.text1.SelStart = 0
            leçon1.text1.SelLength = 0
            Call Sleep(cadencemot)
            leçon1.text1.SelStart = 0
            leçon1.text1.SelLength = Len(leçon1.text1.Text)
            Exit Sub
        End If
        
        'Les demandes d'aide ne sont pas des fautes
        If Keycode = 27 Or Keycode = 112 Or Keycode = 113 Or Keycode = 114 Then Exit Sub

L13:
        'Vraies fautes
        If winstop = 0 Then
            Module_routines.bip leçon1
            iwrong = iwrong + 1: iwrongbis = iwrongbis + 1
            
            ' Pour affichage immédiat du score
            Module_routines.scoreaffich leçon_courante, 0
                    
            ' Reset de iwrongbis si l'utilisateur progresse
            iiprec = iter
            If iiprec > iiante Then
                iiante = iiprec
                iwrongbis = 1
            End If
            
            ' Reset de text2
            leçon1.text1.SelLength = 0
            Call Sleep(150)
            leçon1.text1.SelLength = Len(leçon1.text1)
            leçon1.text2.Text = ""
            leçon1.text2.SelStart = 0
            leçon1.Cls
            
            ' On atteint le nb max de fautes sur le cara
            If iwrongbis >= iwrongbismax Then GoTo L10
        End If
    End If
End If
End Sub


' ******************************  follow  ***********************************************
Private Sub follow()
    iwrongbis = 0 '(septembre 2007)
    If UCase(leçon1.text1.Text) = UCase(vvAlt) Or UCase(leçon1.text1.Text) = UCase(vvAlt) & " " Or UCase(leçon1.text1.Text) = UCase(vvAltGr) Or UCase(leçon1.text1.Text) = UCase(vvAltGr) & " " Or UCase(Left(leçon1.text1.Text, Len(vvWindowsGauche))) = vvWindowsGauche Or UCase(Left(leçon1.text1.Text, Len(vvWindowsDroit))) = vvWindowsDroit Then
    Else
        echapbis = 0
    End If
    iter = iter + 1
    
    ' Msg d'encouragement "Continuez "
    If Not msgtext1(iter) = "" Then
        leçon1.label1.Visible = False
        leçon1.text1.Text = msgtext1(iter)
        currentline = leçon1.text1.Text
        Module_routines.AdjustWidthAndSize leçon_courante, 0
        leçon1.text1.SelStart = 0
        leçon1.text1.SelLength = Len(leçon1.text1)
        If Len(msgtext1(iter)) < 10 Then
            Call Sleep(800)
        Else
            Call Sleep(1200)
        End If
        leçon1.label1.Visible = True
    End If

    ' Msg d'explications
    If Not msgtext2(iter) = "" Then
L11:
        pagenum = 0
        msgtext0 = msgtext2(iter) + pressez
        Msgform.Show 1
        If msgf = 2 Then GoTo L11
        If msgf = 1 Then
            keyinhibit = 0
            'If Dir(vpath & "sonbip2.exe") <> "" Then Module_exec.ExecAndWait vpath & "sonbip2.exe"
            Module_routines.sonbip2tons  ' avril 2008
        End If
    End If
        
    ' Bonne réponse, suite, va chercher la ligne suivante
    Module_routines.lignesuivante 0, 0, 0
    msgtext1(iter) = ""
    msgtext2(iter) = ""
    nbcaras = nbcaras + 1
    If derligne = 2 Then
        derligne = 0
        iter = 0
        Exit Sub
    End If
    
    ' Bonne réponse, suite, va chercher le keyexpect suivant
    Module_routines.cara2ligne1 leçon1
End Sub


' ********************  QUITTER  ***********************************************
Private Sub Quitter_Click()
Module_routines.quit_l
notab = 1
End Sub


' *********** Quitter par le bouton ********************************************
Private Sub Quitter_KeyPress(KeyAscii As Integer)
If KeyAscii = 81 Or KeyAscii = 113 Then Quitter_Click
End Sub

