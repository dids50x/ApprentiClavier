VERSION 4.00
Begin VB.Form le�on13 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   6450
   ClientLeft      =   615
   ClientTop       =   1665
   ClientWidth     =   9810
   ControlBox      =   0   'False
   Height          =   6960
   Left            =   555
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   9810
   Top             =   1215
   Visible         =   0   'False
   Width           =   9930
   Begin VB.Timer Timer9 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9000
      Top             =   4320
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
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   720
      Width           =   5415
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
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   3915
   End
   Begin VB.Timer Timer2 
      Interval        =   2050
      Left            =   9000
      Top             =   3150
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   750
      Picture         =   "le�on13.frx":0000
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   8
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1200
      Visible         =   0   'False
      Width           =   9765
   End
   Begin VB.CommandButton Quitter 
      Caption         =   "  &Quitter vers   Menu   (�chap)"
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
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4920
      Width           =   1965
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   9000
      Top             =   2400
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1500
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   3450
      Width           =   5415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      HideSelection   =   0   'False
      Left            =   1500
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2400
      Width           =   5415
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
      Left            =   300
      TabIndex        =   11
      Top             =   0
      Width           =   8265
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
      Left            =   2640
      TabIndex        =   10
      Top             =   720
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
      TabIndex        =   7
      Top             =   4440
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
      Left            =   3120
      TabIndex        =   6
      Top             =   1800
      Width           =   2445
   End
End
Attribute VB_Name = "le�on13"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' **********************  LOAD  ****************************************
Private Sub Form_Load()
If FullScreenSwitch = 1 Then WindowState = 2

'Param�tres
cadencecara = 200: cadenceligne = 260
typele�on = 1  'mais parfois timevalid <> 0 (le�ons 17C et 17D)
Set le�on_courante = le�on13
Module_routines.Colors le�on13  '12/2011
Module_routines.Dimension le�on13
'If repjawsnames = "" Then Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "."
'If repjawsnames <> "" Then Label2.Caption = msgUser & nom & "." & CRLF & msgLevel & nivo & "." & CRLF & msgSonori & repjawsnames & CRLF & msgSpeedExp & debexplilevel & msgSpeedGen & debgenlevel & "."
Module_routines.mshow le�on13 'avril 2008
label1.Caption = msgTapezTouche
Label3.Caption = bannerVersion & ", " & bannerCopyright
Label4.Caption = msgScore
Label5.Caption = msgF1F2F3
If echapbismax = 0 Then Quitter.Caption = msgQuitterVers + CRLF + bannerMenu + CRLF + msg�chap
If echapbismax = 1 Then Quitter.Caption = msgQuitterVers + CRLF + bannerMenu + CRLF + msg�chap2
If echapbismax > 1 Then Quitter.Caption = msgQuitterVers + CRLF + bannerMenu + CRLF + msg�chap3

' Les 3 touches
Module_routines.SetKeys "CAPSLOCK_OFF"
Module_routines.SetKeys "SCROLLLOCK_OFF"
If numpad = 0 Then Module_routines.SetKeys "NUMLOCK_OFF"
If numpad >= 1 Then ' Important pour les le�ons 13E et 17D sur le pav� num�rique
    keyinhibit = 1
    Module_routines.SetKeys "NUMLOCK_ON"
End If

' Premi�re currentline (apr�s le bip 2 tons, sinon bip inaudible en Win98)
'If Dir(vpath & "sonbip2.exe") <> "" Then Module_exec.ExecAndWait vpath & "sonbip2.exe" 'avril 2008
Module_routines.sonbip2tons 'avril 2008
Module_routines.OpenAndSuffix exo_courant, 1
On Error Resume Next
If EOF(1) Then
    Close #1
    Exit Sub
Else
    Line Input #1, currentline
    le�onfontsize = (24 * le�on13.Width) / 11746
    Module_routines.AdjustWidthAndSize le�on13, 1
    cara1 = "": cara2 = ""
    elapsed = 0: elapsedtot = 0: nbmots = 0
End If

'D�finir le temps si l'utilisateur d�marre trop vite
If timevalid = 0 Then
    Timer2.Enabled = False
Else
    starttop = Now
End If

End Sub

' **********************  LOAD  suite  ****************************************
Private Sub Timer1_Timer()
le�on13.text1.Text = currentline
nbcaras = 1: iwrong = 0: iwrongbis = 0
le�on13.text1.SelStart = 0
le�on13.text1.SelLength = Len(le�on13.text1.Text)

' Comparer
Module_routines.rac2ligne1 le�on13
keyinhibit = 0  ' Pour le�on 13E et 17D sur pav� num�rique
Timer1.Enabled = False
End Sub


' ***************************  TIMER2  *****************************************
Private Sub Timer2_Timer()
' Attendre la premi�re frappe au d�marrage
If nbmots = 0 Then
    starttop = Now

' Puis compter le temps elapsed pour la ligne
Else
    currentdate = Now
    elapsedtot = DateDiff("s", starttop, currentdate)
End If
If timevalid > 0 Then
    scorecourant = CInt(pctt) & " %.      " & nbmots & msgCommandes & elapsedtot & msgSecondes
    le�on13.text5.Text = scorecourant
End If
End Sub


' ******************  SEQUENCE de TOUCHES d�finie au bout du Timer9  ***************
Private Sub Timer9_Timer()
Timer9.Enabled = False

' VALEUR du KeyCode R�EL
' Control
If KeyFirst = 17 And KeySecond = 0 Then Keycode = 17 '(avec ou sans JAWS)

' Alt
If KeyFirst = 18 And KeySecond = 0 Then Keycode = 18 '(sans JAWS)
If KeyFirst = 18 And KeySecond = 18 And KeyThird = 0 Then Keycode = 18 '(avec JAWS)

' AltGr (avec ou sans JAWS)
If KeyFirst = 17 And KeySecond = 18 And KeyThird = 0 Then Keycode = 255 '(avec ou sans JAWS451 et plus)
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

' BONNE R�PONSE (code semblable � la subroutine text2_keyup ci-dessous) � la fin du temps timer9
If Keycode = KeyExpect Then
    
    ' Pour affichage imm�diat du score
    Module_routines.scoreaffich le�on_courante, 0
                    
    ' Reset
    le�on13.text4.Text = ""
    le�on13.text4.Visible = False
    f2link = 0
L10bis:
    ' Bonne r�ponse, traiter jusqu'� obtenir la ligne suivante
    follow
Else
    ' ECHAP r�it�r� (sauf si �chap est la bonne r�ponse)
    If Keycode = 27 Then
        If f2link = 1 Then 'septembre 2007, sortie de F2
            le�on13.text4.Text = ""
            le�on13.text4.Visible = False
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
    If Not Keycode = 20 Then 'D�verr Maj tol�r�
        
        ' Les demandes d'aide ne sont pas des fautes
        If Keycode = 27 Or Keycode = 112 Or Keycode = 113 Or Keycode = 114 Then Exit Sub
        
        ' Alt ou rel�chement de AltGr ne procure pas de faute ; indispensable pour le�on 13E
        If Keycode = 18 Then Exit Sub

        ' Vraies fautes
        If winstop = 0 Or UCase(le�on13.text1.Text) <> vvWindowsGauche Or UCase(le�on13.text1.Text) <> vvWindowsDroit Then
            Module_routines.bip le�on13
            iwrong = iwrong + 1: iwrongbis = iwrongbis + 1
            
            ' Pour affichage imm�diat du score
            Module_routines.scoreaffich le�on_courante, 0
                    
            ' Reset de iwrongbis si l'utilisateur progresse
            iiprec = iter
            If iiprec > iiante Then
                iiante = iiprec
                iwrongbis = 1
            End If
            
            ' Reset de text2
            le�on13.text1.SelLength = 0
            Call Sleep(150)
            le�on13.text1.SelLength = Len(le�on13.text1)
            le�on13.text2.Text = ""
            le�on13.text2.SelStart = 0
            le�on13.Cls
            
            ' On atteint le nb max de fautes sur le cara
            If iwrongbis >= iwrongbismax Then GoTo L10bis
        End If
    End If
End If

End Sub


' *************  TEXT1_KEYUP Events rares sur text1 ***************************
Private Sub text1_KeyUp(Keycode As Integer, shift As Integer)
If Keycode = 27 Then
    Quitter_Click
    Exit Sub
End If

'Curseur sur text1 interdit, passer � text2
On Error Resume Next
le�on13.text2.SetFocus
End Sub


' ***********************  TEXT2 KEY_DOWN  ***********************************
Private Sub Text2_KeyDown(Keycode As Integer, shift As Integer)
'Debug.Print "KeyDown=" & Keycode

' Win 91 et Win 92 (voir en plus Text2_KeyUp)
If Keycode = 91 Or Keycode = 92 Then
    Module_routines.cancelwin 0, le�on13, 0
    Exit Sub
End If

' Pour capturer la touche Tab qui n'est pas sensible � KeyUp
If Keycode = 9 Then Module_routines.pasdetab

'Reset picture � conserver!
Picture1.Visible = False

'RetArr inutile

'Touche Alt+F4 pour quitter
If Keycode = 115 And shift = 4 And UCase(Left(le�on13.text1.Text, 6)) <> "ALT+F4" Then
    altf4 = 1
    If quitactive = 0 Then Module_routines.QuitQuit
End If
End Sub


' ***************  TEXT2_KEYUP  REPONSES UTILISATEUR sur TEXT2  **************
Private Sub Text2_KeyUp(Keycode As Integer, shift As Integer)
'Debug.Print "1KeyUp=" & Keycode & "   KeyInhibit=" & keyinhibit & "   f2link=" & f2link

' TRAITEMENT des "S�QUENCES de TOUCHES" dues � JAWS pour Echap, Alt, Control
' Apr�s appel de F2 (f2link = 1), il faut passer par ce traitement, sinon pb sur AltGr par exemple (septembre 2007) ?
'If (Keycode = 17 Or Keycode = 18 Or Keycode = 27 Or (Keycode = 32 And UCase(le�on13.text1.Text) = "ESPACE�")) And (keyinhibit = 0 Or (keyinhibit = 1 And f2link = 1)) Then 'septembre 2007
If (Keycode = 17 Or Keycode = 18 Or Keycode = 27 Or (Keycode = 32 And UCase(le�on13.text1.Text) = vvEspace)) And (keyinhibit = 0 Or (keyinhibit = 1 And f2link = 1)) Then 'septembre 2007
    If KeyFirst <> 0 And KeySecond <> 0 And KeyThird = 0 Then KeyThird = Keycode
    If KeyFirst <> 0 And KeySecond = 0 Then KeySecond = Keycode
    If KeyFirst = 0 Then
        KeyFirst = Keycode
        ' Ne pas aller dans Timer9 quand le Control ou Alt fait partie d'une demande de combinaison de touches
        If ShiftExpect = 0 Then Timer9.Enabled = True
    End If
End If

' STOP ICI si on a une "S�QUENCE DE TOUCHE"
If Timer9.Enabled = True Then Exit Sub

' Attention, permet de basculer r�ellement la touche NumLock avec Win98 + Jaws !
If numpad = 0 Then Module_routines.SetKeys "NUMLOCK_OFF"
If numpad >= 1 Then Module_routines.SetKeys "NUMLOCK_ON"

' Touche Tab, suite � module_routines.pasdetab, interdite ou engendre �ventuel code 144
If Keycode = 144 And UCase(Left(le�on13.text1.Text, Len(vvVerrouillageNum�rique))) <> vvVerrouillageNum�rique Then
    If numpad <= 0 Then Exit Sub
    If numpad >= 1 And keyinhibit = 1 Then Exit Sub
End If
If keyforce = 9 Then
    le�on_courante.text2.Text = ""
    le�on_courante.text2.SelStart = 0
    keyforce = 0
    keyinhibit = 0
    If notab = 1 Then
        GoTo L13
    Else
        Keycode = 9
    End If
End If

'Touche F2, placer avant le traitement msgbox/keyinhibit=1
If Keycode = 113 Then
    keyinhibit = 2
    avecf2 = 1
    help_f2 le�on13
    avecf2 = 0
End If

' Les msgbox procurent des key-ups ind�sirables
If keyinhibit = 1 Then
    keyinhibit = 0
    If Keycode = 13 Or Keycode = 16 Then Exit Sub 'Entr�e, Majuscule
    'If Keycode = 78 Or Keycode = 79 Then Exit Sub 'N=Non, O=Oui
    le�on13.text4.Visible = False
    le�on13.text1.SelLength = 0
    Call Sleep(cadencecara)
    le�on13.text1.SelLength = Len(le�on13.text1.Text)
    If f2link = 1 And Keycode = 27 Then Exit Sub
    If f2link = 0 Then Exit Sub
End If

' Pour �chap de la touche F2
If keyinhibit = 2 Then keyinhibit = 1

' TOUCHES � PB, annule la commande r�alis�e simultan�ment par Windows
' Win 91 et Win 92 (voir en plus Text2_KeyDown)
If Keycode = 91 Or Keycode = 92 Then
    Module_routines.cancelwin 1, le�on13, 0
    Exit Sub
End If

' AltGauche 18 et Menu-Contextuel 93
If Keycode = 93 Then
    echapbis = -1
    keyinhibit = 1
    forcepause = 2 ' Pour ne pas quitter msgform quand un message d'explications suit, et pour ne pas �chapper du score dans msgform si la le�on se termine par Alt ou AltGr
    'SendKeys "{ESC}"
    'Sendkeys est remplac� par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
    echapbis = -1
    
    ' Menu-Contextuel non demand�
    If UCase(Left(le�on13.text1.Text, Len(vvMenuContextuel))) <> vvMenuContextuel And Keycode = 93 Then GoTo L13
End If

' Sendkeys termine par 145
If Keycode = 145 Then
    iwrong = iwrong - 1
    If iwrong < 0 Then iwrong = 0 'septembre 2007
    iwrongbis = iwrongbis - 1
    If iwrongbis < 0 Then iwrongbis = 0 'septembre 2007
    Exit Sub
End If

'Touche F1
If Keycode = 112 Then
    noechapF1 = 1
    Module_routines.help_f1 le�on13
End If

'Touche F3
If Keycode = 114 Then Module_routines.help_f3 le�on13

'Touche F10 m�ne � la barre menu
If Keycode = 121 Then
    echapbis = echapbis - 1
    'SendKeys "{ESC}", True
    'Sendkeys est remplac� par des actions keybd_event pour Windows Vista juin 2007
    keybd_event VK_ESCAPE, 0, 0, 0
    keybd_event VK_ESCAPE, 0, KEYEVENTF_KEYUP, 0
End If

' BONNE R�PONSE
If Keycode = KeyExpect And shift = ShiftExpect Then
L9:
    ' Pour affichage imm�diat du score
    Module_routines.scoreaffich le�on_courante, 0
    
    ' Reset
    le�on13.text3.Text = ""
    le�on13.text3.Visible = False
    le�on13.text4.Text = ""
    le�on13.text4.Visible = False
    f2link = 0
L10:
    echapbis = 0
    iter = iter + 1
    
    ' Si on va rel�cher Maj, Control, ou Alt
    If shift > 0 Then keyinhibit = 1
    
    ' Msg d'encouragement "Continuez�"
    If Not msgtext1(iter) = "" Then
        le�on13.label1.Visible = False
        le�on13.text1.Text = msgtext1(iter)
        currentline = le�on13.text1.Text
        Module_routines.AdjustWidthAndSize le�on_courante, 0
        le�on13.text1.SelStart = 0
        le�on13.text1.SelLength = Len(le�on13.text1)
        If Len(msgtext1(iter)) < 10 Then
            Call Sleep(800)
        Else
            Call Sleep(1200)
        End If
        le�on13.label1.Visible = True
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
            Module_routines.sonbip2tons 'avril 2008
        End If
    End If
            
    ' Bonne r�ponse, suite, va chercher la ligne suivante
    Module_routines.lignesuivante 0, timevalid, 0
    msgtext1(iter) = ""
    msgtext2(iter) = ""
    nbcaras = nbcaras + 1
    If derligne = 2 Then
        derligne = 0
        iter = 0
        Exit Sub
    End If
    
    ' Bonne r�ponse, suite, va chercher les keyexpect et shiftexpect suivant
    Module_routines.rac2ligne1 le�on13
    
    Else
    ' �chappement r�it�r� (sauf si �chap est la bonne r�ponse)
    If Keycode = 27 Then
        echapbis = echapbis + 1
        If echapbis > echapbismax Then
            echapbis = 0
            If keyinhibit = 0 Then Quitter_Click
            Exit Sub
        End If
    End If
    
    ' FAUTE de FRAPPE
    If Not Keycode = 20 Then 'D�verr Maj tol�r�
        
        ' Demande R�P�TER par ESPACE
        If Keycode = 32 Then
            le�on13.text2.Text = ""
            le�on13.text1.SelStart = 0
            le�on13.text1.SelLength = 0
            Call Sleep(cadencemot)
            le�on13.text1.SelStart = 0
            le�on13.text1.SelLength = Len(le�on13.text1.Text)
            Exit Sub
        End If
        
        ' Les demandes d'aide, Maj, �chap ne sont pas des fautes
        If Keycode = 16 Or Keycode = 27 Or Keycode = 112 Or Keycode = 113 Or Keycode = 114 Then Exit Sub
        ' Les codes renvoy�s par AltGr ne sont pas des fautes
        If Keycode = 17 Or Keycode = 18 Then Exit Sub
L13:
        ' Vraies fautes
        Module_routines.bip le�on13
        iwrong = iwrong + 1: iwrongbis = iwrongbis + 1
    
        ' Pour affichage imm�diat du score
        Module_routines.scoreaffich le�on_courante, 0
        
        ' Reset de iwrongbis si l'utilisateur progresse
        iiprec = iter
        If iiprec > iiante Then
            iiante = iiprec
            iwrongbis = 1
        End If
        
        ' Faute sur cara par erreur Majuscule/Minuscule
        If Keycode = KeyExpect And shift = 0 And ShiftExpect = 1 Then
            On Error Resume Next
            le�on13.text2.SetFocus
            le�on13.text3.Text = le�on13.text1.Text & vvMajuscule
            le�on13.text3.Width = 0.39 * le�on13.Width
            le�on13.text3.SelStart = 0
            le�on13.text3.SelLength = Len(le�on13.text3.Text)
            le�on13.text3.Visible = True
        End If
        If Keycode = KeyExpect And shift = 1 And ShiftExpect = 0 Then
            On Error Resume Next
            le�on13.text2.SetFocus
            le�on13.text3.Text = le�on13.text1.Text & vvMinuscule
            le�on13.text3.Width = 0.39 * le�on13.Width
            le�on13.text3.SelStart = 0
            le�on13.text3.SelLength = Len(le�on13.text3.Text)
            le�on13.text3.Visible = True
        End If
        
        ' Reset de text2
        le�on13.text1.SelLength = 0
        Call Sleep(10)
        le�on13.text1.SelLength = Len(le�on13.text1)
        le�on13.text2.Text = ""
        le�on13.text2.SelStart = 0
        le�on13.Cls
        
        ' On atteint le nombre max de fautes sur le cara, passer
        If iwrongbis >= iwrongbismax Then GoTo L10
    End If
End If
End Sub


' ******************************  follow  ***********************************************
Private Sub follow()
    If UCase(le�on13.text1.Text) = UCase(vvAlt) Or UCase(le�on13.text1.Text) = UCase(vvAlt) & "�" Or UCase(le�on13.text1.Text) = UCase(vvAltGr) Or UCase(le�on13.text1.Text) = UCase(vvAltGr) & "�" Or UCase(Left(le�on13.text1.Text, Len(vvWindowsGauche))) = vvWindowsGauche Or UCase(Left(le�on13.text1.Text, Len(vvWindowsDroit))) = vvWindowsDroit Then
    Else
        echapbis = 0
    End If
    iter = iter + 1
    
    ' Msg d'encouragement "Continuez�"
    If Not msgtext1(iter) = "" Then
        le�on13.label1.Visible = False
        le�on13.text1.Text = msgtext1(iter)
        currentline = le�on13.text1.Text
        Module_routines.AdjustWidthAndSize le�on_courante, 0
        le�on13.text1.SelStart = 0
        le�on13.text1.SelLength = Len(le�on13.text1)
        If Len(msgtext1(iter)) < 10 Then
            Call Sleep(800)
        Else
            Call Sleep(1200)
        End If
        le�on13.label1.Visible = True
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
            Module_routines.sonbip2tons 'avril 2008
        End If
    End If
        
    ' Bonne r�ponse, suite, va chercher la ligne suivante
    Module_routines.lignesuivante 0, 0, 0
    msgtext1(iter) = ""
    msgtext2(iter) = ""
    nbcaras = nbcaras + 1
    If derligne = 2 Then
        derligne = 0
        iter = 0
        Exit Sub
    End If
    
    ' Bonne r�ponse, suite, va chercher le keyexpect suivant
    Module_routines.rac2ligne1 le�on13
End Sub


' ********************  QUITTER  ***********************************************
Private Sub Quitter_Click()
Module_routines.quit_l
ShiftExpect = 0: keyinhibit = 0
End Sub


' *********** Quitter par le bouton ********************************************
Private Sub Quitter_KeyPress(KeyAscii As Integer)
If KeyAscii = 81 Or KeyAscii = 113 Then Quitter_Click
End Sub

