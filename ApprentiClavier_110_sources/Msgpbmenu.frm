VERSION 4.00
Begin VB.Form msgpbmenu 
   Caption         =   "Ouf"
   ClientHeight    =   5880
   ClientLeft      =   600
   ClientTop       =   1935
   ClientWidth     =   9840
   ControlBox      =   0   'False
   Height          =   6690
   KeyPreview      =   -1  'True
   Left            =   540
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9840
   Top             =   1185
   Width           =   9960
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   6000
      Top             =   0
   End
   Begin VB.TextBox Text0 
      BackColor       =   &H00C0C000&
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4965
      HideSelection   =   0   'False
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   450
      Width           =   9465
   End
   Begin VB.Label Label0 
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
      Top             =   150
      Width           =   1815
   End
   Begin VB.Menu Fichier 
      Caption         =   "&Fichier"
      Begin VB.Menu Quitter 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu Aide 
      Caption         =   "&Aide"
      Begin VB.Menu AideGénérale 
         Caption         =   "&Aide générale"
         Shortcut        =   {F1}
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu Aproposde 
         Caption         =   "A &Propos de"
      End
   End
End
Attribute VB_Name = "msgpbmenu"
Attribute VB_Creatable = False
Attribute VB_Exposed = False
' ********************  LOAD  *******************************
Private Sub Form_Load()
If FullScreenSwitch = 1 Then WindowState = 2
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
Text0.SetFocus
Timer1.Enabled = False
Unload msgpbmenu
End Sub

