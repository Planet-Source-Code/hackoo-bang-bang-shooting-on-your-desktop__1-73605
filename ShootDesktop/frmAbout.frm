VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "BangBang"
   ClientHeight    =   4230
   ClientLeft      =   6165
   ClientTop       =   -3285
   ClientWidth     =   7230
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7230
   Begin VB.Timer Timer2 
      Left            =   6720
      Top             =   3720
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   6720
      Top             =   1920
   End
   Begin VB.Label lblProductName 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bang Bang"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   32.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   900
      Left            =   2040
      TabIndex        =   4
      Top             =   1200
      Width           =   3195
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Par : Hackoo Crackoo"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   2280
      TabIndex        =   3
      Top             =   2160
      Width           =   2520
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "© 2009"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   3240
      TabIndex        =   2
      Top             =   3840
      Width           =   600
   End
   Begin VB.Image Image1 
      Height          =   1155
      Left            =   0
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":C84A
      ToolTipText     =   "Si vous voulez m'écrire, cliquez ici!"
      Top             =   0
      Width           =   7245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      X1              =   1920
      X2              =   5160
      Y1              =   2640
      Y2              =   2640
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Appuyez sur une touche pour continuer"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1920
      TabIndex        =   1
      Top             =   2880
      Width           =   3360
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "*Conçu pour une résolution d'écran minimale de 1024 X 768"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   1200
      TabIndex        =   0
      Top             =   3360
      Width           =   4920
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Ferme la fenêtre en appuyant n'importe quelle touche et affiche le menu principal
Private Sub Form_KeyPress(KeyAscii As Integer)
    frmAbout.Hide
    Unload Me
    Form1.Show
End Sub

' Affiche le menu principal
Private Sub Form_Unload(Cancel As Integer)
    Form1.Show
End Sub

'Pour m'envoyer un e-mail
Private Sub Image1_Click()
    Call lblDisclaimer_Click
End Sub

'Pour m'envoyer un e-mail
Private Sub lblDisclaimer_Click()
ShellExecute Me.hWnd, vbNullString, "mailto:hackoofr@yahoo.fr", vbNullString, "", 1
End Sub

Private Sub Timer1_Timer()
    If lblProductName.ForeColor = vbWhite Then
        lblProductName.ForeColor = vbRed
    Else
        lblProductName.ForeColor = vbWhite
    End If
End Sub
Private Sub Form_Load()

'ver.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
'Me.Left = (Screen.Width - Me.Width) / 2

Timer2.Interval = 100
Timer2.Enabled = True

End Sub

Private Sub Timer2_Timer()
Me.Top = Me.Top + 400
If (Me.Top >= (Screen.Height / 2) - (Me.Height / 2)) Then
Timer2.Enabled = False
    
End If
End Sub
