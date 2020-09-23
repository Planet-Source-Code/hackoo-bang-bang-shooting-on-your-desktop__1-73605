VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   510
   ClientTop       =   525
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Form1.frx":C84A
   MousePointer    =   99  'Custom
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Image bang 
      Height          =   765
      Index           =   0
      Left            =   600
      Picture         =   "Form1.frx":C99C
      Top             =   600
      Visible         =   0   'False
      Width           =   810
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SetWindowPos Lib "user32" _
(ByVal handleW1 As Long, _
ByVal handleW1InsertWhere As Long, ByVal w As Long, _
ByVal x As Long, ByVal y As Long, ByVal z As Long, _
ByVal wFlags As Long) As Long

Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" _
                                 (ByVal hdc As Long, _
                                 ByVal x As Long, _
                                 ByVal y As Long, _
                                 ByVal nWidth As Long, _
                                 ByVal nHeight As Long, _
                                 ByVal hSrcDC As Long, _
                                 ByVal xSrc As Long, _
                                 ByVal ySrc As Long, _
                                 ByVal nSrcWidth As Long, _
                                 ByVal nSrcHeight As Long, _
                                 ByVal dwRop As Long) As Long
Dim i As Integer

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

i = i + 1
Load bang(i)
bang(i).Top = y - bang(i).Height / 2
bang(i).Left = x - bang(i).Width / 2
bang(i).Visible = True

JoueWav App.Path & "\" & "bang.wav", 1

End Sub

Private Sub Form_Load()
Dim a, b
MsgBox "Salut,ce petit programme est réalisé dans le but de vous faire débarrasser du Stress du Travail. Donc si vous voulez sortir de ce dernier tapez sur la touche ESC ""échap""", 64, "Copyright © Hackoo Crackoo"
Me.Height = Screen.Height
Me.Width = Screen.Width
    Me.AutoRedraw = True
    Me.ScaleMode = vbPixels
    a = GetDesktopWindow()
    b = GetDC(a)
    
    StretchBlt Me.hdc, 0, 0, Screen.Width, Screen.Height, b, 0, _
                 Screen.Height, Screen.Width, -Screen.Height, vbSrcCopy
    ReleaseDC b, hdc
    DeleteDC b
    HideTaskbar

End Sub
' Quitte avec la touche ESC
Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27 ' ESC
            ShowTaskbar
            Unload Me
    End Select
End Sub



