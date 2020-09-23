Attribute VB_Name = "Module1"
Declare Function JoueWav Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Dim handleW1 As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function FindWindowA Lib "user32" _
(ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Private Declare Function SetWindowPos Lib "user32" _
(ByVal handleW1 As Long, _
ByVal handleW1InsertWhere As Long, ByVal w As Long, _
ByVal x As Long, ByVal y As Long, ByVal z As Long, _
ByVal wFlags As Long) As Long

Const TOGGLE_HIDEWINDOW = &H80
Const TOGGLE_UNHIDEWINDOW = &H40

Function HideTaskbar()
    handleW1 = FindWindowA("Shell_traywnd", "")
    Call SetWindowPos(handleW1, 0, 0, 0, _
         0, 0, TOGGLE_HIDEWINDOW)
End Function

Function ShowTaskbar()
    Call SetWindowPos(handleW1, 0, 0, 0, _
         0, 0, TOGGLE_UNHIDEWINDOW)
End Function


