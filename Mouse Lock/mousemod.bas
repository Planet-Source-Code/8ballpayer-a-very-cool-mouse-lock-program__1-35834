Attribute VB_Name = "Module1"
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, _
ByVal Y As Long, ByVal CX As Long, ByVal CY As Long, _
ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = 2
Public Const SWP_NOSIZE = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function RegisterServiceProcess Lib "kernel32" (ByVal dwProcessID As Long, ByVal dwType As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Const VK_F9 = &H3
Public Const Chr1 = 13
Public Const Chr2 = 10
Public Const Wrap = Chr1 + Chr2

Option Explicit
Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Declare Function ClipCursor Lib "user32" _
(lpRect As Any) As Long

Public Sub DisableTrap(CurForm As Form)
Dim erg As Long
Dim NewRect As RECT
CurForm.Caption = "Mouse Released"
With NewRect
  .Left = 0&
  .Top = 0&
  .Right = Screen.Width / Screen.TwipsPerPixelX
  .Bottom = Screen.Height / Screen.TwipsPerPixelY
End With

erg& = ClipCursor(NewRect)
End Sub

Public Sub EnableTrap(CurForm As Form)
Dim X As Long, Y As Long, erg As Long
Dim NewRect As RECT
X& = Screen.TwipsPerPixelX
Y& = Screen.TwipsPerPixelY
CurForm.Caption = "Mouse Trapped"
With NewRect
  .Left = CurForm.Left / X&
  .Top = CurForm.Top / Y&
  .Right = .Left + CurForm.Width / X&
  .Bottom = .Top + CurForm.Height / Y&
End With
erg& = ClipCursor(NewRect)
End Sub
Public Sub HideTask(Hide As Boolean)
Dim lHandle As Long
Dim lService As Long
lHandle = GetCurrentProcessId()
lService = RegisterServiceProcess(lHandle, Abs(Hide))
End Sub
Public Sub SetOnTop(ByVal hwnd As Long, ByVal bSetOnTop As Boolean)
    Dim lR As Long
    If bSetOnTop Then
        lR = SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    Else
        lR = SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    End If
End Sub




