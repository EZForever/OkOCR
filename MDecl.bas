Attribute VB_Name = "MDecl"
Option Explicit
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public ScrWidth As Single, ScrHeight As Single
Public OrigWndProc As Long
Public OCRObj As Object
Public TmpFile As String

Public Function FTrayWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  If uMsg = 786 And wParam = 1 Then FTray.mnuStart_Click 'WM_HOTKEY
  If uMsg = 2333 And (lParam = 514 Or lParam = 517) Then FTray.PopupMenu FTray.mnuMain 'WM_LBUTTONUP, WM_RBUTTONUP
  FTrayWndProc = CallWindowProc(OrigWndProc, hWnd, uMsg, wParam, lParam)
End Function

