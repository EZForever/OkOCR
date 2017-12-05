VERSION 5.00
Begin VB.Form FTray 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "OkOCR Tray Window"
   ClientHeight    =   600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2535
   ControlBox      =   0   'False
   Icon            =   "FTray.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   40
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   169
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Menu mnuMain 
      Caption         =   "OkOCR"
      Begin VB.Menu mnuStart 
         Caption         =   "&S 开始识别"
      End
      Begin VB.Menu mnuSpr 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&A 关于..."
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&X 退出"
      End
   End
End
Attribute VB_Name = "FTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Type NOTIFYICONDATA
  cbSize As Long
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 64 'MAX_TOOLTIP
End Type

Private Desp As String
Private IconInfo As NOTIFYICONDATA

Private Sub Form_Load()
  Desp = App.Title + " v" & App.Major & "." & App.Minor & "." & App.Revision
  TmpFile = Environ("TMP") + "\OkOCR.bmp"
  On Error GoTo Oops
  Set OCRObj = CreateObject("MODI.Document")
  On Error GoTo 0

  With Screen
    ScrWidth = .Width / .TwipsPerPixelX
    ScrHeight = .Height / .TwipsPerPixelY
  End With

  With IconInfo
    .hWnd = Me.hWnd
    .uID = Me.Icon
    .uFlags = 7 'NIF_ICON Or NIF_MESSAGE Or NIF_TIP
    .uCallbackMessage = 2333 'Self-defined
    .hIcon = Me.Icon.Handle
    .szTip = Desp + vbNullChar
    .cbSize = Len(IconInfo)
  End With
  Shell_NotifyIcon 0, IconInfo 'NIM_ADD

  RegisterHotKey Me.hWnd, 1, 1, Asc("V") 'MOD_ALT
  OrigWndProc = SetWindowLong(Me.hWnd, -4&, AddressOf FTrayWndProc) 'GWL_WNDPROC
  MsgBox "OkOCR 已经启动并运行在托盘区。" + vbCrLf + _
         "按 Alt + V 开始识别，右键菜单中有更多选项。", vbInformation
  Exit Sub
Oops:
  MsgBox "无法加载 OCR 核心组件。请检查 MODI 是否已正确安装。" + vbCrLf + _
         "若要全新安装 MODI，请参阅 http://support.microsoft.com/kb/982760/", vbExclamation
  Unload Me
End Sub

Private Sub mnuExit_Click()
  If Not MsgBox("确认要退出吗？", vbQuestion + vbYesNo) = vbYes Then Exit Sub
  Unload FResult
  'Unload FCapture 'If user can click on this menu, FCapture won't be there
  Shell_NotifyIcon 2, IconInfo 'NIM_DELETE
  SetWindowLong Me.hWnd, -4&, OrigWndProc 'GWL_WNDPROC
  Unload Me
End Sub

Private Sub mnuAbout_Click()
  MsgBox Desp + " (2017-08-30)" + vbCrLf + "版权所有 (C) 2017 EZ Studio", vbInformation
End Sub

Public Sub mnuStart_Click()
  Unload FResult
  FCapture.Show
End Sub
