VERSION 5.00
Begin VB.Form FCapture 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "OkOCR Capture Window"
   ClientHeight    =   3135
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ScaleHeight     =   209
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame frmPrompt 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   735
      Begin VB.CommandButton cmdYes 
         Appearance      =   0  'Flat
         Caption         =   "¡Ì"
         Default         =   -1  'True
         Height          =   375
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdNo 
         Appearance      =   0  'Flat
         Cancel          =   -1  'True
         Caption         =   "¡Á"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Shape shpArea 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderStyle     =   2  'Dash
      Height          =   855
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "FCapture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Private Dragging As Boolean

Private Sub cmdYes_Click()
  Me.Hide
  FResult.Show
End Sub

Private Sub cmdNo_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  SetWindowPos hWnd, -1&, 0, 0, 0, 0, 3& 'HWND_TOPMOST, SWP_NOMOVE Or SWP_NOSIZE
  SetWindowLong Me.hWnd, -20&, &H80000 'GWL_EXSTYLE, WS_EX_LAYERED
  SetLayeredWindowAttributes Me.hWnd, Me.BackColor, 100, 2 'LWA_ALPHA
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dragging = True
  shpArea.Move X, Y, 1, 1
  shpArea.Visible = True
  frmPrompt.Visible = False
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Not Dragging Then Exit Sub
  Dim Tmp As Single
  With shpArea
    If X < .Left Then
      Tmp = .Left + .Width
      .Left = X
      .Width = Tmp - X
    Else
      .Width = X - .Left
    End If
    If Y < .Top Then
      Tmp = .Top + .Height
      .Top = Y
      .Height = Tmp - Y
    Else
      .Height = Y - .Top
    End If
  End With
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dragging = False
  With frmPrompt
    .Left = IIf(X + .Width > ScrWidth, ScrWidth - .Width, X)
    .Top = IIf(Y + .Height > ScrHeight, ScrHeight - .Height, Y)
    .Visible = True
  End With
End Sub
