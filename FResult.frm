VERSION 5.00
Begin VB.Form FResult 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OkOCR"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6990
   Icon            =   "FResult.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   345
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   466
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picScreen 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   2760
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox chkEnglish 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "英文模式"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdDiscard 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   375
      Left            =   4320
      TabIndex        =   3
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "复制并关闭"
      Default         =   -1  'True
      Height          =   375
      Left            =   5640
      TabIndex        =   2
      Top             =   4680
      Width           =   1215
   End
   Begin VB.Frame frmResult 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "识别结果"
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   2400
      Width           =   6735
      Begin VB.TextBox txtResult 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "新宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Frame frmPreview 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "截获图像（部分）"
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.PictureBox picPreview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1815
         Left            =   120
         ScaleHeight     =   1785
         ScaleWidth      =   6465
         TabIndex        =   5
         Top             =   240
         Width           =   6495
      End
   End
   Begin VB.Label lblOCRing 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "正在识别..."
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   3240
      TabIndex        =   4
      Top             =   4800
      Width           =   990
   End
End
Attribute VB_Name = "FResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long

Private Sub chkEnglish_Click()
  lblOCRing.Visible = True
  On Error GoTo Oops
  OCRObj.Images(0).OCR IIf(chkEnglish.Value, 9, 2052), True, True
  txtResult.Text = OCRObj.Images(0).Layout.Text
Oops:
  lblOCRing.Visible = False
  If Err.Number Then MsgBox "识别图像时发生错误：" + vbCrLf & Err.Number & ": " + Err.Description, vbExclamation
End Sub

Private Sub cmdDiscard_Click()
  Unload Me
End Sub

Private Sub cmdDone_Click()
  Clipboard.SetText txtResult.Text
  Unload Me
End Sub

Private Sub Form_Load()
  Dim hdc As Long: hdc = GetWindowDC(GetDesktopWindow)
  With FCapture.shpArea
    picScreen.Width = .Width
    picScreen.Height = .Height
    BitBlt picScreen.hdc, 0, 0, .Width, .Height, hdc, .Left, .Top, &HCC0020 'SRCCOPY
  End With
  ReleaseDC GetDesktopWindow, hdc
  Unload FCapture
  Set picPreview.Picture = picScreen.Image
  If Dir(TmpFile) <> "" Then Kill TmpFile
  SavePicture picScreen.Image, TmpFile
  OCRObj.Create TmpFile
  chkEnglish_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  OCRObj.Close False
End Sub
