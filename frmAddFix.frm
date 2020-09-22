VERSION 5.00
Begin VB.Form frmAddFix 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "MSN Buddy Logger"
   ClientHeight    =   735
   ClientLeft      =   -90
   ClientTop       =   -660
   ClientWidth     =   1845
   ControlBox      =   0   'False
   FillColor       =   &H8000000F&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   49
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   123
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox pb 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   0
      ScaleHeight     =   255
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   240
      Width           =   15
   End
   Begin VB.Timer Timer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1200
      Top             =   120
   End
End
Attribute VB_Name = "frmAddFix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

Public Sub Form_Load()
  Dim hWnd As Long, rctemp As RECT
  hWnd = FindWindow(vbNullString, "MSN Messenger")
  hWnd = FindWindowEx(hWnd, 0, vbNullString, "msmsgs banner")
  If hWnd = 0 Then
    frmMain.boolAddGone = False
    Unload Me
    Exit Sub
  Else
    'Me.Show
    frmMain.boolAddGone = True
    GetWindowRect hWnd, rctemp
    With Me
      .Top = 0
      .Left = 0
      .Height = Me.Height * (rctemp.Bottom - rctemp.Top) / Me.ScaleHeight
      .Width = Me.Width * (rctemp.Right - rctemp.Left) / Me.ScaleWidth
    End With
    Timer.Enabled = True
    SetParent Me.hWnd, hWnd
    pb.Height = Me.ScaleHeight
    pb.Width = Me.ScaleWidth
    pb.ScaleWidth = 100
    pb.DrawMode = 10
  End If
  frmMain.SetFocus
End Sub

Private Sub Form_DblClick()
  'Unload Me
End Sub

Private Sub pb_Click()
  'Unload Me
End Sub

Private Sub Form_Resize()
  pb.Top = 0 '(Me.ScaleHeight - pb.Height) / 2
End Sub

Private Sub Timer_Timer()
  'Dim hWnd As Long
  pb.Cls
  pb.CurrentX = 50 - pb.TextWidth("MSN Buddy Logger" & vbCrLf & "       by R" & Chr$(179) & "Software") / 2
  pb.CurrentY = (pb.ScaleHeight - pb.TextHeight("MSN Buddy Logger" & vbCrLf & "      by R" & Chr$(179) & "Software")) / 2
  pb.Print "MSN Buddy Logger" & vbCrLf & "      by R" & Chr$(179) & "Software"
  pb.Line (0, 0)-(pb.ScaleWidth, pb.ScaleHeight), , BF
  'If hWnd = 0 Then
  '  frmMain.boolAddGone = False
  '  Unload Me
  'End If
End Sub
