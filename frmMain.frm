VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MSN Buddy Logger v3.0 -- For MSN Messenger 5.x"
   ClientHeight    =   2970
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7875
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   0
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox txtLog 
      Height          =   2595
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   0
      Width           =   7875
   End
   Begin VB.Label lblTime 
      Height          =   195
      Left            =   2640
      TabIndex        =   0
      Top             =   2700
      Width           =   3855
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear Log"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Log"
      End
      Begin VB.Menu mnuTray 
         Caption         =   "Sys&Tray"
      End
      Begin VB.Menu mnuBreak1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
   Begin VB.Menu mnuPopupmenu 
      Caption         =   "Popupmenu"
      Visible         =   0   'False
      Begin VB.Menu mnuAbout2 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuShowHide 
         Caption         =   "&Hide"
      End
      Begin VB.Menu mnuBreak2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "E&xit"
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private nid As NOTIFYICONDATA

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_MOUSEMOVE = &H200
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up
Private Const WM_LBUTTONDBLCLK = &H203    'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click

Public boolAddGone As Boolean
Public boolQuit As Boolean
Public WithEvents msn As Messenger
Attribute msn.VB_VarHelpID = -1

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long


Private Function CheckSave()
  Dim intResult As Integer
  
  intResult = MsgBox("Do you want to save the log?", vbYesNoCancel, "Save log?", 0, 0)
    
  Select Case intResult
    Case vbYes
      CheckSave = "yes"
    Case vbNo
      CheckSave = "no"
    Case vbCancel
      CheckSave = "cancel"
  End Select
End Function

Private Sub Form_Load()
  Me.Show
  Me.Refresh
  With nid
    .cbSize = Len(nid)
    .hWnd = Me.hWnd
    .uId = vbNull
    .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    .uCallBackMessage = WM_MOUSEMOVE
    .hIcon = Me.Icon
    .szTip = "MSN Buddy Logger" & vbNullChar
  End With
  Shell_NotifyIcon NIM_ADD, nid
 
  Set msn = New Messenger
  
  txtLog.Text = Now & " log started."
  frmMain.SetFocus
  'frmAddFix.Show
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim Result As Long
  Dim msg As Long

  If Me.ScaleMode = vbPixels Then
    msg = X
  Else
    msg = X / Screen.TwipsPerPixelX
  End If

  Select Case msg
  Case WM_RBUTTONUP
    Result = SetForegroundWindow(Me.hWnd)
    Me.PopupMenu mnuPopupmenu
  Case WM_LBUTTONDBLCLK
    Me.Show
    Result = SetForegroundWindow(Me.hWnd)
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  On Error Resume Next
  If boolQuit = True Then
    Shell_NotifyIcon NIM_DELETE, nid
    Cancel = False
    'If frmAbout.Visible = True Then
      Unload frmAbout
    'End If
    'If frmAddFix.Visible = True Then
      Unload frmAddFix
    'End If
    End
  Else
    mnuShowHide.Caption = "&Show"
    Me.Hide
    Cancel = True
  End If

  'Dim strResult As String
  'strResult = CheckSave
  
  'If strResult = "yes" Then
  '  Call SaveLog
  'ElseIf strResult = "cancel" Then
  '  Cancel = vbYes
  'End If
  
  'If strResult <> "cancel" Then
  '  If frmAbout.Visible = True Then
  '    Unload frmAbout
  '  End If
  '  If frmAddFix.Visible = True Then
  '    Unload frmAddFix
  '  End If
  '  End
  'End If
End Sub

Private Sub mnuAbout_Click()
  frmAbout.Show 1, Me
End Sub

Private Sub mnuAbout2_Click()
  Call mnuAbout_Click
End Sub

Private Sub mnuClear_Click()
  Dim strResult As String
  strResult = CheckSave
  If strResult = "yes" Then
    Call SaveLog
  ElseIf strResult = "cancel" Then
    Exit Sub
  End If
  
  txtLog.Text = ""
  txtLog.Text = Now & " log started."
End Sub

Private Sub mnuExit_Click()
  boolQuit = True
  Unload Me
End Sub

Private Sub mnuQuit_Click(Index As Integer)
  boolQuit = True
  Unload Me
End Sub

Private Sub mnuSave_Click()
  Call SaveLog
End Sub

Private Sub mnuShowHide_Click()
  If mnuShowHide.Caption = "&Hide" Then
    mnuShowHide.Caption = "&Show"
    Me.Hide
  ElseIf mnuShowHide.Caption = "&Show" Then
    mnuShowHide.Caption = "&Hide"
    Me.Show
  End If
End Sub

Private Sub mnuTray_Click()
  mnuShowHide.Caption = "&Show"
  Me.Hide
End Sub

Private Sub SaveLog()
  Dim strSaveName As String
  
  dlgSave.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*|"
  dlgSave.FileName = ""
  dlgSave.ShowSave
    
  strSaveName = dlgSave.FileName
  
  If strSaveName <> "" Then
    Open strSaveName For Output As #1
    Print #1, txtLog.Text
    Close #1
  
    MsgBox "Saved as: " & dlgSave.FileName
  End If
End Sub

Private Sub msn_OnContactFriendlyNameChange(ByVal hr As Long, ByVal pMContact As Object, ByVal bstrPrevFriendlyName As String)
  txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.SigninName & " changed their name to " & pMContact.FriendlyName
End Sub

Private Sub msn_OnContactPhoneChange(ByVal hr As Long, ByVal pContact As Object, ByVal PhoneType As MessengerAPI.MPHONE_TYPE, ByVal bstrNumber As String)
  If PhoneType = MPHONE_TYPE_HOME Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pContact.FriendlyName & "(" & pContact.SigninName & ") changed their home phone # to " & bstrNumber
  End If
  
  If PhoneType = MPHONE_TYPE_MOBILE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pContact.FriendlyName & "(" & pContact.SigninName & ") changed their mobile phone # to " & bstrNumber
  End If
  
  If PhoneType = MPHONE_TYPE_WORK Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pContact.FriendlyName & "(" & pContact.SigninName & ") changed their work phone # to " & bstrNumber
  End If
End Sub

Private Sub msn_OnContactStatusChange(ByVal pMContact As Object, ByVal mStatus As MessengerAPI.MISTATUS)
  If mStatus = MISTATUS_OFFLINE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.FriendlyName & " (" & pMContact.SigninName & ") signed in."
  End If

  If pMContact.Status = MISTATUS_OFFLINE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.FriendlyName & " (" & pMContact.SigninName & ") signed out."
  End If

  If pMContact.Status = MISTATUS_ONLINE And mStatus <> MISTATUS_OFFLINE And mStatus <> MISTATUS_IDLE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.FriendlyName & " (" & pMContact.SigninName & ") changed status to Online."
  End If

  If pMContact.Status = MISTATUS_ONLINE And mStatus <> MISTATUS_OFFLINE And mStatus <> MISTATUS_IDLE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.FriendlyName & " (" & pMContact.SigninName & ") changed status to Online."
  End If
  
  If pMContact.Status = MISTATUS_BUSY And mStatus <> MISTATUS_OFFLINE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.FriendlyName & " (" & pMContact.SigninName & ") changed status to Busy."
  End If
  
  If pMContact.Status = MISTATUS_BE_RIGHT_BACK And mStatus <> MISTATUS_OFFLINE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.FriendlyName & " (" & pMContact.SigninName & ") changed status to Be Right Back."
  End If
  
  If pMContact.Status = MISTATUS_AWAY And mStatus <> MISTATUS_OFFLINE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.FriendlyName & " (" & pMContact.SigninName & ") changed status to Away."
  End If
  
  If pMContact.Status = MISTATUS_ON_THE_PHONE And mStatus <> MISTATUS_OFFLINE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.FriendlyName & " (" & pMContact.SigninName & ") changed status to On The Phone."
  End If
  
  If pMContact.Status = MISTATUS_OUT_TO_LUNCH And mStatus <> MISTATUS_OFFLINE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.FriendlyName & " (" & pMContact.SigninName & ") changed status to Out To Lunch."
  End If
  
  If pMContact.Status = MISTATUS_INVISIBLE And mStatus <> MISTATUS_OFFLINE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.FriendlyName & " (" & pMContact.SigninName & ") changed status to Invisible."
  End If
  
  If pMContact.Status = MISTATUS_IDLE And mStatus <> MISTATUS_OFFLINE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.FriendlyName & " (" & pMContact.SigninName & ") is idle."
  End If
  
  If pMContact.Status <> MISTATUS_OFFLINE And mStatus = MISTATUS_IDLE Then
    txtLog.Text = txtLog.Text & vbCrLf & Now & " " & pMContact.FriendlyName & " (" & pMContact.SigninName & ") has returned to their computer."
  End If
End Sub

Private Sub Timer1_Timer()
  Dim hWnd As Long
  lblTime.Caption = "Current time: " & Now

  If boolAddGone = False Then
    hWnd = FindWindow(vbNullString, "MSN Messenger")
    hWnd = FindWindowEx(hWnd, 0, vbNullString, "msmsgs banner")
    If hWnd <> 0 Then
      Load frmAddFix
      frmAddFix.Show
    End If
  ElseIf boolAddGone = True Then
    hWnd = FindWindow(vbNullString, "MSN Messenger")
    hWnd = FindWindowEx(hWnd, 0, vbNullString, "msmsgs banner")
    If hWnd = 0 Then
      Unload frmAddFix
      boolAddGone = False
    End If
  End If
End Sub
