VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "-=   Fun with System Clock =-"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3720
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   3720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3360
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   3495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3495
      Begin VB.Timer Timer2 
         Interval        =   10
         Left            =   3240
         Top             =   360
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Hide Program"
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Random "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1575
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Rewind"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Fast Forward"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MOD_ALT = &H1
Private Const MOD_CONTROL = &H2
Private Const MOD_SHIFT = &H4
Private Const PM_REMOVE = &H1
Private Const WM_HOTKEY = &H312

Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type Msg
    hWnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Declare Function SetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME) As Long

Dim min As Integer
Dim hour As Integer
Dim day As Integer
Dim month As Integer
Dim dayofweek As Integer
Dim year As Integer
Dim LASTVAL As Integer

Private Declare Function RegisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Private Declare Function UnregisterHotKey Lib "user32" (ByVal hWnd As Long, ByVal id As Long) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As Msg, ByVal hWnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private bCancel As Boolean

Private Sub ProcessMessages()
    Dim Message As Msg
    Do While Not bCancel
        WaitMessage
        If PeekMessage(Message, Me.hWnd, WM_HOTKEY, WM_HOTKEY, PM_REMOVE) Then
            Form1.Show
            Call UnregisterHotKey(Me.hWnd, &HBFFF&)
            Check1.Enabled = True
            Check2.Enabled = True
            Check3.Enabled = True
            Check4.Enabled = True
            Timer1.Enabled = False
            Command1.Caption = "Start"
            Timer2.Enabled = True
        End If
        DoEvents
    Loop
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Check2.Value = 0
   Check3.Value = 0
   LASTVAL = 1
End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   Check3.Value = 0
   Check1.Value = 0
   LASTVAL = 2
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
   Check2.Value = 0
   Check1.Value = 0
   LASTVAL = 3
End If
End Sub

Private Sub Form_Load()
Check1.Value = 1
MsgBox "Many portions of the operating system and 3rd party" _
     & vbCrLf _
     & "programs rely on system time to function correctly." _
     & vbCrLf _
     & "The author is not responsible for any damages that" _
     & vbCrLf _
     & "this program may cause. Tested only on WinXP, may" _
     & vbCrLf _
     & "require administator rights. Restoring the correct" _
     & vbCrLf _
     & "time/date is your job.      YOU HAVE BEEN WARNED." _
     & vbCrLf _
     & "" _
     & vbCrLf _
     & "", , "Disclaimer"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Check4.Value = 1 Then
   If Command1.Caption = "Stop" Then
      Call UnregisterHotKey(Me.hWnd, &HBFFF&)
   End If
End If
End
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Start" Then
Check1.Enabled = False
Check2.Enabled = False
Check3.Enabled = False
Check4.Enabled = False
year = 2004
day = 1
month = 1
Timer1.Enabled = True
Randomize
   If Check4.Value = 1 Then
     Form1.Hide
     Dim ret As Long
     bCancel = False
     MsgBox "The hotkey CONTROL-F will restore hidden window." _
     & vbCrLf _
     & "", , "IMPORTANT INFO"
     ret = RegisterHotKey(Me.hWnd, &HBFFF&, MOD_CONTROL, vbKeyF)
     ProcessMessages
   End If
Timer2.Enabled = False
Command1.Caption = "Stop"
Else
Check1.Enabled = True
Check2.Enabled = True
Check3.Enabled = True
Check4.Enabled = True
Timer1.Enabled = False
Timer2.Enabled = True
Command1.Caption = "Start"
End If

End Sub

Private Sub Timer1_Timer()
    Dim lpSystemTime As SYSTEMTIME
    
    If Check1.Value = 1 Then
        Timer1.Interval = 50
        min = min + 1
        If min > 59 Then
        hour = hour + 1
        min = 0
        End If
        If hour > 23 Then
        day = day + 1
        hour = 0
        End If
        If day > 28 Then
        month = month + 1
        day = 1
        End If
        If month > 12 Then
        month = 1
        year = year + 1
        End If
        If year > 2015 Then year = 2004
    End If
    
    If Check2.Value = 1 Then
        Timer1.Interval = 50
        min = min - 1
        If min < 0 Then
        hour = hour - 1
        min = 59
        End If
        If hour < 0 Then
        day = day - 1
        hour = 23
        End If
        If day < 1 Then
        month = month - 1
        day = 28
        End If
        If month < 1 Then
        month = 12
        year = year - 1
        End If
        If year < 2004 Then year = 2015
    End If
    
    If Check3.Value = 1 Then
        Timer1.Interval = 250
        min = Int(Rnd * 59)
        hour = Int(Rnd * 23)
        day = Int(Rnd * 28)
        month = Int(Rnd * 12)
        year = Int(Rnd * 11) + 2004
    End If
    
    lpSystemTime.wYear = year
    lpSystemTime.wMonth = month
    lpSystemTime.wDayOfWeek = -1
    lpSystemTime.wDay = day
    lpSystemTime.wHour = hour
    lpSystemTime.wMinute = min
    lpSystemTime.wSecond = 0
    lpSystemTime.wMilliseconds = 0
    SetSystemTime lpSystemTime
End Sub

Private Sub Timer2_Timer()

If Check1.Value = 0 Then
   If Check2.Value = 0 Then
      If Check3.Value = 0 Then
         If LASTVAL = 1 Then
            Check1.Value = 1
         End If
         If LASTVAL = 2 Then
            Check2.Value = 1
         End If
         If LASTVAL = 3 Then
            Check3.Value = 1
         End If
      End If
   End If
End If

End Sub
