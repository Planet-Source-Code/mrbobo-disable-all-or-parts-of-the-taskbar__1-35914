VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Enabler"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   3045
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHide 
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton cmdEnable 
      Height          =   495
      Index           =   4
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton cmdEnable 
      Height          =   495
      Index           =   3
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   2655
   End
   Begin VB.CommandButton cmdEnable 
      Height          =   495
      Index           =   2
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdEnable 
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'API to locate the taksbar and it's children
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
'API to disable/enable windows
Private Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
'API to hide/show taskbar
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
'constants used to get window handles
Private Const SWP_SHOWWINDOW = &H40
Private Const SWP_HIDEWINDOW = &H80
Private Const GW_NEXT = 2
Private Const GW_CHILD = 5
Private Function GetTrayHandle(mType As Integer) As Long
    'Get the window handle
    Dim Desktop As Long, mhandle As Long, temp As String * 16, SrchString As String
    Select Case mType
        Case 1 'start button
            SrchString = "Button"
        Case 2 'system tray
            SrchString = "TrayNotifyWnd"
        Case 3 'appbar(where an application minimises to)
            SrchString = "ReBarWindow32"
    End Select
    Desktop = GetDesktopWindow()
    mhandle = GetWindow(Desktop, GW_CHILD)
    Do While mhandle <> 0
        GetClassName mhandle, temp, 14
        If Left(temp, 13) = "Shell_TrayWnd" Then
            If mType = 4 Then 'entire taskbar
                GetTrayHandle = mhandle
                Exit Do
            End If
            mhandle = GetWindow(mhandle, GW_CHILD)
            Do While mhandle <> 0
                GetClassName mhandle, temp, Len(SrchString) + 1
                If Left(temp, Len(SrchString)) = SrchString Then
                    GetTrayHandle = mhandle
                    Exit Function
                End If
                mhandle = GetWindow(mhandle, GW_NEXT)
            Loop
        End If
        mhandle = GetWindow(mhandle, GW_NEXT)
    Loop
End Function
Public Sub ToggleEnabled(mWindow As Long)
    'enable/disable a window given it's handle
    Dim Enabled As Long
    Dim retval As Long
    Enabled = IsWindowEnabled(mWindow)
    If Enabled = 0 Then
        retval = EnableWindow(mWindow, 1)
    Else
        retval = EnableWindow(mWindow, 0)
    End If
    UpdateButtonCaptions
End Sub
Private Sub cmdEnable_Click(Index As Integer)
    'enable/disable a window dependant on which button(index) is pressed
    ToggleEnabled GetTrayHandle(Index)
    UpdateButtonCaptions 'update button captions
End Sub
Private Sub cmdHide_Click()
    ' hide/show taskbar
    Dim Visible As Boolean
    Visible = IsWindowVisible(GetTrayHandle(4))
    If Visible Then
        SetWindowPos GetTrayHandle(4), 0, 0, 0, 0, 0, SWP_HIDEWINDOW
    Else
        SetWindowPos GetTrayHandle(4), 0, 0, 0, 0, 0, SWP_SHOWWINDOW
    End If
    UpdateButtonCaptions

End Sub

Private Sub Form_Load()
    UpdateButtonCaptions 'update button captions
End Sub

Public Sub UpdateButtonCaptions()
    'use API to determine wether a window is enabled and change
    'caption on the appropriate button accordingly
    cmdEnable(1).Caption = IIf(IsWindowEnabled(GetTrayHandle(1)), "Start Buttton Enabled", "Start Buttton Disabled")
    cmdEnable(2).Caption = IIf(IsWindowEnabled(GetTrayHandle(2)), "System Tray Enabled", "System Tray Disabled")
    cmdEnable(3).Caption = IIf(IsWindowEnabled(GetTrayHandle(3)), "AppBar Enabled", "AppBar Disabled")
    cmdEnable(4).Caption = IIf(IsWindowEnabled(GetTrayHandle(4)), "Entire Taskbar Enabled", "Entire Taskbar Disabled")
    cmdHide.Caption = IIf(IsWindowVisible(GetTrayHandle(4)), "Entire Taskbar Visible", "Entire Taskbar Hidden")
End Sub
