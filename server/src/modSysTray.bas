Attribute VB_Name = "modSysTray"
Option Explicit

' Declare a user-defined variable to pass to the Shell_NotifyIcon
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

' Function to add, modify, or delete an icon from the System Tray
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2

' The following constant is the message sent when a mouse event occurs
Public Const WM_MOUSEMOVE = &H200

' The following constants are the flags that indicate the valid
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

' Left-click constants.
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up

' Right-click constants.
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up

' Declare the API function call.
Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pNotifyIcon As NOTIFYICONDATA) As Boolean

' Dimension a variable as the user-defined data type.
Public NotifyIcon As NOTIFYICONDATA

Public Sub DestroySystemTray()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call Shell_NotifyIcon(NIM_DELETE, NotifyIcon) ' Add to the sys tray
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DestroySystemTray", "modSysTray", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub LoadSystemTray()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Create system tray
    NotifyIcon.cbSize = Len(NotifyIcon)
    NotifyIcon.hWnd = frmServer.hWnd
    NotifyIcon.uId = vbNull
    NotifyIcon.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    NotifyIcon.uCallBackMessage = WM_MOUSEMOVE
    NotifyIcon.hIcon = frmServer.Icon
    NotifyIcon.szTip = "Server" & vbNullChar   'You can add your game name or something.
    Call Shell_NotifyIcon(NIM_ADD, NotifyIcon) 'Add to the sys tray
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "LoadSystemTray", "modSysTray", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

