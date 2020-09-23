Attribute VB_Name = "Systray"
Option Explicit

Global blnClick As Boolean
Global vbTray As NOTIFYICONDATA

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long


Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_MOUSEMOVE = &H200
Public Const NIM_ADD = &H0
Public Const NIM_DELETE = &H2
Public Const NIF_ICON = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIM_MODIFY = &H1
Public Const NIF_TIP = &H4
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1

Public Sub SystrayOn(frm As Form, IconTooltipText As String)
 On Error Resume Next
    vbTray.cbSize = Len(vbTray)
    vbTray.hWnd = frm.hWnd
    vbTray.uId = vbNull
    vbTray.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    vbTray.uCallBackMessage = WM_MOUSEMOVE
    vbTray.szTip = Trim(IconTooltipText$) & vbNullChar
    vbTray.hIcon = frm.Icon
    
    Call Shell_NotifyIcon(NIM_ADD, vbTray)
    App.TaskVisible = False
End Sub

Public Sub SystrayOff(frm As Form)
 On Error Resume Next
    vbTray.cbSize = Len(vbTray)
    vbTray.hWnd = frm.hWnd
    vbTray.uId = vbNull
    Call Shell_NotifyIcon(NIM_DELETE, vbTray)
End Sub

Public Sub FormOnTop(frm As Form)
  On Error Resume Next
    Call SetWindowPos(frm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, FLAGS)
End Sub
