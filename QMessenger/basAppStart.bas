Attribute VB_Name = "basAppStart"
Option Explicit
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const WM_USER = &H400

Public Const WM_ICON_MSG As Long = WM_USER + 1

Public Const GWL_WNDPROC = (-4)
Public Const GWL_USERDATA = (-21)

Public Const WM_RBUTTONUP = &H205
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONUP = &H202
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCLBUTTONUP = &HA2

Public Const WM_GETMINMAXINFO = &H24

Public Const NIM_ADD = 0
Public Const NIM_MODIFY = 1
Public Const NIM_DELETE = 2
Public Const NIF_MESSAGE = 1
Public Const NIF_ICON = 2
Public Const NIF_TIP = 4

Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type MINMAXINFO
        ptReserved As POINTAPI
        ptMaxSize As POINTAPI
        ptMaxPosition As POINTAPI
        ptMinTrackSize As POINTAPI
        ptMaxTrackSize As POINTAPI
End Type

Public Const WM_HOTKEY = &H312
Public Const MOD_ALT = &H1
Public Const MOD_CONTROL = &H2
Public Const MOD_SHIFT = &H4

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long

Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function Shell_NotifyIconA Lib "Shell32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Integer

Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public gfnOldWndProc As Long
Public fMain As frmMain
Public gAppPath As String

Sub Main()
    Dim hIns As Long
    If App.PrevInstance Then
        hIns = FindWindow("ThunderRT6FormDC", "Messenger")
        If hIns Then PostMessage hIns, WM_ICON_MSG, 0, WM_LBUTTONDBLCLK
        Exit Sub
    End If

    gAppPath = App.Path
    If Right(gAppPath, 1) <> "\" Then gAppPath = gAppPath + "\"

    Set fMain = New frmMain
    On Error Resume Next
    Load fMain
End Sub

Public Function GetTime() As String
    Dim T As Date
    T = Now
    GetTime = Hour(T) & ":" & Minute(T) & ":" & Second(T)
End Function

Public Function WndProc(ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Msg = WM_ICON_MSG Then
        If lParam = WM_LBUTTONDBLCLK Then
            If Not fMain.Visible Then fMain.Show
        ElseIf lParam = WM_RBUTTONUP Then
            If Not fMain.Visible Then fMain.Show
            fMain.OnIconRightButtonUp
        ElseIf lParam = WM_LBUTTONUP Then
            If Not fMain.Visible Then fMain.Show
            fMain.OnIconLeftButtonUp
        End If
        Exit Function
    ElseIf Msg = WM_HOTKEY Then
        fMain.OnIconLeftButtonUp
        Exit Function
    ElseIf Msg = WM_NCLBUTTONDBLCLK Then
        fMain.Hide
    ElseIf Msg = WM_GETMINMAXINFO Then
        Dim mmiInfo As MINMAXINFO
        Call CopyMemory(mmiInfo, ByVal lParam, Len(mmiInfo))
        mmiInfo.ptMaxTrackSize.X = 96
        mmiInfo.ptMinTrackSize.X = 96
        mmiInfo.ptMinTrackSize.Y = 180
        Call CopyMemory(ByVal lParam, mmiInfo, Len(mmiInfo))
        Exit Function
    End If
    WndProc = CallWindowProc(gfnOldWndProc, hwnd, Msg, wParam, lParam)
End Function

