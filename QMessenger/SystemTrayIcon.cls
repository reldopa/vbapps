VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SystemTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_tip As String
Private m_hWnd As Long
Private Const m_ID As Long = 9011
Private m_hIcon As Long
Private m_Icon As IPictureDisp
Private nid As NOTIFYICONDATA
Private HadAdd As Boolean

'设定Tip
Public Property Get Tip() As String
    Tip = m_tip
End Property

Public Property Let Tip(ByVal vNewValue As String)
    m_tip = vNewValue
End Property

'设定Icon的图标
Public Property Set Icon(ByVal vNewValue As IPictureDisp)
    Set m_Icon = Nothing
    Set m_Icon = vNewValue
    m_hIcon = m_Icon.Handle
End Property

'将原先的Form隐藏，并在右下方加入一个Icon，传入的是待处理的Form.hWnd
Public Function AddSysTrayIcon(ByVal formWnd As Long) As Boolean
    If Not HadAdd Then
        Call Shell_NotifyIconA(NIM_DELETE, nid)
        nid.cbSize = Len(nid)
        nid.hwnd = formWnd
        nid.uID = m_ID
        nid.uFlags = NIF_ICON + NIF_TIP + NIF_MESSAGE
        nid.hIcon = m_hIcon
        nid.szTip = m_tip + Chr(0)
        nid.uCallbackMessage = WM_ICON_MSG
        Dim i As Integer
        If Shell_NotifyIconA(NIM_ADD, nid) = 1 Then '新增成功
            gfnOldWndProc = GetWindowLong(formWnd, GWL_WNDPROC)
            Call SetWindowLong(formWnd, GWL_WNDPROC, AddressOf WndProc)
            AddSysTrayIcon = True
            HadAdd = True
            m_hWnd = formWnd
        End If
    End If
End Function

'删除右下方的Icon
Public Sub DeleteSysTrayIcon()
    If gfnOldWndProc <> 0 Then
        Call SetWindowLong(m_hWnd, GWL_WNDPROC, gfnOldWndProc)
        gfnOldWndProc = 0
    End If
    If HadAdd Then
        Call Shell_NotifyIconA(NIM_DELETE, nid)
        HadAdd = False
    End If
End Sub

'修改Icon的设定 - 图标与Tip
Public Function ModifySysTrayIcon() As Boolean
    If HadAdd Then
        nid.hIcon = m_hIcon
        nid.szTip = m_tip + Chr(0)
        If Shell_NotifyIconA(NIM_MODIFY, nid) = 1 Then
            ModifySysTrayIcon = True
        End If
    End If
End Function

Private Sub Class_Initialize()
    HadAdd = False
End Sub

Private Sub Class_Terminate()
    Call DeleteSysTrayIcon
End Sub

