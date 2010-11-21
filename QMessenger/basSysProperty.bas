Attribute VB_Name = "basSysProperty"
Option Explicit

Public gbAlwaysOnTop As Boolean '������ǰ
Public gbAutoPopup As Boolean   '�Զ���������
Public gnPort As Long           'udp�˿ں�
Public gnFilePort As Long       'tcp�˿ں�
Public gbUseHotkey As Boolean   '�Ƿ�ʹ���ȼ�
Public gnKey As Integer         '�ȼ�
Public gnShift As Integer       '��ϼ�
Public gbNoVoiceChat As Boolean '�Ƿ�ʹ����������

Public Sub LoadSysProp()
    gbAutoPopup = GetSetting(App.Title, "Setting", "Popup", False)
    gbAlwaysOnTop = GetSetting(App.Title, "Setting", "OnTop", True)

    gbUseHotkey = GetSetting(App.Title, "Setting", "UseHotKey", True)
    gnKey = GetSetting(App.Title, "Setting", "HotKey", vbKeyX)
    If gnKey < vbKeyA Or gnKey > vbKeyZ Then gnKey = vbKeyX
    gnShift = GetSetting(App.Title, "Setting", "Shift", 6)
    If gnShift < 0 Or gnShift > 7 Then gnShift = 6

    gnPort = GetSetting(App.Title, "Setting", "Port", 12022)
    If gnPort < 0 Or gnPort > 65535 Then gnPort = 12022

    gnFilePort = GetSetting(App.Title, "Setting", "FilePort", 12090)
    If gnFilePort < 0 Or gnFilePort > 65535 Then gnFilePort = 12090

    gbNoVoiceChat = GetSetting(App.Title, "Setting", "NoVoiceChat", False)
End Sub
