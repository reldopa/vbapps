Attribute VB_Name = "basPublicVars"
Option Explicit

Public Const FRAME_HEADER_SIZE = 100
Public Const SPLIT_CHAR As String = "@"

Public Enum FRAME_HEADER
    'user's message:
    HEADER_USER_SEND_MESSAGE = 404

    'user logon and logoff:
    HEADER_USER_ONLINE = 202
    HEADER_USER_OFFLINE = 204

    'broadcast message:
    HEADER_USER_BROADCAST = 100
    HEADER_SYS_BROADCAST = 101

    'get information:
    HEADER_UPDATE_USER_INFO = 302
    HEADER_SEND_MY_DETAIL = 323
    HEADER_QUERY_INFO = 398
    'file transfer info:
    HEADER_WANT_TO_TRANSFER = 783
    'want start chatting:
    HEADER_CHATTING = 901
End Enum

Public subnetMask As String

Public Sub UnpackChatRequest(ByRef sPacket As String)
    '收到语音聊天请求后：
    If Len(sPacket) <> FRAME_HEADER_SIZE Then Exit Sub
    If gbNoVoiceChat Then Exit Sub

    If Dir(gAppPath + "vdpchat.exe") = "" Then Exit Sub

    Dim s() As String
    s = Split(Trim(sPacket), SPLIT_CHAR, 2)
    If UBound(s) = 1 Then
        Call PlayRingSound
        If MsgBox("用户“" + s(1) + "”(IP: " + s(0) + ")想和你进行语音聊天，是否同意？", vbQuestion Or vbYesNo) = vbYes Then
            '启动语音聊天：
            Err.Clear
            On Error Resume Next
            Call Shell(gAppPath + "vdpchat.exe -wait -name " + s(0), vbNormalFocus)
        End If
    End If
End Sub

Public Function PackChatRequest(ByVal ip As String, ByVal sName As String)
    Dim s As String
    s = ip & SPLIT_CHAR & sName
    PackChatRequest = s + String(FRAME_HEADER_SIZE - Len(s), " ")
End Function

Public Function PackUserInfo(ByRef ui As USER_INFO) As String
    Dim s As String
    s = ui.nFace & SPLIT_CHAR & ui.sHostIP + SPLIT_CHAR + ui.sHostName + SPLIT_CHAR + ui.sName
    PackUserInfo = s + String(FRAME_HEADER_SIZE - Len(s), " ")
End Function

Public Function UnpackUserInfo(ByRef sPacket As String, ByRef ui As USER_INFO) As Boolean
    If Len(sPacket) <> FRAME_HEADER_SIZE Then Exit Function

    Dim s() As String
    s = Split(Trim(sPacket), SPLIT_CHAR, 4)
    If UBound(s) = 3 Then
        ui.nFace = Int(Val(s(0)))
        If ui.nFace < 1 Then ui.nFace = 1
        If ui.nFace > 40 Then ui.nFace = 40
        ui.sHostIP = s(1)
        ui.sHostName = s(2)
        ui.sName = s(3)
        UnpackUserInfo = True
    End If
End Function

Public Function PackFileInfo(ByVal sFileName As String, ByVal nSize As Long, ByVal nPort As Long, ByVal sAbstract As String) As String
   PackFileInfo = Encoding(myInfo.sName + "/" + myInfo.sHostIP) _
                + SPLIT_CHAR + Encoding(sFileName) _
                + SPLIT_CHAR & nSize & SPLIT_CHAR & nPort & SPLIT_CHAR + Encoding(sAbstract)
End Function
