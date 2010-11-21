Attribute VB_Name = "basMyProperty"
Option Explicit

Public Const ENCODING_CHAR As String = "~at$t;"

Public myInfo As USER_INFO

Public gsTrueName As String '真实姓名
Public gsAge As String '年龄
Public gsSex As String '性别
Public gsBlood As String '血型
Public gsEmail As String '电子邮件
Public gsHomepage As String '个人主页
Public gsConstellation As String '星座
Public gsSuxiang As String '属相
Public gsTel As String '固定电话
Public gsMobile As String '移动电话
Public gsAbstract As String '简介

Public Function PackMyDetailInfo() As String '封装我的详细信息
   PackMyDetailInfo = myInfo.nFace & SPLIT_CHAR _
                    + Encoding(myInfo.sName) + SPLIT_CHAR _
                    + Encoding(gsTrueName) + SPLIT_CHAR _
                    + Encoding(gsAge) + SPLIT_CHAR _
                    + Encoding(gsSex) + SPLIT_CHAR _
                    + Encoding(gsBlood) + SPLIT_CHAR _
                    + Encoding(gsEmail) + SPLIT_CHAR _
                    + Encoding(gsHomepage) + SPLIT_CHAR _
                    + Encoding(gsConstellation) + SPLIT_CHAR _
                    + Encoding(gsSuxiang) + SPLIT_CHAR _
                    + Encoding(gsTel) + SPLIT_CHAR _
                    + Encoding(gsMobile) + SPLIT_CHAR _
                    + Encoding(gsAbstract)
End Function

'收到用户详细信息后显示：
Public Sub UnpackUserDetailInfo(ByRef sPacket As String)
    Dim s() As String
    s = Split(sPacket, SPLIT_CHAR, 13)

    If UBound(s) <> 12 Then Exit Sub

    Dim fDetail As frmDetail, n As Long
    Set fDetail = New frmDetail
    Load fDetail
    With fDetail
        n = Int(Val(s(0)))
        If n < 1 Then n = 1
        If n > 40 Then n = 40
        Set .imgFace.Picture = fMain.imglstFace.ListImages(n).ExtractIcon()
        On Error GoTo 0
        .txtName = Decoding(s(1))
        .txtTrueName = Decoding(s(2))
        .txtAge = Decoding(s(3))
        .txtSex = Decoding(s(4))
        .txtBlood = Decoding(s(5))
        .txtEmail = Decoding(s(6))
        .txtHomepage = Decoding(s(7))
        .txtConstellation = Decoding(s(8))
        .txtSuxiang = Decoding(s(9))
        .txtTel = Decoding(s(10))
        .txtMobile = Decoding(s(11))
        .txtAbstract = Decoding(s(12))
    End With
    fDetail.Show vbModeless 'al, fMain
End Sub

Public Function Encoding(ByRef s As String) As String
    Encoding = Replace(s, SPLIT_CHAR, ENCODING_CHAR)
End Function

Public Function Decoding(ByRef s As String) As String
    Decoding = Replace(s, ENCODING_CHAR, SPLIT_CHAR)
End Function

Public Sub LoadMyInfo()
    Dim L As Single, T As Single
    L = GetSetting(App.Title, "Window", "Left", 1000)
    T = GetSetting(App.Title, "Window", "Top", 900)
    fMain.Move L, T

    With myInfo
        .nFace = GetSetting(App.Title, "Setting", "Face", 1)
        If .nFace < 1 Then .nFace = 1
        If .nFace > 40 Then .nFace = 40
        .sName = Trim(GetSetting(App.Title, "Setting", "Name", fMain.wsUDPSocket.LocalHostName))
        If Len(.sName) > 20 Then
            .sName = Left(.sName, 20)
        End If
    End With

    gsTrueName = GetSetting(App.Title, "Setting", "TrueName", "")
    gsAge = GetSetting(App.Title, "Setting", "Age", "")
    gsSex = GetSetting(App.Title, "Setting", "Sex", "")
    gsBlood = GetSetting(App.Title, "Setting", "Blood", "")
    gsEmail = GetSetting(App.Title, "Setting", "Email", "")
    gsHomepage = GetSetting(App.Title, "Setting", "Homepage", "")
    gsConstellation = GetSetting(App.Title, "Setting", "Constellation", "")
    gsSuxiang = GetSetting(App.Title, "Setting", "Suxiang", "")
    gsTel = GetSetting(App.Title, "Setting", "Tel", "")
    gsMobile = GetSetting(App.Title, "Setting", "Mobile", "")
    gsAbstract = GetSetting(App.Title, "Setting", "Abstract", "")
End Sub

Public Sub SaveMyInfo()
    With myInfo
        SaveSetting App.Title, "Window", "Left", fMain.Left
        SaveSetting App.Title, "Window", "Top", fMain.Top

        SaveSetting App.Title, "Setting", "Face", .nFace
        SaveSetting App.Title, "Setting", "Name", Trim(.sName)
    End With

    SaveSetting App.Title, "Setting", "TrueName", gsTrueName
    SaveSetting App.Title, "Setting", "Age", gsAge
    SaveSetting App.Title, "Setting", "Sex", gsSex
    SaveSetting App.Title, "Setting", "Blood", gsBlood
    SaveSetting App.Title, "Setting", "Email", gsEmail
    SaveSetting App.Title, "Setting", "Homepage", gsHomepage
    SaveSetting App.Title, "Setting", "Constellation", gsConstellation
    SaveSetting App.Title, "Setting", "Suxiang", gsSuxiang
    SaveSetting App.Title, "Setting", "Tel", gsTel
    SaveSetting App.Title, "Setting", "Mobile", gsMobile
    SaveSetting App.Title, "Setting", "Abstract", gsAbstract
End Sub

