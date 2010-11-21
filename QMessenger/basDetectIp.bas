Attribute VB_Name = "basDetectIp"
Option Explicit

Private Declare Function GetIpAddrTable Lib "iphlpapi.dll" (ByRef ipAddrTable As MIB_IPADDRTABLE, ByRef dwSize As Long, ByVal n As Long) As Long

Private Type MIB_IPADDRROW
    Addr1 As Byte
    Addr2 As Byte
    Addr3 As Byte
    Addr4 As Byte
    dwIndex As Long
    Mask1 As Byte
    Mask2 As Byte
    Mask3 As Byte
    Mask4 As Byte
    BCast1 As Byte
    BCast2 As Byte
    BCast3 As Byte
    BCast4 As Byte
    dwReasmSize As Long
    unused1 As Integer
    unused2 As Integer
End Type

Private Type MIB_IPADDRTABLE
    dwNumEntries As Long
    IpTable(9) As MIB_IPADDRROW '10 IPs
End Type

Private Sub MakeBCastAddr(ByRef ipAddr As MIB_IPADDRROW)
    With ipAddr
        '获得广播地址：
        .BCast1 = .Addr1 Or (Not .Mask1)
        .BCast2 = .Addr2 Or (Not .Mask2)
        .BCast3 = .Addr3 Or (Not .Mask3)
        .BCast4 = .Addr4 Or (Not .Mask4)
    End With
End Sub

Public Function GetSubNetMask(ByVal localIp As String, ByRef localBroadcast As String) As Boolean
    Dim i As Long, nSize As Long
    Dim ips As MIB_IPADDRTABLE
    Dim IpColl As Collection
    Dim BCastColl As Collection
    Dim s As String

    nSize = LenB(ips)
    If GetIpAddrTable(ips, nSize, 0) <> 0 Then
        MsgBox "无法获得本地网卡的IP地址。请检查本地网络连接。", vbCritical Or vbOKOnly
        Exit Function
    End If

    Set IpColl = New Collection
    Set BCastColl = New Collection

    For i = 0 To ips.dwNumEntries - 1
        Call MakeBCastAddr(ips.IpTable(i))
        Call AddIpCollection(IpColl, BCastColl, _
            (ips.IpTable(i).Addr1 & "." & ips.IpTable(i).Addr2 & "." & ips.IpTable(i).Addr3 & "." & ips.IpTable(i).Addr4), _
            (ips.IpTable(i).BCast1 & "." & ips.IpTable(i).BCast2 & "." & ips.IpTable(i).BCast3 & "." & ips.IpTable(i).BCast4))
    Next i

    If IpColl.Count = 0 Then '没有可用的网络连接：
        MsgBox "网络IP地址无效！请检查网络设置！", vbCritical Or vbOKOnly
    ElseIf IpColl.Count = 1 Then '有一个可用的IP：
        localBroadcast = BCastColl.Item(1)
        GetSubNetMask = True
    ElseIf IpColl.Count > 1 Then '有2个或更多的IP：
        Dim bFound As Boolean
        For i = 1 To IpColl.Count
            If IpColl.Item(i) = localIp Then
                localBroadcast = BCastColl.Item(i)
                bFound = True
                Exit For
            End If
        Next i
        If Not bFound Then localBroadcast = "127.0.0.1"
        GetSubNetMask = True
    End If

    Set BCastColl = Nothing
    Set IpColl = Nothing
End Function

Private Sub AddIpCollection(ByRef cln As Collection, ByRef mskcln As Collection, ByVal sip As String, ByVal smask As String)
    Dim ip
    For Each ip In cln
        If ip = sip Then Exit Sub
    Next ip
    If sip = "127.0.0.1" Then Exit Sub
    If sip = "0.0.0.0" Then Exit Sub
    cln.Add sip
    mskcln.Add smask
End Sub


