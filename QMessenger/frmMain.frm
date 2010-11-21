VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Messenger"
   ClientHeight    =   7815
   ClientLeft      =   165
   ClientTop       =   435
   ClientWidth     =   1320
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   1320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrFlashIcon 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   120
      Top             =   3720
   End
   Begin VB.PictureBox picButtom 
      Align           =   2  'Align Bottom
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H00E0E0E0&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   1320
      TabIndex        =   1
      Top             =   7155
      Width           =   1320
      Begin QMessenger.BitmapButton cmdAbout 
         Height          =   315
         Left            =   10
         TabIndex        =   5
         Top             =   340
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         HotPicture      =   "frmMain.frx":000C
         Picture         =   "frmMain.frx":059E
      End
      Begin QMessenger.BitmapButton cmdBroadcast 
         Height          =   315
         Left            =   350
         TabIndex        =   4
         Top             =   340
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         HotPicture      =   "frmMain.frx":0B30
         Picture         =   "frmMain.frx":1B42
      End
      Begin QMessenger.BitmapButton cmdShowMsg 
         Height          =   315
         Left            =   350
         TabIndex        =   3
         Top             =   20
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         HotPicture      =   "frmMain.frx":2B54
         Picture         =   "frmMain.frx":3B66
      End
      Begin QMessenger.BitmapButton cmdSetting 
         Height          =   315
         Left            =   10
         TabIndex        =   2
         Top             =   20
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   556
         HotPicture      =   "frmMain.frx":4B78
         Picture         =   "frmMain.frx":510A
      End
   End
   Begin MSComctlLib.ImageList imglstFace 
      Left            =   120
      Top             =   6000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   8421376
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":569C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5928
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5BBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5E24
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":609C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6320
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6820
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6D30
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6FA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7230
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":74C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7750
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":79E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C78
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":819C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":841C
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8688
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8900
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8B74
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8DDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9068
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":92F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9584
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":97F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9A74
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9D00
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F90
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A204
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A478
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A6E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A968
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ABD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AE64
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B0D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B36C
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B600
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B87C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsUDPSocket 
      Left            =   840
      Top             =   6120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      RemotePort      =   12022
      LocalPort       =   12022
   End
   Begin MSComctlLib.ListView lvGroup 
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   10186
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      HotTracking     =   -1  'True
      PictureAlignment=   5
      _Version        =   393217
      Icons           =   "imglstFace"
      SmallIcons      =   "imglstFace"
      ForeColor       =   -2147483640
      BackColor       =   -2147483636
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Menu mnuRight 
      Caption         =   "right_menu"
      Visible         =   0   'False
      Begin VB.Menu mnuSendMsg 
         Caption         =   "发送消息(&S)..."
      End
      Begin VB.Menu mnuBar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChat 
         Caption         =   "语音聊天(&C)..."
      End
      Begin VB.Menu mnuTransferFile 
         Caption         =   "传送文件(&T)..."
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBoradcastMsg 
         Caption         =   "发送广播(&B)..."
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDetail 
         Caption         =   "详细资料(&D)..."
      End
   End
   Begin VB.Menu mnuIcon 
      Caption         =   "icon_menu"
      Visible         =   0   'False
      Begin VB.Menu mnuIconBoradcastMsg 
         Caption         =   "发送广播(&B)..."
      End
      Begin VB.Menu mnuBar11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconSetting 
         Caption         =   "参数设置(&S)..."
      End
      Begin VB.Menu mnuBar12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconAbout 
         Caption         =   "关于(&A)..."
      End
      Begin VB.Menu mnuBar14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconExit 
         Caption         =   "退出(&X)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TIP_NORMAL As String = "QQ Messenger - 正在运行"
Private Const TIP_MESSAGE As String = "QQ Messenger - 有新消息啦！"

Dim bDisableIconMenu As Boolean

Dim dblX As Single, dblY As Single  '记录双击时的鼠标位置
Dim fBroadcast As frmBroadcastMsg   '广播窗体
Public fTransfer As frmFileTransfer '传送文件窗体
Public fReceive As frmFileReceive   '接收文件窗体
Dim OnlineUsers As Collection       '记录所有的在线用户
Dim menuSelectItem As ListItem      '鼠标右键选择的项

Dim SysTrayIcon As SystemTrayIcon

Dim msgForms As Collection

'添加Form到消息集合: //////////////////////////////////////////////////////////////////////////////
Public Sub AddMsgForm(ByRef fForm As Form)
    Dim f As Form
    For Each f In msgForms
        If f Is fForm Then Exit Sub
    Next f
    msgForms.Add fForm
    Call FlashIcon(True)
End Sub

Public Function IsReceiveFormExistInMsgForm() As Boolean
    Dim f As Form
    For Each f In msgForms
        If f Is fReceive Then
            IsReceiveFormExistInMsgForm = True
            Exit Function
        End If
    Next f
End Function

Public Function IsTransferFormExistInMsgForm() As Boolean
    Dim f As Form
    For Each f In msgForms
        If f Is fTransfer Then
            IsTransferFormExistInMsgForm = True
            Exit Function
        End If
    Next f
End Function

Public Function IsBroadcastFormExistInMsgForm() As Boolean
    Dim f As Form
    For Each f In msgForms
        If f Is fBroadcast Then
            IsBroadcastFormExistInMsgForm = True
            Exit Function
        End If
    Next f
End Function
'//////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub cmdAbout_Click()
    bDisableIconMenu = True
    frmAbout.Show vbModal, Me
    bDisableIconMenu = False
End Sub

Private Sub cmdBroadcast_Click()
    fBroadcast.Show vbModeless
    fBroadcast.WindowState = 0
End Sub

Private Sub cmdSetting_Click()
    bDisableIconMenu = True
    frmSettings.Show vbModal, Me
    bDisableIconMenu = False
End Sub

Private Sub cmdShowMsg_Click()
    '显示消息：
    Call OnIconLeftButtonUp
End Sub

Private Sub Form_Load()
    Dim dwHeight As Single
    dwHeight = GetSetting(App.Title, "Setting", "WinHeight", 2700)
    Me.Height = dwHeight

    Call LoadSysProp '读系统设置

    Set msgForms = New Collection

    With wsUDPSocket
        '绑定本地 UDP 聊天端口：
        .LocalPort = gnPort
        .RemotePort = gnPort
        Err.Clear
        On Error Resume Next
        .Bind
        If Err.Number Then
            MsgBox "初始化网络失败!" + vbCrLf + "错误: " + Err.Description, vbCritical Or vbOKOnly, "初始化网络失败"
            Unload Me
            Exit Sub
        End If

        '设置子网掩码：
        If Not GetSubNetMask(.localIp, subnetMask) Then
            Unload Me
            Exit Sub
        End If

        '加载文件接收窗体：
        Set fReceive = New frmFileReceive
        Load fReceive
        '加载文件传送窗体：
        Set fTransfer = New frmFileTransfer
        Load fTransfer
        '设置本地TCP端口：
        fTransfer.wsTcpSend.LocalPort = gnFilePort
        fReceive.wsTcpReceive.LocalPort = gnFilePort

        '设置我的信息：
        Set myInfo = New USER_INFO
        Call LoadMyInfo

        myInfo.sHostName = .LocalHostName
        myInfo.sHostIP = .localIp
    End With

    Set OnlineUsers = New Collection
    Set fBroadcast = New frmBroadcastMsg
    Load fBroadcast
    Set fTransfer = New frmFileTransfer
    Load fTransfer
    '上线通知：
    Call BroadcastIAmOnline

    '使窗口总在最前：
    If gbAlwaysOnTop Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If

    'set hotkey:
    If gbUseHotkey Then
        Dim cmod As Long
        Select Case gnShift
        Case 0: cmod = 0
        Case 1: cmod = MOD_SHIFT 'Shift
        Case 2: cmod = MOD_CONTROL 'Ctrl
        Case 3: cmod = MOD_SHIFT + MOD_CONTROL 'Shift+Ctrl
        Case 4: cmod = MOD_ALT 'Alt
        Case 5: cmod = MOD_SHIFT + MOD_ALT 'Shift+Alt
        Case 6: cmod = MOD_CONTROL + MOD_ALT 'Ctrl+Alt
        Case 7: cmod = MOD_SHIFT + MOD_CONTROL + MOD_ALT 'Shift+Ctrl+Alt
        End Select
        If RegisterHotKey(Me.hwnd, 1, cmod, gnKey) = 0 Then
            gbUseHotkey = False
            MsgBox "热键注册失败！请重新设置热键。", vbCritical Or vbOKOnly
        End If
    End If

    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    ' SubClass:
    Set SysTrayIcon = New SystemTrayIcon
    Set SysTrayIcon.Icon = LoadResPicture(101, vbResIcon)
    SysTrayIcon.Tip = TIP_NORMAL
    Call SysTrayIcon.AddSysTrayIcon(Me.hwnd)
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    Me.Visible = True
End Sub

Private Sub Form_Resize()
    lvGroup.Move 0, 0, 1335, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    ' Un-SubClass:
    If Not SysTrayIcon Is Nothing Then Call SysTrayIcon.DeleteSysTrayIcon
    Set SysTrayIcon = Nothing
    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

    If gbUseHotkey Then
        Call UnregisterHotKey(Me.hwnd, 1)
    End If

    Set OnlineUsers = Nothing
    Set msgForms = Nothing

    If Not fBroadcast Is Nothing Then
        '下线通知：
        Call BroadcastIAmOffline
        Unload fBroadcast

        '关闭网络连接：
        wsUDPSocket.Close
        Call SaveMyInfo
        Set myInfo = Nothing
    End If

    '关闭所有窗口：
    Dim f As Form
    For Each f In Forms
        If Not f Is Me Then Unload f
    Next f

    SaveSetting App.Title, "Setting", "WinHeight", Me.Height
End Sub

Private Sub mnuBoradcastMsg_Click()
    fBroadcast.Show vbModeless
    fBroadcast.WindowState = 0
End Sub

Private Sub mnuChat_Click()
    If Not IsItemExist(menuSelectItem) Then Exit Sub

    If Dir(gAppPath + "vdpchat.exe") = "" Then
        MsgBox "无法使用语音聊天。原因：应用程序“vdpchat.exe”不存在。", vbCritical Or vbOKOnly
        Exit Sub
    End If

    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        '发送请求：
        .RemoteHost = menuSelectItem.Tag
        .SendData HEADER_CHATTING & PackChatRequest(menuSelectItem.Tag, menuSelectItem.Text)
    End With
    If Err.Number = 0 Then
        '发送请求成功！准备启动语音聊天：
        Call Shell(gAppPath + "vdpchat.exe -target" + menuSelectItem.Tag + " -name " + menuSelectItem.Text, vbNormalFocus)
    End If
End Sub

Private Sub mnuIconAbout_Click()
    cmdAbout_Click
End Sub

Private Sub mnuIconBoradcastMsg_Click()
    cmdBroadcast_Click
End Sub

Private Sub mnuIconExit_Click()
    Unload Me
End Sub

Private Sub mnuIconSetting_Click()
    cmdSetting_Click
End Sub

Private Sub mnuTransferFile_Click()
    If Not IsItemExist(menuSelectItem) Then Exit Sub

    If fTransfer.Visible Then
        '上一任务尚未完成，无法进行当前任务，退出：
        Exit Sub
    End If

    If fReceive.Visible Then
        Exit Sub
    End If

    '传输前准备：
    Call fTransfer.TransferPrepare(menuSelectItem.Tag, menuSelectItem.Text)
    fTransfer.Show vbModeless
    fTransfer.WindowState = 0
End Sub

Private Sub mnuViewDetail_Click()
    If Not IsItemExist(menuSelectItem) Then Exit Sub

    Call QueryUserDetail(menuSelectItem.Tag)
End Sub

Private Sub lvGroup_DblClick() '双击时弹出聊天对话框
    Dim Item As ListItem
    Set Item = lvGroup.HitTest(dblX, dblY)
    If Item Is Nothing Then Exit Sub

    Dim fMsg As frmMsg
    If FindUserByIP(Item.Tag, fMsg) Then
        fMsg.Show vbModeless
        fMsg.WindowState = 0
        fMsg.txtMsg.SetFocus
    End If
End Sub

Private Function IsItemExist(ByVal refItem As ListItem) As Boolean
    If refItem Is Nothing Then Exit Function

    Dim i As Long
    With lvGroup.ListItems
        For i = 1 To .Count
            If .Item(i) Is refItem Then
                IsItemExist = True
                Exit Function
            End If
        Next i
    End With
End Function

Private Sub mnuSendMsg_Click() '点菜单“发送消息”
    If Not IsItemExist(menuSelectItem) Then Exit Sub

    Dim fMsg As frmMsg
    If FindUserByIP(menuSelectItem.Tag, fMsg) Then
        fMsg.Show vbModeless
        fMsg.WindowState = 0
        fMsg.txtMsg.SetFocus
    End If
End Sub

Private Sub lvGroup_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Item As ListItem
    Set Item = lvGroup.HitTest(X, Y)
    If Item Is Nothing Then
        lvGroup.ToolTipText = ""
    Else
        lvGroup.ToolTipText = "IP: " + Item.Tag
    End If
End Sub

Private Sub lvGroup_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    dblX = X
    dblY = Y
    If Button = vbRightButton Then
        Dim it As ListItem
        Set it = lvGroup.HitTest(X, Y)
        If it Is Nothing Then Exit Sub

        Set menuSelectItem = it
        Me.PopupMenu mnuRight, vbLeftButton Or vbRightButton
    End If
End Sub

'收到数据包:
Private Sub wsUDPSocket_DataArrival(ByVal bytesTotal As Long)
    Err.Clear
    On Error Resume Next
    Dim s As String, Msg As String
    wsUDPSocket.GetData s
    If Len(s) < 3 Then Exit Sub '无效数据包

    Msg = Right(s, Len(s) - 3)
    Select Case Int(Val(Left(s, 3)))
        Case HEADER_USER_SEND_MESSAGE:    '收到用户信息:
            Call OnUserMessageArrival(Msg)
        Case HEADER_USER_ONLINE:          '某个用户上线:
            Call OnUserOnline(Msg)
        Case HEADER_USER_OFFLINE:         '某个用户下线:
            Call OnUserOffline(Msg)
        Case HEADER_USER_BROADCAST:       '收到用户广播:
            Call OnUserBroadcasting(Msg)
        Case HEADER_SYS_BROADCAST:

        Case HEADER_UPDATE_USER_INFO:     '收到某个用户的更新信息:
            Call OnUpdateUserInfo(Msg)
        Case HEADER_SEND_MY_DETAIL:       '收到某个用户的信息
            Call UnpackUserDetailInfo(Msg)
        Case HEADER_QUERY_INFO:           '某个用户要求返回个人信息
            Call OnQueryMyDetailInfo(Msg)
        Case HEADER_WANT_TO_TRANSFER:     '某个用户要求传送文件
            Call OnFileRequest(Msg)
        Case HEADER_CHATTING:             '某个用户要求语音聊天
            Call UnpackChatRequest(Msg)
    End Select
End Sub

'**************************************************************************************************
'** 发送消息 ** 发送消息 ** 发送消息 ** 发送消息 ** 发送消息 ** 发送消息 ** 发送消息 ** 发送消息 **
'**************************************************************************************************

'向其他用户发送消息:
Public Sub SendMessageToUser(ByVal ipAddr As String, ByVal sMsg As String)
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = ipAddr
        .SendData HEADER_USER_SEND_MESSAGE & PackUserInfo(myInfo) + sMsg
    End With
End Sub

'查询用户详细信息：
Public Sub QueryUserDetail(ByVal ipAddr As String)
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = ipAddr
        .SendData HEADER_QUERY_INFO & myInfo.sHostIP
    End With
End Sub

'向所有用户广播消息:
Public Sub SendMessageToAll(ByVal sMsg As String)
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = subnetMask
        .SendData HEADER_USER_BROADCAST & PackUserInfo(myInfo) + sMsg
    End With
End Sub

'发送“传文件”请求：
Public Sub SendFileRequest(ByVal ipAddr As String, ByVal sFileInfo As String)
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = ipAddr
        .SendData HEADER_WANT_TO_TRANSFER & sFileInfo
    End With
End Sub

'通知所有用户“我上线了”:
Public Sub BroadcastIAmOnline()
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = subnetMask
        .SendData HEADER_USER_ONLINE & PackUserInfo(myInfo)
    End With
End Sub

'通知所有用户“我下线了”:
Public Sub BroadcastIAmOffline()
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = subnetMask
        .SendData HEADER_USER_OFFLINE & PackUserInfo(myInfo)
    End With
End Sub

'向所有用户更新我的信息：
Public Sub UpdateMyInfo()
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = subnetMask
        .SendData HEADER_UPDATE_USER_INFO & PackUserInfo(myInfo)
    End With
End Sub

'刷新所有用户：
Public Sub RefreshAllUser()
    'delete all user:
    'While lvGroup.ListItems.Count > 0
    '    lvGroup.ListItems.Remove 1
    'Wend
    'then re-add them:
    'Call UpdateMyInfo
End Sub

'**************************************************************************************************
'** 消息处理 ** 消息处理 ** 消息处理 ** 消息处理 ** 消息处理 ** 消息处理 ** 消息处理 ** 消息处理 **
'**************************************************************************************************

'应请求发送我的详细资料：
Public Sub OnQueryMyDetailInfo(ByVal ipAddr As String)
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = ipAddr
        .SendData HEADER_SEND_MY_DETAIL & PackMyDetailInfo
    End With
End Sub

'收到单个用户消息：
Private Sub OnUserMessageArrival(ByVal s As String)
    If Len(s) < FRAME_HEADER_SIZE Then Exit Sub

    Dim sFrame As String, sMsg As String
    sFrame = Left(s, FRAME_HEADER_SIZE)
    sMsg = Right(s, Len(s) - FRAME_HEADER_SIZE)

    Dim ui As New USER_INFO
    '试图获得用户信息：
    If UnpackUserInfo(sFrame, ui) Then
        If ui.sHostIP <> myInfo.sHostIP Then
            '试图找到与该用户相联系的类窗口：
            Dim fMsg As frmMsg
            If Not FindUserByUI(ui, fMsg) Then
                Set fMsg = AddNewUserByUI(ui)
                Call AddUserIcon(ui)
            End If
            Set ui = Nothing
            '调用 fMsg 显示消息:
            Call PlayMsgSound
            Call fMsg.OnMessageArrival(sMsg)
        End If
    End If
End Sub

'收到用户传送文件请求：
Private Sub OnFileRequest(ByVal sPacket As String)
    If fTransfer.Visible Then Exit Sub '正在传送，因此无法接收文件
    If fReceive.Visible Then Exit Sub  '正在接收，因此无法接收文件
    If IsReceiveFormExistInMsgForm Then Exit Sub '已有请求但用户未相应，因此无法接收文件
    If IsTransferFormExistInMsgForm Then Exit Sub  '已有请求但用户未相应，因此无法接收文件

    Call PlayFileSound
    Call fReceive.QueryReceiveFile(sPacket)
End Sub

'收到用户广播：
Private Sub OnUserBroadcasting(ByVal s As String)
    If Len(s) < FRAME_HEADER_SIZE Then Exit Sub

    Dim sFrame As String, sMsg As String
    sFrame = Left(s, FRAME_HEADER_SIZE)
    sMsg = Right(s, Len(s) - FRAME_HEADER_SIZE)

    Dim ui As New USER_INFO
    If UnpackUserInfo(sFrame, ui) Then
        If ui.sHostIP <> myInfo.sHostIP Then
            Call PlayMsgSound
            Call fBroadcast.OnBroadcastArrival(sMsg, ui.sHostName, ui.sHostIP)
        End If
        Set ui = Nothing
    End If
End Sub

'收到用户上线通知：
Private Sub OnUserOnline(ByVal sFrame As String)
    If Len(sFrame) <> FRAME_HEADER_SIZE Then Exit Sub

    Dim ui As New USER_INFO
    Dim fMsg As frmMsg

    If UnpackUserInfo(sFrame, ui) Then
        If ui.sHostIP <> myInfo.sHostIP Then '该消息如果不是我发的：
            Call PlayUserSound
            If FindUserByUI(ui, fMsg) Then
                fMsg.ui.sHostName = ui.sHostName
                fMsg.ui.sName = ui.sName
            Else
                Call AddNewUserByUI(ui) '就添加一个新用户
            End If
            '添加用户头像：
            Call AddUserIcon(ui)
            '发送 UPDATE_MY_INFO 使该新用户添加我:
            Call UpdateMyInfo

        End If
    End If
    Set ui = Nothing
End Sub

'收到用户下线通知
Private Sub OnUserOffline(ByVal sFrame As String)
    If Len(sFrame) <> FRAME_HEADER_SIZE Then Exit Sub

    Dim ui As New USER_INFO
    Dim fMsg As frmMsg

    If UnpackUserInfo(sFrame, ui) Then
        If ui.sHostIP <> myInfo.sHostIP Then '该消息如果不是我发的：
            Call PlayUserSound
            '删除该用户头像，但保留该用户谈话对话框：
            Dim n As Long
            With lvGroup.ListItems
                For n = 1 To .Count
                    If .Item(n).Tag = ui.sHostIP Then
                        .Remove n
                        Exit For
                    End If
                Next n
            End With
        End If
    End If
    Set ui = Nothing
End Sub

'收到用户更新信息
Private Sub OnUpdateUserInfo(ByVal sFrame As String)
    If Len(sFrame) <> FRAME_HEADER_SIZE Then Exit Sub

    Dim ui As New USER_INFO
    Dim fMsg As frmMsg

    If UnpackUserInfo(sFrame, ui) Then
        If ui.sHostIP <> myInfo.sHostIP Then '该消息如果不是我发的：
            If FindUserByUI(ui, fMsg) Then
                fMsg.ui.sHostName = ui.sHostName
                fMsg.ui.sName = ui.sName
                fMsg.ui.nFace = ui.nFace
                fMsg.UpdateUserInfo

                '更新图标:
                Dim it As ListItem
                For Each it In lvGroup.ListItems
                    If it.Tag = ui.sHostIP Then
                        it.Icon = ui.nFace
                        it.SmallIcon = ui.nFace
                        it.Text = ui.sName
                        Exit For
                    End If
                Next it
                AddUserIcon (ui)
            Else
                Call AddNewUserByUI(ui) '就添加一个新用户
                Call AddUserIcon(ui)
                '发送 UPDATE_MY_INFO 使该新用户添加我:
                Call UpdateMyInfo
            End If
        End If
    End If
    Set ui = Nothing
End Sub

'==============================================================================

Public Sub AddUserIcon(ByRef ui As USER_INFO)
    '添加用户头像：
    Dim it As ListItem, bFound As Boolean
    For Each it In lvGroup.ListItems
        If it.Tag = ui.sHostIP Then
            bFound = True
            Exit For
        End If
    Next it
    If Not bFound Then
        Set it = lvGroup.ListItems.Add(, , ui.sName, ui.nFace, ui.nFace)
        it.Tag = ui.sHostIP
    End If
End Sub

Public Function AddNewUserByUI(ByRef ui As USER_INFO) As frmMsg
    '添加新用户到集合，用 IP 地址作为标识符:
    Dim fM As frmMsg
    If FindUserByUI(ui, fM) Then Exit Function

    Set fM = New frmMsg
    Load fM
    fM.ui.nFace = ui.nFace
    fM.ui.sHostIP = ui.sHostIP
    fM.ui.sHostName = ui.sHostName
    fM.ui.sName = ui.sName
    fM.UpdateUserInfo
    OnlineUsers.Add fM

    Set AddNewUserByUI = fM
End Function

Public Function FindUserByIP(ByRef sip As String, ByRef fRet As frmMsg) As Boolean
    '试图找到与该用户相联系的类窗口：
    Dim fMsg As frmMsg
    For Each fMsg In OnlineUsers
        If fMsg.ui.sHostIP = sip Then
            Set fRet = fMsg
            FindUserByIP = True
            Exit Function
        End If
    Next fMsg
End Function

Public Function FindUserByUI(ByRef ui As USER_INFO, ByRef fRet As frmMsg) As Boolean
    '试图找到与该用户相联系的类窗口：
    Dim fMsg As frmMsg
    For Each fMsg In OnlineUsers
        If fMsg.ui.sHostIP = ui.sHostIP Then
            Set fRet = fMsg
            FindUserByUI = True
            Exit Function
        End If
    Next fMsg
End Function

Public Sub OnIconRightButtonUp()
    If Not bDisableIconMenu Then
        Me.SetFocus
        PopupMenu mnuIcon, vbLeftButton Or vbRightButton
    End If
End Sub

Public Sub OnIconLeftButtonUp()
    If msgForms.Count > 0 Then
        Dim f As Form
        Set f = msgForms(1)
        f.Show vbModeless
        f.WindowState = 0
        msgForms.Remove 1
        '如果没有窗口了:
        If msgForms.Count = 0 Then Call FlashIcon(False)
    End If
    If Not bDisableIconMenu Then
        Me.SetFocus
    End If
End Sub

Private Sub FlashIcon(ByVal b As Boolean)
    Static bFlashing As Boolean
    If bFlashing = b Then Exit Sub '状态相同

    bFlashing = b
    If Not bFlashing Then
        '停止闪烁:
        tmrFlashIcon.Enabled = False
        Set SysTrayIcon.Icon = LoadResPicture(101, vbResIcon)
        SysTrayIcon.Tip = TIP_NORMAL
        SysTrayIcon.ModifySysTrayIcon
    Else
        SysTrayIcon.Tip = TIP_MESSAGE
        SysTrayIcon.ModifySysTrayIcon
        tmrFlashIcon.Enabled = True
    End If
End Sub

Private Sub tmrFlashIcon_Timer()
    Static b As Boolean
    If b Then
        Set SysTrayIcon.Icon = LoadResPicture(102, vbResIcon)
        SysTrayIcon.ModifySysTrayIcon
    Else
        Set SysTrayIcon.Icon = LoadResPicture(103, vbResIcon)
        SysTrayIcon.ModifySysTrayIcon
    End If
    b = Not b
End Sub
