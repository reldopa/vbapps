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
      Name            =   "����"
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
         Name            =   "����"
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
         Caption         =   "������Ϣ(&S)..."
      End
      Begin VB.Menu mnuBar8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChat 
         Caption         =   "��������(&C)..."
      End
      Begin VB.Menu mnuTransferFile 
         Caption         =   "�����ļ�(&T)..."
      End
      Begin VB.Menu mnuBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBoradcastMsg 
         Caption         =   "���͹㲥(&B)..."
      End
      Begin VB.Menu mnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewDetail 
         Caption         =   "��ϸ����(&D)..."
      End
   End
   Begin VB.Menu mnuIcon 
      Caption         =   "icon_menu"
      Visible         =   0   'False
      Begin VB.Menu mnuIconBoradcastMsg 
         Caption         =   "���͹㲥(&B)..."
      End
      Begin VB.Menu mnuBar11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconSetting 
         Caption         =   "��������(&S)..."
      End
      Begin VB.Menu mnuBar12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconAbout 
         Caption         =   "����(&A)..."
      End
      Begin VB.Menu mnuBar14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIconExit 
         Caption         =   "�˳�(&X)"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const TIP_NORMAL As String = "QQ Messenger - ��������"
Private Const TIP_MESSAGE As String = "QQ Messenger - ������Ϣ����"

Dim bDisableIconMenu As Boolean

Dim dblX As Single, dblY As Single  '��¼˫��ʱ�����λ��
Dim fBroadcast As frmBroadcastMsg   '�㲥����
Public fTransfer As frmFileTransfer '�����ļ�����
Public fReceive As frmFileReceive   '�����ļ�����
Dim OnlineUsers As Collection       '��¼���е������û�
Dim menuSelectItem As ListItem      '����Ҽ�ѡ�����

Dim SysTrayIcon As SystemTrayIcon

Dim msgForms As Collection

'���Form����Ϣ����: //////////////////////////////////////////////////////////////////////////////
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
    '��ʾ��Ϣ��
    Call OnIconLeftButtonUp
End Sub

Private Sub Form_Load()
    Dim dwHeight As Single
    dwHeight = GetSetting(App.Title, "Setting", "WinHeight", 2700)
    Me.Height = dwHeight

    Call LoadSysProp '��ϵͳ����

    Set msgForms = New Collection

    With wsUDPSocket
        '�󶨱��� UDP ����˿ڣ�
        .LocalPort = gnPort
        .RemotePort = gnPort
        Err.Clear
        On Error Resume Next
        .Bind
        If Err.Number Then
            MsgBox "��ʼ������ʧ��!" + vbCrLf + "����: " + Err.Description, vbCritical Or vbOKOnly, "��ʼ������ʧ��"
            Unload Me
            Exit Sub
        End If

        '�����������룺
        If Not GetSubNetMask(.localIp, subnetMask) Then
            Unload Me
            Exit Sub
        End If

        '�����ļ����մ��壺
        Set fReceive = New frmFileReceive
        Load fReceive
        '�����ļ����ʹ��壺
        Set fTransfer = New frmFileTransfer
        Load fTransfer
        '���ñ���TCP�˿ڣ�
        fTransfer.wsTcpSend.LocalPort = gnFilePort
        fReceive.wsTcpReceive.LocalPort = gnFilePort

        '�����ҵ���Ϣ��
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
    '����֪ͨ��
    Call BroadcastIAmOnline

    'ʹ����������ǰ��
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
            MsgBox "�ȼ�ע��ʧ�ܣ������������ȼ���", vbCritical Or vbOKOnly
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
        '����֪ͨ��
        Call BroadcastIAmOffline
        Unload fBroadcast

        '�ر��������ӣ�
        wsUDPSocket.Close
        Call SaveMyInfo
        Set myInfo = Nothing
    End If

    '�ر����д��ڣ�
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
        MsgBox "�޷�ʹ���������졣ԭ��Ӧ�ó���vdpchat.exe�������ڡ�", vbCritical Or vbOKOnly
        Exit Sub
    End If

    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        '��������
        .RemoteHost = menuSelectItem.Tag
        .SendData HEADER_CHATTING & PackChatRequest(menuSelectItem.Tag, menuSelectItem.Text)
    End With
    If Err.Number = 0 Then
        '��������ɹ���׼�������������죺
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
        '��һ������δ��ɣ��޷����е�ǰ�����˳���
        Exit Sub
    End If

    If fReceive.Visible Then
        Exit Sub
    End If

    '����ǰ׼����
    Call fTransfer.TransferPrepare(menuSelectItem.Tag, menuSelectItem.Text)
    fTransfer.Show vbModeless
    fTransfer.WindowState = 0
End Sub

Private Sub mnuViewDetail_Click()
    If Not IsItemExist(menuSelectItem) Then Exit Sub

    Call QueryUserDetail(menuSelectItem.Tag)
End Sub

Private Sub lvGroup_DblClick() '˫��ʱ��������Ի���
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

Private Sub mnuSendMsg_Click() '��˵���������Ϣ��
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

'�յ����ݰ�:
Private Sub wsUDPSocket_DataArrival(ByVal bytesTotal As Long)
    Err.Clear
    On Error Resume Next
    Dim s As String, Msg As String
    wsUDPSocket.GetData s
    If Len(s) < 3 Then Exit Sub '��Ч���ݰ�

    Msg = Right(s, Len(s) - 3)
    Select Case Int(Val(Left(s, 3)))
        Case HEADER_USER_SEND_MESSAGE:    '�յ��û���Ϣ:
            Call OnUserMessageArrival(Msg)
        Case HEADER_USER_ONLINE:          'ĳ���û�����:
            Call OnUserOnline(Msg)
        Case HEADER_USER_OFFLINE:         'ĳ���û�����:
            Call OnUserOffline(Msg)
        Case HEADER_USER_BROADCAST:       '�յ��û��㲥:
            Call OnUserBroadcasting(Msg)
        Case HEADER_SYS_BROADCAST:

        Case HEADER_UPDATE_USER_INFO:     '�յ�ĳ���û��ĸ�����Ϣ:
            Call OnUpdateUserInfo(Msg)
        Case HEADER_SEND_MY_DETAIL:       '�յ�ĳ���û�����Ϣ
            Call UnpackUserDetailInfo(Msg)
        Case HEADER_QUERY_INFO:           'ĳ���û�Ҫ�󷵻ظ�����Ϣ
            Call OnQueryMyDetailInfo(Msg)
        Case HEADER_WANT_TO_TRANSFER:     'ĳ���û�Ҫ�����ļ�
            Call OnFileRequest(Msg)
        Case HEADER_CHATTING:             'ĳ���û�Ҫ����������
            Call UnpackChatRequest(Msg)
    End Select
End Sub

'**************************************************************************************************
'** ������Ϣ ** ������Ϣ ** ������Ϣ ** ������Ϣ ** ������Ϣ ** ������Ϣ ** ������Ϣ ** ������Ϣ **
'**************************************************************************************************

'�������û�������Ϣ:
Public Sub SendMessageToUser(ByVal ipAddr As String, ByVal sMsg As String)
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = ipAddr
        .SendData HEADER_USER_SEND_MESSAGE & PackUserInfo(myInfo) + sMsg
    End With
End Sub

'��ѯ�û���ϸ��Ϣ��
Public Sub QueryUserDetail(ByVal ipAddr As String)
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = ipAddr
        .SendData HEADER_QUERY_INFO & myInfo.sHostIP
    End With
End Sub

'�������û��㲥��Ϣ:
Public Sub SendMessageToAll(ByVal sMsg As String)
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = subnetMask
        .SendData HEADER_USER_BROADCAST & PackUserInfo(myInfo) + sMsg
    End With
End Sub

'���͡����ļ�������
Public Sub SendFileRequest(ByVal ipAddr As String, ByVal sFileInfo As String)
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = ipAddr
        .SendData HEADER_WANT_TO_TRANSFER & sFileInfo
    End With
End Sub

'֪ͨ�����û����������ˡ�:
Public Sub BroadcastIAmOnline()
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = subnetMask
        .SendData HEADER_USER_ONLINE & PackUserInfo(myInfo)
    End With
End Sub

'֪ͨ�����û����������ˡ�:
Public Sub BroadcastIAmOffline()
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = subnetMask
        .SendData HEADER_USER_OFFLINE & PackUserInfo(myInfo)
    End With
End Sub

'�������û������ҵ���Ϣ��
Public Sub UpdateMyInfo()
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = subnetMask
        .SendData HEADER_UPDATE_USER_INFO & PackUserInfo(myInfo)
    End With
End Sub

'ˢ�������û���
Public Sub RefreshAllUser()
    'delete all user:
    'While lvGroup.ListItems.Count > 0
    '    lvGroup.ListItems.Remove 1
    'Wend
    'then re-add them:
    'Call UpdateMyInfo
End Sub

'**************************************************************************************************
'** ��Ϣ���� ** ��Ϣ���� ** ��Ϣ���� ** ��Ϣ���� ** ��Ϣ���� ** ��Ϣ���� ** ��Ϣ���� ** ��Ϣ���� **
'**************************************************************************************************

'Ӧ�������ҵ���ϸ���ϣ�
Public Sub OnQueryMyDetailInfo(ByVal ipAddr As String)
    Err.Clear
    On Error Resume Next
    With wsUDPSocket
        .RemoteHost = ipAddr
        .SendData HEADER_SEND_MY_DETAIL & PackMyDetailInfo
    End With
End Sub

'�յ������û���Ϣ��
Private Sub OnUserMessageArrival(ByVal s As String)
    If Len(s) < FRAME_HEADER_SIZE Then Exit Sub

    Dim sFrame As String, sMsg As String
    sFrame = Left(s, FRAME_HEADER_SIZE)
    sMsg = Right(s, Len(s) - FRAME_HEADER_SIZE)

    Dim ui As New USER_INFO
    '��ͼ����û���Ϣ��
    If UnpackUserInfo(sFrame, ui) Then
        If ui.sHostIP <> myInfo.sHostIP Then
            '��ͼ�ҵ�����û�����ϵ���ര�ڣ�
            Dim fMsg As frmMsg
            If Not FindUserByUI(ui, fMsg) Then
                Set fMsg = AddNewUserByUI(ui)
                Call AddUserIcon(ui)
            End If
            Set ui = Nothing
            '���� fMsg ��ʾ��Ϣ:
            Call PlayMsgSound
            Call fMsg.OnMessageArrival(sMsg)
        End If
    End If
End Sub

'�յ��û������ļ�����
Private Sub OnFileRequest(ByVal sPacket As String)
    If fTransfer.Visible Then Exit Sub '���ڴ��ͣ�����޷������ļ�
    If fReceive.Visible Then Exit Sub  '���ڽ��գ�����޷������ļ�
    If IsReceiveFormExistInMsgForm Then Exit Sub '���������û�δ��Ӧ������޷������ļ�
    If IsTransferFormExistInMsgForm Then Exit Sub  '���������û�δ��Ӧ������޷������ļ�

    Call PlayFileSound
    Call fReceive.QueryReceiveFile(sPacket)
End Sub

'�յ��û��㲥��
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

'�յ��û�����֪ͨ��
Private Sub OnUserOnline(ByVal sFrame As String)
    If Len(sFrame) <> FRAME_HEADER_SIZE Then Exit Sub

    Dim ui As New USER_INFO
    Dim fMsg As frmMsg

    If UnpackUserInfo(sFrame, ui) Then
        If ui.sHostIP <> myInfo.sHostIP Then '����Ϣ��������ҷ��ģ�
            Call PlayUserSound
            If FindUserByUI(ui, fMsg) Then
                fMsg.ui.sHostName = ui.sHostName
                fMsg.ui.sName = ui.sName
            Else
                Call AddNewUserByUI(ui) '�����һ�����û�
            End If
            '����û�ͷ��
            Call AddUserIcon(ui)
            '���� UPDATE_MY_INFO ʹ�����û������:
            Call UpdateMyInfo

        End If
    End If
    Set ui = Nothing
End Sub

'�յ��û�����֪ͨ
Private Sub OnUserOffline(ByVal sFrame As String)
    If Len(sFrame) <> FRAME_HEADER_SIZE Then Exit Sub

    Dim ui As New USER_INFO
    Dim fMsg As frmMsg

    If UnpackUserInfo(sFrame, ui) Then
        If ui.sHostIP <> myInfo.sHostIP Then '����Ϣ��������ҷ��ģ�
            Call PlayUserSound
            'ɾ�����û�ͷ�񣬵��������û�̸���Ի���
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

'�յ��û�������Ϣ
Private Sub OnUpdateUserInfo(ByVal sFrame As String)
    If Len(sFrame) <> FRAME_HEADER_SIZE Then Exit Sub

    Dim ui As New USER_INFO
    Dim fMsg As frmMsg

    If UnpackUserInfo(sFrame, ui) Then
        If ui.sHostIP <> myInfo.sHostIP Then '����Ϣ��������ҷ��ģ�
            If FindUserByUI(ui, fMsg) Then
                fMsg.ui.sHostName = ui.sHostName
                fMsg.ui.sName = ui.sName
                fMsg.ui.nFace = ui.nFace
                fMsg.UpdateUserInfo

                '����ͼ��:
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
                Call AddNewUserByUI(ui) '�����һ�����û�
                Call AddUserIcon(ui)
                '���� UPDATE_MY_INFO ʹ�����û������:
                Call UpdateMyInfo
            End If
        End If
    End If
    Set ui = Nothing
End Sub

'==============================================================================

Public Sub AddUserIcon(ByRef ui As USER_INFO)
    '����û�ͷ��
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
    '������û������ϣ��� IP ��ַ��Ϊ��ʶ��:
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
    '��ͼ�ҵ�����û�����ϵ���ര�ڣ�
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
    '��ͼ�ҵ�����û�����ϵ���ര�ڣ�
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
        '���û�д�����:
        If msgForms.Count = 0 Then Call FlashIcon(False)
    End If
    If Not bDisableIconMenu Then
        Me.SetFocus
    End If
End Sub

Private Sub FlashIcon(ByVal b As Boolean)
    Static bFlashing As Boolean
    If bFlashing = b Then Exit Sub '״̬��ͬ

    bFlashing = b
    If Not bFlashing Then
        'ֹͣ��˸:
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
