VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSettings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   5775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5775
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList imglstFace 
      Left            =   240
      Top             =   5040
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
            Picture         =   "frmSettings.frx":0CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":0F56
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":11EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":1452
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":16CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":194E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":1BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":1E4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":20E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":235E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":25CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":285E
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2AEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":2D7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3012
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":32A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3536
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":37CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3A4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":3F2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":41A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":440A
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":4696
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":4926
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":4BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":4E26
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":50A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":532E
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":55BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":5832
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":5AA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":5D16
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":5F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":6202
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":6492
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":6706
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":699A
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":6C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSettings.frx":6EAA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "系统设置"
      Height          =   1815
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   5415
      Begin VB.CheckBox chkNoVoiceChat 
         Caption         =   "拒绝任何用户的语音聊天请求"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   1320
         Width           =   2655
      End
      Begin VB.CheckBox chkUseHotkey 
         Caption         =   "使用系统热键"
         Height          =   255
         Left            =   240
         TabIndex        =   38
         Top             =   960
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.ComboBox cmbFilePort 
         Height          =   300
         ItemData        =   "frmSettings.frx":713E
         Left            =   3720
         List            =   "frmSettings.frx":7145
         TabIndex        =   37
         Text            =   "12090"
         Top             =   600
         Width           =   1575
      End
      Begin VB.ComboBox cmbPort 
         Height          =   300
         ItemData        =   "frmSettings.frx":7151
         Left            =   3720
         List            =   "frmSettings.frx":7158
         TabIndex        =   35
         Text            =   "12022"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdModifySystem 
         Caption         =   "修改"
         Height          =   375
         Left            =   3960
         TabIndex        =   22
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtHotKey 
         Height          =   270
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "Ctrl+Alt+X"
         Top             =   960
         Width           =   1575
      End
      Begin VB.CheckBox chkAutoPopup 
         Caption         =   "自动弹出消息"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkAlwaysOnTop 
         Caption         =   "窗口总在最前"
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   240
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "文件传送端口"
         Height          =   180
         Left            =   2520
         TabIndex        =   36
         Top             =   640
         Width           =   1080
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "聊天广播端口"
         Height          =   180
         Left            =   2520
         TabIndex        =   29
         Top             =   280
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "窗口弹出热键"
         Height          =   180
         Left            =   2520
         TabIndex        =   20
         Top             =   1000
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭"
      Height          =   375
      Left            =   4080
      TabIndex        =   4
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "我的信息"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox cmbBlood 
         Height          =   300
         ItemData        =   "frmSettings.frx":7164
         Left            =   4440
         List            =   "frmSettings.frx":717D
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtAge 
         Height          =   270
         Left            =   4440
         MaxLength       =   5
         TabIndex        =   32
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtHomepage 
         Height          =   270
         Left            =   960
         MaxLength       =   50
         TabIndex        =   31
         Top             =   1320
         Width           =   2895
      End
      Begin VB.TextBox txtAbstract 
         Height          =   615
         Left            =   960
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Top             =   2400
         Width           =   2895
      End
      Begin VB.ComboBox cmbSex 
         Height          =   300
         ItemData        =   "frmSettings.frx":7196
         Left            =   4440
         List            =   "frmSettings.frx":71A3
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox cmbConstellation 
         Height          =   300
         ItemData        =   "frmSettings.frx":71B1
         Left            =   4440
         List            =   "frmSettings.frx":71DC
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cmbSuxiang 
         Height          =   300
         ItemData        =   "frmSettings.frx":722A
         Left            =   4440
         List            =   "frmSettings.frx":7255
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtMobile 
         Height          =   270
         Left            =   960
         MaxLength       =   20
         TabIndex        =   14
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox txtTel 
         Height          =   270
         Left            =   960
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1680
         Width           =   2895
      End
      Begin VB.TextBox txtEmail 
         Height          =   270
         Left            =   960
         MaxLength       =   30
         TabIndex        =   12
         Top             =   960
         Width           =   2895
      End
      Begin VB.TextBox txtTrueName 
         Height          =   270
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   11
         Top             =   600
         Width           =   2295
      End
      Begin VB.CommandButton cmdModifyUser 
         Caption         =   "修改"
         Height          =   375
         Left            =   3960
         TabIndex        =   5
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtName 
         Height          =   270
         Left            =   1560
         MaxLength       =   20
         TabIndex        =   3
         Top             =   240
         Width           =   2295
      End
      Begin MSComctlLib.ImageCombo imgcmbFace 
         Height          =   570
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   1005
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         ImageList       =   "imglstFace"
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "血型"
         Height          =   180
         Left            =   3960
         TabIndex        =   33
         Top             =   1000
         Width           =   360
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "个人主页"
         Height          =   180
         Left            =   120
         TabIndex        =   30
         Top             =   1360
         Width           =   720
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "个人简介"
         Height          =   180
         Left            =   120
         TabIndex        =   27
         Top             =   2440
         Width           =   720
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "性别"
         Height          =   180
         Left            =   3960
         TabIndex        =   25
         Top             =   640
         Width           =   360
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "星座"
         Height          =   180
         Left            =   3960
         TabIndex        =   23
         Top             =   1360
         Width           =   360
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "属相"
         Height          =   180
         Left            =   3960
         TabIndex        =   15
         Top             =   1720
         Width           =   360
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "移动电话"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   2080
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "固定电话"
         Height          =   180
         Left            =   120
         TabIndex        =   9
         Top             =   1720
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "电子邮件"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   1000
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "年龄"
         Height          =   180
         Left            =   3960
         TabIndex        =   7
         Top             =   280
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "真名"
         Height          =   180
         Left            =   1080
         TabIndex        =   6
         Top             =   640
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "昵称"
         Height          =   180
         Left            =   1080
         TabIndex        =   2
         Top             =   280
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim tmpKey As Integer
Dim tmpShift As Integer

Private Sub chkUseHotkey_Click()
    txtHotKey.Enabled = IIf(chkUseHotkey.Value = 1, True, False)
End Sub

Private Sub cmbFilePort_Click()
    cmbFilePort.Text = "12090"
End Sub

Private Sub cmbPort_Click()
    cmbPort.Text = "12022"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdModifySystem_Click()
    MsgBox "新的设置将在下次启动时生效！", vbInformation Or vbOKOnly

    SaveSetting App.Title, "Setting", "UseHotKey", IIf(chkUseHotkey.Value = 1, "True", "False")
    SaveSetting App.Title, "Setting", "HotKey", tmpKey
    SaveSetting App.Title, "Setting", "Shift", tmpShift

    SaveSetting App.Title, "Setting", "Popup", IIf(chkAutoPopup.Value = 1, "True", "False")
    SaveSetting App.Title, "Setting", "OnTop", IIf(chkAlwaysOnTop.Value = 1, "True", "False")
    SaveSetting App.Title, "Setting", "Port", Int(Val(cmbPort.Text))
    SaveSetting App.Title, "Setting", "FilePort", Int(Val(cmbFilePort.Text))

    SaveSetting App.Title, "Setting", "NoVoiceChat", IIf(chkNoVoiceChat.Value = 1, "True", "False")
End Sub

Private Sub cmdModifyUser_Click()
    myInfo.nFace = imgcmbFace.SelectedItem.Index
    myInfo.sName = Trim(txtName.Text)

    gsTrueName = txtTrueName
    gsAge = txtAge
    gsEmail = txtEmail
    gsHomepage = txtHomepage
    gsTel = txtTel
    gsMobile = txtMobile
    gsAbstract = txtAbstract

    On Error Resume Next
    gsSex = cmbSex.Text
    gsBlood = cmbBlood.Text
    gsSuxiang = cmbSuxiang.Text
    gsConstellation = cmbConstellation.Text

    'send REFRESH message:
    Call fMain.UpdateMyInfo
    MsgBox "您已经成功地更改了您的个人信息！", vbInformation Or vbOKOnly
End Sub

Private Sub Form_Load()
    '读用户设置 *********************************
    Dim i As Long
    With imgcmbFace.ComboItems
        For i = 1 To 40
            .Add , , , i
        Next i
    End With
    imgcmbFace.ComboItems(myInfo.nFace).Selected = True
    txtName.Text = myInfo.sName

    txtTrueName = gsTrueName
    txtAge = gsAge
    txtEmail = gsEmail
    txtHomepage = gsHomepage
    txtTel = gsTel
    txtMobile = gsMobile
    txtAbstract = gsAbstract

    On Error Resume Next
    cmbSex.Text = gsSex
    cmbBlood.Text = gsBlood
    cmbSuxiang.Text = gsSuxiang
    cmbConstellation.Text = gsConstellation

    '读系统设置 *********************************
    Dim nPort As Long, nFilePort As Long
    nPort = GetSetting(App.Title, "Setting", "Port", 12022)
    If nPort < 0 Or nPort > 65535 Then nPort = 12022

    nFilePort = GetSetting(App.Title, "Setting", "FilePort", 12090)
    If nFilePort < 0 Or nFilePort > 65535 Then nFilePort = 12090

    tmpKey = GetSetting(App.Title, "Setting", "HotKey", vbKeyX)
    If tmpKey < vbKeyA Or tmpKey > vbKeyZ Then tmpKey = vbKeyX

    tmpShift = GetSetting(App.Title, "Setting", "Shift", 6)
    If tmpShift < 0 Or tmpShift > 7 Then tmpShift = 6

    chkAlwaysOnTop.Value = IIf(GetSetting(App.Title, "Setting", "OnTop", "True") = "True", 1, 0)
    chkAutoPopup.Value = IIf(GetSetting(App.Title, "Setting", "Popup", "False") = "True", 1, 0)
    chkUseHotkey.Value = IIf(GetSetting(App.Title, "Setting", "UseHotKey", "True") = "True", 1, 0)

    chkNoVoiceChat.Value = IIf(GetSetting(App.Title, "Setting", "NoVoiceChat", "False") = "False", 0, 1)

    Call txtHotKey_KeyUp(tmpKey, tmpShift)
    cmbPort.Text = nPort
    cmbFilePort.Text = nFilePort

    '使窗口总在最前：
    If gbAlwaysOnTop Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Private Sub txtHotKey_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim sShift As String
    If KeyCode >= vbKeyA And KeyCode <= vbKeyZ Then
        tmpKey = KeyCode
        tmpShift = Shift
        Select Case Shift
        Case 0: sShift = ""
        Case 1: sShift = "Shift+"
        Case 2: sShift = "Ctrl+"
        Case 3: sShift = "Shift+Ctrl+"
        Case 4: sShift = "Alt+"
        Case 5: sShift = "Shift+Alt+"
        Case 6: sShift = "Ctrl+Alt+"
        Case 7: sShift = "Shift+Ctrl+Alt+"
        End Select
        txtHotKey.Text = sShift + Chr(KeyCode)
    End If
End Sub
