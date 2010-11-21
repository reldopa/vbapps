VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "发送文件"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileTransfer.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAbstract 
      Height          =   270
      Left            =   960
      MaxLength       =   100
      TabIndex        =   0
      Top             =   840
      Width           =   4935
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton cmdTransfer 
      Caption         =   "开始传送(&T)"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar prgBar 
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   1260
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSWinsockLib.Winsock wsTcpSend 
      Left            =   240
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   12090
      LocalPort       =   12090
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      TabIndex        =   1
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox txtFile 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   480
      Width           =   4575
   End
   Begin VB.TextBox txtIP 
      BackColor       =   &H8000000F&
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "userName/192.168.0.9"
      Top             =   120
      Width           =   4935
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "准备开始传送文件。"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1620
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "文件说明"
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   880
      Width           =   720
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "传送进度"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   1300
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "发送文件"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   520
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "发送目标"
      Height          =   180
      Left            =   120
      TabIndex        =   6
      Top             =   160
      Width           =   720
   End
End
Attribute VB_Name = "frmFileTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nFileSize As Long
Dim sRemoteIP As String
Public bComplete As Boolean

Private Sub CloseAndHide()
    wsTcpSend.Close
    If FreeFile() <> 1 Then Close #1
    Me.Hide
End Sub

Private Sub cmdBrowse_Click()
    Dim dlg As CommonDlg
    Set dlg = New CommonDlg

    txtFile.Text = dlg.ShowOpen(Me.hwnd, "", "所有文件 (*.*)" + Chr(0) + "*.*")
End Sub

Private Sub cmdCancel_Click() '取消并关闭窗口：
    Call CloseAndHide
End Sub

Private Sub cmdClose_Click() '关闭窗口：
    Call CloseAndHide
End Sub

'准备传送：
Public Sub TransferPrepare(ByVal sIP As String, ByVal sUserName As String)
    bComplete = False

    sRemoteIP = sIP
    txtIP.Text = sUserName + "/" + sIP
    wsTcpSend.RemoteHost = sIP
    txtFile.Text = ""
    txtAbstract.Text = ""
    lblInfo.Caption = "准备开始传送文件。"
    cmdTransfer.Enabled = False
    cmdCancel.Enabled = True
    cmdBrowse.Enabled = True
    txtAbstract.Enabled = True
    prgBar.Value = 0
End Sub

'开始传送:
Private Sub cmdTransfer_Click()
    If FileLen(txtFile.Text) > 10485760 Then
        MsgBox "文件太大(超过10MB)，无法传送！", vbInformation Or vbOKOnly
        Exit Sub
    End If

    '开始本地侦听:
    With wsTcpSend
        Err.Clear
        On Error Resume Next
        .Listen
        If Err.Number Then
            '侦听失败:
            MsgBox "无法绑定本地端口！请重新选择端口号。", vbCritical Or vbOKOnly, "建立本地连接失败"
            lblInfo.Caption = "无法绑定本地端口。"
            Exit Sub
        End If
    End With
    lblInfo.Caption = "正在等待对方响应..."
    'try to get the file name(without path):
    Dim n As Long, s As String
    n = 1
    s = txtFile.Text
    While n > 0
        n = InStr(1, s, "\")
        If n > 0 Then s = Right(s, Len(s) - n)
    Wend
    fMain.SendFileRequest sRemoteIP, PackFileInfo(s, FileLen(txtFile.Text), gnFilePort, txtAbstract.Text)
    cmdTransfer.Enabled = False
    cmdBrowse.Enabled = False
    txtAbstract.Enabled = False
End Sub

Private Sub Form_Load()
    '使窗口总在最前：
    If gbAlwaysOnTop Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        Cancel = True
        Call CloseAndHide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    wsTcpSend.Close
End Sub

Private Sub txtFile_Change()
    If Trim(txtFile) <> "" Then
        cmdTransfer.Enabled = True
    Else
        cmdTransfer.Enabled = False
    End If
End Sub

Private Sub wsTcpSend_Close()
    Call CloseAndHide
End Sub

Private Sub wsTcpSend_ConnectionRequest(ByVal requestID As Long)
    Dim s() As Byte, nSize As Long

    If wsTcpSend.RemoteHostIP = sRemoteIP Then
        '读文件：
        Err.Clear
        On Error Resume Next
        Open txtFile.Text For Binary Access Read Lock Write As #1
        If Err.Number Then '无法打开文件
            Call CloseAndHide
            MsgBox "错误：无法读取文件！", vbCritical Or vbOKOnly
            Exit Sub
        End If
        nSize = LOF(1)
        If nSize > 10485760 Then 'File>10M
            Close #1
            Call CloseAndHide
            MsgBox "文件太大(超过10MB)，无法传送！", vbInformation Or vbOKOnly
            Exit Sub
        End If
        If nSize = 0 Then
            Close #1
            Call CloseAndHide
            MsgBox "文件长度为0，请重新选择文件！", vbInformation Or vbOKOnly
            Exit Sub
        End If
        ReDim s(nSize - 1)
        Get #1, , s
        Close #1
        wsTcpSend.Close
        wsTcpSend.Accept requestID
        lblInfo.Caption = "正在传送文件..."
        '已连接上，发送数据：
        wsTcpSend.SendData s
    End If
End Sub

Private Sub wsTcpSend_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "传输错误：" + Description, vbCritical Or vbOKOnly
    Call CloseAndHide
End Sub

Private Sub wsTcpSend_SendComplete()
    '传送完毕：
    bComplete = True
    wsTcpSend.Close

    prgBar.Value = 100 '填充进度条
    lblInfo.Caption = "文件传送完毕。"
    MsgBox "文件传送完毕。", vbInformation Or vbOKOnly
    Call CloseAndHide
End Sub

Private Sub wsTcpSend_SendProgress(ByVal bytesSent As Long, ByVal bytesRemaining As Long)
    Dim n As Long
    n = bytesSent * 100 \ (bytesSent + bytesRemaining)
    prgBar.Value = n
End Sub
