VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFileReceive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "接收文件"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5280
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFileReceive.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5280
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtSaveFile 
      Height          =   270
      Left            =   1440
      TabIndex        =   16
      Top             =   1800
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Frame frameAsk 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   5175
      Begin VB.TextBox txtFileSize 
         BackColor       =   &H8000000F&
         Height          =   270
         Left            =   4200
         TabIndex        =   15
         Text            =   "XXX KB"
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdRefuse 
         Caption         =   "拒绝(&R)"
         Height          =   375
         Left            =   3240
         TabIndex        =   14
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdReceive 
         Caption         =   "接收(&G)"
         Height          =   375
         Left            =   1800
         TabIndex        =   13
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtUserInfo 
         BackColor       =   &H8000000F&
         Height          =   270
         Left            =   2520
         TabIndex        =   12
         Text            =   "userName/192.168.0.99"
         Top             =   120
         Width           =   2655
      End
      Begin VB.TextBox txtAbstract 
         BackColor       =   &H8000000F&
         Height          =   270
         Left            =   960
         TabIndex        =   9
         Text            =   "%FileDescription%"
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtFileName 
         BackColor       =   &H8000000F&
         Height          =   270
         Left            =   960
         TabIndex        =   7
         Text            =   "%FileName%"
         Top             =   480
         Width           =   3135
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "文件名称"
         Height          =   180
         Left            =   120
         TabIndex        =   11
         Top             =   520
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "是否接收？"
         Height          =   180
         Left            =   120
         TabIndex        =   10
         Top             =   1240
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "文件说明"
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   880
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "某位用户试图向您传送文件："
         Height          =   180
         Left            =   120
         TabIndex        =   6
         Top             =   160
         Width           =   2340
      End
   End
   Begin VB.Frame frameReceiving 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消(&C)"
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin MSComctlLib.ProgressBar prgBar 
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "传送进度"
         Height          =   180
         Left            =   120
         TabIndex        =   2
         Top             =   640
         Width           =   720
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "正在接收文件，请等待..."
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   160
         Width           =   2070
      End
   End
   Begin MSWinsockLib.Winsock wsTcpReceive 
      Left            =   480
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "frmFileReceive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sRemoteIP As String
Dim nRemotePort As Long
Dim sSaveFile As String
Dim nFileSize As Long
Dim nReceived As Long

Dim bComplete As Boolean

Private Sub CloseAndHide()
    wsTcpReceive.Close
    If FreeFile() <> 1 Then Close #1
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Call CloseAndHide
End Sub

Private Sub cmdReceive_Click()
    '选择保存位置：
    Dim dlg As CommonDlg
    Set dlg = New CommonDlg
    sSaveFile = dlg.ShowSave(Me.hwnd, "", "All Files(*.*)" + Chr(0) + "*.*", sSaveFile)
    txtSaveFile.Text = sSaveFile
    sSaveFile = txtSaveFile.Text
    If sSaveFile = "" Then Exit Sub

    With wsTcpReceive
        Err.Clear
        On Error Resume Next
        .Connect sRemoteIP, nRemotePort
        If Err.Number Then
            MsgBox "连接失败！原因：" + Err.Description, vbCritical Or vbOKOnly
            Call CloseAndHide
        End If

        '连接成功：
        frameReceiving.Visible = True
        frameAsk.Visible = False
    End With
End Sub

Private Sub cmdRefuse_Click()
    'tell the sender?
    '???

    Call CloseAndHide
End Sub

Public Sub QueryReceiveFile(ByVal sPacket As String) '询问是否接收文件
    Dim s() As String, n As Long

    nReceived = 0
    s = Split(sPacket, SPLIT_CHAR, 5)
    If UBound(s) <> 4 Then Exit Sub

    txtUserInfo.Text = s(0) 's(0) = userInfo
    'try to get the IP:
    n = 1
    While n > 0
        n = InStr(1, s(0), "/")
        If n > 0 Then
            s(0) = Right(s(0), Len(s(0)) - n)
        End If
    Wend
    sRemoteIP = s(0)
    txtFileName.Text = s(1)    's(1) = fileName
    sSaveFile = s(1)

    nFileSize = Int(Val(s(2))) 's(2) = fileSize
    If nFileSize >= 1048576 Then
        txtFileSize.Text = Int(nFileSize / (1024 * 10.24)) / 100 & " MB"
    ElseIf nFileSize >= 1024 Then
        txtFileSize.Text = Int(nFileSize / 10.24) / 100 & " KB"
    Else
        txtFileSize.Text = nFileSize & " B"
    End If
    nRemotePort = Int(Val(s(3))) 's(3) = remotePort
    txtAbstract.Text = s(4)      's(4) = fileDescription

    bComplete = False
    frameAsk.Visible = True
    frameReceiving.Visible = False

    If gbAutoPopup Then
        Me.Show vbModeless
    Else
        fMain.AddMsgForm Me
    End If
End Sub

Private Sub Form_Load()
    '使窗口总在最前：
    If gbAlwaysOnTop Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Private Sub wsTcpReceive_Close()
    Close #1
    If bComplete Then
        prgBar.Value = 100
        MsgBox "文件传送完毕。", vbInformation Or vbOKOnly
    End If
    Call CloseAndHide
End Sub

Private Sub wsTcpReceive_Connect()
    '已连接上：
    Err.Clear
    On Error Resume Next
    Open sSaveFile For Binary Access Write Lock Write As #1
    If Err.Number Then
        wsTcpReceive.Close
        MsgBox "错误：无法保存文件！", vbCritical Or vbOKOnly
        Call CloseAndHide
        Exit Sub
    End If
    nReceived = 0
End Sub

Private Sub wsTcpReceive_DataArrival(ByVal bytesTotal As Long)
    Dim by() As Byte

    nReceived = nReceived + bytesTotal
    If nReceived = nFileSize Then bComplete = True

    prgBar.Value = nReceived * 100 \ nFileSize

    wsTcpReceive.GetData by, vbByte, bytesTotal
    Put #1, , by
End Sub

Private Sub wsTcpReceive_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "传送错误：" + Description, vbCritical Or vbOKOnly
    Close #1
    Call CloseAndHide
End Sub

