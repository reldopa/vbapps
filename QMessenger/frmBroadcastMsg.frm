VERSION 5.00
Begin VB.Form frmBroadcastMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "用户广播消息"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBroadcastMsg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   6720
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭(&C)"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "发送(&S)"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtMsg 
      Height          =   735
      HideSelection   =   0   'False
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2160
      Width           =   6495
   End
   Begin VB.TextBox txtLog 
      BackColor       =   &H8000000F&
      Height          =   1455
      HideSelection   =   0   'False
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   360
      Width           =   6495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "发送广播(最多100个字符，按 Ctrl+Enter 快速发送)："
      Height          =   180
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   4410
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "广播记录："
      Height          =   180
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   900
   End
End
Attribute VB_Name = "frmBroadcastMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    txtMsg.Text = ""
    Me.Hide
End Sub

Private Sub cmdSend_Click()
    If txtMsg.Text <> "" Then
        Call fMain.SendMessageToAll(txtMsg.Text)
        txtLog.Text = txtLog.Text + "[ " + myInfo.sName + " 在 " + GetTime + " 时发送广播 ] " + txtMsg.Text + vbCrLf + vbCrLf
        txtLog.SelStart = Len(txtLog.Text)
        txtMsg.Text = ""
    End If
    txtMsg.SetFocus
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
        Me.Hide
    End If
End Sub

Public Sub OnBroadcastArrival(ByRef sMsg As String, ByVal sName As String, ByVal sip As String)
    txtLog.Text = txtLog.Text + "[ " + sName + " @ " + sip + " 在 " + GetTime + " 时发送广播 ] " + sMsg + vbCrLf + vbCrLf
    txtLog.SelStart = Len(txtLog.Text)

    If gbAutoPopup Then
        Me.Show vbModeless
    Else
        fMain.AddMsgForm Me
    End If
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Then 'Ctrl+Enter pressed
        cmdSend_Click
        KeyAscii = 0
    End If
End Sub

