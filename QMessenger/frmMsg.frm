VERSION 5.00
Begin VB.Form frmMsg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���� / ������Ϣ"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "����"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMsg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "�ر�(&C)"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Timer tmrAD 
      Interval        =   5000
      Left            =   2760
      Top             =   3720
   End
   Begin VB.PictureBox picAD 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "����"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   615
      Left            =   3480
      ScaleHeight     =   555
      ScaleWidth      =   2595
      TabIndex        =   6
      ToolTipText     =   "���Ͷ�ţ�100�����/�죻���̵绰��010-62286253"
      Top             =   120
      Width           =   2655
   End
   Begin VB.TextBox txtMsg 
      Height          =   735
      HideSelection   =   0   'False
      Left            =   120
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2880
      Width           =   6735
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "����(&S)"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3720
      Width           =   1335
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
      Top             =   1080
      Width           =   6735
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "��ϸ����"
      Height          =   615
      Left            =   6240
      TabIndex        =   7
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox txtIP 
      BackColor       =   &H8000000F&
      Height          =   270
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "x.x.x.x/_pc_name_"
      Top             =   480
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H8000000F&
      Height          =   270
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "_qq_name_"
      Top             =   120
      Width           =   2055
   End
   Begin VB.Image imgFace 
      Height          =   480
      Left            =   120
      Picture         =   "frmMsg.frx":0CCA
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "�ڴ�������Ϣ(���100���ַ����� Ctrl+Enter ���ٷ���)��"
      Height          =   180
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   4770
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "�����¼��"
      Height          =   180
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "IP��ַ"
      Height          =   180
      Left            =   720
      TabIndex        =   9
      Top             =   520
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "������"
      Height          =   180
      Left            =   720
      TabIndex        =   8
      Top             =   160
      Width           =   540
   End
End
Attribute VB_Name = "frmMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ui As USER_INFO
Dim sPoem(3) As String

'�յ�ĳ�û���Ϣ:
Public Sub OnMessageArrival(ByRef sMsg As String)
    txtLog.SelStart = Len(txtLog.Text)
    txtLog.Text = txtLog.Text + "[ " + ui.sName + " �� " + GetTime + " ʱ˵ ] " + sMsg + vbCrLf + vbCrLf
    txtLog.SelStart = Len(txtLog.Text)
    If gbAutoPopup Then
        Me.Show vbModeless
    Else
        fMain.AddMsgForm Me
    End If
End Sub

Private Sub cmdClose_Click()
    txtMsg.Text = ""
    Me.Hide
End Sub

Private Sub cmdDetail_Click()
    Call fMain.QueryUserDetail(ui.sHostIP)
End Sub

Private Sub cmdSend_Click()
    If txtMsg.Text <> "" Then
        Call fMain.SendMessageToUser(ui.sHostIP, txtMsg.Text)
        txtLog.Text = txtLog.Text + "[ " + myInfo.sName + " �� " + GetTime + " ʱ˵ ] " + txtMsg.Text + vbCrLf + vbCrLf
        txtLog.SelStart = Len(txtLog.Text)
        txtMsg.Text = ""
    End If
    txtMsg.SetFocus
End Sub

Private Sub Form_Load()
    Set ui = New USER_INFO
    'ʹ����������ǰ��
    If gbAlwaysOnTop Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If

    Call Rnd(-Second(Now) * 90.1809 * Minute(Now))
    Select Case Int(Rnd() * 4)
    Case 0
        sPoem(1) = " ������ɽѩ���޻�ֻ�к���"
        sPoem(2) = " ��������������ɫδ������"
        sPoem(3) = " ��ս���ģ����߱��񰰡�"
        sPoem(0) = " Ը�����½���ֱΪն¥����"
    Case 1
        sPoem(1) = " ��ڷ���Σ���³��δ�ˡ�"
        sPoem(2) = " �컯�����㣬�����������"
        sPoem(3) = " ���������ƣ����������"
        sPoem(0) = " �ᵱ�������һ����ɽС��"
    Case 2
        sPoem(1) = " ��ɽ��������������"
        sPoem(2) = " �����ɼ��գ���Ȫʯ������"
        sPoem(3) = " �������Ů�����������ۡ�"
        sPoem(0) = " ���ⴺ��Ъ�������Կ�����"
    Case 3
        sPoem(1) = " �����п�����������ɽׯ��"
        sPoem(2) = " �������ˮ��ɣ����������"
        sPoem(3) = " һ�贺���죬ʮ�ﵾ���㡣"
        sPoem(0) = " ʢ���޼��٣������֯æ��"
    End Select
    Call tmrAD_Timer
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set ui = Nothing
End Sub

'������ʾ�û���Ϣ��
Public Sub UpdateUserInfo()
    txtName.Text = ui.sName
    txtIP.Text = ui.sHostIP + "/" + ui.sHostName
    Set imgFace.Picture = fMain.imglstFace.ListImages(ui.nFace).ExtractIcon()
End Sub

Private Sub tmrAD_Timer()
    Static r As Long
    r = r + 1
    If r = 4 Then r = 0

    picAD.Cls
    picAD.CurrentY = 160
    picAD.Print sPoem(r)
End Sub

Private Sub txtMsg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Then 'Ctrl+Enter pressed
        cmdSend_Click
        KeyAscii = 0
    End If
End Sub


