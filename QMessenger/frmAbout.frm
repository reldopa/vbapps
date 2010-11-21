VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于 QMessenger"
   ClientHeight    =   3885
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   6615
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2681.496
   ScaleMode       =   0  'User
   ScaleWidth      =   6211.828
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3390
      Width           =   1455
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      Caption         =   "在此留言"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   5520
      MouseIcon       =   "frmAbout.frx":08CA
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   1800
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "有任何意见或建议，请"
      Height          =   180
      Left            =   3720
      TabIndex        =   5
      Top             =   1800
      Width           =   1800
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "免费软件"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   3720
      TabIndex        =   4
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "作者：廖雪峰"
      Height          =   180
      Left            =   3720
      TabIndex        =   3
      Top             =   1440
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "版本：1.1.0408"
      Height          =   180
      Left            =   5040
      TabIndex        =   2
      Top             =   600
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "QMessenger"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3720
      TabIndex        =   1
      Top             =   240
      Width           =   1515
   End
   Begin VB.Image imgLogo 
      BorderStyle     =   1  'Fixed Single
      Height          =   2925
      Left            =   120
      Picture         =   "frmAbout.frx":0A1C
      ToolTipText     =   "一只可爱的大猫正在欺骗一只善良的小老鼠"
      Top             =   143
      Width           =   3360
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   112.686
      X2              =   6085.056
      Y1              =   2236.306
      Y2              =   2236.306
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   6085.056
      Y1              =   2257.012
      Y2              =   2257.012
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    '使窗口总在最前：
    If gbAlwaysOnTop Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Private Sub lblEmail_Click()
    ShellExecute Me.hwnd, "open", "http://code.google.com/p/vbapps/wiki/QMessenger", "", "", 5
End Sub
