VERSION 5.00
Begin VB.UserControl BitmapButton 
   CanGetFocus     =   0   'False
   ClientHeight    =   1245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1245
   ForwardFocus    =   -1  'True
   ScaleHeight     =   1245
   ScaleWidth      =   1245
   Begin VB.Image imgDown 
      Height          =   975
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Image imgHot 
      Enabled         =   0   'False
      Height          =   975
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "BitmapButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

'Event Declarations:
Public Event Click()
Dim bPressed As Boolean

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UserControl.Enabled = False Then Exit Sub
    'if button has been pressed down:
    If bPressed Then Exit Sub

    'button is not pressed down:
    If X >= 0 And Y >= 0 And X < UserControl.Width And Y < UserControl.Height Then
        If Button = vbLeftButton Then
            imgDown.Visible = True
        Else
            imgHot.Visible = True
        End If
        SetCapture UserControl.hwnd
    Else
        If X < 0 Or Y < 0 Or X >= UserControl.Width Or Y >= UserControl.Height Then
            imgHot.Visible = False
            imgDown.Visible = False
            ReleaseCapture
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgHot.Visible = False
    If UserControl.Enabled Then
        If Button = vbLeftButton Then RaiseEvent Click
        If bPressed Then
            imgDown.Visible = True
        End If
    End If
    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub UserControl_Resize()
    imgHot.Refresh
    imgDown.Refresh
End Sub

Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set UserControl.Picture = New_Picture
    PropertyChanged "Picture"
End Property

Public Property Get HotPicture() As Picture
    Set HotPicture = imgHot.Picture
End Property

Public Property Set HotPicture(ByVal New_Picture As Picture)
    Set imgHot.Picture = New_Picture
    PropertyChanged "HotPicture"
End Property

Public Property Get DownPicture() As Picture
    Set DownPicture = imgDown.Picture
End Property

Public Property Set DownPicture(ByVal New_Picture As Picture)
    Set imgDown.Picture = New_Picture
    PropertyChanged "DownPicture"
End Property

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    Enabled = PropBag.ReadProperty("Enabled", True)
    IsPressed = PropBag.ReadProperty("IsPressed", False)
    Set HotPicture = PropBag.ReadProperty("HotPicture", Nothing)
    Set DownPicture = PropBag.ReadProperty("DownPicture", Nothing)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", Enabled, True)
    Call PropBag.WriteProperty("IsPressed", bPressed, False)
    Call PropBag.WriteProperty("HotPicture", HotPicture, Nothing)
    Call PropBag.WriteProperty("DownPicture", DownPicture, Nothing)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
End Sub

Public Property Get IsPressed() As Boolean
    IsPressed = bPressed
End Property

Public Property Let IsPressed(ByVal vNewValue As Boolean)
    bPressed = vNewValue
    imgDown.Visible = bPressed
    PropertyChanged "IsPressed"
End Property
