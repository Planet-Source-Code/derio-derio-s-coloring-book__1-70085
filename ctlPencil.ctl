VERSION 5.00
Begin VB.UserControl ctlPencil 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   3405
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2595
   FillStyle       =   0  'Solid
   MaskColor       =   &H00FF00FF&
   MaskPicture     =   "ctlPencil.ctx":0000
   Picture         =   "ctlPencil.ctx":12F2
   ScaleHeight     =   227
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   173
   Begin VB.Image imgLib 
      Height          =   1380
      Index           =   2
      Left            =   780
      Picture         =   "ctlPencil.ctx":25E4
      Top             =   1500
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgLib 
      Height          =   1380
      Index           =   1
      Left            =   360
      Picture         =   "ctlPencil.ctx":48A6
      Top             =   1500
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Image imgLib 
      Height          =   1380
      Index           =   0
      Left            =   60
      Picture         =   "ctlPencil.ctx":62C8
      Top             =   1500
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "ctlPencil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private Declare Function ExtFloodFill _
        Lib "gdi32" _
        (ByVal hdc As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal crColor As Long, _
        ByVal wFillType As Long) As Long


Private vColor As OLE_COLOR
Private vSharp As Integer

Public Event LeftClick()
Public Event RightClick()

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    RaiseEvent LeftClick
  ElseIf Button = vbRightButton Then
    RaiseEvent RightClick
  End If
End Sub

Private Sub UserControl_Resize()
  Select Case vSharp
  Case 1
    UserControl.Width = 255
  Case 2
    UserControl.Width = 360
  Case 3
    UserControl.Width = 480
  End Select
  UserControl.Height = 1380
End Sub

Public Property Get Color() As OLE_COLOR
  Color = vColor
End Property

Public Property Let Color(ByVal NewValue As OLE_COLOR)
Dim Result As Long
Dim R As Long
Dim G As Long
Dim B As Long
Dim BorderColor As Long

  UserControl.FillColor = NewValue
  Select Case vSharp
  Case 1
    Result = ExtFloodFill(UserControl.hdc, 8, 3, vColor, 1)
    
    GetRGB NewValue, R, G, B
    
    'draw the light pencil tip
    BorderColor = RGB(R \ 2, G \ 2, B \ 2)
    UserControl.PSet (9, 1), BorderColor
    UserControl.PSet (10, 2), BorderColor
    UserControl.PSet (10, 3), BorderColor
    UserControl.PSet (11, 4), BorderColor
    UserControl.PSet (11, 5), BorderColor
    
    'draw the dark pencil tip
    BorderColor = RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.PSet (8, 0), BorderColor
    UserControl.PSet (7, 1), BorderColor
    UserControl.PSet (6, 2), BorderColor
    UserControl.PSet (6, 3), BorderColor
    UserControl.PSet (5, 4), BorderColor
    UserControl.PSet (5, 5), BorderColor
    
    'draw the body
    UserControl.Line (4, 18)-(12, 91), NewValue, BF
    UserControl.Line (4, 17)-(5, 17), NewValue
    UserControl.Line (8, 17)-(10, 17), NewValue
    
    'draw the dark pencil body
    UserControl.Line (13, 18)-(13, 91), RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.Line (14, 18)-(14, 91), RGB(3 * R \ 5, 3 * G \ 5, 3 * B \ 5)
    UserControl.Line (15, 17)-(15, 91), RGB(3 * R \ 7, 3 * G \ 7, 3 * B \ 7)
    
    'draw the light pencil body
    UserControl.Line (3, 18)-(3, 91), RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.Line (2, 17)-(2, 91), RGB(4 * R \ 5, 4 * G \ 5, 4 * B \ 5)
    UserControl.Line (1, 17)-(1, 91), RGB(6 * R \ 7, 6 * G \ 7, 6 * B \ 7)
    UserControl.Line (0, 18)-(0, 91), RGB(R, G, B)
    
  Case 2
    Result = ExtFloodFill(UserControl.hdc, 12, 3, vColor, 1)
    
    GetRGB NewValue, R, G, B
    BorderColor = RGB(R \ 2, G \ 2, B \ 2)
    UserControl.PSet (14, 2), BorderColor
    UserControl.PSet (14, 3), BorderColor
    UserControl.PSet (15, 4), BorderColor
    UserControl.PSet (15, 5), BorderColor
    
    BorderColor = RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.PSet (10, 1), BorderColor
    UserControl.PSet (11, 1), BorderColor
    UserControl.PSet (12, 1), BorderColor
    UserControl.PSet (13, 1), BorderColor
    UserControl.PSet (9, 2), BorderColor
    UserControl.PSet (9, 3), BorderColor
    UserControl.PSet (8, 4), BorderColor
    UserControl.PSet (8, 5), BorderColor
  
    'draw the body
    UserControl.Line (6, 22)-(17, 91), NewValue, BF
    UserControl.Line (14, 21)-(17, 21), NewValue
  
    'draw the dark pencil body
    UserControl.Line (18, 21)-(18, 91), RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.Line (19, 21)-(19, 91), RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.Line (20, 21)-(20, 91), RGB(3 * R \ 5, 3 * G \ 5, 3 * B \ 5)
    UserControl.Line (21, 22)-(21, 91), RGB(3 * R \ 5, 3 * G \ 5, 3 * B \ 5)
    UserControl.Line (22, 22)-(22, 91), RGB(3 * R \ 7, 3 * G \ 7, 3 * B \ 7)
    
    'draw the light pencil body
    UserControl.Line (4, 22)-(4, 91), RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.Line (5, 22)-(5, 91), RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.Line (2, 21)-(2, 91), RGB(4 * R \ 5, 4 * G \ 5, 4 * B \ 5)
    UserControl.Line (3, 21)-(3, 91), RGB(4 * R \ 5, 4 * G \ 5, 4 * B \ 5)
    UserControl.Line (1, 22)-(1, 91), RGB(6 * R \ 7, 6 * G \ 7, 6 * B \ 7)
    UserControl.Line (0, 22)-(0, 91), RGB(R, G, B)
    
  Case 3
    Result = ExtFloodFill(UserControl.hdc, 16, 4, vColor, 1)
    GetRGB NewValue, R, G, B
    BorderColor = RGB(R \ 2, G \ 2, B \ 2)
    UserControl.PSet (19, 3), BorderColor
    UserControl.PSet (20, 4), BorderColor
    UserControl.PSet (20, 5), BorderColor
    UserControl.PSet (21, 6), BorderColor
    
    BorderColor = RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.PSet (13, 2), BorderColor
    UserControl.PSet (14, 2), BorderColor
    UserControl.PSet (15, 2), BorderColor
    UserControl.PSet (16, 2), BorderColor
    UserControl.PSet (17, 2), BorderColor
    UserControl.PSet (18, 2), BorderColor
    UserControl.PSet (12, 3), BorderColor
    UserControl.PSet (11, 4), BorderColor
    UserControl.PSet (11, 5), BorderColor
    UserControl.PSet (10, 6), BorderColor
  
    'draw the body
    UserControl.Line (8, 32)-(23, 91), NewValue, BF
    UserControl.Line (8, 31)-(12, 31), NewValue
    UserControl.Line (17, 31)-(23, 31), NewValue
    UserControl.Line (8, 30)-(10, 30), NewValue
    UserControl.Line (19, 30)-(23, 30), NewValue
  
    'draw the dark pencil body
    UserControl.Line (24, 29)-(24, 91), RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.Line (25, 29)-(25, 91), RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.Line (26, 29)-(26, 91), RGB(3 * R \ 5, 3 * G \ 5, 3 * B \ 5)
    UserControl.Line (27, 29)-(27, 91), RGB(3 * R \ 5, 3 * G \ 5, 3 * B \ 5)
    UserControl.Line (28, 30)-(28, 91), RGB(3 * R \ 5, 3 * G \ 5, 3 * B \ 5)
    UserControl.Line (29, 30)-(29, 91), RGB(3 * R \ 5, 3 * G \ 5, 3 * B \ 5)
    UserControl.Line (30, 30)-(30, 91), RGB(3 * R \ 7, 3 * G \ 7, 3 * B \ 7)
    
    'draw the light pencil body
    UserControl.Line (6, 31)-(6, 91), RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.Line (7, 31)-(7, 91), RGB(3 * R \ 4, 3 * G \ 4, 3 * B \ 4)
    UserControl.Line (4, 30)-(4, 91), RGB(4 * R \ 5, 4 * G \ 5, 4 * B \ 5)
    UserControl.Line (5, 31)-(5, 91), RGB(4 * R \ 5, 4 * G \ 5, 4 * B \ 5)
    UserControl.Line (3, 30)-(3, 91), RGB(4 * R \ 5, 4 * G \ 5, 4 * B \ 5)
    UserControl.Line (2, 31)-(2, 91), RGB(6 * R \ 7, 6 * G \ 7, 6 * B \ 7)
    UserControl.Line (1, 31)-(1, 91), RGB(6 * R \ 7, 6 * G \ 7, 6 * B \ 7)
    UserControl.Line (0, 31)-(0, 91), RGB(R, G, B)
  
  End Select
  UserControl.Refresh
  vColor = NewValue
  PropertyChanged "Color"
End Property

Private Sub UserControl_Initialize()
  vColor = RGB(255, 255, 255)
  vSharp = 1
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  Color = PropBag.ReadProperty("Color", RGB(192, 192, 192))
  Sharp = PropBag.ReadProperty("Sharp", 1)
  MousePointer = PropBag.ReadProperty("MousePointer", vbArrow)
  Set MouseIcon = PropBag.ReadProperty("MouseIcon", UserControl.MouseIcon)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "Color", vColor
  PropBag.WriteProperty "Sharp", vSharp
  PropBag.WriteProperty "MousePointer", UserControl.MousePointer
  PropBag.WriteProperty "MouseIcon", UserControl.MouseIcon
End Sub

Private Sub GetRGB(Color As Long, R As Long, G As Long, B As Long)
  R = Color And 255
  G = (Color And 65280) \ 256
  B = Color \ 65536
End Sub

Public Property Get Sharp() As Integer
  Sharp = vSharp
End Property

Public Property Let Sharp(ByVal vNewValue As Integer)
Dim tmpColor As OLE_COLOR

  vSharp = vNewValue
  UserControl.MaskPicture = UserControl.imgLib(vSharp - 1).Picture
  UserControl.Picture = UserControl.imgLib(vSharp - 1).Picture
  UserControl_Resize
  tmpColor = vColor
  vColor = RGB(255, 255, 255)
  Me.Color = tmpColor
  UserControl.Refresh
  PropertyChanged "Sharp"
End Property

Public Property Get MousePointer() As MousePointerConstants
  MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal vNewValue As MousePointerConstants)
  UserControl.MousePointer = vNewValue
  PropertyChanged "MousePointer"
End Property

Public Property Get MouseIcon() As IPictureDisp
  Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal vNewValue As IPictureDisp)
  Set UserControl.MouseIcon = vNewValue
  PropertyChanged "MouseIcon"
End Property

