VERSION 5.00
Begin VB.UserControl ctlTool 
   AutoRedraw      =   -1  'True
   BackColor       =   &H0000FF00&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   1740
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6285
   FillStyle       =   0  'Solid
   MaskColor       =   &H00FF00FF&
   MaskPicture     =   "ctlTool.ctx":0000
   Picture         =   "ctlTool.ctx":DDBE
   ScaleHeight     =   1740
   ScaleWidth      =   6285
   Begin VB.Image imgLib 
      Height          =   645
      Index           =   7
      Left            =   5100
      Picture         =   "ctlTool.ctx":1BB7C
      Top             =   720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgLib 
      Height          =   645
      Index           =   6
      Left            =   720
      Picture         =   "ctlTool.ctx":1D49A
      Top             =   780
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgLib 
      Height          =   780
      Index           =   5
      Left            =   4380
      Picture         =   "ctlTool.ctx":1EDB8
      Top             =   780
      Visible         =   0   'False
      Width           =   930
   End
   Begin VB.Image imgLib 
      Height          =   975
      Index           =   4
      Left            =   3360
      Picture         =   "ctlTool.ctx":2142A
      Top             =   780
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image imgLib 
      Height          =   720
      Index           =   3
      Left            =   1260
      Picture         =   "ctlTool.ctx":24018
      Top             =   780
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Image imgLib 
      Height          =   825
      Index           =   2
      Left            =   2160
      Picture         =   "ctlTool.ctx":25A9A
      Top             =   840
      Visible         =   0   'False
      Width           =   825
   End
   Begin VB.Image imgLib 
      Height          =   675
      Index           =   1
      Left            =   1380
      Picture         =   "ctlTool.ctx":27EF4
      Top             =   960
      Visible         =   0   'False
      Width           =   6285
   End
   Begin VB.Image imgLib 
      Height          =   735
      Index           =   0
      Left            =   360
      Picture         =   "ctlTool.ctx":35CB2
      Top             =   900
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "ctlTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum Type_ToolsName
  Eraser = 1
  PencilBox = 2
  Disk = 3
  Binder = 4
  DownArrow = 5
  RightArrow = 6
  FancyLeftArrow = 7
  FancyRightArrow = 8
End Enum

Private vToolsName As Type_ToolsName
Private NewWidth As Integer
Private NewHeight As Integer

Public Event Click()

Private Sub UserControl_Click()
  RaiseEvent Click
End Sub


Public Property Get ToolsName() As Type_ToolsName
  ToolsName = vToolsName
End Property

Public Property Let ToolsName(ByVal vNewValue As Type_ToolsName)
Dim Index As Integer

  If vNewValue <= 0 Then Exit Property
  vToolsName = vNewValue
  Index = vToolsName
  Index = Index - 1
  UserControl.MaskPicture = UserControl.imgLib(Index).Picture
  UserControl.Picture = UserControl.imgLib(Index).Picture
  NewWidth = UserControl.imgLib(Index).Width
  NewHeight = UserControl.imgLib(Index).Height
  UserControl_Resize
  UserControl.Refresh
  PropertyChanged "ToolsName"
End Property

Private Sub UserControl_Initialize()
  NewWidth = UserControl.imgLib(1).Width
  NewHeight = UserControl.imgLib(1).Height
End Sub

Private Sub UserControl_InitProperties()
  vToolsName = PencilBox
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  ToolsName = PropBag.ReadProperty("ToolsName", Type_ToolsName.PencilBox)
  MousePointer = PropBag.ReadProperty("MousePointer", vbArrow)
  Set MouseIcon = PropBag.ReadProperty("MouseIcon", UserControl.MouseIcon)
End Sub

Private Sub UserControl_Resize()
  UserControl.Width = NewWidth
  UserControl.Height = NewHeight
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  PropBag.WriteProperty "ToolsName", vToolsName
  PropBag.WriteProperty "MousePointer", UserControl.MousePointer
  PropBag.WriteProperty "MouseIcon", UserControl.MouseIcon
End Sub

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
