VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "*\APencil.vbp"
Object = "*\ATools.vbp"
Begin VB.Form FMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Derio's Coloring Book"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10290
   Icon            =   "FMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   530
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   686
   StartUpPosition =   2  'CenterScreen
   Begin Tools.ctlTool ctlLoading 
      Height          =   780
      Left            =   9300
      ToolTipText     =   " Open coloring image ..."
      Top             =   1320
      Width           =   930
      _ExtentX        =   1640
      _ExtentY        =   1376
      ToolsName       =   6
      MousePointer    =   99
      MouseIcon       =   "FMain.frx":0442
   End
   Begin Tools.ctlTool ctlSaving 
      Height          =   975
      Left            =   9240
      ToolTipText     =   " Save as ..."
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1720
      ToolsName       =   5
      MousePointer    =   99
      MouseIcon       =   "FMain.frx":05A4
   End
   Begin Tools.ctlTool ctlDisk 
      Height          =   825
      Left            =   9420
      Top             =   660
      Width           =   825
      _ExtentX        =   1455
      _ExtentY        =   1455
      ToolsName       =   3
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":0706
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   720
      Index           =   13
      Left            =   150
      Top             =   7125
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":0722
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   12
      Left            =   135
      Top             =   6555
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":073E
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   11
      Left            =   135
      Top             =   6015
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":075A
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   10
      Left            =   135
      Top             =   5475
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":0776
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   9
      Left            =   135
      Top             =   4935
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":0792
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   8
      Left            =   135
      Top             =   4395
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":07AE
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   7
      Left            =   135
      Top             =   3855
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":07CA
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   6
      Left            =   135
      Top             =   3315
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":07E6
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   4
      Left            =   135
      Top             =   2775
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":0802
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   3
      Left            =   135
      Top             =   2235
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":081E
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   2
      Left            =   135
      Top             =   1695
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":083A
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   1
      Left            =   135
      Top             =   1155
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":0856
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   5
      Left            =   135
      Top             =   615
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":0872
   End
   Begin Tools.ctlTool ctlBinder 
      Height          =   630
      Index           =   0
      Left            =   135
      Top             =   75
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   1270
      ToolsName       =   4
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":088E
   End
   Begin Tools.ctlTool ctlEraser 
      Height          =   735
      Left            =   8340
      ToolTipText     =   " Eraser "
      Top             =   7140
      Width           =   645
      _ExtentX        =   1138
      _ExtentY        =   1296
      ToolsName       =   1
      MousePointer    =   99
      MouseIcon       =   "FMain.frx":08AA
   End
   Begin VB.PictureBox pctPencilBox 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1140
      MouseIcon       =   "FMain.frx":0A0C
      MousePointer    =   99  'Custom
      Picture         =   "FMain.frx":0B5E
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   419
      TabIndex        =   2
      ToolTipText     =   " Changes color table "
      Top             =   7740
      Width           =   6285
   End
   Begin Pencil.ctlPencil ctlPencilList 
      Height          =   1380
      Index           =   0
      Left            =   1200
      Top             =   6360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2434
      Color           =   16777215
      Sharp           =   1
      MousePointer    =   99
      MouseIcon       =   "FMain.frx":81C0
   End
   Begin Tools.ctlTool ctlPencilBox 
      Height          =   675
      Left            =   1140
      Top             =   7140
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   1191
      ToolsName       =   2
      MousePointer    =   1
      MouseIcon       =   "FMain.frx":8322
   End
   Begin Pencil.ctlPencil ctlActivePencil 
      Height          =   1380
      Index           =   2
      Left            =   9780
      ToolTipText     =   " Duller pencil "
      Top             =   7020
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   2434
      Color           =   16777215
      Sharp           =   3
      MousePointer    =   99
      MouseIcon       =   "FMain.frx":833E
   End
   Begin Pencil.ctlPencil ctlActivePencil 
      Height          =   1380
      Index           =   1
      Left            =   9360
      ToolTipText     =   " Dull pencil "
      Top             =   7020
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   2434
      Color           =   16777215
      Sharp           =   2
      MousePointer    =   99
      MouseIcon       =   "FMain.frx":84A0
   End
   Begin Pencil.ctlPencil ctlActivePencil 
      Height          =   1380
      Index           =   0
      Left            =   9060
      ToolTipText     =   " Pointed pencil"
      Top             =   7020
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   2434
      Color           =   16777215
      Sharp           =   1
      MousePointer    =   99
      MouseIcon       =   "FMain.frx":8602
   End
   Begin MSComDlg.CommonDialog cmdFile 
      Left            =   7560
      Top             =   -60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pctBack 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7650
      Left            =   -45
      ScaleHeight     =   7620
      ScaleWidth      =   10125
      TabIndex        =   0
      Top             =   120
      Width           =   10155
      Begin VB.PictureBox pctTarget 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   7230
         Left            =   960
         Picture         =   "FMain.frx":8764
         ScaleHeight     =   482
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   576
         TabIndex        =   1
         Top             =   60
         Width           =   8640
         Begin VB.Shape shpBrush 
            Height          =   195
            Left            =   600
            Top             =   1860
            Visible         =   0   'False
            Width           =   135
         End
      End
   End
   Begin VB.Image imgTemplates 
      Height          =   7200
      Left            =   10440
      Picture         =   "FMain.frx":D2826
      Top             =   360
      Visible         =   0   'False
      Width           =   8610
   End
   Begin VB.Shape shpPage 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   7635
      Index           =   0
      Left            =   525
      Top             =   165
      Width           =   9555
   End
   Begin VB.Shape shpPage 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   7635
      Index           =   1
      Left            =   495
      Top             =   195
      Width           =   9555
   End
   Begin VB.Shape shpPage 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   7635
      Index           =   2
      Left            =   480
      Top             =   210
      Width           =   9555
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************************
'* Title  : Derio's Coloring Book        *
'* Type   : Education Application        *
'* Stamp  : 8 Feb 2008                   *
'*****************************************

Private Declare Sub Sleep _
        Lib "kernel32" _
        (ByVal dwMilliseconds As Long)
Dim ColorTables(1 To 64) As Long

Dim BackGroundColor As Long
Dim BorderLineColor As Long
Dim BrushSize As Integer
Dim BrushWidth As Integer
Dim ActiveColor As Long
Dim ActiveFileName As String
Dim ActivePencilIndex As Integer
Dim ActiveColorTable As Integer
Dim DrawingMode As Integer
Dim PencilSharp As Integer
Dim ImageChange As Boolean

Const MAX_PENCILS = 24
Const APP_TITLE = "Derio's Coloring Book"



Private Sub CreateColorTables(ByVal Shift As Integer)
'* Create 64 (4*4*4) colors and store to ColorTables array
'  using Formula : Color = Shift + Factor * ColorComponent

Dim Factor As Integer
Dim Red As Integer
Dim Green As Integer
Dim Blue As Integer
Dim Index As Integer

  'calculate the Factor and the Shift
  Factor = (256 - Shift) \ 3
  Shift = 256 - Factor * 3
  
  'create colors
  For Red = 0 To 3
    For Green = 0 To 3
      For Blue = 0 To 3
        Index = Index + 1
        ColorTables(Index) = RGB(Shift + Factor * Red, _
                                 Shift + Factor * Green, _
                                 Shift + Factor * Blue)
      Next Blue
    Next Green
  Next Red
End Sub

Private Function GetColor(ByVal TableIndex As Integer, ByVal Index As Integer) As Long
'* Get color from ColorTables base on TableIndex and Index
'  TableIndex range between 0 and 2
'  Assumption: 24 colors on the same TableIndex

Dim StartPos As Integer

  Select Case TableIndex
  Case 0
    StartPos = 0
  Case 1
    StartPos = 18
  Case 2
    StartPos = 40
  End Select
  GetColor = ColorTables(StartPos + Index)
End Function

Private Sub ChangeAllPencilsColor()
'* Change all the pencils color base on turn

Dim I As Integer
Dim J As Integer
Dim arrRandom() As Integer
Dim TmpIndex As Integer

  'get the random seq
  ReDim arrRandom(Me.ctlPencilList.Count - 1)
  GetRandom arrRandom()
  
  'turn active pencil down
  If ActivePencilIndex <> -1 Then
    With ctlPencilList(ActivePencilIndex)
    For I = 1 To 25
      .Top = .Top + 1
      DoEvents
    Next I
    End With
    TmpIndex = ActivePencilIndex
    ActivePencilIndex = -1
    
  Else
    TmpIndex = -1
  End If
  
  'set next color table that will be used
  ActiveColorTable = (ActiveColorTable + 1) Mod 3
  
  'turn all of the pencil down base on random seq
  For J = 1 To 10
    For I = 0 To Me.ctlPencilList.Count - 1
      With Me.ctlPencilList(arrRandom(I))
        .Top = .Top + 6
      End With
    Next I
    DoEvents
  Next J
   
  'change the color
  For I = 0 To Me.ctlPencilList.Count - 1
    With Me.ctlPencilList(arrRandom(I))
      .Color = GetColor(ActiveColorTable + 1, arrRandom(I) + 1)
    End With
  Next I
   
  'show up the new pencil with new random seq
  GetRandom arrRandom()
  For I = Me.ctlPencilList.Count - 1 To 0 Step -1
    For J = 1 To 15
      With Me.ctlPencilList(arrRandom(I))
        .Top = .Top - 4
      End With
      DoEvents
    Next J
    DoEvents
  Next I
  
  'setup the active pencil
  If TmpIndex <> -1 Then
    ctlPencilList_LeftClick TmpIndex
  End If
End Sub

Private Sub GetRandom(ArrTable() As Integer)
'* Get random seq

Dim I As Integer
Dim J As Integer
Dim X As Integer
Dim Y As Integer
Dim Z As Integer

  'create the seq number
  J = UBound(ArrTable) + 1
  For I = 0 To J - 1
    ArrTable(I) = I
  Next I
  
  'randomize the number
  For I = 1 To J * 4
    X = Int(Rnd * 32) Mod J
    Y = Int(Rnd * 32) Mod J
    Z = ArrTable(X)
    ArrTable(X) = ArrTable(Y)
    ArrTable(Y) = Z
  Next I
End Sub

Private Sub ctlActivePencil_LeftClick(Index As Integer)
'* Select the sharpness of the pencil

Dim I As Integer
Dim J As Integer

  'turn down prev sharpness pencil
  For I = 0 To ctlActivePencil.Count - 1
    With ctlActivePencil(I)
      If .Tag <> "" Then
        For J = 1 To 7
          .Top = .Top + 4
        Next J
      End If
      .Tag = ""
    End With
  Next I
  
  'push it up
  PencilSharp = ctlActivePencil(Index).Sharp
  With ctlActivePencil(Index)
    For I = 1 To 7
      .Top = .Top - 4
    Next I
    .Tag = "1"
  End With
  
  'activate the selected sharp pencil
  If DrawingMode = 1 Then BrushSize = GetBrushSize()
End Sub

Private Sub ctlEraser_Click()
'* Activate the eraser

Dim I As Integer

  If DrawingMode <> 2 Then
    'turn the active pencil down
    DownActivePencil
    
    'push the eraser up
    For I = 1 To 20
      With Me.ctlEraser
        .Top = .Top - 1
        If I Mod 2 = 0 Then
          .Left = .Left - 1
        End If
      End With
      DoEvents
    Next I
  
    'activate the eraser
    SetDrawingMode 2
  End If
End Sub

Private Sub ctlLoading_Click()
'* Open the saved image

  If ImageChange Then
    Select Case MsgBox("The images has changed." & vbCrLf & _
                        "Do you want to save it before open the new one?", _
                        vbQuestion + vbYesNoCancel, _
                        APP_TITLE)
    Case vbYes
      ctlSaving_Click
      If ImageChange Then Exit Sub
      
    Case vbCancel
      Exit Sub
    End Select
  End If
  
  On Local Error Resume Next
  With Me.cmdFile
    .CancelError = True
    .DialogTitle = "Open image"
    .Filter = "Bitmap (*.BMP)|*.bmp"
    .FilterIndex = 0
    .Flags = cdlOFNFileMustExist
    .FileTitle = ActiveFileName
    .ShowOpen
    If Err <> 32755 Then
      If .FileName <> "" Then
        Me.pctTarget = LoadPicture(.FileName)
        SavePicture Me.pctTarget.Image, .FileName
        ActiveFileName = .FileTitle
        Caption = APP_TITLE & " - " & ActiveFileName
        ImageChange = False
      End If
    End If
  End With
  On Local Error GoTo 0
End Sub

Private Sub ctlPencilList_LeftClick(Index As Integer)
'* Select the pencil

Dim I As Integer

  If DrawingMode = 1 Then
    DownActivePencil
  
  ElseIf DrawingMode = 2 Then
    For I = 1 To 20
      With Me.ctlEraser
        .Top = .Top + 1
        If I Mod 2 = 0 Then
          .Left = .Left + 1
        End If
      End With
    Next I
  End If
  
  'push the selected pencil up
  With ctlPencilList(Index)
    For I = 1 To 25
      .Top = .Top - 1
      DoEvents
    Next I
    SetActivePencilColor .Color
  End With
  ActivePencilIndex = Index
  
  'activate the drawing mode
  SetDrawingMode 1
End Sub

Private Sub SetActivePencilColor(ByVal Color As Long)
'* Setup the sharpness pencil color base on the selected color

Dim I As Integer

  For I = 0 To Me.ctlActivePencil.Count - 1
    Me.ctlActivePencil(I).Color = Color
  Next I
End Sub

Private Sub DownActivePencil()
Dim I As Integer

  If ActivePencilIndex <> -1 Then
    With ctlPencilList(ActivePencilIndex)
    For I = 1 To 25
      .Top = .Top + 1
      DoEvents
    Next I
    End With
  End If
  ActivePencilIndex = -1
End Sub

Private Sub SetDrawingMode(ByVal Mode As Integer)
  Select Case Mode
  Case 1 'using pencil
    DrawingMode = 1
    BrushSize = GetBrushSize()
    BrushWidth = 1
    ActiveColor = Me.ctlPencilList(ActivePencilIndex).Color
    
  Case 2 'using eraser
    DrawingMode = 2
    BrushSize = 5
    BrushWidth = BrushSize
    ActiveColor = BackGroundColor
  End Select
End Sub

Private Function GetBrushSize() As Integer
  Select Case PencilSharp
  Case 1
    GetBrushSize = 1
  Case 2
    GetBrushSize = 3
  Case 3
    GetBrushSize = 5
  End Select
End Function

Private Sub ctlSaving_Click()

  On Local Error Resume Next
  
  With Me.cmdFile
    .CancelError = True
    .DialogTitle = "Save image"
    .Filter = "Bitmap (*.BMP)|*.bmp"
    .FilterIndex = 0
    .Flags = cdlOFNOverwritePrompt
   .FileTitle = ActiveFileName
    .ShowSave
    If Err <> 32755 Then
      If .FileName <> "" Then
        SavePicture Me.pctTarget.Image, .FileName
        ActiveFileName = .FileTitle
        Caption = APP_TITLE & " - " & ActiveFileName
        ImageChange = False
      End If
    End If
  End With
  
  On Local Error GoTo 0
End Sub

Private Sub Form_Load()
  Caption = APP_TITLE
  BackGroundColor = RGB(255, 255, 255)
  BorderLineColor = 0
  CreateColorTables 64
  CreatePencilList
  ActiveColorTable = -1
  ChangeAllPencilsColor
  ctlPencilList_LeftClick 0
  ctlActivePencil_LeftClick 0
End Sub

Private Sub CreatePencilList()
Dim I As Integer

  Me.ctlPencilList(0).Top = Me.ScaleHeight - Me.ctlPencilList(0).Height + 40
  For I = 1 To MAX_PENCILS - 1
    Load Me.ctlPencilList(I)
    With Me.ctlPencilList(I)
      .Left = Me.ctlPencilList(I - 1).Left + .Width
      .Top = Me.pctPencilBox.Top - .Height + 50 + 10 * Rnd
      .Visible = True
      .ZOrder
    End With
  Next I
  Me.pctPencilBox.ZOrder
  ActivePencilIndex = -1
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If ImageChange Then
    Select Case MsgBox("The images has changed." & vbCrLf & _
                       "Do you want to save it before exit?", _
                       vbQuestion + vbYesNoCancel, _
                       APP_TITLE)
    Case vbYes
      ctlSaving_Click
      If ImageChange Then
        Cancel = True
        Exit Sub
      End If
      
    Case vbCancel
      Cancel = True
      Exit Sub
    End Select
  End If
End Sub

Private Sub pctPencilBox_Click()
  ChangeAllPencilsColor
  DoEvents
End Sub

Private Sub pctTarget_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = vbLeftButton Then
    ShowBrush X, Y
    DrawDot X, Y, ActiveColor
  Else
    HideBrush
  End If
End Sub

Private Sub ShowBrush(ByVal X As Integer, ByVal Y As Integer)
  With Me.shpBrush
    If Not .Visible Then
      .Visible = True
      .Width = BrushWidth * 2
      .Height = BrushSize * 2
    End If
    .Top = Y - .Height \ 2
    .Left = X - .Width \ 2
  End With
End Sub

Private Sub HideBrush()
  With Me.shpBrush
    If .Visible Then .Visible = False
  End With
End Sub

Private Sub GetRGB(Color As Long, R As Long, G As Long, B As Long)
'* Extract the Color into R, G, B components

  R = Color And 255
  G = (Color And 65280) \ 256
  B = Color \ 65536
End Sub

Private Function BlandColor(Color1 As Long, Color2 As Long, _
                            ByVal dX1 As Integer, ByVal dX2 As Integer) As Long
'* Bland Color1 and Color2 with distribution of weight

Dim R1 As Long
Dim R2 As Long
Dim G1 As Long
Dim G2 As Long
Dim B1 As Long
Dim B2 As Long

  GetRGB Color1, R1, G1, B1
  GetRGB Color2, R2, G2, B2
  
  BlandColor = RGB((dX1 * R1 + dX2 * R2) \ (dX1 + dX2), _
                   (dX1 * G1 + dX2 * G2) \ (dX1 + dX2), _
                   (dX1 * B1 + dX2 * B2) \ (dX1 + dX2))
End Function

Private Function DrawDot(ByVal X As Integer, _
                         ByVal Y As Integer, _
                         ByVal Color As Long)
'* Draw the area from X,Y as the origin with specific Color

Dim I As Integer
Dim J As Integer
Dim BackColor As Long

  ImageChange = True
  Select Case DrawingMode
  Case 1 'pencil
    For J = Y - BrushSize To Y + BrushSize
      For I = X - BrushWidth To X + BrushWidth
        If Rnd < 0.75 Then
          BackColor = Me.pctTarget.Point(I, J)
          If BackColor <> BorderLineColor Then
            If BackColor <> BackGroundColor And _
               BackColor <> Color Then
              Me.pctTarget.PSet (I, J), _
                                BlandColor(Color, BackColor, 1, 10)
            Else
              Me.pctTarget.PSet (I, J), Color
            End If
          End If
        End If
      Next I
    Next J

  Case 2 'eraser
    For J = Y - BrushSize To Y + BrushSize
      For I = X - BrushWidth To X + BrushWidth
        BackColor = Me.pctTarget.Point(I, J)
        If BackColor <> BorderLineColor Then
          Me.pctTarget.PSet (I, J), Color
        End If
      Next I
    Next J
  End Select
End Function

