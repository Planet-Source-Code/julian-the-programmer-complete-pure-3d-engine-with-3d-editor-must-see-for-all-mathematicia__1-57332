VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Main 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Model Maker"
   ClientHeight    =   10485
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   ScaleHeight     =   699
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   698
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Backbuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5250
      Left            =   0
      ScaleHeight     =   346
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   5250
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   600
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer UpdateTimer 
      Interval        =   100
      Left            =   120
      Top             =   240
   End
   Begin VB.PictureBox PicYZ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00C0C0C0&
      Height          =   5250
      Left            =   5280
      ScaleHeight     =   346
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   3
      Top             =   5280
      Width           =   5250
   End
   Begin VB.PictureBox PicXZ 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000C000&
      Height          =   5250
      Left            =   0
      ScaleHeight     =   346
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   2
      Top             =   5280
      Width           =   5250
   End
   Begin VB.PictureBox PicXY 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   5250
      Left            =   5280
      ScaleHeight     =   346
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   346
      TabIndex        =   1
      Top             =   0
      Width           =   5250
   End
   Begin VB.PictureBox PicXYZ 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   5370
      Left            =   -120
      ScaleHeight     =   354
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   354
      TabIndex        =   0
      Top             =   -120
      Width           =   5370
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Index           =   0
      Begin VB.Menu New 
         Caption         =   "New"
         Index           =   1
      End
      Begin VB.Menu Open 
         Caption         =   "Open"
         Index           =   2
      End
      Begin VB.Menu Save 
         Caption         =   "Save"
         Index           =   3
      End
      Begin VB.Menu Quit 
         Caption         =   "Quit"
         Index           =   4
      End
   End
   Begin VB.Menu Options 
      Caption         =   "Options"
      Index           =   5
      Begin VB.Menu SetupMatrix 
         Caption         =   "Matrix Setup"
         Index           =   6
      End
   End
   Begin VB.Menu Objects 
      Caption         =   "Objects"
      Index           =   7
      Begin VB.Menu AddTriangle 
         Caption         =   "Add Triangle"
         Index           =   8
      End
      Begin VB.Menu DeleteTriangle 
         Caption         =   "Delete Triangle"
         Index           =   9
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SXP As Single
Dim SYP As Single
Dim SP As Boolean

Dim XYPos As Coordinates2D
Dim XZPos As Coordinates2D
Dim YZPos As Coordinates2D

Dim FillMode As FillMode

Dim CameraPosSelected As Boolean
Dim CameraViewSelected As Boolean

Dim WindowFocus As Integer

Private HHdc As Long

Dim RotObj As Boolean

Private Sub MatrixSetup_Click(Index As Integer)
 MatrixSetup.Show
End Sub

Private Sub AddTriangle_Click(Index As Integer)
 
 If Not Mesh.Triangles = 0 Then
  Dim Trigs() As ObjectTriangle
  ReDim Trigs(Mesh.Triangles)
 
  Dim I As Integer
  For I = 1 To Mesh.Triangles
   Trigs(I) = Mesh.Triangle(I)
  Next
 
  Dim CT As Integer
 
  Mesh.Triangles = Mesh.Triangles + 1
  CT = Mesh.Triangles
 
  ReDim Mesh.Triangle(CT)
 
 
  I = 0
  For I = 1 To CT - 1
   Mesh.Triangle(I) = Trigs(I)
  Next
 
 
  Mesh.Triangle(CT) = ResetTriangle()
 
  Mesh.Triangle(CT).Coordinates(1).X = 15
  Mesh.Triangle(CT).Coordinates(1).Y = 25
  Mesh.Triangle(CT).Coordinates(1).Z = 15


  Mesh.Triangle(CT).Coordinates(2).X = 15
  Mesh.Triangle(CT).Coordinates(2).Y = 15
  Mesh.Triangle(CT).Coordinates(2).Z = 5


  Mesh.Triangle(CT).Coordinates(3).X = 5
  Mesh.Triangle(CT).Coordinates(3).Y = 15
  Mesh.Triangle(CT).Coordinates(3).Z = 15
 Else
  ReDim Mesh.Triangle(1)
  
  Mesh.Triangle(1) = ResetTriangle()
 
  Mesh.Triangle(1).Coordinates(1).X = 15
  Mesh.Triangle(1).Coordinates(1).Y = 25
  Mesh.Triangle(1).Coordinates(1).Z = 15


  Mesh.Triangle(1).Coordinates(2).X = 15
  Mesh.Triangle(1).Coordinates(2).Y = 15
  Mesh.Triangle(1).Coordinates(2).Z = 5


  Mesh.Triangle(1).Coordinates(3).X = 5
  Mesh.Triangle(1).Coordinates(3).Y = 15
  Mesh.Triangle(1).Coordinates(3).Z = 15
  
  DoEvents
  Mesh.Triangles = 1
 End If
 DoEvents
 Update2DPictures
 Update3DPictures
End Sub

Private Sub DeleteTriangle_Click(Index As Integer)
 Dim I As Integer
 
 If Not Trig = 0 Then
  If Mesh.Triangles = 1 Then
   ReDim Mesh.Triangle(0)
   Mesh.Triangles = 0
  ElseIf Mesh.Triangles = 0 Then
    
  ElseIf Mesh.Triangles = Mesh.Triangles Then
   Dim Trigs() As ObjectTriangle
   ReDim Trigs(Mesh.Triangles - 1)
   
   For I = 1 To Mesh.Triangles - 1
    If Not I = Trig Then
     Trigs(I) = Mesh.Triangle(I)
    End If
   Next
   
   Trigs(Trig - 1) = Mesh.Triangle(Mesh.Triangles)
   
   I = 0
   ReDim Mesh.Triangle(Mesh.Triangles - 1)
   Mesh.Triangles = Mesh.Triangles - 1
   For I = 1 To Mesh.Triangles
    Mesh.Triangle(I) = Trigs(I)
   Next
  Else
   Dim Trigss() As ObjectTriangle
   ReDim Trigss(Mesh.Triangles - 1)
   
   For I = 1 To Mesh.Triangles - 1
    If Not I = Trig Then
     Trigss(I) = Mesh.Triangle(I)
    End If
   Next
    
   Trigss(Trig) = Mesh.Triangle(Mesh.Triangles)
   
   I = 0
   ReDim Mesh.Triangle(Mesh.Triangles - 1)
   Mesh.Triangles = Mesh.Triangles - 1
   For I = 1 To Mesh.Triangles
    Mesh.Triangle(I) = Trigss(I)
   Next
  End If
 Trig = 0
 End If
End Sub

Private Sub Form_Load()
 FillMode = Wireframe

' HHdc = CreateHdc(350, 350)
' DoEvents
' HHdc = GetCurrentHdc()
 
 Cam = ResetCamera
 Cam.MaxScreen.X = 350
 Cam.MaxScreen.Y = 350
 Mesh = ResetMesh
 
 DoEvents
 
End Sub

Private Sub Form_Resize()
    Dim NW As Long
    Dim NH As Long
    
    NW = Me.Width / Screen.TwipsPerPixelX
    NH = Me.Height / Screen.TwipsPerPixelY
    
    PicXYZ.Left = 0
    PicXYZ.Top = 0
    
    PicXYZ.Width = NW / 2
    PicXYZ.Height = NH / 2

    PicXY.Left = NW / 2 + 2
    PicXY.Top = 0
    
    PicXY.Width = NW / 2 - 2
    PicXY.Height = NH / 2
    
    PicXZ.Left = 0
    PicXZ.Top = NH / 2 + 2
    
    PicXZ.Width = NW / 2
    PicXZ.Height = NH / 2
    
    PicYZ.Left = NW / 2 + 2
    PicYZ.Top = NH / 2 + 2
    
    PicYZ.Width = NW / 2
    PicYZ.Height = NH / 2
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
 DeleteHdc HHdc
 CleanUp
 End
End Sub

Private Sub New_Click(Index As Integer)
 CleanUp
End Sub

Private Sub Open_Click(Index As Integer)
 CommonDialog.Filter = "Model File|*GBS"
 CommonDialog.ShowOpen
 If Not CommonDialog.FileName = "" Then
  Mesh = OpenModelFile(CommonDialog.FileName)
 End If
End Sub

Private Sub PicXY_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 77 Then
  TriangleOptions.Show
 End If
 
 If Not Trig = 0 Then
  If KeyCode = 38 Then
   Mesh.Triangle(Trig).Coordinates(Coord).Y = Mesh.Triangle(Trig).Coordinates(Coord).Y + 1
  End If
  If KeyCode = 40 Then
   Mesh.Triangle(Trig).Coordinates(Coord).Y = Mesh.Triangle(Trig).Coordinates(Coord).Y - 1
  End If
  If KeyCode = 39 Then
   Mesh.Triangle(Trig).Coordinates(Coord).X = Mesh.Triangle(Trig).Coordinates(Coord).X + 1
  End If
  If KeyCode = 37 Then
   Mesh.Triangle(Trig).Coordinates(Coord).X = Mesh.Triangle(Trig).Coordinates(Coord).X - 1
  End If
 ElseIf CameraPosSelected = True Then
  If KeyCode = 38 Then
   Cam.Position.Y = Cam.Position.Y + 1
  End If
  If KeyCode = 40 Then
   Cam.Position.Y = Cam.Position.Y - 1
  End If
  If KeyCode = 39 Then
   Cam.Position.X = Cam.Position.X + 1
  End If
  If KeyCode = 37 Then
   Cam.Position.X = Cam.Position.X - 1
  End If
 ElseIf CameraViewSelected = True Then
  If KeyCode = 38 Then
   Cam.ViewPosition.Y = Cam.ViewPosition.Y + 1
  End If
  If KeyCode = 40 Then
   Cam.ViewPosition.Y = Cam.ViewPosition.Y - 1
  End If
  If KeyCode = 39 Then
   Cam.ViewPosition.X = Cam.ViewPosition.X + 1
  End If
  If KeyCode = 37 Then
   Cam.ViewPosition.X = Cam.ViewPosition.X - 1
  End If
 End If
 
 If KeyCode = 78 Then
  SP = True
  If Mesh.Triangles = 0 Then
   SP = False
  ElseIf Trig = 0 Then
   Trig = 1
   Coord = 1
  ElseIf Coord = 3 Then
   If Trig = Mesh.Triangles Then
    Trig = 1
    Coord = 1
   Else
    Trig = Trig + 1
    Coord = 1
   End If
  Else
   Coord = Coord + 1
  End If
 End If
 
 If KeyCode = 80 Then
  SP = True
  If Trig = 0 Then
   Trig = Mesh.Triangles
   Coord = 3
  ElseIf Coord = 1 Then
   If Trig = 1 Then
    Trig = Mesh.Triangles
    Coord = 3
   Else
    Trig = Trig - 1
    Coord = 1
   End If
  ElseIf Mesh.Triangles = 0 Then
   SP = False
  Else
   Coord = Coord - 1
  End If
 End If
 
 
 If KeyCode = 45 Then
  Call AddTriangle_Click(1)
 End If

 Dim I As Integer
 If KeyCode = 46 Then
  If Not Trig = 0 Then
   If Mesh.Triangles = 1 Then
    ReDim Mesh.Triangle(0)
    Mesh.Triangles = 0
   ElseIf Mesh.Triangles = 0 Then
    
   ElseIf Mesh.Triangles = Mesh.Triangles Then
    Dim Trigs() As ObjectTriangle
    ReDim Trigs(Mesh.Triangles - 1)
   
    For I = 1 To Mesh.Triangles - 1
     If Not I = Trig Then
      Trigs(I) = Mesh.Triangle(I)
     End If
    Next
    
    Trigs(Trig - 1) = Mesh.Triangle(Mesh.Triangles)
   
    I = 0
    ReDim Mesh.Triangle(Mesh.Triangles - 1)
    Mesh.Triangles = Mesh.Triangles - 1
    For I = 1 To Mesh.Triangles
     Mesh.Triangle(I) = Trigs(I)
    Next
   Else
    Dim Trigss() As ObjectTriangle
    ReDim Trigss(Mesh.Triangles - 1)
   
    For I = 1 To Mesh.Triangles - 1
     If Not I = Trig Then
      Trigss(I) = Mesh.Triangle(I)
     End If
    Next
    
    Trigss(Trig) = Mesh.Triangle(Mesh.Triangles)
   
    I = 0
    ReDim Mesh.Triangle(Mesh.Triangles - 1)
    Mesh.Triangles = Mesh.Triangles - 1
    For I = 1 To Mesh.Triangles
     Mesh.Triangle(I) = Trigss(I)
    Next
   End If
  Trig = 0
  End If
 End If
End Sub

Private Sub PicXY_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 WindowFocus = 2
 
 Dim PX As Single
 Dim PY As Single
 PX = X + XYPos.X
 PY = Y + XYPos.Y
  
 If Button = 1 Then
  If Not WindowFocus = 2 Then
   WindowFocus = 2
   SP = False
  Else
   If SP = True Then
    If CameraPosSelected = True Then
     Cam.Position.X = Round((PX - 20) / 10)
     Cam.Position.Y = Round((PY - 20) / 10)
    ElseIf CameraViewSelected = True Then
     Cam.ViewPosition.X = Round((PX - 20) / 10)
     Cam.ViewPosition.Y = Round((PY - 20) / 10)
    Else
     If Not Trig = 0 Then
      Mesh.Triangle(Trig).Coordinates(Coord).X = Round((PX - 20) / 10)
      Mesh.Triangle(Trig).Coordinates(Coord).Y = Round((PY - 20) / 10)
     End If
    End If
   End If
  End If
 End If
 
 If Button = 0 Then
 
  Dim I As Integer
 
  Dim SD As Double
  Dim CSD As Double
 
  SD = 20
 
  For I = 1 To Mesh.Triangles
   
    CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Mesh.Triangle(I).Coordinates(1).X * 10) + 20, (Mesh.Triangle(I).Coordinates(1).Y * 10) + 20))
    If SD > CSD Then
     SD = CSD
     Trig = I
     Coord = 1
    
    End If
   
    CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Mesh.Triangle(I).Coordinates(2).X * 10) + 20, (Mesh.Triangle(I).Coordinates(2).Y * 10) + 20))
    If SD > CSD Then
     SD = CSD
     Trig = I
     Coord = 2
    End If
   
    CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Mesh.Triangle(I).Coordinates(3).X * 10) + 20, (Mesh.Triangle(I).Coordinates(3).Y * 10) + 20))
    If SD > CSD Then
     SD = CSD
     Trig = I
     Coord = 3
    End If
   
  Next
 
   CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Cam.Position.X * 10) + 20, (Cam.Position.Y * 10) + 20))
   
   If Not Trig = 0 Then
    If CSD < SD Then
     CameraPosSelected = True
     CameraViewSelected = False
     SD = CSD
    Else
     CameraPosSelected = False
    End If
   Else
    If CSD < 20 Then
     CameraPosSelected = True
     CameraViewSelected = False
     SD = CSD
    End If
   End If
   
   CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Cam.ViewPosition.X * 10) + 20, (Cam.ViewPosition.Y * 10) + 20))
   
   If Not Trig = 0 Then
    If CSD < SD Then
     CameraViewSelected = True
     CameraPosSelected = False
     SD = CSD
    Else
     CameraViewSelected = False
    End If
   Else
    If CSD < 20 Then
     CameraViewSelected = True
     CameraPosSelected = False
     SD = CSD
    End If
   End If

  If CameraPosSelected = True Then
   Trig = 0
   Coord = 0
  End If
  
  If CameraViewSelected = True Then
   Trig = 0
   Coord = 0
  End If
    
  If Not Trig = 0 Then
   SXP = (CInt(Mesh.Triangle(Trig).Coordinates(Coord).X) * 10) + 20
   SYP = (CInt(Mesh.Triangle(Trig).Coordinates(Coord).Y) * 10) + 20
   SP = True
  ElseIf CameraPosSelected = True Then
   SP = True
  ElseIf CameraViewSelected = True Then
   SP = True
  ElseIf TriangleOriginSelected = True Then
   SP = True
  Else
   SXP = 0
   SYP = 0
   SP = False
  End If
  
 End If
 If Button = 2 Then
  If X < 15 Then
   XYPos.X = XYPos.X - 10
  ElseIf X > 285 Then
   XYPos.X = XYPos.X + 10
  End If
  
  If Y < 15 Then
   XYPos.Y = XYPos.Y - 10
  ElseIf Y > 285 Then
   XYPos.Y = XYPos.Y + 10
  End If
 End If
End Sub

Private Sub PicXY_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If WindowFocus = 2 Then
  SP = False
  Trig = 0
  Coord = 0
 End If
End Sub

Private Sub PicXYZ_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = 79 Then
  If RotObj = False Then
   RotObj = True
  Else
   RotObj = False
  End If
 End If
 
 Static AngleX As Single
 Static AngleY As Single
 Dim AngPos As Coordinates2D
   
 If KeyCode = 39 Then
  If AngleX < 350 Then
   AngleX = AngleX + 0.2
  Else
   AngleX = 0
  End If
  If RotObj = False Then
   AngPos = VectorAngleToPosition2D(Cam.Position.X, Cam.Position.Y, CSng(AngleX), 10)
   Cam.ViewPosition.X = AngPos.X
   Cam.ViewPosition.Y = AngPos.Y
  Else
   AngPos = VectorAngleToPosition2D(Cam.ViewPosition.X, Cam.ViewPosition.Y, CSng(AngleX), 10)
   Cam.Position.X = AngPos.X
   Cam.Position.Y = AngPos.Y
  End If
 End If
 
 If KeyCode = 37 Then
  If AngleX > 10 Then
   AngleX = AngleX - 0.2
  Else
   AngleX = 360
  End If
  If RotObj = False Then
   AngPos = VectorAngleToPosition2D(Cam.Position.X, Cam.Position.Y, CSng(AngleX), 10)
   Cam.ViewPosition.X = AngPos.X
   Cam.ViewPosition.Y = AngPos.Y
  Else
   AngPos = VectorAngleToPosition2D(Cam.ViewPosition.X, Cam.ViewPosition.Y, CSng(AngleX), 10)
   Cam.Position.X = AngPos.X
   Cam.Position.Y = AngPos.Y
  End If
 End If
 
 If KeyCode = 38 Then
  If AngleY < 350 Then
   AngleY = AngleY + 0.2
  Else
   AngleY = 0
  End If
  If RotObj = False Then
   AngPos = VectorAngleToPosition2D(Cam.Position.Y, Cam.Position.Z, CSng(AngleY), 10)
   Cam.ViewPosition.Y = AngPos.X
   Cam.ViewPosition.Z = AngPos.Y
  Else
   AngPos = VectorAngleToPosition2D(Cam.ViewPosition.Y, Cam.ViewPosition.Z, CSng(AngleY), 10)
   Cam.Position.Y = AngPos.X
   Cam.Position.Z = AngPos.Y
  End If
 End If
 
 If KeyCode = 40 Then
  If AngleX > 10 Then
   AngleY = AngleY - 0.2
  Else
   AngleY = 360
  End If
  If RotObj = False Then
   AngPos = VectorAngleToPosition2D(Cam.Position.Y, Cam.Position.Z, CSng(AngleY), 10)
   Cam.ViewPosition.Y = AngPos.X
   Cam.ViewPosition.Z = AngPos.Y
  Else
   AngPos = VectorAngleToPosition2D(Cam.ViewPosition.Y, Cam.ViewPosition.Z, CSng(AngleY), 10)
   Cam.Position.Y = AngPos.X
   Cam.Position.Z = AngPos.Y
  End If
 End If
 
 Dim pa As Coordinates2D
 
 If KeyCode = 32 Then
  pa = VectorAngleToPosition2D(Cam.Position.Y, Cam.Position.Z, AngleY, 2)
  Cam.Position.Z = pa.Y
 End If
 
 If KeyCode = 17 Then
  pa = VectorAngleToPosition2D(Cam.Position.Y, Cam.Position.Z, AngleY, -2)
  Cam.Position.Z = pa.Y
 End If
 
 If KeyCode = 68 Then
  pa = VectorAngleToPosition2D(Cam.Position.X, Cam.Position.Y, AngleX, 2)
  Cam.Position.X = pa.X
 End If
 
 If KeyCode = 65 Then
  pa = VectorAngleToPosition2D(Cam.Position.X, Cam.Position.Y, AngleX, -2)
  Cam.Position.X = pa.X
 End If
 
 If KeyCode = 87 Then
  Cam.Position.Y = Cam.Position.Y + 2
 End If
 
 If KeyCode = 83 Then
  Cam.Position.Y = Cam.Position.Y - 2
 End If
 
 If KeyCode = 107 Then
  If Cam.ScaleSize < 10 Then
   Cam.ScaleSize = Cam.ScaleSize + 10
  ElseIf Cam.ScaleSize > -10 Then
   Cam.ScaleSize = Cam.ScaleSize + 10
  Else
   Cam.ScaleSize = Cam.ScaleSize + 10
  End If
 End If
 
 If KeyCode = 109 Then
  If Cam.ScaleSize < 10 Then
   Cam.ScaleSize = Cam.ScaleSize - 10
  ElseIf Cam.ScaleSize > -10 Then
   Cam.ScaleSize = Cam.ScaleSize - 10
  Else
   Cam.ScaleSize = Cam.ScaleSize - 10
  End If
 End If
 
 Dim A As Long
 Dim B As Long
 
 If KeyCode = 82 Then
  If FillMode = Wireframe Then
   FillMode = Solid
  Else
   FillMode = Wireframe
  End If
 End If
 Update2DPictures
End Sub

Private Sub PicXYZ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
  If Not WindowFocus = 1 Then
   WindowFocus = 1
   SP = False
  End If
 End If
 
 If Button = 0 Then WindowFocus = 1
End Sub

Private Sub PicXZ_KeyDown(KeyCode As Integer, Shift As Integer)
 
 If KeyCode = 77 Then
  TriangleOptions.Show
 End If
 
 If Not Trig = 0 Then
  If KeyCode = 38 Then
   Mesh.Triangle(Trig).Coordinates(Coord).Z = Mesh.Triangle(Trig).Coordinates(Coord).Z + 1
  End If
  If KeyCode = 40 Then
   Mesh.Triangle(Trig).Coordinates(Coord).Z = Mesh.Triangle(Trig).Coordinates(Coord).Z - 1
  End If
  If KeyCode = 39 Then
   Mesh.Triangle(Trig).Coordinates(Coord).X = Mesh.Triangle(Trig).Coordinates(Coord).X + 1
  End If
  If KeyCode = 37 Then
   Mesh.Triangle(Trig).Coordinates(Coord).X = Mesh.Triangle(Trig).Coordinates(Coord).X - 1
  End If
 ElseIf CameraPosSelected = True Then
  If KeyCode = 38 Then
   Cam.Position.Z = Cam.Position.Z + 1
  End If
  If KeyCode = 40 Then
   Cam.Position.Z = Cam.Position.Z - 1
  End If
  If KeyCode = 39 Then
   Cam.Position.X = Cam.Position.X + 1
  End If
  If KeyCode = 37 Then
   Cam.Position.X = Cam.Position.X - 1
  End If
 ElseIf CameraViewSelected = True Then
  If KeyCode = 38 Then
   Cam.ViewPosition.Z = Cam.ViewPosition.Z + 1
  End If
  If KeyCode = 40 Then
   Cam.ViewPosition.Z = Cam.ViewPosition.Z - 1
  End If
  If KeyCode = 39 Then
   Cam.ViewPosition.X = Cam.ViewPosition.X + 1
  End If
  If KeyCode = 37 Then
   Cam.ViewPosition.X = Cam.ViewPosition.X - 1
  End If
 End If
 
 If KeyCode = 78 Then
  SP = True
  If Mesh.Triangles = 0 Then
   SP = False
  ElseIf Trig = 0 Then
   Trig = 1
   Coord = 1
  ElseIf Coord = 3 Then
   If Trig = Mesh.Triangles Then
    Trig = 1
    Coord = 1
   Else
    Trig = Trig + 1
    Coord = 1
   End If
  Else
   Coord = Coord + 1
  End If
 End If
 
 If KeyCode = 80 Then
  SP = True
  If Trig = 0 Then
   Trig = Mesh.Triangles
   Coord = 3
  ElseIf Coord = 1 Then
   If Trig = 1 Then
    Trig = Mesh.Triangles
    Coord = 3
   Else
    Trig = Trig - 1
    Coord = 1
   End If
  ElseIf Mesh.Triangles = 0 Then
   SP = False
  Else
   Coord = Coord - 1
  End If
 End If
 
 
 If KeyCode = 45 Then
  Call AddTriangle_Click(1)
 End If
 
 Dim I As Integer
 If KeyCode = 46 Then
  If Not Trig = 0 Then
   If Mesh.Triangles = 1 Then
    ReDim Mesh.Triangle(0)
    Mesh.Triangles = 0
   ElseIf Mesh.Triangles = 0 Then
    
   ElseIf Mesh.Triangles = Mesh.Triangles Then
    Dim Trigs() As ObjectTriangle
    ReDim Trigs(Mesh.Triangles - 1)
   
    For I = 1 To Mesh.Triangles - 1
     If Not I = Trig Then
      Trigs(I) = Mesh.Triangle(I)
     End If
    Next
    
    Trigs(Trig - 1) = Mesh.Triangle(Mesh.Triangles)
   
    I = 0
    ReDim Mesh.Triangle(Mesh.Triangles - 1)
    Mesh.Triangles = Mesh.Triangles - 1
    For I = 1 To Mesh.Triangles
     Mesh.Triangle(I) = Trigs(I)
    Next
   Else
    Dim Trigss() As ObjectTriangle
    ReDim Trigss(Mesh.Triangles - 1)
   
    For I = 1 To Mesh.Triangles - 1
     If Not I = Trig Then
      Trigss(I) = Mesh.Triangle(I)
     End If
    Next
    
    Trigss(Trig) = Mesh.Triangle(Mesh.Triangles)
   
    I = 0
    ReDim Mesh.Triangle(Mesh.Triangles - 1)
    Mesh.Triangles = Mesh.Triangles - 1
    For I = 1 To Mesh.Triangles
     Mesh.Triangle(I) = Trigss(I)
    Next
   End If
  Trig = 0
  End If
 End If
End Sub

Private Sub PicXZ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 WindowFocus = 2
 
 Dim PX As Single
 Dim PY As Single
 PX = X + XZPos.X
 PY = Y + XZPos.Y
  
 If Button = 1 Then
  If Not WindowFocus = 2 Then
   WindowFocus = 2
   SP = False
  Else
   If SP = True Then
    If CameraPosSelected = True Then
     Cam.Position.X = Round((PX - 20) / 10)
     Cam.Position.Z = Round((PY - 20) / 10)
    ElseIf CameraViewSelected = True Then
     Cam.ViewPosition.X = Round((PX - 20) / 10)
     Cam.ViewPosition.Z = Round((PY - 20) / 10)
    Else
     Mesh.Triangle(Trig).Coordinates(Coord).X = Round((PX - 20) / 10)
     Mesh.Triangle(Trig).Coordinates(Coord).Z = Round((PY - 20) / 10)
    End If
   End If
  End If
 End If
 
 If Button = 0 Then
 
  Dim I As Integer
 
  Dim SD As Double
  Dim CSD As Double
 
  SD = 20
 
  For I = 1 To Mesh.Triangles
   
    CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Mesh.Triangle(I).Coordinates(1).X * 10) + 20, (Mesh.Triangle(I).Coordinates(1).Z * 10) + 20))
    If SD > CSD Then
     SD = CSD
     Trig = I
     Coord = 1
    
    End If
   
    CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Mesh.Triangle(I).Coordinates(2).X * 10) + 20, (Mesh.Triangle(I).Coordinates(2).Z * 10) + 20))
    If SD > CSD Then
     SD = CSD
     Trig = I
     Coord = 2
    End If
   
    CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Mesh.Triangle(I).Coordinates(3).X * 10) + 20, (Mesh.Triangle(I).Coordinates(3).Z * 10) + 20))
    If SD > CSD Then
     SD = CSD
     Trig = I
     Coord = 3
    End If
   
  Next
 
   CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Cam.Position.X * 10) + 20, (Cam.Position.Z * 10) + 20))
   
   If Not Trig = 0 Then
    If CSD < SD Then
     CameraPosSelected = True
     CameraViewSelected = False
     SD = CSD
    Else
     CameraPosSelected = False
    End If
   Else
    If CSD < 20 Then
     CameraPosSelected = True
     CameraViewSelected = False
     SD = CSD
    End If
   End If
   
   CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Cam.ViewPosition.X * 10) + 20, (Cam.ViewPosition.Z * 10) + 20))
   
   If Not Trig = 0 Then
    If CSD < SD Then
     CameraViewSelected = True
     CameraPosSelected = False
     SD = CSD
    Else
     CameraViewSelected = False
    End If
   Else
    If CSD < 20 Then
     CameraViewSelected = True
     CameraPosSelected = False
     SD = CSD
    End If
   End If
   
  If CameraPosSelected = True Then
    Trig = 0
    Coord = 0
  End If
  
  If CameraViewSelected = True Then
    Trig = 0
    Coord = 0
  End If
  
  If Not Trig = 0 Then
   SXP = (CInt(Mesh.Triangle(Trig).Coordinates(Coord).X) * 10) + 20
   SYP = (CInt(Mesh.Triangle(Trig).Coordinates(Coord).Z) * 10) + 20
   SP = True
  ElseIf CameraPosSelected = True Then
   SP = True
  ElseIf CameraViewSelected = True Then
   SP = True
  Else
   SXP = 0
   SYP = 0
   SP = False
  End If
 
 End If
 
 If Button = 2 Then
  If X < 15 Then
    XZPos.X = XZPos.X - 10
   ElseIf X > 285 Then
    XZPos.X = XZPos.X + 10
   End If
  
   If Y < 15 Then
    XZPos.Y = XZPos.Y - 10
   ElseIf Y > 285 Then
    XZPos.Y = XZPos.Y + 10
   End If
  End If
End Sub

Private Sub PicXZ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If WindowFocus = 3 Then
  SP = False
  Trig = 0
  Coord = 0
 End If
End Sub

Private Sub PicYZ_KeyDown(KeyCode As Integer, Shift As Integer)

 If KeyCode = 77 Then
  TriangleOptions.Show
 End If
 
 If Not Trig = 0 Then
  If KeyCode = 38 Then
   Mesh.Triangle(Trig).Coordinates(Coord).Z = Mesh.Triangle(Trig).Coordinates(Coord).Z + 1
  End If
  If KeyCode = 40 Then
   Mesh.Triangle(Trig).Coordinates(Coord).Z = Mesh.Triangle(Trig).Coordinates(Coord).Z - 1
  End If
  If KeyCode = 39 Then
   Mesh.Triangle(Trig).Coordinates(Coord).Y = Mesh.Triangle(Trig).Coordinates(Coord).Y + 1
  End If
  If KeyCode = 37 Then
   Mesh.Triangle(Trig).Coordinates(Coord).Y = Mesh.Triangle(Trig).Coordinates(Coord).Y - 1
  End If
 ElseIf CameraPosSelected = True Then
  If KeyCode = 38 Then
   Cam.Position.Z = Cam.Position.Z + 1
  End If
  If KeyCode = 40 Then
   Cam.Position.Z = Cam.Position.Z - 1
  End If
  If KeyCode = 39 Then
   Cam.Position.Y = Cam.Position.Y + 1
  End If
  If KeyCode = 37 Then
   Cam.Position.Y = Cam.Position.Y - 1
  End If
 ElseIf CameraViewSelected = True Then
  If KeyCode = 38 Then
   Cam.ViewPosition.Z = Cam.ViewPosition.Z + 1
  End If
  If KeyCode = 40 Then
   Cam.ViewPosition.Z = Cam.ViewPosition.Z - 1
  End If
  If KeyCode = 39 Then
   Cam.ViewPosition.Y = Cam.ViewPosition.Y + 1
  End If
  If KeyCode = 37 Then
   Cam.ViewPosition.Y = Cam.ViewPosition.Y - 1
  End If
 End If
 
 If KeyCode = 78 Then
  SP = True
  If Mesh.Triangles = 0 Then
   SP = False
  ElseIf Trig = 0 Then
   Trig = 1
   Coord = 1
  ElseIf Coord = 3 Then
   If Trig = Mesh.Triangles Then
    Trig = 1
    Coord = 1
   Else
    Trig = Trig + 1
    Coord = 1
   End If
  Else
   Coord = Coord + 1
  End If
 End If
 
 If KeyCode = 80 Then
  SP = True
  If Trig = 0 Then
   Trig = Mesh.Triangles
   Coord = 3
  ElseIf Coord = 1 Then
   If Trig = 1 Then
    Trig = Mesh.Triangles
    Coord = 3
   Else
    Trig = Trig - 1
    Coord = 1
   End If
  ElseIf Mesh.Triangles = 0 Then
   SP = False
  Else
   Coord = Coord - 1
  End If
 End If
 
 If KeyCode = 45 Then
  Call AddTriangle_Click(1)
 End If
 
 Dim I As Integer
 If KeyCode = 46 Then
  If Not Trig = 0 Then
   If Mesh.Triangles = 1 Then
    ReDim Mesh.Triangle(0)
    Mesh.Triangles = 0
   ElseIf Mesh.Triangles = 0 Then
    
   ElseIf Mesh.Triangles = Mesh.Triangles Then
    Dim Trigs() As ObjectTriangle
    ReDim Trigs(Mesh.Triangles - 1)
   
    For I = 1 To Mesh.Triangles - 1
     If Not I = Trig Then
      Trigs(I) = Mesh.Triangle(I)
     End If
    Next
    
    Trigs(Trig - 1) = Mesh.Triangle(Mesh.Triangles)
   
    I = 0
    ReDim Mesh.Triangle(Mesh.Triangles - 1)
    Mesh.Triangles = Mesh.Triangles - 1
    For I = 1 To Mesh.Triangles
     Mesh.Triangle(I) = Trigs(I)
    Next
   Else
    Dim Trigss() As ObjectTriangle
    ReDim Trigss(Mesh.Triangles - 1)
   
    For I = 1 To Mesh.Triangles - 1
     If Not I = Trig Then
      Trigss(I) = Mesh.Triangle(I)
     End If
    Next
    
    Trigss(Trig) = Mesh.Triangle(Mesh.Triangles)
   
    I = 0
    ReDim Mesh.Triangle(Mesh.Triangles - 1)
    Mesh.Triangles = Mesh.Triangles - 1
    For I = 1 To Mesh.Triangles
     Mesh.Triangle(I) = Trigss(I)
    Next
   End If
  Trig = 0
  End If
 End If
End Sub

Private Sub PicYZ_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 WindowFocus = 4
 
 Dim PX As Single
 Dim PY As Single
 PX = X + YZPos.X
 PY = Y + YZPos.Y
  
 If Button = 1 Then
  If Not WindowFocus = 4 Then
   WindowFocus = 4
   SP = False
  Else
   If SP = True Then
    If CameraPosSelected = True Then
     Cam.Position.Y = Round((PX - 20) / 10)
     Cam.Position.Z = Round((PY - 20) / 10)
    ElseIf CameraViewSelected = True Then
     Cam.ViewPosition.Y = Round((PX - 20) / 10)
     Cam.ViewPosition.Z = Round((PY - 20) / 10)
    Else
     Mesh.Triangle(Trig).Coordinates(Coord).Y = Round((PX - 20) / 10)
     Mesh.Triangle(Trig).Coordinates(Coord).Z = Round((PY - 20) / 10)
    End If
   End If
  End If
 End If
 
 If Button = 0 Then
 
  Dim I As Integer
 
  Dim SD As Double
  Dim CSD As Double
 
  SD = 20
 
  For I = 1 To Mesh.Triangles
   
    CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Mesh.Triangle(I).Coordinates(1).Y * 10) + 20, (Mesh.Triangle(I).Coordinates(1).Z * 10) + 20))
    If SD > CSD Then
     SD = CSD
     Trig = I
     Coord = 1
    
    End If
   
    CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Mesh.Triangle(I).Coordinates(2).Y * 10) + 20, (Mesh.Triangle(I).Coordinates(2).Z * 10) + 20))
    If SD > CSD Then
     SD = CSD
     Trig = I
     Coord = 2
    End If
   
    CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Mesh.Triangle(I).Coordinates(3).Y * 10) + 20, (Mesh.Triangle(I).Coordinates(3).Z * 10) + 20))
    If SD > CSD Then
     SD = CSD
     Trig = I
     Coord = 3
    End If

  Next
 
   CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Cam.Position.Y * 10) + 20, (Cam.Position.Z * 10) + 20))
   
   If Not Trig = 0 Then
    If CSD < SD Then
     CameraPosSelected = True
     CameraViewSelected = False
     SD = CSD
    Else
     CameraPosSelected = False
    End If
   Else
    If CSD < 20 Then
     CameraPosSelected = True
     CameraViewSelected = False
     SD = CSD
    End If
   End If
   
   CSD = VectorDistance2D(Make2DCoordinate(PX, PY), Make2DCoordinate((Cam.ViewPosition.Y * 10) + 20, (Cam.ViewPosition.Z * 10) + 20))
   
   If Not Trig = 0 Then
    If CSD < SD Then
     CameraViewSelected = True
     CameraPosSelected = False
     SD = CSD
    Else
     CameraViewSelected = False
    End If
   Else
    If CSD < 20 Then
     CameraViewSelected = True
     CameraPosSelected = False
     SD = CSD
    End If
   End If
   
  If CameraPosSelected = True Then
    Trig = 0
    Coord = 0
  End If
  
  If CameraViewSelected = True Then
    Trig = 0
    Coord = 0
  End If
  
  If Not Trig = 0 Then
   SXP = (CInt(Mesh.Triangle(Trig).Coordinates(Coord).Y) * 10) + 20
   SYP = (CInt(Mesh.Triangle(Trig).Coordinates(Coord).Z) * 10) + 20
   SP = True
  ElseIf CameraPosSelected = True Then
   SP = True
  ElseIf CameraViewSelected = True Then
   SP = True
  Else
   SXP = 0
   SYP = 0
   SP = False
  End If
 End If
 If Button = 2 Then
   If X < 15 Then
    YZPos.X = YZPos.X - 10
   ElseIf X > 285 Then
    YZPos.X = YZPos.X + 10
   End If
   If Y < 15 Then
    YZPos.Y = YZPos.Y - 10
   ElseIf Y > 285 Then
    YZPos.Y = YZPos.Y + 10
   End If
  End If
End Sub

Private Sub PicYZ_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If WindowFocus = 4 Then
  SP = False
  Trig = 0
  Coord = 0
 End If
End Sub

Private Sub Quit_Click(Index As Integer)
 DeleteHdc HHdc
 CleanUp
 End
End Sub

Private Sub Save_Click(Index As Integer)
 CommonDialog.Filter = "Model File|*GBS"
 CommonDialog.ShowSave
 If Not CommonDialog.FileName = "" Then
  If Not LCase(Right(CommonDialog.FileName, 4)) = ".gbs" Then CommonDialog.FileName = CommonDialog.FileName & ".gbs"
  SaveModelFile CommonDialog.FileName, Mesh
 End If
End Sub

Private Sub SetupMatrix_Click(Index As Integer)
 Load MatrixSetup
 MatrixSetup.Show
End Sub

Private Sub Timer1_Timer()

End Sub

Private Sub UpdateTimer_Timer()
  Update2DPictures
  Update3DPictures
End Sub

Sub Update2DPictures()
 PicXY.Cls
 PicXZ.Cls
 PicYZ.Cls
 
 PicXY.Forecolor = &H404040
 DrawGrid 35 * PicXY.Width / 350, PicXY.Width, 35 * PicXY.Height / 350, PicXY.Height, PicXY.HDc
 PicXY.Forecolor = &HC0C0C0
 
 PicXZ.Forecolor = &H404040
 DrawGrid 35 * PicXZ.Width / 350, PicXZ.Width, 35 * PicXZ.Height / 350, PicXZ.Height, PicXZ.HDc
 PicXZ.Forecolor = &HC0C0C0
 
 PicYZ.Forecolor = &H404040
 DrawGrid 35 * PicYZ.Width / 350, PicYZ.Width, 35 * PicYZ.Height / 350, PicYZ.Height, PicYZ.HDc
 PicYZ.Forecolor = &HC0C0C0

 PrintText "2D - X, Y: " & XYPos.X / 10 & ", " & XYPos.Y / 10, 1, 1, PicXY.HDc
 PrintText "2D - X, Z: " & XZPos.X / 10 & ", " & XZPos.Y / 10, 1, 1, PicXZ.HDc
 PrintText "2D - Y, Z: " & YZPos.X / 10 & ", " & YZPos.Y / 10, 1, 1, PicYZ.HDc

 Dim I As Integer
 
 For I = 1 To Mesh.Triangles
  If I = Trig Then
   PicXY.Forecolor = &HFF0000         '&HFFFFFF
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(1).X), Int(Mesh.Triangle(I).Coordinates(1).Y), Int(Mesh.Triangle(I).Coordinates(2).X), Int(Mesh.Triangle(I).Coordinates(2).Y), 10, 20 - XYPos.X, 20 - XYPos.Y, PicXY.HDc
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(2).X), Int(Mesh.Triangle(I).Coordinates(2).Y), Int(Mesh.Triangle(I).Coordinates(3).X), Int(Mesh.Triangle(I).Coordinates(3).Y), 10, 20 - XYPos.X, 20 - XYPos.Y, PicXY.HDc
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(3).X), Int(Mesh.Triangle(I).Coordinates(3).Y), Int(Mesh.Triangle(I).Coordinates(1).X), Int(Mesh.Triangle(I).Coordinates(1).Y), 10, 20 - XYPos.X, 20 - XYPos.Y, PicXY.HDc
  Else
   PicXY.Forecolor = &HC0C0C0
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(1).X), Int(Mesh.Triangle(I).Coordinates(1).Y), Int(Mesh.Triangle(I).Coordinates(2).X), Int(Mesh.Triangle(I).Coordinates(2).Y), 10, 20 - XYPos.X, 20 - XYPos.Y, PicXY.HDc
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(2).X), Int(Mesh.Triangle(I).Coordinates(2).Y), Int(Mesh.Triangle(I).Coordinates(3).X), Int(Mesh.Triangle(I).Coordinates(3).Y), 10, 20 - XYPos.X, 20 - XYPos.Y, PicXY.HDc
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(3).X), Int(Mesh.Triangle(I).Coordinates(3).Y), Int(Mesh.Triangle(I).Coordinates(1).X), Int(Mesh.Triangle(I).Coordinates(1).Y), 10, 20 - XYPos.X, 20 - XYPos.Y, PicXY.HDc
  End If
 Next
 
 If CameraPosSelected = True Then
  PicXY.Forecolor = &HFF0000
 Else
  PicXY.Forecolor = &HFF&
 End If
 PicXY.DrawWidth = 6
 PicXY.PSet ((Cam.Position.X * 10) + (17 - XYPos.X), (Cam.Position.Y * 10) + (17 - XYPos.Y))
 PicXY.DrawWidth = 1
 If CameraViewSelected = True Then
  PicXY.Forecolor = &HFF0000
 Else
  PicXY.Forecolor = &HFF&
 End If
 DrawLineScaled Int(Cam.Position.X), Int(Cam.Position.Y), Int(Cam.ViewPosition.X), Int(Cam.ViewPosition.Y), 10, 20 - XYPos.X, 20 - XYPos.Y, PicXY.HDc
 PicXY.DrawWidth = 1
 PicXY.Forecolor = &HC0C0C0
 
 
 
 I = 0
 
 For I = 1 To Mesh.Triangles
  If I = Trig Then
   PicXZ.Forecolor = &HFF0000   '&HFFFFFF
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(1).X), Int(Mesh.Triangle(I).Coordinates(1).Z), Int(Mesh.Triangle(I).Coordinates(2).X), Int(Mesh.Triangle(I).Coordinates(2).Z), 10, 20 - XZPos.X, 20 - XZPos.Y, PicXZ.HDc
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(2).X), Int(Mesh.Triangle(I).Coordinates(2).Z), Int(Mesh.Triangle(I).Coordinates(3).X), Int(Mesh.Triangle(I).Coordinates(3).Z), 10, 20 - XZPos.X, 20 - XZPos.Y, PicXZ.HDc
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(3).X), Int(Mesh.Triangle(I).Coordinates(3).Z), Int(Mesh.Triangle(I).Coordinates(1).X), Int(Mesh.Triangle(I).Coordinates(1).Z), 10, 20 - XZPos.X, 20 - XZPos.Y, PicXZ.HDc
  Else
   PicXZ.Forecolor = &HC0C0C0
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(1).X), Int(Mesh.Triangle(I).Coordinates(1).Z), Int(Mesh.Triangle(I).Coordinates(2).X), Int(Mesh.Triangle(I).Coordinates(2).Z), 10, 20 - XZPos.X, 20 - XZPos.Y, PicXZ.HDc
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(2).X), Int(Mesh.Triangle(I).Coordinates(2).Z), Int(Mesh.Triangle(I).Coordinates(3).X), Int(Mesh.Triangle(I).Coordinates(3).Z), 10, 20 - XZPos.X, 20 - XZPos.Y, PicXZ.HDc
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(3).X), Int(Mesh.Triangle(I).Coordinates(3).Z), Int(Mesh.Triangle(I).Coordinates(1).X), Int(Mesh.Triangle(I).Coordinates(1).Z), 10, 20 - XZPos.X, 20 - XZPos.Y, PicXZ.HDc
  End If
 Next
 
 If CameraPosSelected = True Then
  PicXZ.Forecolor = &HFF0000
 Else
  PicXZ.Forecolor = &HFF&
 End If
 PicXZ.DrawWidth = 6
 PicXZ.PSet ((Cam.Position.X * 10) + (17 - XZPos.X), (Cam.Position.Z * 10) + (17 - XZPos.Y))
 PicXZ.DrawWidth = 1
 If CameraViewSelected = True Then
  PicXZ.Forecolor = &HFF0000
 Else
  PicXZ.Forecolor = &HFF&
 End If
 DrawLineScaled Int(Cam.Position.X), Int(Cam.Position.Z), Int(Cam.ViewPosition.X), Int(Cam.ViewPosition.Z), 10, 20 - XZPos.X, 20 - XZPos.Y, PicXZ.HDc
 PicXZ.DrawWidth = 1
 PicXZ.Forecolor = &HC0C0C0
 
 I = 0
 
 For I = 1 To Mesh.Triangles
  If I = Trig Then
   PicYZ.Forecolor = &HFF0000   '&HFFFFFF
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(1).Y), Int(Mesh.Triangle(I).Coordinates(1).Z), Int(Mesh.Triangle(I).Coordinates(2).Y), Int(Mesh.Triangle(I).Coordinates(2).Z), 10, 20 - YZPos.X, 20 - YZPos.Y, PicYZ.HDc
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(2).Y), Int(Mesh.Triangle(I).Coordinates(2).Z), Int(Mesh.Triangle(I).Coordinates(3).Y), Int(Mesh.Triangle(I).Coordinates(3).Z), 10, 20 - YZPos.X, 20 - YZPos.Y, PicYZ.HDc
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(3).Y), Int(Mesh.Triangle(I).Coordinates(3).Z), Int(Mesh.Triangle(I).Coordinates(1).Y), Int(Mesh.Triangle(I).Coordinates(1).Z), 10, 20 - YZPos.X, 20 - YZPos.Y, PicYZ.HDc
  Else
   PicYZ.Forecolor = &HC0C0C0
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(1).Y), Int(Mesh.Triangle(I).Coordinates(1).Z), Int(Mesh.Triangle(I).Coordinates(2).Y), Int(Mesh.Triangle(I).Coordinates(2).Z), 10, 20 - YZPos.X, 20 - YZPos.Y, PicYZ.HDc
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(2).Y), Int(Mesh.Triangle(I).Coordinates(2).Z), Int(Mesh.Triangle(I).Coordinates(3).Y), Int(Mesh.Triangle(I).Coordinates(3).Z), 10, 20 - YZPos.X, 20 - YZPos.Y, PicYZ.HDc
   DrawLineScaled Int(Mesh.Triangle(I).Coordinates(3).Y), Int(Mesh.Triangle(I).Coordinates(3).Z), Int(Mesh.Triangle(I).Coordinates(1).Y), Int(Mesh.Triangle(I).Coordinates(1).Z), 10, 20 - YZPos.X, 20 - YZPos.Y, PicYZ.HDc
  End If
 Next

 If CameraPosSelected = True Then
  PicYZ.Forecolor = &HFF0000
 Else
  PicYZ.Forecolor = &HFF&
 End If
 PicYZ.DrawWidth = 6
 PicYZ.PSet ((Cam.Position.Y * 10) + (17 - YZPos.X), (Cam.Position.Z * 10) + (17 - YZPos.Y))
 PicYZ.DrawWidth = 1
 
 
 
 If CameraViewSelected = True Then
  PicYZ.Forecolor = &HFF0000
 Else
  PicYZ.Forecolor = &HFF&
 End If
 DrawLineScaled Int(Cam.Position.Y), Int(Cam.Position.Z), Int(Cam.ViewPosition.Y), Int(Cam.ViewPosition.Z), 10, 20 - YZPos.X, 20 - YZPos.Y, PicYZ.HDc
 PicYZ.DrawWidth = 1
 PicYZ.Forecolor = &HC0C0C0
End Sub

Sub Update3DPictures()
 PicXYZ.BackColor = 0
 Backbuffer.BackColor = 0
 
 RenderMesh3D Mesh, Cam, Software, FillMode, Backbuffer.ScaleLeft, Backbuffer.ScaleWidth, Backbuffer.ScaleHeight, Backbuffer.ScaleTop, Backbuffer.HDc
 DrawHdcOnHdc Backbuffer.HDc, PicXYZ.HDc, Backbuffer.Width, Backbuffer.Height, 0, 0, 0, 0
End Sub

