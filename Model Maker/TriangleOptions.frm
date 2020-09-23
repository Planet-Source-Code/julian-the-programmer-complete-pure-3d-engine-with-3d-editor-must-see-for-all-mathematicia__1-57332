VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TriangleOptions 
   Caption         =   "Triangle Options"
   ClientHeight    =   5985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3570
   LinkTopic       =   "Form1"
   ScaleHeight     =   399
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   238
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   120
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox TriangleTexture 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   1920
      Left            =   840
      ScaleHeight     =   124
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   124
      TabIndex        =   23
      Top             =   3600
      Width           =   1920
   End
   Begin VB.TextBox SolidColorText 
      Height          =   285
      Left            =   840
      TabIndex        =   21
      Text            =   "&H00FFFFFF&"
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton ARTButton 
      Caption         =   "Apply Rotation"
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   2040
      Width           =   2775
   End
   Begin VB.CommandButton ATTButton 
      Caption         =   "Apply Translation"
      Height          =   255
      Left            =   360
      TabIndex        =   19
      Top             =   2400
      Width           =   2775
   End
   Begin VB.CommandButton ASTMButton 
      Caption         =   "Apply Scale"
      Height          =   255
      Left            =   360
      TabIndex        =   18
      Top             =   2760
      Width           =   2775
   End
   Begin VB.CommandButton ResetTButton 
      Caption         =   "Reset"
      Height          =   255
      Left            =   360
      TabIndex        =   17
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton SaveButton 
      Caption         =   "Save"
      Height          =   255
      Left            =   240
      TabIndex        =   16
      Top             =   5640
      Width           =   3015
   End
   Begin VB.TextBox T14 
      Height          =   285
      Left            =   2640
      TabIndex        =   15
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox T13 
      Height          =   285
      Left            =   1800
      TabIndex        =   14
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox T12 
      Height          =   285
      Left            =   960
      TabIndex        =   13
      Text            =   "0"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox T11 
      Height          =   285
      Left            =   120
      TabIndex        =   12
      Text            =   "1"
      Top             =   600
      Width           =   735
   End
   Begin VB.TextBox T24 
      Height          =   285
      Left            =   2640
      TabIndex        =   11
      Text            =   "0"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox T23 
      Height          =   285
      Left            =   1800
      TabIndex        =   10
      Text            =   "0"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox T22 
      Height          =   285
      Left            =   960
      TabIndex        =   9
      Text            =   "1"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox T21 
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Text            =   "0"
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox T34 
      Height          =   285
      Left            =   2640
      TabIndex        =   7
      Text            =   "0"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox T33 
      Height          =   285
      Left            =   1800
      TabIndex        =   6
      Text            =   "1"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox T32 
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Text            =   "0"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox T31 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Text            =   "0"
      Top             =   1320
      Width           =   735
   End
   Begin VB.TextBox T44 
      Height          =   285
      Left            =   2640
      TabIndex        =   3
      Text            =   "1"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox T43 
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Text            =   "0"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox T42 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Text            =   "0"
      Top             =   1680
      Width           =   735
   End
   Begin VB.TextBox T41 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "0"
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label SolidColorLabel 
      Caption         =   "Solid Color"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   120
      Width           =   855
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   232
      Y1              =   232
      Y2              =   232
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   32
      Y2              =   232
   End
   Begin VB.Line Line2 
      X1              =   232
      X2              =   232
      Y1              =   32
      Y2              =   232
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   232
      Y1              =   32
      Y2              =   32
   End
End
Attribute VB_Name = "TriangleOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function SetT(Matrix As Matrix4x4)
 T11.Text = Matrix.rc11
 T12.Text = Matrix.rc12
 T13.Text = Matrix.rc13
 T14.Text = Matrix.rc14
 
 T21.Text = Matrix.rc21
 T22.Text = Matrix.rc22
 T23.Text = Matrix.rc23
 T24.Text = Matrix.rc24
 
 T31.Text = Matrix.rc31
 T32.Text = Matrix.rc32
 T33.Text = Matrix.rc33
 T34.Text = Matrix.rc34
 
 T41.Text = Matrix.rc41
 T42.Text = Matrix.rc42
 T43.Text = Matrix.rc43
 T44.Text = Matrix.rc44
End Function

Private Function GetT() As Matrix4x4
 GetT.rc11 = T11.Text
 GetT.rc12 = T12.Text
 GetT.rc13 = T13.Text
 GetT.rc14 = T14.Text
 
 GetT.rc21 = T21.Text
 GetT.rc22 = T22.Text
 GetT.rc23 = T23.Text
 GetT.rc24 = T24.Text
 
 GetT.rc31 = T31.Text
 GetT.rc32 = T32.Text
 GetT.rc33 = T33.Text
 GetT.rc34 = T34.Text
 
 GetT.rc41 = T41.Text
 GetT.rc42 = T42.Text
 GetT.rc43 = T43.Text
 GetT.rc44 = T44.Text
End Function

Private Sub ARTButton_Click()
 On Error Resume Next
 Dim CurrentMatrix As Matrix4x4
 Dim RotationMatrixX As Matrix4x4
 Dim RotationMatrixY As Matrix4x4
 Dim RotationMatrixZ As Matrix4x4
 Dim Output As Matrix4x4
 
 Dim X As Single
 Dim Y As Single
 Dim Z As Single
 Dim I As Integer
 Dim A As Integer
 Dim Text As String
 
 Text = InputBox("Enter rotation in degrees, seperated by commas: X, Y, Z", "Rotation", "0, 0, 0")
 
 X = Val(Left(Text, Len(Text)))
    
 For I = 1 To Len(Text)
  If Mid(Text, I, 1) = "," Then
   If Not A = 0 Then
    If Y = 0 Then
     Y = Val(Mid(Text, A))
    End If
   End If
   A = I + 1
  End If
 Next
 Z = Val(Mid(Text, A))
 
 CurrentMatrix = GetT()
 RotationMatrixX = MatrixRotationX(ConvertDegToRad(X))
 RotationMatrixY = MatrixRotationY(ConvertDegToRad(Y))
 RotationMatrixZ = MatrixRotationZ(ConvertDegToRad(Z))

 
 Output = MatrixIdentity()
 Output = MatrixMultiply(Output, CurrentMatrix)
 Output = MatrixMultiply(Output, RotationMatrixX)
 Output = MatrixMultiply(Output, RotationMatrixY)
 Output = MatrixMultiply(Output, RotationMatrixZ)
 
 SetT Output
 
End Sub

Private Sub ASTMButton_Click()
 On Error Resume Next
 Dim CurrentMatrix As Matrix4x4
 Dim ScaledMatrix As Matrix4x4
 Dim Output As Matrix4x4
 
 Dim X As Single
 Dim Y As Single
 Dim Z As Single
 Dim I As Integer
 Dim A As Integer
 Dim Text As String
 
 Text = InputBox("Enter new scale, seperated by commas: X, Y, Z", "Scale", "1, 1, 1")
 
 X = Val(Left(Text, Len(Text)))
    
 For I = 1 To Len(Text)
  If Mid(Text, I, 1) = "," Then
   If Not A = 0 Then
    If Y = 0 Then
     Y = Val(Mid(Text, A))
    End If
   End If
   A = I + 1
  End If
 Next
 Z = Val(Mid(Text, A))
 
 CurrentMatrix = GetT()
 ScaledMatrix = MatrixScale(X, Y, Z)
 Output = MatrixMultiply(CurrentMatrix, ScaledMatrix)
 SetT Output
End Sub

Private Sub ATTButton_Click()
 On Error Resume Next
 Dim CurrentMatrix As Matrix4x4
 Dim TranslatedMatrix As Matrix4x4
 Dim Output As Matrix4x4
 
 Dim X As Single
 Dim Y As Single
 Dim Z As Single
 Dim I As Integer
 Dim A As Integer
 Dim Text As String
 
 Text = InputBox("Enter new position, seperated by commas: X, Y, Z", "Translation", "0, 0, 0")
 
 X = Val(Left(Text, Len(Text)))
    
 For I = 1 To Len(Text)
  If Mid(Text, I, 1) = "," Then
   If Not A = 0 Then
    If Y = 0 Then
     Y = Val(Mid(Text, A))
    End If
   End If
   A = I + 1
  End If
 Next
 Z = Val(Mid(Text, A))
 
 CurrentMatrix = GetT()
 TranslatedMatrix = MatrixTranslation(X, Y, Z)
 Output = MatrixMultiply(CurrentMatrix, TranslatedMatrix)
 SetT Output
End Sub

Private Sub Form_Activate()
 If Mesh.Triangles = 0 Then
  Unload Me
 Else
  SetT Mesh.Triangle(Trig).IdentityMatrix
  SolidColorText.Text = Val(Mesh.Triangle(Trig).SolidColor)
  If Not Mesh.Triangle(Trig).Texture.TextureHdc = 0 Then
   Draw TriangleTexture.Hdc, Mesh.Triangle(Trig).Texture.TextureHdc
   DoEvents
  Else
   TriangleTexture.Cls
  End If
 End If
End Sub

Private Sub ResetTButton_Click()
 SetT MatrixIdentity()
End Sub

Private Sub SaveButton_Click()
 Mesh.Triangle(Trig).SolidColor = Val(SolidColorText.Text)
 Mesh.Triangle(Trig).IdentityMatrix = GetT()
 If Not Mesh.Triangle(Trig).Texture.TextureHdc = 0 Then
  ClearHdc Mesh.Triangle(Trig).Texture.TextureHdc, 128, 128
 Else
  Mesh.Triangle(Trig).Texture.TextureHdc = CreateHdc(128, 128)
  DoEvents
  Mesh.Triangle(Trig).Texture.TextureHdc = GetCurrentHdc
 End If
 DoEvents
 Draw Mesh.Triangle(Trig).Texture.TextureHdc, TriangleTexture.Hdc
 Mesh.Triangle(Trig).Texture.TextureWidth = 128
 Mesh.Triangle(Trig).Texture.TextureHeight = 128
 
 DoEvents
 Unload Me
End Sub

Private Sub TriangleTexture_Click()
 CommonDialog.ShowOpen
 If Not CommonDialog.FileName = "" Then
  TriangleTexture.Picture = LoadPicture(CommonDialog.FileName)
 End If
End Sub
