VERSION 5.00
Begin VB.Form MatrixSetup 
   Caption         =   "Matrix Setup"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton SaveButton 
      Caption         =   "Save"
      Height          =   255
      Left            =   240
      TabIndex        =   42
      Top             =   3240
      Width           =   5295
   End
   Begin VB.TextBox MMM44 
      Height          =   285
      Left            =   5160
      TabIndex        =   41
      Text            =   "1"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox MMM41 
      Height          =   285
      Left            =   3000
      TabIndex        =   40
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox MMM42 
      Height          =   285
      Left            =   3720
      TabIndex        =   39
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox MMM43 
      Height          =   285
      Left            =   4440
      TabIndex        =   38
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox MMM14 
      Height          =   285
      Left            =   5160
      TabIndex        =   37
      Text            =   "0"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox MMM24 
      Height          =   285
      Left            =   5160
      TabIndex        =   36
      Text            =   "0"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox MMM34 
      Height          =   285
      Left            =   5160
      TabIndex        =   35
      Text            =   "0"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox CVM43 
      Height          =   285
      Left            =   1440
      TabIndex        =   34
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox CVM14 
      Height          =   285
      Left            =   2160
      TabIndex        =   33
      Text            =   "0"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox CVM24 
      Height          =   285
      Left            =   2160
      TabIndex        =   32
      Text            =   "0"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox CVM34 
      Height          =   285
      Left            =   2160
      TabIndex        =   31
      Text            =   "0"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox CVM41 
      Height          =   285
      Left            =   0
      TabIndex        =   30
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox CVM42 
      Height          =   285
      Left            =   720
      TabIndex        =   29
      Text            =   "0"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox CVM44 
      Height          =   285
      Left            =   2160
      TabIndex        =   28
      Text            =   "1"
      Top             =   1440
      Width           =   615
   End
   Begin VB.CommandButton ResetMMMButton 
      Caption         =   "Reset"
      Height          =   255
      Left            =   3000
      TabIndex        =   27
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton ResetCVMButton 
      Caption         =   "Reset"
      Height          =   255
      Left            =   0
      TabIndex        =   26
      Top             =   2880
      Width           =   2775
   End
   Begin VB.CommandButton ASMMMButton 
      Caption         =   "Apply Scale"
      Height          =   255
      Left            =   3000
      TabIndex        =   25
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton ASCVMButton 
      Caption         =   "Apply Scale"
      Height          =   255
      Left            =   0
      TabIndex        =   24
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton ATMMMButton 
      Caption         =   "Apply Translation"
      Height          =   255
      Left            =   3000
      TabIndex        =   23
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton ATMCVMButton 
      Caption         =   "Apply Translation"
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   2160
      Width           =   2775
   End
   Begin VB.CommandButton ARMMMButton 
      Caption         =   "Apply Rotation"
      Height          =   255
      Left            =   3000
      TabIndex        =   21
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton ARCVMButton 
      Caption         =   "Apply Rotation"
      Height          =   255
      Left            =   0
      TabIndex        =   20
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox MMM33 
      Height          =   285
      Left            =   4440
      TabIndex        =   19
      Text            =   "1"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox MMM32 
      Height          =   285
      Left            =   3720
      TabIndex        =   18
      Text            =   "0"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox MMM31 
      Height          =   285
      Left            =   3000
      TabIndex        =   17
      Text            =   "0"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox MMM23 
      Height          =   285
      Left            =   4440
      TabIndex        =   16
      Text            =   "0"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox MMM22 
      Height          =   285
      Left            =   3720
      TabIndex        =   15
      Text            =   "1"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox MMM21 
      Height          =   285
      Left            =   3000
      TabIndex        =   14
      Text            =   "0"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox MMM13 
      Height          =   285
      Left            =   4440
      TabIndex        =   13
      Text            =   "0"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox MMM12 
      Height          =   285
      Left            =   3720
      TabIndex        =   12
      Text            =   "0"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox MMM11 
      Height          =   285
      Left            =   3000
      TabIndex        =   11
      Text            =   "1"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox CVM33 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Text            =   "1"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox CVM32 
      Height          =   285
      Left            =   720
      TabIndex        =   8
      Text            =   "0"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox CVM31 
      Height          =   285
      Left            =   0
      TabIndex        =   7
      Text            =   "0"
      Top             =   1080
      Width           =   615
   End
   Begin VB.TextBox CVM23 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Text            =   "0"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox CVM22 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "1"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox CVM21 
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Text            =   "0"
      Top             =   720
      Width           =   615
   End
   Begin VB.TextBox CVM13 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Text            =   "0"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox CVM12 
      Height          =   285
      Left            =   720
      TabIndex        =   2
      Text            =   "0"
      Top             =   360
      Width           =   615
   End
   Begin VB.TextBox CVM11 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "1"
      Top             =   360
      Width           =   615
   End
   Begin VB.Line MiddleLine 
      X1              =   2880
      X2              =   2880
      Y1              =   0
      Y2              =   3120
   End
   Begin VB.Label MMMLabel 
      Caption         =   "Mesh Matrix"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   10
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label CWMLabel 
      Caption         =   "Camera View Matrix"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "MatrixSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ARCVMButton_Click()
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
 
 CurrentMatrix = GetCVM()
 RotationMatrixX = MatrixRotationX(ConvertDegToRad(X))
 RotationMatrixY = MatrixRotationY(ConvertDegToRad(Y))
 RotationMatrixZ = MatrixRotationZ(ConvertDegToRad(Z))

 
 Output = MatrixIdentity()
 Output = MatrixMultiply(Output, CurrentMatrix)
 Output = MatrixMultiply(Output, RotationMatrixX)
 Output = MatrixMultiply(Output, RotationMatrixY)
 Output = MatrixMultiply(Output, RotationMatrixZ)
 
 SetCVM Output
 
End Sub

Private Function SetCVM(Matrix As Matrix4x4)
 CVM11.Text = Matrix.rc11
 CVM12.Text = Matrix.rc12
 CVM13.Text = Matrix.rc13
 CVM14.Text = Matrix.rc14
 
 CVM21.Text = Matrix.rc21
 CVM22.Text = Matrix.rc22
 CVM23.Text = Matrix.rc23
 CVM24.Text = Matrix.rc24
 
 CVM31.Text = Matrix.rc31
 CVM32.Text = Matrix.rc32
 CVM33.Text = Matrix.rc33
 CVM34.Text = Matrix.rc34
 
 CVM41.Text = Matrix.rc41
 CVM42.Text = Matrix.rc42
 CVM43.Text = Matrix.rc43
 CVM44.Text = Matrix.rc44
End Function

Private Function SetMMM(Matrix As Matrix4x4)
 MMM11.Text = Matrix.rc11
 MMM12.Text = Matrix.rc12
 MMM13.Text = Matrix.rc13
 MMM14.Text = Matrix.rc14
 
 MMM21.Text = Matrix.rc21
 MMM22.Text = Matrix.rc22
 MMM23.Text = Matrix.rc23
 MMM24.Text = Matrix.rc24
 
 MMM31.Text = Matrix.rc31
 MMM32.Text = Matrix.rc32
 MMM33.Text = Matrix.rc33
 MMM34.Text = Matrix.rc34
 
 MMM41.Text = Matrix.rc41
 MMM42.Text = Matrix.rc42
 MMM43.Text = Matrix.rc43
 MMM44.Text = Matrix.rc44
End Function

Private Function GetCVM() As Matrix4x4
 On Error Resume Next
 GetCVM.rc11 = CVM11.Text
 GetCVM.rc12 = CVM12.Text
 GetCVM.rc13 = CVM13.Text
 GetCVM.rc14 = CVM14.Text
 
 GetCVM.rc21 = CVM21.Text
 GetCVM.rc22 = CVM22.Text
 GetCVM.rc23 = CVM23.Text
 GetCVM.rc24 = CVM24.Text
 
 GetCVM.rc31 = CVM31.Text
 GetCVM.rc32 = CVM32.Text
 GetCVM.rc33 = CVM33.Text
 GetCVM.rc34 = CVM34.Text
 
 GetCVM.rc41 = CVM41.Text
 GetCVM.rc42 = CVM42.Text
 GetCVM.rc43 = CVM43.Text
 GetCVM.rc44 = CVM44.Text
End Function

Private Function GetMMM() As Matrix4x4
 On Error Resume Next
 GetMMM.rc11 = MMM11.Text
 GetMMM.rc12 = MMM12.Text
 GetMMM.rc13 = MMM13.Text
 GetMMM.rc14 = MMM14.Text
 
 GetMMM.rc21 = MMM21.Text
 GetMMM.rc22 = MMM22.Text
 GetMMM.rc23 = MMM23.Text
 GetMMM.rc24 = MMM24.Text
 
 GetMMM.rc31 = MMM31.Text
 GetMMM.rc32 = MMM32.Text
 GetMMM.rc33 = MMM33.Text
 GetMMM.rc34 = MMM34.Text
 
 GetMMM.rc41 = MMM41.Text
 GetMMM.rc42 = MMM42.Text
 GetMMM.rc43 = MMM43.Text
 GetMMM.rc44 = MMM44.Text
End Function

Private Sub ARMMMButton_Click()
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
 
 CurrentMatrix = GetMMM()
 RotationMatrixX = MatrixRotationX(ConvertDegToRad(X))
 RotationMatrixY = MatrixRotationY(ConvertDegToRad(Y))
 RotationMatrixZ = MatrixRotationZ(ConvertDegToRad(Z))

 
 Output = MatrixIdentity()
 Output = MatrixMultiply(Output, CurrentMatrix)
 Output = MatrixMultiply(Output, RotationMatrixX)
 Output = MatrixMultiply(Output, RotationMatrixY)
 Output = MatrixMultiply(Output, RotationMatrixZ)
 
 SetMMM Output
End Sub

Private Sub ASCVMButton_Click()
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
 
 CurrentMatrix = GetCVM()
 ScaledMatrix = MatrixScale(X, Y, Z)
 Output = MatrixMultiply(CurrentMatrix, ScaledMatrix)
 SetCVM Output
End Sub

Private Sub ASMMMButton_Click()
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
 
 CurrentMatrix = GetMMM()
 ScaledMatrix = MatrixScale(X, Y, Z)
 Output = MatrixMultiply(CurrentMatrix, ScaledMatrix)
 SetMMM Output
End Sub

Private Sub ATMCVMButton_Click()
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
 
 CurrentMatrix = GetCVM()
 TranslatedMatrix = MatrixTranslation(X, Y, Z)
 Output = MatrixMultiply(CurrentMatrix, TranslatedMatrix)
 SetCVM Output
End Sub

Private Sub ATMMMButton_Click()
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
 
 CurrentMatrix = GetMMM()
 TranslatedMatrix = MatrixTranslation(X, Y, Z)
 Output = MatrixMultiply(CurrentMatrix, TranslatedMatrix)
 SetMMM Output
End Sub

Private Sub Form_Activate()
 SetMMM Mesh.IdentityMatrix
 SetCVM Cam.ViewMatrix
End Sub

Private Sub ResetCVMButton_Click()
 SetCVM MatrixIdentity()
End Sub

Private Sub ResetMMMButton_Click()
 SetMMM MatrixIdentity()
End Sub

Private Sub SaveButton_Click()
 Mesh.IdentityMatrix = GetMMM()
 Cam.ViewMatrix = GetCVM()
 DoEvents
 Unload Me
End Sub
