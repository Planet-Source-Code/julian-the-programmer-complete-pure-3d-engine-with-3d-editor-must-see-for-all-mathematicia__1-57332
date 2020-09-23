Attribute VB_Name = "RenderMod"
Function RenderMesh3D(Mesh3D As Object3DMesh, ViewCamera As ObjectCamera, Rendering As RenderingType, Filling As FillMode, ScaleLeft As Single, ScaleWidth As Single, ScaleHeight As Single, ScaleTop As Single, HDc As Long)
' SetParamaters ViewCamera
 If Mesh3D.Triangles = 0 Then Exit Function 'Exit if none
 
 Dim MatrixViewPort As Matrix4x4 'Viewspace matrix
 Dim Mesh3DMatrix As Matrix4x4 'Mesh3D matrix
 Dim CameraViewMatrix As Matrix4x4
 Dim CameraMatrixViewOrientation As Matrix4x4
 
 Dim MiddleX As Single
 Dim MiddleY As Single
 
 Dim CurrentScreen As Coordinates4D
 
 
 
 MiddleX = ((ViewCamera.MaxScreen.X - ViewCamera.MinScreen.X) / 2) + ViewCamera.MinScreen.X
 MiddleY = ((ViewCamera.MaxScreen.Y - ViewCamera.MinScreen.Y) / 2) + ViewCamera.MinScreen.Y
 
 
 CameraViewMatrix = ViewCamera.ViewMatrix 'Set the CameraViewMatrix to the Camera View Matrix
 
 CameraMatrixViewOrientation = GetMatrixViewOrientation(ViewCamera) 'Set view orientation based on the camera

 MatrixViewPort = MatrixView3D(ScaleLeft, ScaleWidth, ScaleHeight, ScaleTop, -1, 0) 'Set viewspace on scale
 Mesh3DMatrix = Mesh3D.IdentityMatrix 'Set Mesh3DMatrix to the Mesh3D Matrix
 
 Dim I As Integer
 
 For I = 1 To Mesh3D.Triangles
  Dim MatrixOutput As Matrix4x4 'Matrix Output
  Dim MatrixWorld As Matrix4x4 'Temporarly matrix for triangle matrix storage
  
  Dim TempX(3) As Long
  Dim TempY(3) As Long
 
  Dim A As Integer 'Counter
  
  MatrixOutput = MatrixIdentity() 'Reset it

  'Multiply with all the other matrices, combining them all to one
  MatrixOutput = MatrixMultiply(MatrixOutput, Mesh3DMatrix) 'Multiply with the main Identity matrix
  MatrixOutput = MatrixMultiply(MatrixOutput, Mesh3D.Triangle(I).IdentityMatrix) 'Combine with the sub Identity matrix for the triangle
  MatrixOutput = MatrixMultiply(MatrixOutput, CameraViewMatrix) 'Combine with the viewmatrix
  MatrixOutput = MatrixMultiply(MatrixOutput, CameraMatrixViewOrientation) 'Combine with the orientation matrix
  
  For A = 1 To 3
   'Rotate so Z is just the deepth in the screen
   CurrentScreen = MatrixMultiplyVector(MatrixOutput, Mesh3D.Triangle(I).Coordinates(A))
   
   
   'Convert from 4D to 3D, by dividing with W: X/W, Y/W, Z/W
   CurrentScreen = Vector4DTo3D(CurrentScreen)
   
   'Divide by Z and multiply with the Scale
   On Error GoTo Make4DTo3D
   CurrentScreen.X = (CurrentScreen.X / CurrentScreen.Z) * ViewCamera.ScaleSize + MiddleY
   CurrentScreen.Y = (CurrentScreen.Y / CurrentScreen.Z) * ViewCamera.ScaleSize + MiddleX

   TempX(A) = CLng(CurrentScreen.X)
   TempY(A) = CLng(CurrentScreen.Y)
Make4DTo3D:
 CurrentScreen = Vector4DTo3D(CurrentScreen)
'Resume Next

  Next
  
     
   'Draw a line from the last point to the current
   If Filling = Solid Then
    FillSolidTriangle Mesh3D.Triangle(I).SolidColor, Make2DCoordinate(CSng(TempX(1)), CSng(TempY(1))), Make2DCoordinate(CSng(TempX(2)), CSng(TempY(2))), Make2DCoordinate(CSng(TempX(3)), CSng(TempY(3))), 1024, HDc
   End If
    DrawLine TempX(1), TempY(1), TempX(2), TempY(2), HDc
    DrawLine TempX(2), TempY(2), TempX(3), TempY(3), HDc
    DrawLine TempX(3), TempY(3), TempX(1), TempY(1), HDc
 Next
End Function
