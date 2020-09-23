Attribute VB_Name = "Math3DMod"

'Constants

'PI / 180
Private Const ConstPIDivideBy180 As Double = 1.74532925199433E-02
'180 / 180
Private Const Const180DivideByPI As Double = 57.2957795130823

'Functions

'Conversions

'Convets from Degrees to Radians
Function ConvertDegToRad(Degress As Single) As Single
 ConvertDegToRad = Degress * (ConstPIDivideBy180)
End Function
'Convets from Radians to Degrees
Function ConvertRadToDeg(Radians As Single) As Single
 ConvertRadToDeg = Radians * (Const180DivideByPI)
End Function


'Matrices

'Shifts the Z Axe on X and Y with the values, allowing one to disort
Function MatrixShear(ShearX As Single, ShearY As Single) As Matrix4x4
 'Reset MatrixShear
 MatrixShear = MatrixIdentity()
 'Set third colum on X and Y to ShearX and ShearY
 MatrixShear.rc13 = ShearX
 MatrixShear.rc23 = ShearY
End Function

'Converts a Camera into a Matrix, used to optimize Camera operations
Function MatrixViewMappingPerspective(Camera As ObjectCamera) As Matrix4x4
 Dim VectorCW As Coordinates4D            '   Centre of Window
 Dim VectorDOP As Coordinates4D           '   Direction Of Projection
 Dim MatrixTranslate As Matrix4x4
    
 Dim MShearX As Single
 Dim MShearY As Single
 Dim SMatrixShear As Matrix4x4
    
 Dim MScaleX As Single
 Dim MScaleY As Single
 Dim MScaleZ As Single
 Dim SMatrixScale As Matrix4x4
    
 Dim MatrixPerspective As Matrix4x4
 
 'Set MatrixTranslate origin to Camera.PRP(Projection Reference Point)
 MatrixTranslate = MatrixTranslation(-Camera.PRP.X, -Camera.PRP.Y, -Camera.PRP.Z)
 
 'Calculate the center of the window
 VectorCW.X = (Camera.MaxScreen.X + Camera.MinScreen.X) / 2
 VectorCW.Y = (Camera.MaxScreen.Y + Camera.MinScreen.Y) / 2
 'Since the screen only is 2D(X, Y) it cannot use the Z and W Axes
 VectorCW.Z = 0
 VectorCW.W = 1
 
 'Subtract VectorCW(Center Of Screen) from Camera.PRP(Projection Reference Point)
 VectorDOP = VectorSubtract(VectorCW, Camera.PRP)
 
 'Dividing by zero returns errors, so check if it's zero first
 If VectorDOP.Z <> 0 Then
  'Calculate the Shear Matrix
  MShearX = -(VectorDOP.X / VectorDOP.Z)
  MShearY = -(VectorDOP.Y / VectorDOP.Z)
 End If
 
 'Shear MatrixShear with the values calculated above, MShearX and MShearY
 SMatrixShear = MatrixShear(MShearX, MShearY)
    
 Dim MTemp As Double
 'Calculate the Perspective Scale transformation based on -Camera.PRP(Projection Reference Point),
 'MaxScreen, MinScreen, ClipNear and ClipFar
 MScaleX = (2 * -Camera.PRP.Z) / ((Camera.MaxScreen.X - Camera.MinScreen.X) * (-Camera.PRP.Z + Camera.ClipFar))
 MScaleY = (2 * -Camera.PRP.Z) / ((Camera.MaxScreen.Y - Camera.MinScreen.Y) * (-Camera.PRP.Z + Camera.ClipFar))
 MScaleZ = -1 / (-Camera.PRP.Z + Camera.ClipFar)
 
 'Scale MatrixScale by MScale
 SMatrixScale = MatrixScale(MScaleX, MScaleY, MScaleZ)
    
 Dim MZmin As Double
 
 MZmin = -((-Camera.PRP.Z + Camera.ClipNear) / (-Camera.PRP.Z + Camera.ClipFar))
 MatrixPerspective = MatrixIdentity
 
 'Minus one will cause errors, therefore check it
 If MZmin <> -1 Then
  MatrixPerspective.rc33 = 1 / (1 + MZmin)
  MatrixPerspective.rc34 = -MZmin / (1 + MZmin)
  MatrixPerspective.rc43 = -1
  MatrixPerspective.rc44 = 0
 End If
 
 'Reset MatrixViewMappingPerspective
 MatrixViewMappingPerspective = MatrixIdentity()
 'Multiply with MatrixTranslate, setting it's origin
 MatrixViewMappingPerspective = MatrixMultiply(MatrixViewMappingPerspective, MatrixTranslate)
 'Multiply with MatrixShear, setting it's Shear
 MatrixViewMappingPerspective = MatrixMultiply(MatrixViewMappingPerspective, SMatrixShear)
 'Multiply with MatrixScale, scalling it
 MatrixViewMappingPerspective = MatrixMultiply(MatrixViewMappingPerspective, SMatrixScale)
 'Multiply with MatrixPerspective, setting it's perspective
 MatrixViewMappingPerspective = MatrixMultiply(MatrixViewMappingPerspective, MatrixPerspective)
End Function

'Returns View Orientation based on vectors
Function MatrixViewOrientation(VectorVPN As Coordinates4D, VectorVUP As Coordinates4D, VectorVRP As Coordinates4D) As Matrix4x4
 Dim RotateVRC As Matrix4x4
 Dim TranslateVRP As Matrix4x4
 
 Dim VectorN As Coordinates4D
 Dim VectorU As Coordinates4D
 Dim VectorV As Coordinates4D
 
 'Normalize VectorN
 VectorN = VectorNormalize(VectorVPN)
    
 'Get CrossProduct of VectorVUP and VectorN(Which is normalized) and then Normalize the results
 VectorU = CrossProduct(VectorVUP, VectorN)
 VectorU = VectorNormalize(VectorU)
 
 'Get CrossProduct bettwen the Normalized Vectors: VectorU & VectorN
 VectorV = CrossProduct(VectorN, VectorU)

 'Reset RotateVRC Matrix
 RotateVRC = MatrixIdentity()
 
 'Define so that VectorU becomes first row, VectorV second and VectorN third
 With RotateVRC
  .rc11 = VectorU.X: .rc12 = VectorU.Y: .rc13 = VectorU.Z
  .rc21 = VectorV.X: .rc22 = VectorV.Y: .rc23 = VectorV.Z
  .rc31 = VectorN.X: .rc32 = VectorN.Y: .rc33 = VectorN.Z
 End With

 'Set TranslateVRP to have it's origin on -VectorVRP
 TranslateVRP = MatrixTranslation(-VectorVRP.X, -VectorVRP.Y, -VectorVRP.Z)

 'Reset MatrixViewOrientation
 MatrixViewOrientation = MatrixIdentity()
 'Multiply MatrixViewOrientation with TranslateVRP(Origin)
 MatrixViewOrientation = MatrixMultiply(MatrixViewOrientation, TranslateVRP)
 'Multiply MatrixViewOrientation with RotateVRC(Rotation)
 MatrixViewOrientation = MatrixMultiply(MatrixViewOrientation, RotateVRC)
End Function

Function MatrixView3D(Xmin As Single, Xmax As Single, Ymin As Single, Ymax As Single, Zmin As Single, Zmax As Single, Optional blnKeepSquare As Boolean = True) As Matrix4x4
 Dim MatrixTranslateA As Matrix4x4
 Dim MatrixTranslateB As Matrix4x4
 Dim SAspectRaio As Single
 Dim SScaleX As Single
 Dim SScaleY As Single
 Dim SScaleZ As Single
 Dim SMatrixScale As Matrix4x4
    
 Dim SDelta1 As Single
    
 'Sets MatrixTranslateA's origin to 1, 1, 1
 MatrixTranslateA = MatrixTranslation(1, 1, 1)
 
 'Set's SScale to Man - Min / 2(/ 1 with Z)
 SScaleX = (Xmax - Xmin) / 2
 SScaleY = (Ymax - Ymin) / 2
 SScaleZ = (Zmax - Zmin) / 1
 
 'If blnKeepSquare = True then use SAspectRatio when scalling height
 If blnKeepSquare = True Then
  
  SAspectRaio = Abs((Xmax - Xmin) / (Ymax - Ymin))  ' X pixels are SAspectRaio times bigger than Y pixels.
  SMatrixScale = MatrixScale(SScaleX, SScaleY * SAspectRaio, SScaleZ)
  SDelta1 = (Xmax - Ymin) / 2
 Else
 'Else just scale with the usual scale
  SMatrixScale = MatrixScale(SScaleX, SScaleY, SScaleZ)
  SDelta1 = 0
 End If
 
 'Sets MatrixTranslateA's origin to Xmin, Ymin + SDelta1, Zmin
 MatrixTranslateB = MatrixTranslation(Xmin, Ymin + SDelta1, Zmin)
 
 'Reset MatrixView3D
 MatrixView3D = MatrixIdentity()
 'Multiply with MatrixTranslateA, which is used to calculate the aspect/ratio
 MatrixView3D = MatrixMultiply(MatrixView3D, MatrixTranslateA)
 'Multiply with MatrixScale, which is used to calculate the size
 MatrixView3D = MatrixMultiply(MatrixView3D, SMatrixScale)
 'Multiply with MatrixTranslateB, which is used to calculate the position
 MatrixView3D = MatrixMultiply(MatrixView3D, MatrixTranslateB)
End Function

'Returns a Matrix Scaled by ScaleX, ScaleY and ScaleZ
Function MatrixScale(ScaleX As Single, ScaleY As Single, ScaleZ As Single) As Matrix4x4
 'Reset MatrixScale
 MatrixScale = MatrixIdentity()
 'Set 1,1 2,2 3,3 to ScaleX, ScaleY and ScaleZ because when
 'Row=Col then it's the multiplication value
 MatrixScale.rc11 = ScaleX
 MatrixScale.rc22 = ScaleY
 MatrixScale.rc33 = ScaleZ
End Function

'Switches bettwen Rows and Cols
Function MatrixTranspose(Matrix As Matrix4x4) As Matrix4x4
 With MatrixTranspose
  .rc11 = Matrix.rc11: .rc12 = Matrix.rc21: .rc13 = Matrix.rc31: .rc14 = Matrix.rc41
  .rc21 = Matrix.rc12: .rc22 = Matrix.rc22: .rc23 = Matrix.rc32: .rc24 = Matrix.rc42
  .rc31 = Matrix.rc13: .rc32 = Matrix.rc23: .rc33 = Matrix.rc33: .rc34 = Matrix.rc43
  .rc41 = Matrix.rc14: .rc42 = Matrix.rc24: .rc43 = Matrix.rc34: .rc44 = Matrix.rc44
 End With
End Function

'Returns a Matrix set to 0, but when Row=Col then set it to 1,
'because this is the multiplication value, if set to 0 the hole matrix is equal to nothing
Function MatrixIdentity() As Matrix4x4
 With MatrixIdentity
  .rc11 = 1: .rc12 = 0: .rc13 = 0: .rc14 = 0
  .rc21 = 0: .rc22 = 1: .rc23 = 0: .rc24 = 0
  .rc31 = 0: .rc32 = 0: .rc33 = 1: .rc34 = 0
  .rc41 = 0: .rc42 = 0: .rc43 = 0: .rc44 = 1
 End With
End Function

'Returns a new Matrix at the Offset positions
Function MatrixTranslation(OffsetX As Single, OffsetY As Single, OffsetZ As Single) As Matrix4x4
 'Reset the Matrix
 MatrixTranslation = MatrixIdentity()
 'The 4 col is the values(position) in the Matrix,
 'so all one have to do to change position is to change these values
 With MatrixTranslation
  .rc14 = OffsetX
  .rc24 = OffsetY
  .rc34 = OffsetZ
 End With
End Function

'The core function behind Matrices, because it allows one to combine several matrices into one,
'henze speeding up the calculations
Function MatrixMultiply(FirstMatrix As Matrix4x4, SecondMatrix As Matrix4x4) As Matrix4x4
 On Error GoTo err
 Dim TempFirstMatrix As Matrix4x4
 Dim TempSecondMatrix As Matrix4x4
 TempFirstMatrix = FirstMatrix
 TempSecondMatrix = SecondMatrix
 
 'Reset MatrixMultiply
 MatrixMultiply = MatrixIdentity()
 
 'Calculate Multiplication
 With MatrixMultiply
  
  .rc11 = (TempFirstMatrix.rc11 * TempSecondMatrix.rc11) + (TempFirstMatrix.rc21 * TempSecondMatrix.rc12) + (TempFirstMatrix.rc31 * TempSecondMatrix.rc13) + (TempFirstMatrix.rc41 * TempSecondMatrix.rc14)
  .rc12 = (TempFirstMatrix.rc12 * TempSecondMatrix.rc11) + (TempFirstMatrix.rc22 * TempSecondMatrix.rc12) + (TempFirstMatrix.rc32 * TempSecondMatrix.rc13) + (TempFirstMatrix.rc42 * TempSecondMatrix.rc14)
  .rc13 = (TempFirstMatrix.rc13 * TempSecondMatrix.rc11) + (TempFirstMatrix.rc23 * TempSecondMatrix.rc12) + (TempFirstMatrix.rc33 * TempSecondMatrix.rc13) + (TempFirstMatrix.rc43 * TempSecondMatrix.rc14)
  .rc14 = (TempFirstMatrix.rc14 * TempSecondMatrix.rc11) + (TempFirstMatrix.rc24 * TempSecondMatrix.rc12) + (TempFirstMatrix.rc34 * TempSecondMatrix.rc13) + (TempFirstMatrix.rc44 * TempSecondMatrix.rc14)
        
  .rc21 = (TempFirstMatrix.rc11 * TempSecondMatrix.rc21) + (TempFirstMatrix.rc21 * TempSecondMatrix.rc22) + (TempFirstMatrix.rc31 * TempSecondMatrix.rc23) + (TempFirstMatrix.rc41 * TempSecondMatrix.rc24)
  .rc22 = (TempFirstMatrix.rc12 * TempSecondMatrix.rc21) + (TempFirstMatrix.rc22 * TempSecondMatrix.rc22) + (TempFirstMatrix.rc32 * TempSecondMatrix.rc23) + (TempFirstMatrix.rc42 * TempSecondMatrix.rc24)
  .rc23 = (TempFirstMatrix.rc13 * TempSecondMatrix.rc21) + (TempFirstMatrix.rc23 * TempSecondMatrix.rc22) + (TempFirstMatrix.rc33 * TempSecondMatrix.rc23) + (TempFirstMatrix.rc43 * TempSecondMatrix.rc24)
  .rc24 = (TempFirstMatrix.rc14 * TempSecondMatrix.rc21) + (TempFirstMatrix.rc24 * TempSecondMatrix.rc22) + (TempFirstMatrix.rc34 * TempSecondMatrix.rc23) + (TempFirstMatrix.rc44 * TempSecondMatrix.rc24)
        
  .rc31 = (TempFirstMatrix.rc11 * TempSecondMatrix.rc31) + (TempFirstMatrix.rc21 * TempSecondMatrix.rc32) + (TempFirstMatrix.rc31 * TempSecondMatrix.rc33) + (TempFirstMatrix.rc41 * TempSecondMatrix.rc34)
  .rc32 = (TempFirstMatrix.rc12 * TempSecondMatrix.rc31) + (TempFirstMatrix.rc22 * TempSecondMatrix.rc32) + (TempFirstMatrix.rc32 * TempSecondMatrix.rc33) + (TempFirstMatrix.rc42 * TempSecondMatrix.rc34)
  .rc33 = (TempFirstMatrix.rc13 * TempSecondMatrix.rc31) + (TempFirstMatrix.rc23 * TempSecondMatrix.rc32) + (TempFirstMatrix.rc33 * TempSecondMatrix.rc33) + (TempFirstMatrix.rc43 * TempSecondMatrix.rc34)
  .rc34 = (TempFirstMatrix.rc14 * TempSecondMatrix.rc31) + (TempFirstMatrix.rc24 * TempSecondMatrix.rc32) + (TempFirstMatrix.rc34 * TempSecondMatrix.rc33) + (TempFirstMatrix.rc44 * TempSecondMatrix.rc34)
        
  .rc41 = (TempFirstMatrix.rc11 * TempSecondMatrix.rc41) + (TempFirstMatrix.rc21 * TempSecondMatrix.rc42) + (TempFirstMatrix.rc31 * TempSecondMatrix.rc43) + (TempFirstMatrix.rc41 * TempSecondMatrix.rc44)
  .rc42 = (TempFirstMatrix.rc12 * TempSecondMatrix.rc41) + (TempFirstMatrix.rc22 * TempSecondMatrix.rc42) + (TempFirstMatrix.rc32 * TempSecondMatrix.rc43) + (TempFirstMatrix.rc42 * TempSecondMatrix.rc44)
  .rc43 = (TempFirstMatrix.rc13 * TempSecondMatrix.rc41) + (TempFirstMatrix.rc23 * TempSecondMatrix.rc42) + (TempFirstMatrix.rc33 * TempSecondMatrix.rc43) + (TempFirstMatrix.rc43 * TempSecondMatrix.rc44)
  .rc44 = (TempFirstMatrix.rc14 * TempSecondMatrix.rc41) + (TempFirstMatrix.rc24 * TempSecondMatrix.rc42) + (TempFirstMatrix.rc34 * TempSecondMatrix.rc43) + (TempFirstMatrix.rc44 * TempSecondMatrix.rc44)
 End With
Exit Function
err:
 MatrixMultiply = FirstMatrix
End Function

'Multiplies a Matrix by Vector, the second core function for using matrices,
'because it's the bridge bettwen vectors and matrices
Function MatrixMultiplyVector(Matrix As Matrix4x4, Vector As Coordinates4D) As Coordinates4D
 'Calculate multiplication of Matrix and Vector
 With MatrixMultiplyVector
  .X = (Matrix.rc11 * Vector.X) + (Matrix.rc12 * Vector.Y) + (Matrix.rc13 * Vector.Z) + (Matrix.rc14 * Vector.W)
  .Y = (Matrix.rc21 * Vector.X) + (Matrix.rc22 * Vector.Y) + (Matrix.rc23 * Vector.Z) + (Matrix.rc24 * Vector.W)
  .Z = (Matrix.rc31 * Vector.X) + (Matrix.rc32 * Vector.Y) + (Matrix.rc33 * Vector.Z) + (Matrix.rc34 * Vector.W)
  .W = (Matrix.rc41 * Vector.X) + (Matrix.rc42 * Vector.Y) + (Matrix.rc43 * Vector.Z) + (Matrix.rc44 * Vector.W)
 End With
End Function

' Matrix Rotations

'Returns a Matrix rotated around the X Axe based on number of Radians
Function MatrixRotationX(Radians As Single) As Matrix4x4
 Dim sngCosine As Double
 Dim sngSine As Double
    
 'Calculate Cosine and Sine based on Radians
 sngCosine = Round(Cos(Radians), 6)
 sngSine = Round(Sin(Radians), 6)
    
 'Reset Matrix
 MatrixRotationX = MatrixIdentity()
 
 'When rotating about the X Axe you don't really change the X Axe, but rather the Y and Z Axes
 With MatrixRotationX
  .rc22 = sngCosine
  .rc23 = -sngSine
  .rc32 = sngSine
  .rc33 = sngCosine
 End With
End Function

'Returns a Matrix rotated around the Y Axe based on number of Radians
Function MatrixRotationY(Radians As Single) As Matrix4x4
 Dim sngCosine As Double
 Dim sngSine As Double

 'Calculate Cosine and Sine based on Radians
 sngCosine = Round(Cos(Radians), 6)
 sngSine = Round(Sin(Radians), 6)
    
 'Reset Matrix
 MatrixRotationY = MatrixIdentity()
 
 'When rotating about the Y Axe you don't really change the Y Axe, but rather the X and Y Axes
 With MatrixRotationY
  .rc11 = sngCosine
  .rc31 = -sngSine
  .rc13 = sngSine
  .rc33 = sngCosine
 End With
End Function

'Returns a Matrix rotated around the Z Axe based on number of Radians
Function MatrixRotationZ(Radians As Single) As Matrix4x4
 Dim sngCosine As Double
 Dim sngSine As Double

 'Calculate Cosine and Sine based on Radians
 sngCosine = Round(Cos(Radians), 6)
 sngSine = Round(Sin(Radians), 6)
    
 'Reset Matrix
 MatrixRotationZ = MatrixIdentity()
 
 'When rotating about the Z Axe you don't really change the Z Axe, but rather the X and Y Axes
 With MatrixRotationZ
  .rc11 = sngCosine
  .rc21 = sngSine
  .rc12 = -sngSine
  .rc22 = sngCosine
 End With
End Function


' Vectors

'Returns the normalized vector, meaning dividing it by it's length
Function VectorNormalize(Vector As Coordinates4D) As Coordinates4D
 Dim VecLength As Single
 VecLength = VectorLength(Vector)
 If VecLength = 0 Then VecLength = 1
 With VectorNormalize
  .X = Vector.X / VecLength
  .Y = Vector.Y / VecLength
  .Z = Vector.Z / VecLength
  'Ignore W
  .W = Vector.W
 End With
End Function

Function VectorDistance(FirstVector As Coordinates4D, SecondVector As Coordinates4D) As Single
 'Calculates the length based on Phytagoras theory
 VectorDistance = VectorLength(VectorSubtract(FirstVector, SecondVector))
 'Ignore W
End Function

'Returns the vector length
Function VectorLength(Vector As Coordinates4D) As Single
 'Calculates the length based on Phytagoras theory
 VectorLength = Sqr((Vector.X ^ 2) + (Vector.Y ^ 2) + (Vector.Z ^ 2))
 'Ignore W
End Function

'Returns two vectors added together
Function VectorAddition(FirstVector As Coordinates4D, SecondVector As Coordinates4D) As Coordinates4D
 With VectorAddition
  'Add
  .X = FirstVector.X + SecondVector.X
  .Y = FirstVector.Y + SecondVector.Y
  .Z = FirstVector.Z + SecondVector.Z
  .W = 1 'Ignore W
 End With
End Function

'Returns the FirstVector subtracted by the SecondVector
Function VectorSubtract(FirstVector As Coordinates4D, SecondVector As Coordinates4D) As Coordinates4D
 With VectorSubtract
  'Subtract
  .X = FirstVector.X - SecondVector.X
  .Y = FirstVector.Y - SecondVector.Y
  .Z = FirstVector.Z - SecondVector.Z
  .W = 1 'Ignore W
 End With
End Function

'Convert 4D Vectors to 2D Vectors
Function Vector4DTo3D(Vector As Coordinates4D) As Coordinates4D
 On Error Resume Next
 With Vector4DTo3D
  If Vector.W <> 0 Then
    'Divide X, Y and Z by W
    .X = Vector.X / Vector.W
    .Y = Vector.Y / Vector.W
    .Z = Vector.Z / Vector.W
  Else
    .X = Vector.X
    .Y = Vector.Y
    .Z = Vector.Z
  End If
 End With
End Function

'Returns VectorIN multiplied by Scalar
Function VectorMultiplyByScalar(VectorIn As Coordinates4D, Scalar As Single) As Coordinates4D
 With VectorMultiplyByScalar
  .X = CSng(VectorIn.X) * CSng(Scalar)
  .Y = CSng(VectorIn.Y) * CSng(Scalar)
  .Z = CSng(VectorIn.Z) * CSng(Scalar)
  .W = VectorIn.W ' Ignore W
 End With
End Function

'Products

'Returns the CrossProduct of the FirstVector and the SecondVector, which is perpendicular to the two vectors
Function CrossProduct(FirstVector As Coordinates4D, SecondVector As Coordinates4D) As Coordinates4D
 'Calculate the CrossProduct
 With CrossProduct
  .X = (FirstVector.Y * SecondVector.Z) - (FirstVector.Z * SecondVector.Y)
  .Y = (FirstVector.Z * SecondVector.X) - (FirstVector.X * SecondVector.Z)
  .Z = (FirstVector.X * SecondVector.Y) - (FirstVector.Y * SecondVector.X)
  .W = 1 ' Ignore W
 End With
End Function

Function DotProduct3D(FirstVector As Coordinates4D, SecondVector As Coordinates4D) As Single
 'Calculate the DotProduct
 DotProduct3D = (FirstVector.X * SecondVector.X) + (FirstVector.Y * SecondVector.Y) + (FirstVector.Z * SecondVector.Z)
 'Ignore W
End Function

