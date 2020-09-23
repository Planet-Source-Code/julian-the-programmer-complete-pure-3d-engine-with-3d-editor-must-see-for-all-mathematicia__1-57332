Attribute VB_Name = "Math2DMod"
Function VectorDistance2D(FirstVector As Coordinates2D, SecondVector As Coordinates2D) As Single
 'Calculates the length based on Phytagoras theory
 VectorDistance2D = VectorLength2D(VectorSubtract2D(FirstVector, SecondVector))
End Function

Function DotProduct2D(FirstVector As Coordinates2D, SecondVector As Coordinates2D) As Single
 'Calculate the DotProduct based on FX*SX + FY*SY
 DotProduct2D = (FirstVector.X * SecondVector.X) + (FirstVector.Y * SecondVector.Y)
 If DotProduct2D = 0 Then DotProduct = 1
End Function

Function VectorLength2D(Vector As Coordinates2D) As Single
 'Calculates the length based on Phytagoras theory
 VectorLength2D = Sqr((Vector.X ^ 2) + (Vector.Y ^ 2))
End Function

'Returns two vectors added together
Function VectorAddition2D(FirstVector As Coordinates2D, SecondVector As Coordinates2D) As Coordinates2D
 With VectorAddition2D
  'Add
  .X = FirstVector.X + SecondVector.X
  .Y = FirstVector.Y + SecondVector.Y
 End With
End Function

'Returns the FirstVector subtracted by the SecondVector
Function VectorSubtract2D(FirstVector As Coordinates2D, SecondVector As Coordinates2D) As Coordinates2D
 With VectorSubtract2D
  'Subtract
  .X = FirstVector.X - SecondVector.X
  .Y = FirstVector.Y - SecondVector.Y
 End With
End Function

Function VectorNormalize2D(Vector As Coordinates2D) As Coordinates2D
 Dim VecLength As Single
 VecLength = VectorLength2D(Vector)
 If VecLength = 0 Then VecLength = 1
 With VectorNormalize2D
  .X = Vector.X / VecLength
  .Y = Vector.Y / VecLength
 End With
End Function

Function IsInTriangle2D(Position As Coordinates2D, FirstVector As Coordinates2D, SecondVector As Coordinates2D, ThirdVector As Coordinates2D) As Boolean
    Dim bc, ca, ab, ap, bp, cp, abc As Double
    
    bc = SecondVector.X * ThirdVector.Y - SecondVector.Y * ThirdVector.X
    ca = ThirdVector.X * FirstVector.Y - ThirdVector.Y * FirstVector.X
    ab = FirstVector.X * SecondVector.Y - FirstVector.Y * SecondVector.X
    ap = FirstVector.X * Position.Y - FirstVector.Y * Position.X
    bp = SecondVector.X * Position.Y - SecondVector.Y * Position.X
    cp = ThirdVector.X * Position.Y - ThirdVector.Y * Position.X
    abc = Sgn(bc + ca + ab)

    If (abc * (bc - bp + cp) > 0) And (abc * (ca - cp + ap) > 0) And (abc * (ab - ap + bp) > 0) Then IsInTriangle2D = True
End Function

Function VectorAngleToPosition2D(X As Single, Y As Single, Angle As Single, Steps As Single) As Coordinates2D
 VectorAngleToPosition2D.X = X + Round(Steps * Cos(Angle))
 VectorAngleToPosition2D.Y = Y + Round(Steps * Sin(Angle))
End Function

Function GetXByYInLine(FirstX As Single, FirstY As Single, SecondX As Single, SecondY As Single, Y As Single) As Single
 GetXByYInLine = ((FirstX - SecondX) / (FirstY - SecondY)) * (Y - FirstY + ((FirstY - SecondY) / (FirstX - SecondX)) * FirstX)
End Function

