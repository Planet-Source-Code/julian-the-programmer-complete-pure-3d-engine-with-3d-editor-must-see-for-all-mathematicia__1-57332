Attribute VB_Name = "TypesMod"
Option Explicit

'Simple 2D Axes coordinate system: X, Y
Type Coordinates2D
 X As Single
 Y As Single
End Type

'Simple 3D Axes coordinate system: X, Y, Z
Type Coordinates3D
 X As Single
 Y As Single
 Z As Single
End Type

'Simple 3D Axes coordinate system: X, Y, Z, W - W because we ran out of characters
Type Coordinates4D
 X As Single
 Y As Single
 Z As Single
 W As Single
End Type

'Matrix on 4x4, used to speed up rendering
'One could remove it yes, and just use the calculations yes,
'but Matrices has the ability to perform a calculation on a single addition and multiplication
'instead of having to include the subtracting and divition.
'Plus Matrices also speeds up when it comes to rotation.

'RC = Rows-Cols
Type Matrix4x4
 rc11 As Single: rc12 As Single: rc13 As Single: rc14 As Single
 rc21 As Single: rc22 As Single: rc23 As Single: rc24 As Single
 rc31 As Single: rc32 As Single: rc33 As Single: rc34 As Single
 rc41 As Single: rc42 As Single: rc43 As Single: rc44 As Single
End Type

'The texture, but without positions, this enables each texture to be independent
'and duplicated easy
Type ObjectTexture
 TextureHdc As Long
 TextureWidth As Long
 TextureHeight As Long
End Type

'Most simple polygon, with three positions and a matrix
'describing it's origin position, rotation and size, and also with texture
Type ObjectTriangle
 TextureFirstPosition As Coordinates2D
 TextureSecondPosition As Coordinates2D
 TextureThirdPosition As Coordinates2D
 Texture As ObjectTexture
 SolidColor As Long
 Coordinates(3) As Coordinates4D
 IdentityMatrix As Matrix4x4
End Type

'Basicly just alot of triangles and a matrix to describe
'the overall origin position, rotation and size
Type Object3DMesh
 Position As Coordinates4D
 Triangle() As ObjectTriangle
 Triangles As Integer
 IdentityMatrix As Matrix4x4
End Type

'The camera contains it's own world position, when it comes to standing and viewing
'The camera also has a value that decide which way is UP (VUP),
'a Projection Reference Point (PRP),
'a ClipFar and a ClipNear, that decides when objects are too close or too far to be rendered
'and also a ViewMatrix
Type ObjectCamera
 Position As Coordinates4D
 ViewPosition As Coordinates4D
 VUP As Coordinates4D
 PRP As Coordinates4D

 MinScreen As Coordinates2D
 MaxScreen As Coordinates2D
 ViewMatrix As Matrix4x4
 
 ClipFar As Single
 ClipNear As Single
 
 ScaleSize As Single
End Type

Enum RenderingType
 OpenGL = 1
 DirectX = 2
 Software = 3
End Enum

Function Make2DCoordinate(X As Single, Y As Single) As Coordinates2D
 Make2DCoordinate.X = X
 Make2DCoordinate.Y = Y
End Function

Function Make3DCoordinate(X As Single, Y As Single, Z As Single) As Coordinates3D
 Make3DCoordinate.X = X
 Make3DCoordinate.Y = Y
 Make3DCoordinate.Z = Z
End Function

Function Make4DCoordinate(X As Single, Y As Single, Z As Single, W As Single) As Coordinates4D
 Make4DCoordinate.X = X
 Make4DCoordinate.Y = Y
 Make4DCoordinate.Z = Z
 Make4DCoordinate.W = W
End Function
