Attribute VB_Name = "ObjectMod"
Function ResetMesh() As Object3DMesh
 ResetMesh.IdentityMatrix = MatrixIdentity()
 
 ResetMesh.Position.X = 0#
 ResetMesh.Position.Y = 0#
 ResetMesh.Position.Z = 0#
 ResetMesh.Position.W = 1#
 
 ResetMesh.Triangles = 0
End Function

Function ResetTriangle() As ObjectTriangle
 ResetTriangle.IdentityMatrix = MatrixIdentity()
 
 ResetTriangle.Coordinates(1).X = 0
 ResetTriangle.Coordinates(1).Y = 0
 ResetTriangle.Coordinates(1).Z = 0
 ResetTriangle.Coordinates(1).W = 1#
 
 ResetTriangle.Coordinates(2).X = 0
 ResetTriangle.Coordinates(2).Y = 0
 ResetTriangle.Coordinates(2).Z = 0
 ResetTriangle.Coordinates(2).W = 1#
 
 ResetTriangle.Coordinates(3).X = 0
 ResetTriangle.Coordinates(3).Y = 0
 ResetTriangle.Coordinates(3).Z = 0
 ResetTriangle.Coordinates(3).W = 1#
 
 ResetTriangle.SolidColor = &HFFFFFF
End Function
