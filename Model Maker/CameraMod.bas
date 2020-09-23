Attribute VB_Name = "CameraMod"
Function GetMatrixViewOrientation(Camera As ObjectCamera) As Matrix4x4
 Dim VectorVRP As Coordinates4D   ' View Reference Point (VRP) - The world position of the virtual camera AND the virtual film!
 Dim VectorVPN As Coordinates4D   ' View Plane Normal (VPN) - The direction that the virtual camera is pointing "away from"!
 Dim VectorPRP As Coordinates4D   ' Projection Reference Point (PRP), also known as Centre Of Projection (COP) - This is the distance between the virtual camera's film, and the pin-hole lens of the virtual camera.
 Dim VectorVUP As Coordinates4D   ' View UP direction (VUP) - Which way is up? This is used for tilting (or not tilting) the camera.
    
    
    ' Define the View Reference Point (VRP)
    ' This is defined in the World Coordinate (WC) system.
    VectorVRP = Camera.Position
    
    
    ' Subtract the Camera's world position (VRP) from the 'LookingAt' point to give us the View Plane Normal (VPN).
    ' VPN means different things to different 3D packages, ie. PHIGS and OpenGL do not agree on this one.
    ' In this application, the VPN points in the opposite direction that the camera is facing!! I said, Opposite!
    VectorVPN = VectorSubtract(VectorVRP, Camera.ViewPosition)
    If (VectorVPN.X = 0) And (VectorVPN.y = 0) And (VectorVPN.Z = 0) Then
        VectorVPN.X = 0# ' Do not allow VPN to be all zero's (shouldn't happen anyway, but still check)
        VectorVPN.y = 0#
        VectorVPN.Z = 1#
    End If
    
    
    ' PRP is specified in the View Reference Coordinate system (and NOT the world coordinate system)
    VectorPRP = Camera.PRP
'    VectorPRP.x = 0#
'    VectorPRP.y = 0#
'    VectorPRP.z = 1#  ' << Change this value for perspective distortion (any positive value).
'    VectorPRP.w = 1#
    
    
    ' The VUP vector is usually x=0,y=1,z=0. This is used to tilt the camera.
    VectorVUP.X = 0#
    VectorVUP.y = 1#
    VectorVUP.Z = 0#
    VectorVUP.W = 1#
    
    ' ============================================================================
    ' View Orientation.
    ' ============================================================================
    GetMatrixViewOrientation = MatrixViewOrientation(VectorVPN, VectorVUP, VectorVRP)
End Function

Function ResetCamera() As ObjectCamera
 
 ResetCamera.Position.X = 25#
 ResetCamera.Position.y = 15#
 ResetCamera.Position.Z = 25#
 ResetCamera.Position.W = 1#
 
 ResetCamera.ViewMatrix = MatrixIdentity()
 
 ResetCamera.ViewPosition.X = 15#
 ResetCamera.ViewPosition.y = 15#
 ResetCamera.ViewPosition.Z = 15#
 ResetCamera.ViewPosition.W = 1#
       
 ResetCamera.VUP.X = 1#
 ResetCamera.VUP.y = 0#
 ResetCamera.VUP.Z = 0#
 ResetCamera.VUP.W = 1#
        
 ResetCamera.PRP.X = 0#
 ResetCamera.PRP.y = 0#
 ResetCamera.PRP.Z = 1#
 ResetCamera.PRP.W = 1#
        
 ResetCamera.MinScreen.X = 1#
 ResetCamera.MinScreen.y = 1#
 ResetCamera.MaxScreen.X = 250#
 ResetCamera.MaxScreen.y = 250#
        

 ResetCamera.ScaleSize = 100#
 
 ResetCamera.ClipNear = -100
 ResetCamera.ClipFar = -1
End Function
