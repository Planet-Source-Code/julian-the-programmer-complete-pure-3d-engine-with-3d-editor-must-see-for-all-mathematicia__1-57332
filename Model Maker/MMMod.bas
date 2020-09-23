Attribute VB_Name = "MMMod"
Public Mesh As Object3DMesh
Public Cam As ObjectCamera
Public Trig As Integer
Public Coord As Integer

Function OpenModelFile(FileName As String) As Object3DMesh
 Dim FilePointer As Long
 Dim FileLength As Long
 FilePointer = MemAlloc(OpenFile(FileName))
 FileLength = Len(RetMemory(FilePointer))
 
 Dim Triangles As Integer
 Dim Textures As Integer
 
 Dim OpenBMP As IPictureDisp
 
 Dim TriangleStart As Long
 Dim LastLine As Long
 
 
 Dim CurrentTriangle As Integer
 Dim IType As Integer
 
 Dim PhytonSign As Long
 Dim FileStartSign As Long
 
 Dim I As Long
 Dim A As Long
 
 For I = 1 To FileLength
  If Mid(RetMemory(FilePointer), I, 1) = "!" Then
   FileStartSign = I + 1
  End If
  If Mid(RetMemory(FilePointer), I, 1) = "-" Then
   If Not FileStartSign = 0 Then
    Triangles = Val(Mid(RetMemory(FilePointer), FileStartSign, I))
    PhytonSign = I + 1
   End If
  End If
  If Mid(RetMemory(FilePointer), I, 2) = Chr(13) & Chr(10) Then
   If Not PhytonSign = 0 Then
    Textures = Val(Mid(RetMemory(FilePointer), PhytonSign, I - PhytonSign))
    I = FileLength
   End If
  End If
 Next
 
 OpenModelFile = ResetMesh()
 
 ReDim OpenModelFile.Triangle(Triangles)
 OpenModelFile.Triangles = Triangles
 
 For I = 1 To FileLength
  If Mid(RetMemory(FilePointer), I, 1) = "{" Then
   CurrentTriangle = CurrentTriangle + 1
   OpenModelFile.Triangle(CurrentTriangle) = ResetTriangle()
   IType = 1
  End If
  If Mid(RetMemory(FilePointer), I, 1) = "}" Then
   IType = 0
  End If
  If Mid(RetMemory(FilePointer), I, 2) = Chr(13) & Chr(10) Then
   If IType = 1 Then
    If Mid(RetMemory(FilePointer), LastLine, 10) = "Coord(1).X" Then
     OpenModelFile.Triangle(CurrentTriangle).Coordinates(1).X = Val(Mid(RetMemory(FilePointer), LastLine + 11, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 10) = "Coord(1).Y" Then
     OpenModelFile.Triangle(CurrentTriangle).Coordinates(1).Y = Val(Mid(RetMemory(FilePointer), LastLine + 11, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 10) = "Coord(1).Z" Then
     OpenModelFile.Triangle(CurrentTriangle).Coordinates(1).Z = Val(Mid(RetMemory(FilePointer), LastLine + 11, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 10) = "Coord(1).W" Then
     OpenModelFile.Triangle(CurrentTriangle).Coordinates(1).W = Val(Mid(RetMemory(FilePointer), LastLine + 11, 10))
     
    ElseIf Mid(RetMemory(FilePointer), LastLine, 10) = "Coord(2).X" Then
     OpenModelFile.Triangle(CurrentTriangle).Coordinates(2).X = Val(Mid(RetMemory(FilePointer), LastLine + 11, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 10) = "Coord(2).Y" Then
     OpenModelFile.Triangle(CurrentTriangle).Coordinates(2).Y = Val(Mid(RetMemory(FilePointer), LastLine + 11, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 10) = "Coord(2).Z" Then
     OpenModelFile.Triangle(CurrentTriangle).Coordinates(2).Z = Val(Mid(RetMemory(FilePointer), LastLine + 11, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 10) = "Coord(2).W" Then
     OpenModelFile.Triangle(CurrentTriangle).Coordinates(2).W = Val(Mid(RetMemory(FilePointer), LastLine + 11, 10))
     
    ElseIf Mid(RetMemory(FilePointer), LastLine, 10) = "Coord(3).X" Then
     OpenModelFile.Triangle(CurrentTriangle).Coordinates(3).X = Val(Mid(RetMemory(FilePointer), LastLine + 11, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 10) = "Coord(3).Y" Then
     OpenModelFile.Triangle(CurrentTriangle).Coordinates(3).Y = Val(Mid(RetMemory(FilePointer), LastLine + 11, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 10) = "Coord(3).Z" Then
     OpenModelFile.Triangle(CurrentTriangle).Coordinates(3).Z = Val(Mid(RetMemory(FilePointer), LastLine + 11, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 10) = "Coord(3).W" Then
     OpenModelFile.Triangle(CurrentTriangle).Coordinates(3).W = Val(Mid(RetMemory(FilePointer), LastLine + 11, 10))
               
     
    ElseIf Mid(RetMemory(FilePointer), LastLine, 2) = "SC" Then
     OpenModelFile.Triangle(CurrentTriangle).SolidColor = Val(Mid(RetMemory(FilePointer), LastLine + 3, 10))
   
    ElseIf Mid(RetMemory(FilePointer), LastLine, 13) = "TexCoord(1).X" Then
     OpenModelFile.Triangle(CurrentTriangle).TextureFirstPosition.X = Val(Mid(RetMemory(FilePointer), LastLine + 14, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 13) = "TexCoord(1).Y" Then
     OpenModelFile.Triangle(CurrentTriangle).TextureFirstPosition.Y = Val(Mid(RetMemory(FilePointer), LastLine + 14, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 13) = "TexCoord(2).X" Then
     OpenModelFile.Triangle(CurrentTriangle).TextureSecondPosition.X = Val(Mid(RetMemory(FilePointer), LastLine + 14, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 13) = "TexCoord(2).Y" Then
     OpenModelFile.Triangle(CurrentTriangle).TextureSecondPosition.Y = Val(Mid(RetMemory(FilePointer), LastLine + 14, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 13) = "TexCoord(3).X" Then
     OpenModelFile.Triangle(CurrentTriangle).TextureThirdPosition.X = Val(Mid(RetMemory(FilePointer), LastLine + 14, 10))
    ElseIf Mid(RetMemory(FilePointer), LastLine, 13) = "TexCoord(3).Y" Then
     OpenModelFile.Triangle(CurrentTriangle).TextureThirdPosition.Y = Val(Mid(RetMemory(FilePointer), LastLine + 14, 10))
    
    ElseIf Mid(RetMemory(FilePointer), LastLine, 6) = "TexBmp" Then
'     Set OpenBMP = Nothing
     
'     OpenBMP = LoadPicture(Mid(RetMemory(FilePointer), LastLine + 7, I - LastLine - 7))
     If Not OpenModelFile.Triangle(I).Texture.TextureWidth = 0 And Not OpenModelFile.Triangle(I).Texture.TextureHeight = 0 Then
      OpenModelFile.Triangle(I).Texture.TextureHdc = CreateHdc(OpenModelFile.Triangle(I).Texture.TextureWidth, OpenModelFile.Triangle(I).Texture.TextureHeight)
      OpenModelFile.Triangle(I).Texture.TextureHdc = GetCurrentHdc()
      DoEvents
      CopyHdc OpenBMP.Handle, OpenModelFile.Triangle(I).Texture.TextureHdc
     End If
    
    ElseIf Mid(RetMemory(FilePointer), LastLine, 8) = "TexWidth" Then
     OpenModelFile.Triangle(CurrentTriangle).Texture.TextureWidth = Val(Mid(RetMemory(FilePointer), LastLine + 8, 10))
    
    ElseIf Mid(RetMemory(FilePointer), LastLine, 9) = "TexHeight" Then
     OpenModelFile.Triangle(CurrentTriangle).Texture.TextureHeight = Val(Mid(RetMemory(FilePointer), LastLine + 10, 10))
    End If
   End If
   LastLine = I + 2
  End If
 Next
 
 FreeMemory FilePointer
End Function

Function SaveModelFile(FileName As String, Mesh3D As Object3DMesh) As Boolean
 Dim FileString As String
 Dim L As String
 Dim BMPSave As IPictureDisp
 
 L = Chr(13) & Chr(10)
 
 FileString = "!" & Mesh3D.Triangles & "-0" & L
 
 For I = 1 To Mesh3D.Triangles
  FileString = FileString & "{" & L

  FileString = FileString & "Coord(1).X " & Mesh3D.Triangle(I).Coordinates(1).X & L
  FileString = FileString & "Coord(1).Y " & Mesh3D.Triangle(I).Coordinates(1).Y & L
  FileString = FileString & "Coord(1).Z " & Mesh3D.Triangle(I).Coordinates(1).Z & L
  FileString = FileString & "Coord(1).W " & Mesh3D.Triangle(I).Coordinates(1).W & L
  
  FileString = FileString & L
  
  FileString = FileString & "Coord(2).X " & Mesh3D.Triangle(I).Coordinates(2).X & L
  FileString = FileString & "Coord(2).Y " & Mesh3D.Triangle(I).Coordinates(2).Y & L
  FileString = FileString & "Coord(2).Z " & Mesh3D.Triangle(I).Coordinates(2).Z & L
  FileString = FileString & "Coord(2).W " & Mesh3D.Triangle(I).Coordinates(2).W & L

  FileString = FileString & L

  FileString = FileString & "Coord(3).X " & Mesh3D.Triangle(I).Coordinates(3).X & L
  FileString = FileString & "Coord(3).Y " & Mesh3D.Triangle(I).Coordinates(3).Y & L
  FileString = FileString & "Coord(3).Z " & Mesh3D.Triangle(I).Coordinates(3).Z & L
  FileString = FileString & "Coord(3).W " & Mesh3D.Triangle(I).Coordinates(3).W & L
  
  FileString = FileString & L
  
  FileString = FileString & "SC " & Mesh3D.Triangle(I).SolidColor & L
  
  FileString = FileString & L
  
  FileString = FileString & "TexCoord(1).X " & Mesh3D.Triangle(I).TextureFirstPosition.X & L
  FileString = FileString & "TexCoord(1).Y " & Mesh3D.Triangle(I).TextureFirstPosition.Y & L
  
  FileString = FileString & L
  
  FileString = FileString & "TexCoord(2).X " & Mesh3D.Triangle(I).TextureSecondPosition.X & L
  FileString = FileString & "TexCoord(2).Y " & Mesh3D.Triangle(I).TextureSecondPosition.Y & L
  
  FileString = FileString & L
  
  FileString = FileString & "TexCoord(3).X " & Mesh3D.Triangle(I).TextureThirdPosition.X & L
  FileString = FileString & "TexCoord(3).Y " & Mesh3D.Triangle(I).TextureThirdPosition.Y & L

  FileString = FileString & L
  
  FileString = FileString & "TexWidth " & Mesh3D.Triangle(I).Texture.TextureWidth & L
  FileString = FileString & "TexHeight " & Mesh3D.Triangle(I).Texture.TextureWidth & L

  If Not Mesh3D.Triangle(I).Texture.TextureHdc = 0 Then
   Set BMPSave = New StdPicture
   Dim ca As PictureBox
   Set ca = Main.PicXY
   CopyHdc Mesh3D.Triangle(I).Texture.TextureHdc, BMPSave.Handle
   DoEvents
   SavePicture ca.Image, GetFolder(FileName) & "TG" & I & ".bmp"
   FileString = FileString & "TexBmp " & GetFolder(FileName) & "TG" & I & ".bmp" & L
  End If
  
  FileString = FileString & "}" & L & L
 Next
 
 Open FileName For Binary Access Write As #1
  Put #1, , FileString
 Close #1
End Function

Function GetFolder(FileName As String) As String
 Dim I As Integer
 Dim A As Integer
 For I = 1 To Len(FileName)
  If Mid(FileName, I, 1) = "/" Then
  A = I
  ElseIf Mid(FileName, I, 1) = "\" Then
  A = I
  End If
 Next
 GetFolder = Mid(FileName, 1, A)
End Function

Function CleanUp()
 Dim I As Integer
 For I = 1 To Mesh.Triangles
  If Not Mesh.Triangle(I).Texture.TextureHdc = 0 Then
   DeleteHdc Mesh.Triangle(I).Texture.TextureHdc
  End If
  Mesh.Triangle(I) = ResetTriangle()
 Next
 Mesh.Triangles = 0
 ReDim Mesh.Triangle(0)
End Function
