Attribute VB_Name = "DrawingMod"
'Backbuffer
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal HDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal HDc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal HDc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal HDc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function FillRect Lib "user32" (ByVal HDc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long


'Colors
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

'Basic
Private Declare Function LineTo Lib "gdi32" (ByVal HDc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal HDc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpPoint As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal HDc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal HDc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal HDc As Long, ByVal X As Long, ByVal Y As Long) As Long

'Types

Private Type BITMAP
 bmType As Long
 bmWidth As Long
 bmHeight As Long
 bmWidthBytes As Long
 bmPlanes As Long
 bmBitsPixel As Integer
 bmBits As Long
End Type

Enum FillMode
 Wireframe = 1
 Solid = 2
 Texture = 3
End Enum

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private CurrentHdc As Long
Private CurrentBMP As Long
Private OldBMP As Long

Private Const PS_SOLID = 0

Function CreateHdc(Width As Long, Height As Long) As Long
 Dim lHdcC As Long
  lHdcC = CreateDC("DISPLAY", "", "", ByVal 0&)
  If Not lHdcC = 0 Then
   CurrentHdc = CreateCompatibleDC(lHdcC)
   If Not CurrentHdc = 0 Then
    CurrentBMP = CreateCompatibleBitmap(lHdcC, Width, Height)
    If Not CurrentBMP = 0 Then
     OldBMP = SelectObject(CurrentHdc, CurrentBMP)
     If Not OldBMP = 0 Then
      DeleteDC lHdcC
      CreateHdc = CurrentHdc
      Exit Function
     End If
    End If
   End If
  DeleteDC lHdcC
 End If
End Function

Function DeleteHdc(HDc As Long) As Long
 DeleteDC HDc
End Function

Function GetCurrentHdc() As Long
 GetCurrentHdc = CurrentHdc
End Function

Function DrawHdcOnHdc(SourceHdc As Long, DestinationHdc As Long, Width As Long, Height As Long, xDst As Long, yDst As Long, xSrc As Long, ySrc As Long)
 BitBlt DestinationHdc, xDst, yDst, Width, Height, SourceHdc, xSrc, ySrc, vbSrcCopy
End Function

Function ClearHdc(HDc As Long, Width As Long, Heigth As Long)
 Dim hBr As Long
 Dim RectDraw As RECT
 RectDraw.Bottom = 1
 RectDraw.Left = 1
 RectDraw.Right = Width
 RectDraw.Top = Heigth
 hBr = CreateSolidBrush(vbBlack) '&HF0000015 And &H1F& 'GetSysColorBrush(&HF0000015 And &H1F&)
 FillRect HDc, RectDraw, hBr
 DeleteObject hBr
End Function

Public Sub DrawCurrentHdc(ByVal HDc As Long, Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0)
   If WidthSrc <= 0 Then WidthSrc = 800
   If HeightSrc <= 0 Then HeightSrc = 640
   BitBlt HDc, xDst, yDst, WidthSrc, HeightSrc, GetCurrentHdc(), xSrc, ySrc, vbSrcCopy
End Sub

Public Sub Draw( _
      ByVal HDc As Long, Optional SrcHdc, _
      Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
      Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
      Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0 _
   )
   If WidthSrc <= 0 Then WidthSrc = 800
   If HeightSrc <= 0 Then HeightSrc = 640
   If SrcHdc = 0 Then
    BitBlt HDc, xDst, yDst, WidthSrc, HeightSrc, CurrentHdc, xSrc, ySrc, vbSrcCopy
   Else
    BitBlt HDc, xDst, yDst, WidthSrc, HeightSrc, SrcHdc, xSrc, ySrc, vbSrcCopy
   End If
End Sub

Public Sub CopyHdc( _
      ByVal HDc As Long, Optional DestHdc As Long, _
      Optional ByVal xSrc As Long = 0, Optional ByVal ySrc As Long = 0, _
      Optional ByVal WidthSrc As Long = 0, Optional ByVal HeightSrc As Long = 0, _
      Optional ByVal xDst As Long = 0, Optional ByVal yDst As Long = 0 _
   )
   If WidthSrc <= 0 Then WidthSrc = 800
   If HeightSrc <= 0 Then HeightSrc = 640
   If DestHdc = 0 Then
    BitBlt CurrentHdc, xDst, yDst, WidthSrc, HeightSrc, HDc, xSrc, ySrc, vbSrcCopy
   Else
    BitBlt DestHdc, xDst, yDst, WidthSrc, HeightSrc, HDc, xSrc, ySrc, vbSrcCopy
   End If
End Sub

'

Function PrintText(Text As String, X As Long, Y As Long, HDc As Long)
 TextOut HDc, X, Y, Text, Len(Text)
End Function

Function DrawLine(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, HDc As Long)
 MoveToEx HDc, X1, Y1, 0
 LineTo HDc, X2, Y2
End Function

Function DrawLineScaled(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long, Scallation As Integer, minx As Integer, MinY As Integer, HDc As Long)
 MoveToEx HDc, (X1 * Scallation) + minx, (Y1 * Scallation) + MinY, 0
 LineTo HDc, (X2 * Scallation) + minx, (Y2 * Scallation) + MinY
End Function

Function FillSolidTriangle(Color As Long, FirstVector As Coordinates2D, SecondVector As Coordinates2D, ThirdVector As Coordinates2D, Max As Long, HDc As Long)
 Dim A As Single
 Dim B As Single
 
 Dim N As Long
 Dim M As Long
 
 Dim SmallX As Integer
 Dim BigX As Integer
 Dim SmallY As Integer
 Dim BigY As Integer
 
 SmallX = FirstVector.X
 If SmallX > SecondVector.X Then SmallX = SecondVector.X
 If SmallX > ThirdVector.X Then SmallX = ThirdVector.X
 If SmallX < 0 Then SmallX = 0

 SmallY = FirstVector.Y
 If SmallY > SecondVector.Y Then SmallY = SecondVector.Y
 If SmallY > ThirdVector.Y Then SmallY = ThirdVector.Y
 If SmallY < 0 Then SmallY = 0

 BigX = FirstVector.X
 If BigX < SecondVector.X Then BigX = SecondVector.X
 If BigX < ThirdVector.X Then BigX = ThirdVector.X
 If BigX > Max Then BigX = Max
 
 BigY = FirstVector.Y
 If BigY < SecondVector.Y Then BigY = SecondVector.Y
 If BigY < ThirdVector.Y Then BigY = ThirdVector.Y
 If BigY > Max Then BigY = Max
  
 Dim GC As Long
 
  
  For A = SmallX To BigX
   For B = SmallY To BigY
    If IsInTriangle2D(Make2DCoordinate(A, B), FirstVector, SecondVector, ThirdVector) = True Then
        SetPixelV HDc, Round(A), Round(B), Color
    End If
   Next
  Next
End Function

Function FillTextureTriangle(Texture As ObjectTexture, FirstVector As Coordinates2D, SecondVector As Coordinates2D, ThirdVector As Coordinates2D, Max As Long, HDc As Long) 'Optional UsePerspectiveTexturing As Boolean = False, Optional Triangle3D As ObjectTriangle)
 Dim A As Single
 Dim B As Single
 
 Dim N As Long
 Dim M As Long
 
 Dim SmallX As Integer
 Dim BigX As Integer
 Dim SmallY As Integer
 Dim BigY As Integer
 
 SmallX = FirstVector.X
 If SmallX > SecondVector.X Then SmallX = SecondVector.X
 If SmallX > ThirdVector.X Then SmallX = ThirdVector.X
 If SmallX < 0 Then SmallX = 0

 SmallY = FirstVector.Y
 If SmallY > SecondVector.Y Then SmallY = SecondVector.Y
 If SmallY > ThirdVector.Y Then SmallY = ThirdVector.Y
 If SmallY < 0 Then SmallY = 0

 BigX = FirstVector.X
 If BigX < SecondVector.X Then BigX = SecondVector.X
 If BigX < ThirdVector.X Then BigX = ThirdVector.X
 If BigX > Max Then BigX = Max
 
 BigY = FirstVector.Y
 If BigY < SecondVector.Y Then BigY = SecondVector.Y
 If BigY < ThirdVector.Y Then BigY = ThirdVector.Y
 If BigY > Max Then BigY = Max
  
 Dim GC As Long
 
  
  For A = SmallX To BigX
   For B = SmallY To BigY
    If IsInTriangle2D(Make2DCoordinate(A, B), FirstVector, SecondVector, ThirdVector) = True Then
     N = Abs((Texture.TextureWidth / (GetXByYInLine(FirstVector.X, FirstVector.Y, SecondVector.X, SecondVector.Y, B) - GetXByYInLine(SecondVector.X, SecondVector.Y, ThirdVector.X, ThirdVector.Y, (B / Texture.TextureWidth)))) * (A / Texture.TextureWidth))
     M = Abs((Texture.TextureHeight / (GetXByYInLine(FirstVector.Y, FirstVector.X, SecondVector.Y, SecondVector.X, A) - GetXByYInLine(SecondVector.Y, SecondVector.X, ThirdVector.Y, ThirdVector.X, (A / Texture.TextureHeight)))) * (B / Texture.TextureHeight))
     
     GC = GetPixel(Texture.TextureHdc, N, M)
'     If GC = 0 Then GC = &HFFFFFF
     
'     SetPixelV Hdc, Round(A), Round(B), GC
    End If
   Next
  Next
End Function

Function ChangeForecolor(HDc As Long, Forecolor As Long)
 Dim hPen As Long
 Dim hPenOld As Long
 hPen = CreatePen(PS_SOLID, 1, Forecolor)
 hPenOld = SelectObject(HDc, hPen)
End Function

Function DrawGrid(GridsX As Long, Width As Long, GridsY As Long, Height As Long, HDc As Long)
 Dim I As Integer
 For I = 1 To GridsX
  DrawLine (I * (Width / GridsX)), 1, (I * (Width / GridsX)), Height, HDc
 Next
 I = 0
 For I = 1 To GridsX
  DrawLine 1, (I * (Height / GridsY)), Width, (I * (Height / GridsY)), HDc
 Next
End Function
