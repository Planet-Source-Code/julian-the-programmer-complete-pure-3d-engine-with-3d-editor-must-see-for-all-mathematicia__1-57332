Attribute VB_Name = "IPALibrary"
'Brightness Demo ©2002 by Tanner "DemonSpectre" Helland

'Source code for "Graphics Programming in Visual Basic - Part 3: Advanced API Pixel Routines"

'This simple program demonstrates how to adjust an image's brightness using the API calls of
'GetDIBits and StretchDIBits.  As a good exercise, try rewriting this program using the
'GetBitmapBits and SetBitmapBits calls (they can be found back in the tutorial).  This demonstrates
'some fast graphics, but they can be made even faster!  Read Tutorial 4 for more information
'about optimizing graphics functions (or download my Brightness program from
'VacantBrains.tripod.com

'The CG graphic in the picture box is ©1998 by SquareSoft
'(it's from Final Fantasy VIII, if you care)

'For additional cool code, check out my website at
'http://tannerhelland.tripod.com/VBStuff.htm

'All of the DIB types
Private Type BITMAP
bmType As Long
bmWidth As Long
bmHeight As Long
bmWidthBytes As Long
bmPlanes As Integer
bmBitsPixel As Integer
bmBits As Long
End Type
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
 
Private Type RGBQUAD
rgbBlue As Byte
rgbGreen As Byte
rgbRed As Byte
rgbAlpha As Byte
End Type
 
Private Type BITMAPINFOHEADER
bmSize As Long
bmWidth As Long
bmHeight As Long
bmPlanes As Integer
bmBitCount As Integer
bmCompression As Long
bmSizeImage As Long
bmXPelsPerMeter As Long
bmYPelsPerMeter As Long
bmClrUsed As Long
bmClrImportant As Long
End Type
 
Private Type BITMAPINFO
bmHeader As BITMAPINFOHEADER
bmColors(0 To 255) As RGBQUAD
End Type

Type POINTAPI
    x As Long
    y As Long
End Type

'The magical API DIB function calls (they're long!)
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dWidth As Long, ByVal dHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long, ByVal RasterOp As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

'The array that will hold our pixel data
Dim ImageData() As Byte

'Temporary brightness variable
Dim tBrightness As Single

Type RGBType
 R As Byte
 G As Byte
 B As Byte
End Type

Function GetRGB(ByVal Color As Long) As RGBType
  Dim sColor As String

  sColor = Right("000000" & Hex(Color), 6)
  With GetRGB
    .R = Val("&h" & Right(sColor, 2))
    .G = Val("&h" & Mid(sColor, 3, 2))
    .B = Val("&h" & Left(sColor, 2))
  End With
End Function

'A simple subroutine that will change the brightness of a picturebox using simple API routines.
Sub DrawBrightness(DstPicture As PictureBox, SrcPicture As PictureBox, ByVal Brightness As Single)
    'Coordinate variables
    Dim x As Long, y As Long
    'Build a look-up table for all possible brightness values
    Dim bTable(0 To 255) As Long
    Dim TempColor As Long
    For x = 0 To 255
        'Calculate the brightness for pixel value x
        TempColor = x * Brightness
        'Make sure that the calculated value is between 0 and 255 (so we don't get an error)
        ByteMe TempColor
        'Place the corrected value into its array spot
        bTable(x) = TempColor
    Next x
    'Get the pixel data into our ImageData array
    GetImageData SrcPicture, ImageData()
    'Temporary width and height variables are faster than accessing the Scale properties over and over again
    Dim TempWidth As Long, TempHeight As Long
    TempWidth = DstPicture.ScaleWidth - 1
    TempHeight = DstPicture.ScaleHeight
    'run a loop through the picture to change every pixel
    For x = 0 To TempWidth
    For y = 0 To TempHeight
        'Use the values in the look-up table to quickly change the brightness values
        'of each color.  The look-up table is much faster than doing the math
        'over and over for each individual pixel.
        ImageData(2, x, y) = bTable(ImageData(2, x, y))   'Change the red
        ImageData(1, x, y) = bTable(ImageData(1, x, y))   'Change the green
        ImageData(0, x, y) = bTable(ImageData(0, x, y))   'Change the blue
    Next y
        'refresh the picture box every 25 lines (a nice progress bar effect if AutoRedraw is set)
        If DstPicture.AutoRedraw = True And (x Mod 25) = 0 Then
            SetImageData DstPicture, ImageData()
        End If
    Next x
    'final picture refresh
    SetImageData DstPicture, ImageData()
End Sub

'Routine to get an image's pixel information into an array dimensioned (rgb, x, y)
Sub GetImageData(ByRef SrcPictureBox As PictureBox, ByRef ImageData() As Byte)
    'Declare us some variables of the necessary bitmap types
    Dim bm As BITMAP
    Dim bmi As BITMAPINFO
    'Now we fill up the bmi (Bitmap information variable) with all of the appropriate data
    bmi.bmHeader.bmSize = 40 'Size, in bytes, of the header (always 40)
    bmi.bmHeader.bmPlanes = 1 'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24 'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0 'Compression: standard/none or RLE
    'Calculate the size of the bitmap type (in bytes)
    Dim bmLen As Long
    bmLen = Len(bm)
    'Get the picture box information from SrcPictureBox and put it into our 'bm' variable
    GetObject SrcPictureBox.Image, bmLen, bm
    'Build a correctly sized array
    ReDim ImageData(0 To 2, 0 To bm.bmWidth - 1, 0 To bm.bmHeight)
    'Finish building the 'bmi' variable we want to pass to the GetDIBits call (the same one we used above)
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    'Now that we've completely filled up the 'bmi' variable, we use GetDIBits to take the data from
    'SrcPictureBox and put it into the ImageData() array using the settings we specified in 'bmi'
    GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, ImageData(0, 0, 0), bmi, 0
End Sub

'Standardized routine for converting to absolute byte values
Sub ByteMe(ByRef TempVar As Long)
    If TempVar > 255 Then TempVar = 255: Exit Sub
    If TempVar < 0 Then TempVar = 0: Exit Sub
End Sub

