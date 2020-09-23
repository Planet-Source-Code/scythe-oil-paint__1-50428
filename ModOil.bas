Attribute VB_Name = "ModOil"
Option Explicit

'Oil Paining Module
'© Scythe 2003
'Created for LM-X

Private Type BITMAP
 bmType As Long
 bmWidth As Long
 bmHeight As Long
 bmWidthBytes As Long
 bmPlanes As Integer
 bmBitsPixel As Integer
 bmBits As Long
End Type

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

Private Type RGBQUAD
 rgbBlue As Byte
 rgbGreen As Byte
 rgbRed As Byte
 rgbReserved As Byte
End Type

'ColorTypes
Public Type RGBcolors
 Blue As Byte
 Green As Byte
 Red As Byte
End Type
Dim PicInfo As BITMAP

'Oil Painting
'Original from Jason Waltman
'He wrote the C++ version of this code
'(Great Idea but to slow thru some bugs)

'Big THANKS to Robert Raimond
'who gave me the his Page
'He´s also one of the best coders on PSC !

'This code need to compiled to be fast !!!!!!

Public Sub PicOilPaint(PicBox As PictureBox, StartX As Long, StartY As Long, LenghtX As Long, LenghtY As Long, Rad As Long, Maxintensity As Long)
 Dim x As Long
 Dim y As Long
 Dim i As Long
 Dim f As Long
 
 Dim col As Long
 Dim PicAr1() As RGBQUAD
 Dim PicAr2() As RGBQUAD
 Dim PicAr3() As RGBQUAD
 
 Dim r As Long
 Dim Skale As Double
 
 Dim IntensityCount(255) As Long
 Dim Intensity As Long
 Dim AvColor(255) As RGBcolors
 
 Skale = Maxintensity / 255
 
 Pic2Array PicBox, PicAr1 'The intensity Picture
 Pic2Array PicBox, PicAr2 'The Pictures Source
 Pic2Array PicBox, PicAr3 'The Result Picture
 
'Create a BW Picture (Intensity Picture)
 For x = StartX To StartX + LenghtX
  For y = StartY To StartY + LenghtY
   'calculate the Colors
   'Red * 0,3 + Green * 0,59 + Blue * 0,11 gives us the Graycolor
   'The Maximum result is 255
   col = 0.3 * CLng(PicAr1(x, y).rgbRed) + 0.59 * CLng(PicAr1(x, y).rgbGreen) + 0.11 * CLng(PicAr1(x, y).rgbBlue)
   PicAr1(x, y).rgbBlue = col * Skale
  Next y
 Next x

 For x = StartX To StartX + LenghtX
  For y = StartY To StartY + LenghtY

   
   'Find the Most frequent color in this block
   For i = x - Rad To x + Rad
    For f = y - Rad To y + Rad
     If i > StartX And i < LenghtX Then
      If f > StartY And f < LenghtY Then
        Intensity = PicAr1(i, f).rgbBlue
        IntensityCount(Intensity) = IntensityCount(Intensity) + 1
        
        If IntensityCount(Intensity) = 1 Then
          AvColor(Intensity).Red = PicAr2(i, f).rgbRed
          AvColor(Intensity).Green = PicAr2(i, f).rgbGreen
          AvColor(Intensity).Blue = PicAr2(i, f).rgbBlue
        End If
      End If
     End If
    Next f
   Next i
  
  'Now find the Most frequent color
  Intensity = 0
  f = 0
  For i = 0 To Maxintensity
   If IntensityCount(i) > f Then
    Intensity = i
    f = IntensityCount(i)
   End If
   IntensityCount(i) = 0
  Next i
  
  'Set it as new color
  PicAr3(x, y).rgbRed = AvColor(Intensity).Red '\ f
  PicAr3(x, y).rgbGreen = AvColor(Intensity).Green ' \ f
  PicAr3(x, y).rgbBlue = AvColor(Intensity).Blue '\ f
 Next y
Next x

Array2Pic PicBox, PicAr3

End Sub


'Get a Picture as Array
Private Sub Pic2Array(PicBox As PictureBox, ByRef PicArray() As RGBQUAD)
 GetObject PicBox.Image, Len(PicInfo), PicInfo
 ReDim PicArray(0 To PicInfo.bmWidth - 1, 0 To PicInfo.bmHeight - 1) As RGBQUAD
 GetBitmapBits PicBox.Image, PicInfo.bmWidth * PicInfo.bmHeight * 4, PicArray(0, 0)
End Sub

'Write a Array to a Picture
Private Sub Array2Pic(PicBox As PictureBox, ByRef PicArray() As RGBQUAD)
 GetObject PicBox.Image, Len(PicInfo), PicInfo
 SetBitmapBits PicBox.Image, PicInfo.bmWidth * PicInfo.bmHeight * 4, PicArray(0, 0)
End Sub
