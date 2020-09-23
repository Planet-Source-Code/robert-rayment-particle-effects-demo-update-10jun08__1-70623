Attribute VB_Name = "ModPublics"
'ModPublics.bas

Option Explicit

Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
'Public Const SM_CXSCREEN = 0 'X Size of screen
'Public Const SM_CYSCREEN = 1 'Y Size of Screen
Public Const SM_CYCAPTION = 4 'Height of windows caption
Public Const SM_CYMENU = 15 'Height of menu
Public Const SM_CXBORDER = 5 'Width of no-sizable borders
Public Const SM_CYBORDER = 6 'Height of non-sizable borders
'Public Const SM_CXDLGFRAME = 7 'Width of dialog box borders
'Public Const SM_CYDLGFRAME = 8 'Height of dialog box borders
'Public Const SM_CYSMCAPTION = 51 'Height of windows 95 small caption


Public Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" _
 (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public PicInfo As BITMAP

Public Declare Function GetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long) As Long
Public Declare Function SetStretchBltMode Lib "gdi32" _
   (ByVal hdc As Long, ByVal nStretchMode As Long) As Long

'Public Const COLORONCOLOR = 3
Public Const HALFTONE = 4

Public Type BITMAPINFOHEADER ' 40 bytes
   biSize As Long
   biWidth As Long
   biHeight As Long
   biPlanes As Integer
   biBitCount As Integer
   biCompression As Long
   biSizeImage As Long
   biXPelsPerMeter As Long
   biYPelsPerMeter As Long
   biClrUsed As Long
   biClrImportant As Long
End Type
Public BHI As BITMAPINFOHEADER

Public Declare Function StretchDIBits Lib "gdi32.dll" _
   (ByVal hdc As Long, _
   ByVal X As Long, ByVal Y As Long, _
   ByVal dx As Long, ByVal dy As Long, _
   ByVal SrcX As Long, ByVal SrcY As Long, _
   ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, _
   ByRef lpBits As Any, _
   ByRef BInfo As BITMAPINFOHEADER, _
   ByVal wUsage As Long, _
   ByVal dwRop As Long) As Long

' For fitting a bitmap into picturebox Pic(1) from Pic(0)
Public Declare Function StretchBlt Lib "gdi32" _
   (ByVal hdc As Long, ByVal X As Long, _
    ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
    ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
' rs&=StretchBlt(hdcD,xD,yD,wiD,htD,hdcS,xS,yS,wiS,htS,SRCCOPY)

Public Const SRCCOPY = &HCC0020

Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Public Declare Function GetDIBits Lib "gdi32" _
   (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, _
    ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFOHEADER, _
    ByVal wUsage As Long) As Long

Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'Public Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" _
'   (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Source data
Public picDATAORG() As Long
' Destination data to Display
Public picDATA() As Long

' TwipsPerPixel
Public STX As Long, STY As Long

Public Const pi# = 3.14159265
Public Const d2r# = pi# / 180

Public Sub DISPLAY(PIC As PictureBox, DATARRAY() As Long)
' Public BHI As BITMAPINFOHEADER
' DATARRAY(0,0) ' 0 To W-1,0 To H-1)
Dim GetMode As Long
' Offsets
Dim xlo As Long, ylo As Long
   xlo = 0
   ylo = 0
   GetMode = GetStretchBltMode(PIC.hdc)

   SetStretchBltMode PIC.hdc, HALFTONE
   
   With BHI
      .biSize = 40
      .biPlanes = 1
      .biWidth = PIC.Width 'picwidth
      .biHeight = PIC.Height 'picheight
      .biBitCount = 32
   End With
   
   Call StretchDIBits(PIC.hdc, _
   0, 0, _
   PIC.Width, PIC.Height, _
   xlo, ylo, _
   PIC.Width, PIC.Height, _
   DATARRAY(0, 0), _
   BHI, 0, vbSrcCopy)
   
   SetStretchBltMode PIC.hdc, GetMode
   
   PIC.Refresh
   'Pic.Picture = Pic.Image
End Sub

Public Sub LngToRGB(LCul As Long, R As Byte, G As Byte, B As Byte)
   R = LCul And &HFF&
   G = (LCul And &HFF00&) \ &H100&
   B = (LCul And &HFF0000) \ &H10000
End Sub

' Not used
'Public Function zATan2(Y As Single, X As Single) As Single
'' Const pi# = 3.14159265
'' Input:  deltay(Y),deltax(X) - real
'' Output: atan(Y/X) for -pi#/2 to pi#/2
'   If X <> 0 Then
'      zATan2 = Atn(Y / X)
'      If (X < 0) Then
'         If (Y < 0) Then zATan2 = zATan2 - pi# Else zATan2 = zATan2 + pi#
'      End If
'   Else  ' x=0
'      If Abs(Y) > 0 Then   ' Must be an overflow
'         If Y > 0 Then zATan2 = pi# / 2 Else zATan2 = -pi# / 2
'      Else  ' 0/0
'         zATan2 = 0
'      End If
'   End If
'End Function

