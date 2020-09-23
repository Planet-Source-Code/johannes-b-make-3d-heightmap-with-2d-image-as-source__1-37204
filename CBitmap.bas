Attribute VB_Name = "CBitmap"
'Autor: ALKO
'e-mail: alfred.koppold@freenet.de

Option Explicit

' Constants
Private Const SRCCOPY = &HCC0020
Private Const BI_RGB = 0&
Private Const CBM_INIT = &H4
Private Const DIB_RGB_COLORS = 0
' Types
Public Type RGBTriple
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Private Type BITMAPINFOHEADER
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

Private Type RGBQUAD
rgbBlue As Byte
rgbGreen As Byte
rgbRed As Byte
rgbReserved As Byte
End Type

Private Type BITMAPINFO_1
bmiHeader As BITMAPINFOHEADER
bmiColors(1) As RGBQUAD
End Type
Private Type BITMAPINFO_4
bmiHeader As BITMAPINFOHEADER
bmiColors(15) As RGBQUAD
End Type
Private Type BITMAPINFO_8
bmiHeader As BITMAPINFOHEADER
bmiColors(255) As RGBQUAD
End Type
Private Type BITMAPINFO_16
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type
Private Type BITMAPINFO_24
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

' Functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateDIBitmap_1 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_1, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_4 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_4, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_8 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_8, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_16 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_16, ByVal wUsage As Long) As Long
Private Declare Function CreateDIBitmap_24 Lib "gdi32" Alias "CreateDIBitmap" (ByVal hdc As Long, lpInfoHeader As BITMAPINFOHEADER, ByVal dwUsage As Long, lpInitBits As Any, lpInitInfo As BITMAPINFO_24, ByVal wUsage As Long) As Long



'header
Private bm1 As BITMAPINFO_1
Private bm4 As BITMAPINFO_4
Private bm8 As BITMAPINFO_8
Private bm16 As BITMAPINFO_16
Private bm24 As BITMAPINFO_24
'bitmap handle.
Private hBmp As Long

Private Type ScTw
Width As Long
Height As Long
End Type

Public Sub InitColorTable_1()
bm1.bmiColors(0).rgbRed = 0
bm1.bmiColors(0).rgbGreen = 0
bm1.bmiColors(0).rgbBlue = 0
bm1.bmiColors(0).rgbReserved = 0
bm1.bmiColors(1).rgbRed = 255
bm1.bmiColors(1).rgbGreen = 255
bm1.bmiColors(1).rgbBlue = 255
bm1.bmiColors(1).rgbReserved = 0

End Sub
Public Sub InitColorTable_8(ByteArray() As Byte)
'Construct the palette
'==================================================
    Dim Palette8() As RGBTriple
        ReDim Palette8(UBound(ByteArray) / 3)
        CopyMemory Palette8(0), ByteArray(0), UBound(ByteArray)

    Dim nCount As Long
    
    'Create Palette
    For nCount = 0 To 255
    bm8.bmiColors(nCount).rgbBlue = Palette8(nCount).Blue
    bm8.bmiColors(nCount).rgbGreen = Palette8(nCount).Green
    bm8.bmiColors(nCount).rgbRed = Palette8(nCount).Red
    bm8.bmiColors(nCount).rgbReserved = 0
    Next nCount

End Sub
Public Sub InitColorTable_4(ByteArray() As Byte)
    Dim Palette4() As RGBTriple
        ReDim Palette4(UBound(ByteArray) / 3)
        CopyMemory Palette4(0), ByteArray(0), UBound(ByteArray)

Dim i As Integer
' Create a color table
For i = 0 To 15
bm4.bmiColors(i).rgbRed = Palette4(i).Red
bm4.bmiColors(i).rgbGreen = Palette4(i).Green
bm4.bmiColors(i).rgbBlue = Palette4(i).Blue
bm4.bmiColors(i).rgbReserved = 0
Next i

End Sub


Public Sub CreateBitmap_1(ByteArray() As Byte, BMPWidth As Long, BMPHeight As Long, Orientation As Integer)
' Create a 1bit Bitmap
Dim hdc As Long
With bm1.bmiHeader
.biSize = Len(bm1.bmiHeader)
.biWidth = BMPWidth
        If Orientation = 0 Then
        .biHeight = BMPHeight                    'Bitmap Height, bitmap is top down.
        Else
        .biHeight = -BMPHeight
        End If
.biPlanes = 1
.biBitCount = 1
.biCompression = BI_RGB
.biSizeImage = 0
.biXPelsPerMeter = 0
.biYPelsPerMeter = 0
.biClrUsed = 0
.biClrImportant = 0
End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_1(hdc, bm1.bmiHeader, CBM_INIT, ByteArray(0), bm1, DIB_RGB_COLORS)
End Sub
Public Sub CreateBitmap_4(ByteArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer)
' Creates a device independent bitmap
' from the pixel data in Data().
Dim hdc As Long
With bm4.bmiHeader
.biSize = Len(bm1.bmiHeader)
.biWidth = PicWidth
        If Orientation = 0 Then
        .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
        .biHeight = -PicHeight
        End If
.biPlanes = 1
.biBitCount = 4
.biCompression = BI_RGB
.biSizeImage = 0
.biXPelsPerMeter = 0
.biYPelsPerMeter = 0
.biClrUsed = 0
.biClrImportant = 0
End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_4(hdc, bm4.bmiHeader, CBM_INIT, ByteArray(0), bm4, DIB_RGB_COLORS)
End Sub

Public Sub CreateBitmap_8(BitmapArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer)
' Creates a device independent bitmap
' from the pixel data in BitmapArry().
Dim hdc As Long
With bm8.bmiHeader
.biSize = Len(bm8.bmiHeader)
.biWidth = PicWidth
        If Orientation = 0 Then
        .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
        .biHeight = -PicHeight
        End If
.biPlanes = 1
.biBitCount = 8
.biCompression = BI_RGB
.biSizeImage = 0
.biXPelsPerMeter = 0
.biYPelsPerMeter = 0
.biClrUsed = 0
.biClrImportant = 0
End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_8(hdc, bm8.bmiHeader, CBM_INIT, BitmapArray(0), bm8, DIB_RGB_COLORS)
End Sub

Public Sub DrawBitmap(PicWidth As Long, PicHeight As Long, PicObject As Object, Scalierung As Boolean)
Dim cDC As Long
Dim a As Long
Dim Übergabe As ScTw
Dim realheight As Long
Dim realwidth As Long
PicObject.Cls
a = PicObject.ScaleMode
PicObject.ScaleMode = 1
Select Case Scalierung
Case True
Übergabe = PixelToTwips(PicWidth, PicHeight)
PicObject.Height = Übergabe.Height / 15
PicObject.Width = Übergabe.Width / 15
Case False
End Select

'If PicObject.Height <> PicObject.ScaleHeight Then 'with Boarders

'Übergabe = Twipstopixel(PicObject.Width, PicObject.Height)
'realheight = Übergabe.Height
'realwidth = Übergabe.Width
'PicObject.Height = PicObject.Height + (PicObject.Height - PicObject.ScaleHeight)
'PicObject.Width = PicObject.Width + (PicObject.Width - PicObject.ScaleWidth)

'Else
PicObject.ScaleMode = 3
realheight = PicObject.ScaleHeight
realwidth = PicObject.ScaleWidth
'End If
If hBmp Then
cDC = CreateCompatibleDC(PicObject.hdc)
SelectObject cDC, hBmp
Call StretchBlt(PicObject.hdc, 0, 0, realwidth, realheight, cDC, 0, 0, PicWidth, PicHeight, SRCCOPY)
DeleteDC cDC
End If
PicObject.ScaleMode = a
End Sub






Public Sub CreateBitmap_24(ByteArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer)
' Creates a device independent bitmap
' from the pixel data in BitmapArray().

Dim hdc As Long
Dim Bits() As RGBQUAD
ReDim Bits(UBound(ByteArray) / 4)
CopyMemory Bits(0), ByteArray(0), UBound(ByteArray)
    With bm24.bmiHeader
        .biSize = Len(bm24.bmiHeader)        'SizeOf Struct
        .biWidth = PicWidth        'Bitmap Width
        If Orientation = 0 Then
        .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
        .biHeight = -PicHeight
        End If
        .biPlanes = 1                           'Single plane
        .biBitCount = 32                        '32 bit alignment
        .biCompression = BI_RGB                 'No Compression
        .biSizeImage = 0                        'Default
        .biXPelsPerMeter = 0                    'Default
        .biYPelsPerMeter = 0                    'Default
        .biClrUsed = 0                          'Default
        .biClrImportant = 0                     'Default
    End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_24(hdc, bm24.bmiHeader, CBM_INIT, Bits(0), bm24, DIB_RGB_COLORS)
End Sub
Public Sub CreateBitmap_16(ByteArray() As Byte, PicWidth As Long, PicHeight As Long, Orientation As Integer)
' Creates a device independent bitmap
' from the pixel data in BitmapArray().
Dim hdc As Long

    With bm16.bmiHeader
        .biSize = Len(bm16.bmiHeader)        'SizeOf Struct
        .biWidth = PicWidth                       'Bitmap Width
        If Orientation = 0 Then
        .biHeight = PicHeight                    'Bitmap Height, bitmap is top down.
        Else
        .biHeight = -PicHeight
        End If
        .biPlanes = 1                           'Single plane
        .biBitCount = 16                        '32 bit alignment
        .biCompression = BI_RGB                 'No Compression
        .biSizeImage = 0                        'Default
        .biXPelsPerMeter = 0                    'Default
        .biYPelsPerMeter = 0                    'Default
        .biClrUsed = 0                          'Default
        .biClrImportant = 0                     'Default
    End With
' Get the DC.
hdc = GetDC(0)
hBmp = CreateDIBitmap_16(hdc, bm16.bmiHeader, CBM_INIT, ByteArray(0), bm16, DIB_RGB_COLORS)
End Sub

Private Function PixelToTwips(xwert As Long, ywert As Long) As ScTw
Dim ux As Long
Dim uy As Long
Dim XWert1 As Long
Dim yWert1 As Long
ux = Screen.TwipsPerPixelX
PixelToTwips.Width = xwert * ux
uy = Screen.TwipsPerPixelY
PixelToTwips.Height = ywert * uy
End Function



Public Function Twipstopixel(xwert As Long, ywert As Long) As ScTw
Dim ux As Long
Dim uy As Long
Dim XWert1 As Long
Dim yWert1 As Long
ux = Screen.TwipsPerPixelX
Twipstopixel.Width = xwert / ux
uy = Screen.TwipsPerPixelY
Twipstopixel.Height = ywert / uy
End Function
