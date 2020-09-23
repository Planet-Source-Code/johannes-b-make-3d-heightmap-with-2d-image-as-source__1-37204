Attribute VB_Name = "Module1"
Declare Function GetPixel Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function SetPixelV Lib "gdi32" (ByVal HDC As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long

'RGB to HSL module
Public Const HSLMAX As Integer = 240 '***
    'H, S and L values can be 0 - HSLMAX
    '240 matches what is used by MS Win;
    'any number less than 1 byte is OK;
    'works best if it is evenly divisible by 6
Const RGBMAX As Integer = 255 '***
    'R, G, and B value can be 0 - RGBMAX
Const UNDEFINED As Integer = (HSLMAX * 2 / 3) '***
    'Hue is undefined if Saturation = 0 (greyscale)

Public Type HSLCol 'Datatype used to pass HSL Color values
    Hue As Integer
    Sat As Integer
    Lum As Integer
End Type

Dim HHH As Double, sSS As Double, LLL As Double
Public Function RGBtoHSL(RGBCol As Long) As HSLCol '***
'Returns an HSLCol datatype containing Hue, Luminescence
'and Saturation; given an RGB Color value

Dim R As Integer, G As Integer, B As Integer
Dim cMax As Integer, cMin As Integer
Dim RDelta As Double, GDelta As Double, _
    BDelta As Double

Dim cMinus As Long, cPlus As Long
    
    R = RGBRed(RGBCol)
    G = RGBGreen(RGBCol)
    B = RGBBlue(RGBCol)
    
    cMax = iMax(iMax(R, G), B) 'Highest and lowest
    cMin = iMin(iMin(R, G), B) 'color values
    
    cMinus = cMax - cMin 'Used to simplify the
    cPlus = cMax + cMin  'calculations somewhat.
    
    'Calculate luminescence (lightness)
    LLL = ((cPlus * HSLMAX) + RGBMAX) / (2 * RGBMAX)
    
    If cMax = cMin Then 'achromatic (r=g=b, greyscale)
        sSS = 0 'Saturation 0 for greyscale
        HHH = UNDEFINED 'Hue undefined for greyscale
    Else
        'Calculate color saturation
        If L <= (HSLMAX / 2) Then
            sSS = ((cMinus * HSLMAX) + 0.5) / cPlus
        Else
            sSS = ((cMinus * HSLMAX) + 0.5) / (2 * RGBMAX - cPlus)
        End If
    
        'Calculate hue
        RDelta = (((cMax - R) * (HSLMAX / 6)) + 0.5) / cMinus
        GDelta = (((cMax - G) * (HSLMAX / 6)) + 0.5) / cMinus
        BDelta = (((cMax - B) * (HSLMAX / 6)) + 0.5) / cMinus
    
        Select Case cMax
            Case CLng(R)
                HHH = BDelta - GDelta
            Case CLng(G)
                HHH = (HSLMAX / 3) + RDelta - BDelta
            Case CLng(B)
                HHH = ((2 * HSLMAX) / 3) + GDelta - RDelta
        End Select
        
        If HHH < 0 Then HHH = HHH + HSLMAX
    End If
    
    Form1.SetH (HHH)
    Form1.SetS (sSS)
    Form1.SetL (LLL)
    

End Function
Private Function iMax(a As Integer, B As Integer) _
    As Integer
'Return the Larger of two values
    iMax = IIf(a > B, a, B)
End Function

Private Function iMin(a As Integer, B As Integer) _
    As Integer
'Return the smaller of two values
    iMin = IIf(a < B, a, B)
End Function
Public Function RGBRed(RGBCol As Long) As Integer
'Return the Red component from an RGB Color
    RGBRed = RGBCol And &HFF
End Function

Public Function RGBGreen(RGBCol As Long) As Integer
'Return the Green component from an RGB Color
    RGBGreen = ((RGBCol And &H100FF00) / &H100)
End Function

Public Function RGBBlue(RGBCol As Long) As Integer
'Return the Blue component from an RGB Color
    RGBBlue = (RGBCol And &HFF0000) / &H10000
End Function
