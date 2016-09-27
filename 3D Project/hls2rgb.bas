Attribute VB_Name = "ColorsAndLightness"
'Public lightns As Single
'Public H As Single, L As Single, S As Single
Const HLSMAX = 360
Const UNDEFINED = HLSMAX * 2 / 3
Const RGBMAX = 255

Public Sub hls2rgb(theHue As Double, Lightns As Double, theSat As Double, r As Single, g As Single, b As Single)
Dim hue As Single, sat As Single
Dim m1 As Single, m2 As Single
'RGB2HLS Form4.HScroll1.Value, Form4.HScroll2.Value, Form4.HScroll3.Value
hue = theHue
'hue = H
sat = theSat
If Lightns <= 0.5 Then
    m1 = Lightns - Lightns * sat
    m2 = Lightns + Lightns * sat
Else
    m1 = Lightns - (1 - Lightns) * sat
    m2 = Lightns + (1 - Lightns) * sat
End If
If sat = 0 Then
    r = Lightns
    g = Lightns
    b = Lightns
Else
    r = myRGB(hue, m1, m2)
    g = myRGB(hue - 120, m1, m2)
    b = myRGB(hue + 120, m1, m2)
End If
End Sub

Private Function myRGB(hue As Single, Value1 As Single, Value2 As Single) As Single
Dim hue1 As Single
If hue > 360 Then
    hue1 = hue - 360
ElseIf hue < 0 Then
    hue1 = hue + 360
Else
    hue1 = hue
End If
If hue1 < 60 Then
    myRGB = (Value2 - Value1) * hue1 / 60 + Value1
ElseIf hue1 < 180 Then
    myRGB = Value2
ElseIf hue1 < 240 Then
    myRGB = (Value2 - Value1) * (240 - hue1) / 60 + Value1
Else
    myRGB = Value1
End If
End Function

Public Function Max(ByVal a As Single, ByVal b As Single) As Single
If a > b Then
    Max = a
Else
    Max = b
End If
End Function

Public Function Min(ByVal a As Single, ByVal b As Single) As Single
If a < b Then
    Min = a
Else
    Min = b
End If
End Function

Public Sub RGB2HLS(r As Single, g As Single, b As Single, H As Double, L As Double, S As Double)

'Dim R, G, B As Single        ' input RGB values
Dim cMax As Single     ' max and min RGB values
Dim cMin As Single     ' max and min RGB values
Dim Rdelta, Gdelta, Bdelta As Single ' intermediate value: % of spread from max


' get R, G, and B out of DWORD
'R = GetRValue(lRGBColor);
'G = GetGValue(lRGBColor);
'B = GetBValue(lRGBColor);
   
' calculate lightness
cMax = Max(Max(r, g), b)
cMin = Min(Min(r, g), b)
L = ((((cMax + cMin) * HLSMAX) + RGBMAX) / (2 * RGBMAX)) / RGBMAX

If (cMax = cMin) Then           ' r=g=b --> achromatic case
    S = 0                     ' saturation
    H = UNDEFINED             ' hue
    ' chromatic case
Else
    If (L <= (HLSMAX / 2)) Then   ' saturation
        S = (((cMax - cMin) * HLSMAX) + ((cMax + cMin) / 2)) / (cMax + cMin)
    Else
        S = (((cMax - cMin) * HLSMAX) + ((2 * RGBMAX - cMax - cMin) / 2)) / (2 * RGBMAX - cMax - cMin)
    End If
      ' hue
    Rdelta = (((cMax - r) * (HLSMAX / 6)) + ((cMax - cMin) / 2)) / (cMax - cMin)
    Gdelta = (((cMax - g) * (HLSMAX / 6)) + ((cMax - cMin) / 2)) / (cMax - cMin)
    Bdelta = (((cMax - b) * (HLSMAX / 6)) + ((cMax - cMin) / 2)) / (cMax - cMin)

    If (r = cMax) Then
        H = Bdelta - Gdelta
    Else
        If (g = cMax) Then
            H = (HLSMAX / 3) + Rdelta - Bdelta
        Else  ' B == cMax
            H = ((2 * HLSMAX) / 3) + Gdelta - Rdelta
        End If
    End If
    If (H < 0) Then
        H = H + HLSMAX
    End If
    If (H > HLSMAX) Then
        H = H - HLSMAX
    End If
End If
End Sub

Public Function GetRed(colorVal As Long) As Integer
GetRed = colorVal Mod 256
End Function

Public Function GetGreen(colorVal As Long) As Integer
GetGreen = ((colorVal And &HFF00FF00) / 256&)
End Function

Public Function GetBlue(colorVal As Long) As Integer
GetBlue = ((colorVal And &HFF0000) / 65536)
End Function

