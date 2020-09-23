Attribute VB_Name = "mHSVColourSpace"
Option Explicit

' Define the name of this class/module for error-trap reporting.
Private Const m_strModuleName As String = "mHSVColourSpace"

Public Function HSV(Optional ByVal Hue As Single = -1, Optional ByVal Saturation As Single = 1, Optional ByVal Brightness As Single = 1) As Long
Attribute HSV.VB_Description = "Returns a Long value representing the RGB equivalent of the given Hue, Saturation and Brightness. Hue Range: -1 to 360. Saturation Range: 0 to 1. Brightness Range: 0 to 1. (Use as a direct replacement for VB's internal RGB function.)"

    ' =================================================================================================
    ' Given a Hue, Saturation and Brightness, return the Red-Green_Blue equivalent as a Long data type.
    ' This funtion is intended to replace VB's RGB function.
    '
    ' Ranges:
    '   Hue -1 (no hue)
    '       or
    '   Hue 0 to 360
    '
    '   Saturation 0 to 1
    '   Brightness 0 to 1
    '
    ' ie. Bright-RED = (Hue=0, Saturation=1, Brightness=1)
    '
    ' Example:
    '   Picture1.ForeColor = HSV(0,1,1)
    '
    ' ==============================================================================================
    
    Dim Red As Single
    Dim Green As Single
    Dim Blue As Single
    
    Dim i As Single
    Dim f As Single
    Dim p As Single
    Dim q As Single
    Dim t As Single
    
    If Saturation = 0 Then  '   The colour is on the black-and-white center line.
        If Hue = -1 Then    '   Achromatic color: There is no hue.
            Red = Brightness
            Green = Brightness
            Blue = Brightness
        Else
            ' *** Make sure you've turned on 'Break on unhandled Errors' ***
            Err.Raise vbObjectError + 1000, "HSV_to_RGB", "A Hue was given with no Saturation. This is invalid."
        End If
    Else
        Hue = (Hue Mod 360) / 60
        i = Int(Hue)    ' Return largest integer
        f = Hue - i     ' f is the fractional part of Hue
        p = Brightness * (1 - Saturation)
        q = Brightness * (1 - (Saturation * f))
        t = Brightness * (1 - (Saturation * (1 - f)))
        Select Case i
            Case 0
                Red = Brightness
                Green = t
                Blue = p
            Case 1
                Red = q
                Green = Brightness
                Blue = p
            Case 2
                Red = p
                Green = Brightness
                Blue = t
            Case 3
                Red = p
                Green = q
                Blue = Brightness
            Case 4
                Red = t
                Green = p
                Blue = Brightness
            Case 5
                Red = Brightness
                Green = p
                Blue = q
        End Select
    End If
    
    HSV = RGB(255 * Red, 255 * Green, 255 * Blue)
        
End Function

Public Function HSV2(Red As Single, Green As Single, Blue As Single, Optional ByVal Hue As Single, Optional ByVal Saturation As Single = 1, Optional ByVal Brightness As Single = 1) As Long
Attribute HSV2.VB_Description = "Returns the Red, Green & Blue components of the given Hue, Saturation and Brightness. Hue Range: -1 to 360. Saturation Range: 0 to 1. Brightness Range: 0 to 1. Red, Green & Blue Ranges: 0 to 1."

    ' ==============================================================================================
    ' Given a Hue, Saturation and Brightness, return the separate Red, Green and Blue values having
    ' ranges between 0 and 1.
    '
    ' Ranges:
    '   Hue -1 (no hue)
    '       or
    '   Hue 0 to 360
    '
    '   Saturation 0 to 1
    '   Brightness 0 to 1
    '
    ' ie. Bright-RED = (Hue=0, Saturation=1, Brightness=1)
    '       returns
    '     Red=1, Green=0, Blue=0
    '
    ' Example:
    '
    '   Dim myRed As Single, myGreen As Single, myBlue As Single
    '   Call HSV2(myRed, myGreen, myBlue, 0, 1, 1)
    '   Picture1.ForeColour = RGB(255*myRed, 255*myGreen, 255*myBlue)
    '
    ' ==============================================================================================
    
    Dim i As Single
    Dim f As Single
    Dim p As Single
    Dim q As Single
    Dim t As Single
    
    If Saturation = 0 Then  '   The colour is on the black-and-white center line.
        If Hue = -1 Then    '   Achromatic color: There is no hue.
            Red = Brightness
            Green = Brightness
            Blue = Brightness
        Else
            Err.Raise vbObjectError + 1000, "HSV_to_RGB", "A Hue was given with no Saturation. This is invalid."
        End If
    Else
        Hue = (Hue Mod 360) / 60
        i = Int(Hue)    ' Return largest integer
        f = Hue - i     ' f is the fractional part of Hue
        p = Brightness * (1 - Saturation)
        q = Brightness * (1 - (Saturation * f))
        t = Brightness * (1 - (Saturation * (1 - f)))
        Select Case i
            Case 0
                Red = Brightness
                Green = t
                Blue = p
            Case 1
                Red = q
                Green = Brightness
                Blue = p
            Case 2
                Red = p
                Green = Brightness
                Blue = t
            Case 3
                Red = p
                Green = q
                Blue = Brightness
            Case 4
                Red = t
                Green = p
                Blue = Brightness
            Case 5
                Red = Brightness
                Green = p
                Blue = q
        End Select
    End If

End Function

Public Function RGBtoHSV(Red As Integer, Green As Integer, Blue As Integer, Hue As Single, Saturation As Single, Brightness As Single)
Attribute RGBtoHSV.VB_Description = "Returns the Hue, Saturation and Brightness, given the Red, Green and Blue components of a colour. RGB ranges: 0 to 255."

    ' ======================================================================
    ' Converts Red, Green & Blue back into a Hue, Saturation and Brightness.
    ' ======================================================================
    
    Dim sngRed As Single
    Dim sngGreen As Single
    Dim sngBlue As Single
    Dim sngMaxBrightness As Single
    Dim sngMinBrightness As Single
    Dim sngDelta As Single
    
    ' Clamp to safe values.
    ' =====================
    If Red > 255 Then Red = 255
    If Green > 255 Then Green = 255
    If Blue > 255 Then Blue = 255
    If Red < 0 Then Red = 0
    If Green < 0 Then Green = 0
    If Blue < 0 Then Blue = 0
    
    
    ' Convert values from the range 0-255 to 0-1
    ' ==========================================
    sngRed = 255 / Red
    sngGreen = 255 / Green
    sngBlue = 255 / Blue
    
    
    ' Find the Min & Max Brightness values.
    ' ====================================
    sngMaxBrightness = 0
    If sngRed > sngMaxBrightness Then sngMaxBrightness = sngRed
    If sngGreen > sngMaxBrightness Then sngMaxBrightness = sngGreen
    If sngBlue > sngMaxBrightness Then sngMaxBrightness = sngBlue
    Brightness = sngMaxBrightness
    
    sngMinBrightness = 1
    If sngRed < sngMinBrightness Then sngMinBrightness = sngRed
    If sngGreen < sngMinBrightness Then sngMinBrightness = sngGreen
    If sngBlue < sngMinBrightness Then sngMinBrightness = sngBlue
    
    
    ' Calculate Saturation
    ' ====================
    If sngMaxBrightness = 0 Then
        Saturation = 0                                      ' << Saturation is 0 if Red, Green and Blue are all zero.
    Else
        Saturation = (sngMaxBrightness - sngMinBrightness) / sngMaxBrightness
    End If
    
    
    ' Calculate Hue.
    ' ==============
    If Saturation = 0 Then
        Hue = -1 ' Undefined.
    Else
        sngDelta = (sngMaxBrightness - sngMinBrightness)
        
        If sngRed = sngMaxBrightness Then
            Hue = (sngGreen - sngBlue) / sngDelta           '   << Resulting colour is between yellow and magenta.
        ElseIf sngGreen = sngMaxBrightness Then
            Hue = 2 + (sngBlue - sngRed) / sngDelta         '   << Resulting colour is between cyan and yellow.
        ElseIf sngBlue = sngMaxBrightness Then
            Hue = 4 + (sngRed - sngGreen) / sngDelta        '   << Resulting colour is between magenta and cyan.
        End If
        
        Hue = Hue * 60                                      '   << Convert Hue to degrees in the range 0 to 360.
        
        If Hue < 0 Then Hue = Hue + 360                     '   << Make sure Hue is non-negative.
        
    End If ' Is Chromatic?
    
    
End Function

