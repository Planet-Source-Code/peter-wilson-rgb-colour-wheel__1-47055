Attribute VB_Name = "mHSVtoRGB"
Option Explicit

Public Function HSV(Optional ByVal Hue As Single = -1, Optional ByVal Saturation As Single = 1, Optional ByVal Lightness As Single = 1) As Long

    ' ==============================================================================================
    ' Given a Hue, Saturation and Lightness, return the Red-Green_Blue equivalent as a Long data type.
    ' This funtion is intended to replace VB's RGB function.
    '
    ' Ranges:
    '   Hue -1 (no hue)
    '       or
    '   Hue 0 to 360
    '
    '   Saturation 0 to 1
    '   Lightness 0 to 1
    '
    ' ie. Bright-RED = (Hue=0, Saturation=1, Lightness=1)
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
            Red = Lightness
            Green = Lightness
            Blue = Lightness
        Else
            ' *** Make sure you've turned on 'Break on unhandled Errors' ***
            Err.Raise vbObjectError + 1000, "HSV_to_RGB", "A Hue was given with no Saturation. This is invalid."
        End If
    Else
        Hue = (Hue Mod 360) / 60
        i = Int(Hue)    ' Return largest integer
        f = Hue - i     ' f is the fractional part of Hue
        p = Lightness * (1 - Saturation)
        q = Lightness * (1 - (Saturation * f))
        t = Lightness * (1 - (Saturation * (1 - f)))
        Select Case i
            Case 0
                Red = Lightness
                Green = t
                Blue = p
            Case 1
                Red = q
                Green = Lightness
                Blue = p
            Case 2
                Red = p
                Green = Lightness
                Blue = t
            Case 3
                Red = p
                Green = q
                Blue = Lightness
            Case 4
                Red = t
                Green = p
                Blue = Lightness
            Case 5
                Red = Lightness
                Green = p
                Blue = q
        End Select
    End If
    
    HSV = RGB(255 * Red, 255 * Green, 255 * Blue)
        
End Function


Public Function HSV2(Red As Single, Green As Single, Blue As Single, Optional ByVal Hue As Single, Optional ByVal Saturation As Single = 1, Optional ByVal Lightness As Single = 1) As Long

    ' ==============================================================================================
    ' Given a Hue, Saturation and Lightness, return the separate Red, Green and Blue values having
    ' ranges between 0 and 1.
    '
    ' Ranges:
    '   Hue -1 (no hue)
    '       or
    '   Hue 0 to 360
    '
    '   Saturation 0 to 1
    '   Lightness 0 to 1
    '
    ' ie. Bright-RED = (Hue=0, Saturation=1, Lightness=1)
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
            Red = Lightness
            Green = Lightness
            Blue = Lightness
        Else
            Err.Raise vbObjectError + 1000, "HSV_to_RGB", "A Hue was given with no Saturation. This is invalid."
        End If
    Else
        Hue = (Hue Mod 360) / 60
        i = Int(Hue)    ' Return largest integer
        f = Hue - i     ' f is the fractional part of Hue
        p = Lightness * (1 - Saturation)
        q = Lightness * (1 - (Saturation * f))
        t = Lightness * (1 - (Saturation * (1 - f)))
        Select Case i
            Case 0
                Red = Lightness
                Green = t
                Blue = p
            Case 1
                Red = q
                Green = Lightness
                Blue = p
            Case 2
                Red = p
                Green = Lightness
                Blue = t
            Case 3
                Red = p
                Green = q
                Blue = Lightness
            Case 4
                Red = t
                Green = p
                Blue = Lightness
            Case 5
                Red = Lightness
                Green = p
                Blue = q
        End Select
    End If

End Function

Public Function RGBHTML(ByVal Red As Single, ByVal Green As Single, ByVal Blue As Single) As String

    ' Assumes that Red, Green and Blue values range from 0 to 1
    ' Returns HTML colour code.

    Dim lngRed As Long
    Dim lngGreen As Long
    Dim lngBlue As Long
    
    Dim lngHTMLCode As Single
    Dim strHexValue As String
    Dim strHexFormatted As String
        
    lngRed = Int(Red * 255) * CLng(256) * CLng(256) ' Don't allow fractional parts of Red, Green or Blue (ie. Use Int() )
    lngGreen = Int(Green * 255) * CLng(256)
    lngBlue = Int(Blue * 255)
    
    lngHTMLCode = lngRed + lngGreen + lngBlue
    
    strHexValue = Hex(lngHTMLCode)
    
    strHexFormatted = Format(strHexValue, "@@@@@@")
    
    RGBHTML = "#" & Replace(strHexFormatted, " ", "0")
    

End Function

