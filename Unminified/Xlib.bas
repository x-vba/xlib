Attribute VB_Name = "Xlib"
'The MIT License (MIT)
'Copyright © 2020 Anthony Mancini
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Option Explicit
Option Private Module

'@Module: This module contains a set of functions for manipulating and working with arrays.



Public Function CountUnique( _
    ParamArray array1() As Variant) _
As Integer
    
    '@Description: This function counts the number of unique occurances of values within a range or multiple ranges
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: array1 is the group of cells we are counting the unique values of
    '@Returns: Returns the number of unique values
    '@Example: =CountUnique(1, 2, 2, 3) -> 3;
    '@Example: =CountUnique("a", "a", "a") -> 1;
    '@Example: =CountUnique(arr) -> 3; Where arr = [1, 2, 4, 4, 1]
    
    Dim individualElement As Variant
    Dim individualValue As Variant
    Dim uniqueDictionary As Object
    Dim uniqueCount As Integer
    
    Set uniqueDictionary = CreateObject("Scripting.Dictionary")
    
    For Each individualElement In array1
        If IsArray(individualElement) Then
            For Each individualValue In individualElement
                If Not uniqueDictionary.exists(individualValue) Then
                    uniqueDictionary.Add individualValue, 0
                    uniqueCount = uniqueCount + 1
                End If
            Next
        Else
            If Not uniqueDictionary.exists(individualElement) Then
                uniqueDictionary.Add individualElement, 0
                uniqueCount = uniqueCount + 1
            End If
        End If
    Next
    
    CountUnique = uniqueCount
    
End Function


Public Function Sort( _
    ByVal sortableArray As Variant, _
    Optional ByVal descendingFlag As Boolean) _
As Variant

    '@Description: This function is an implementation of Bubble Sort, allowing the user to sort an array, optionally allowing the user to specify the array to be sorted in descending order
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: sortableArray is the array that will be sorted
    '@Param: descendingFlag changes the sort to descending
    '@Returns: Returns the a sorted array
    '@Example: =Sort({1,3,2}) -> {1,2,3}
    '@Example: =Sort({1,3,2}, True) -> {3,2,1}

    Dim i As Integer
    Dim swapOccuredBool As Boolean
    Dim arrayLength As Integer
    arrayLength = UBound(sortableArray) - LBound(sortableArray) + 1
    
    Dim sortedArray() As Variant
    ReDim sortedArray(arrayLength)
    
    For i = 0 To arrayLength - 1
        sortedArray(i) = sortableArray(i)
    Next
    
    Dim temporaryValue As Variant
    
    Do
        swapOccuredBool = False
        For i = 0 To arrayLength - 1
            If (sortedArray(i)) < sortedArray(i + 1) Then
                temporaryValue = sortedArray(i)
                sortedArray(i) = sortedArray(i + 1)
                sortedArray(i + 1) = temporaryValue
                swapOccuredBool = True
            End If
        Next
    Loop While swapOccuredBool
    
    If descendingFlag = True Then
        Sort = sortedArray
    Else
        Dim ascendingArray() As Variant
        ReDim ascendingArray(arrayLength)
        
        For i = 0 To arrayLength - 1
            ascendingArray(i) = sortedArray(arrayLength - i - 1)
        Next
        
        Sort = ascendingArray
    End If
    
End Function


Public Function Reverse( _
    ByVal array1 As Variant) _
As Variant

    '@Description: This function takes an array and reverses all its elements
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: array1 is the array that will be reversed
    '@Returns: Returns the a reversed array
    '@Example: =Reverse({1,2,3}) -> {3,2,1}

    Dim i As Integer
    Dim arrayLength As Integer
    Dim reversedArray() As Variant
    
    arrayLength = UBound(array1) - LBound(array1)
    ReDim reversedArray(arrayLength)
    
    For i = LBound(array1) To UBound(array1)
        reversedArray(arrayLength - i) = array1(i)
    Next
    
    Reverse = reversedArray

End Function


Public Function SumHigh( _
    ByVal array1 As Variant, _
    ByVal numberSummed As Integer) _
As Variant

    '@Description: This function returns the sum of the top values of the number specified in the second argument. For example, if the second argument is 3, only the top 3 values will be summed
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: array1 is the range that will be summed
    '@Param: numberSummed is the number of the top values that will be summed
    '@Returns: Returns the sum of the top numbers specified
    '@Example: =SumHigh({1,2,3,4}, 2) -> 7; as 3 and 4 will be summed
    '@Example: =SumHigh({1,2,3,4}, 3) -> 9; as 2, 3, and 4 will be summed

    Dim i As Integer
    Dim sumValue As Double
    
    For i = 1 To numberSummed
        sumValue = sumValue + Large(array1, i)
    Next
    
    SumHigh = sumValue

End Function


Public Function SumLow( _
    ByVal array1 As Variant, _
    ByVal numberSummed As Integer) _
As Variant

    '@Description: This function returns the sum of the bottom values of the number specified in the second argument. For example, if the second argument is 3, only the bottom 3 values will be summed
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: array1 is the range that will be summed
    '@Param: numberSummed is the number of the bottom values that will be summed
    '@Returns: Returns the sum of the bottom numbers specified
    '@Example: =SumLow({1,2,3,4}, 2) -> 3; as 1 and 2 will be summed
    '@Example: =SumLow({1,2,3,4}, 3) -> 6; as 1, 2, and 3 will be summed

    Dim i As Integer
    Dim sumValue As Double
    
    For i = 1 To numberSummed
        sumValue = sumValue + Small(array1, i)
    Next
    
    SumLow = sumValue

End Function


Public Function AverageHigh( _
    ByVal array1 As Variant, _
    ByVal numberAveraged As Integer) _
As Variant

    '@Description: This function returns the average of the top values of the number specified in the second argument. For example, if the second argument is 3, only the top 3 values will be averaged
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: array1 is the range that will be averaged
    '@Param: numberAveraged is the number of the top values that will be averaged
    '@Returns: Returns the average of the top numbers specified
    '@Example: =AverageHigh({1,2,3,4}, 2) -> 3.5; as 3 and 4 will be averaged
    '@Example: =AverageHigh({1,2,3,4}, 3) -> 3; as 2, 3, and 4 will be averaged

    Dim i As Integer
    Dim sumValue As Double
    
    For i = 1 To numberAveraged
        sumValue = sumValue + Large(array1, i)
    Next
    
    AverageHigh = sumValue / numberAveraged

End Function


Public Function AverageLow( _
    ByVal array1 As Variant, _
    ByVal numberAveraged As Integer) _
As Variant

    '@Description: This function returns the average of the bottom values of the number specified in the second argument. For example, if the second argument is 3, only the bottom 3 values will be averaged
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: array1 is the range that will be averaged
    '@Param: numberAveraged is the number of the bottom values that will be averaged
    '@Returns: Returns the average of the bottom numbers specified
    '@Example: =AverageLow({1,2,3,4}, 2) -> 1.5; as 1 and 2 will be averaged
    '@Example: =AverageLow({1,2,3,4}, 3) -> 2; as 1, 2, and 3 will be averaged

    Dim i As Integer
    Dim sumValue As Double
    
    For i = 1 To numberAveraged
        sumValue = sumValue + Small(array1, i)
    Next
    
    AverageLow = sumValue / numberAveraged

End Function


Public Function Large( _
    ByVal array1 As Variant, _
    ByVal nthNumber As Integer) _
As Variant

    '@Description: This function returns the nth highest number an in array, similar to Excel's LARGE function.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: array1 is the array that the number will be pulled from
    '@Param: nthNumber is the number of the top value that will be chosen. For example, a nthNumber of 1 results in the 1st highest value being chosen, when a number of 2 results in the 2nd, etc.
    '@Returns: Returns the nth highest number
    '@Example: =Large({1,2,3,4}, 1) -> 4
    '@Example: =Large({1,2,3,4}, 2) -> 3

    Large = Sort(array1)(UBound(array1) - (nthNumber - 1))

End Function


Public Function Small( _
    ByVal array1 As Variant, _
    ByVal nthNumber As Integer) _
As Variant

    '@Description: This function returns the nth lowest number an in array, similar to Excel's SMALL function.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: array1 is the array that the number will be pulled from
    '@Param: nthNumber is the number of the bottom value that will be chosen. For example, a nthNumber of 1 results in the 1st smallest value being chosen, when a number of 2 results in the 2nd, etc.
    '@Returns: Returns the nth smallest number
    '@Example: =Small({1,2,3,4}, 1) -> 1
    '@Example: =Small({1,2,3,4}, 2) -> 2

    Small = Sort(array1, True)(UBound(array1) - (nthNumber - 1))

End Function

Public Function IsInArray( _
    ByVal value1 As Variant, _
    ByVal array1 As Variant) _
As Boolean

    '@Description: This function checks if a value is in an array
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: value1 is the value that will be checked if its in the array
    '@Param: array1 is the array
    '@Returns: Returns boolean True if the value is in the array, and false otherwise
    '@Example: =IsInArray("hello", {"one", 2, "hello"}) -> True
    '@Example: =IsInArray("hello", {1, "two", "three"}) -> False

    Dim individualElement As Variant
    
    For Each individualElement In array1
        If individualElement = value1 Then
            IsInArray = True
            Exit Function
        End If
    Next

    IsInArray = False

End Function



'@Module: This module contains a set of functions for working with colors



Public Function Rgb2Hex( _
    ByVal redColorInteger As Integer, _
    ByVal greenColorInteger As Integer, _
    ByVal blueColorInteger As Integer) _
As String

    '@Description: This function converts an RGB color value into a HEX color value
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: redColorInteger is the red value
    '@Param: greenColorInteger is the green value
    '@Param: blueColorInteger is the blue value
    '@Returns: Returns a string with the HEX value of the color
    '@Example: =Rgb2Hex(255, 255, 255) -> "FFFFFF"

    Rgb2Hex = Dec2Hex(redColorInteger, 2) & Dec2Hex(greenColorInteger, 2) & Dec2Hex(blueColorInteger, 2)
    
End Function

Public Function Hex2Rgb( _
    ByVal hexColorString As String, _
    Optional ByVal singleColorNumberOrName As Variant = -1) _
As Variant

    '@Description: This function converts a HEX color value into an RGB color value, or optionally a single value from the RGB value.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: hexColorString is the color in HEX format
    '@Param: singleColorNumberOrName is a number or string that specifies which of the single color values to return. If this is set to 0 or "Red", the red value will be returned. If this is set to 1 or "Green", the green value will be returned. If this is set to 2 or "Blue", the blue value will be returned.
    '@Returns: Returns a string with the RGB value of the color or the number of the individual color chosen
    '@Example: =Hex2Rgb("FFFFFF") -> "(255, 255, 255)"
    '@Example: =Hex2Rgb("FF0109", 0) -> 255; The red color
    '@Example: =Hex2Rgb("FF0109", "Red") -> 255; The red color
    '@Example: =Hex2Rgb("FF0109", 1) -> 1; The green color
    '@Example: =Hex2Rgb("FF0109", "Green") -> 1; The green color
    '@Example: =Hex2Rgb("FF0109", 2) -> 9; The blue color
    '@Example: =Hex2Rgb("FF0109", "Blue") -> 9; The blue color

    hexColorString = Replace(hexColorString, "#", "")

    If singleColorNumberOrName = 0 Or singleColorNumberOrName = "Red" Then
        Hex2Rgb = Hex2Dec(Left(hexColorString, 2))
    ElseIf singleColorNumberOrName = 1 Or singleColorNumberOrName = "Green" Then
        Hex2Rgb = Hex2Dec(Mid(hexColorString, 3, 2))
    ElseIf singleColorNumberOrName = 2 Or singleColorNumberOrName = "Blue" Then
        Hex2Rgb = Hex2Dec(Right(hexColorString, 2))
    Else
        Hex2Rgb = "(" & Hex2Dec(Left(hexColorString, 2)) & ", " & Hex2Dec(Mid(hexColorString, 3, 2)) & ", " & Hex2Dec(Right(hexColorString, 2)) & ")"
    End If

End Function


Public Function Rgb2Hsl( _
    ByVal redColorInteger As Integer, _
    ByVal greenColorInteger As Integer, _
    ByVal blueColorInteger As Integer, _
    Optional ByVal singleColorNumberOrName As Variant = -1) _
As Variant

    '@Description: This function converts an RGB color value into an HSL color value and returns a string of the HSL value, or optionally a single value from the HSL value.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: redColorInteger is the red value
    '@Param: greenColorInteger is the green value
    '@Param: blueColorInteger is the blue value
    '@Param: singleColorNumberOrName is a number or string that specifies which of the single color values to return. If this is set to 0 or "Hue", the hue value will be returned. If this is set to 1 or "Saturation", the saturation value will be returned. If this is set to 2 or "Lightness", the lightness value will be returned.
    '@Returns: Returns a string with the HSL value of the color
    '@Example: =Rgb2Hsl(8, 64, 128) -> "(212.0ï¿½, 88.2%, 26.7%)"
    '@Example: =Rgb2Hsl(8, 64, 128, 0) -> 212
    '@Example: =Rgb2Hsl(8, 64, 128, "Hue") -> 212
    '@Example: =Rgb2Hsl(8, 64, 128, 1) -> .882
    '@Example: =Rgb2Hsl(8, 64, 128, "Saturation") -> .882
    '@Example: =Rgb2Hsl(8, 64, 128, 2) -> .267
    '@Example: =Rgb2Hsl(8, 64, 128, "Lightness") -> .267

    ' Calculating values needed to calculate HSL
    Dim redPrime As Double
    Dim greenPrime As Double
    Dim bluePrime As Double
    
    redPrime = redColorInteger / 255
    greenPrime = greenColorInteger / 255
    bluePrime = blueColorInteger / 255
    
    Dim colorMax As Double
    Dim colorMin As Double
    
    colorMax = Max(redPrime, greenPrime, bluePrime)
    colorMin = Min(redPrime, greenPrime, bluePrime)
    
    Dim deltaValue As Double
    
    deltaValue = colorMax - colorMin
    
    Dim hueValue As Double
    Dim saturationValue As Double
    Dim lightnessValue As Double
    
    
    ' Calculating Hue
    If deltaValue = 0 Then
        hueValue = 0
    Else
        Select Case colorMax
            Case redPrime
                hueValue = 60 * (((greenPrime - bluePrime) / deltaValue) Mod 6)
            Case greenPrime
                hueValue = 60 * (((bluePrime - redPrime) / deltaValue) + 2)
            Case bluePrime
                hueValue = 60 * (((redPrime - greenPrime) / deltaValue) + 4)
        End Select
    End If
    
    
    ' Calculating Lightness
    lightnessValue = (colorMax + colorMin) / 2
    
    
    ' Calculating Saturation
    If deltaValue = 0 Then
        saturationValue = 0
    Else
        saturationValue = deltaValue / (1 - Abs((2 * lightnessValue - 1)))
    End If


    If singleColorNumberOrName = 0 Or singleColorNumberOrName = "Hue" Then
        Rgb2Hsl = hueValue
    ElseIf singleColorNumberOrName = 1 Or singleColorNumberOrName = "Saturation" Then
        Rgb2Hsl = saturationValue
    ElseIf singleColorNumberOrName = 2 Or singleColorNumberOrName = "Lightness" Then
        Rgb2Hsl = lightnessValue
    Else
        Rgb2Hsl = "(" & Format(hueValue, "#.0") & ", " & Format(saturationValue * 100, "#.0") & "%, " & Format(lightnessValue * 100, "#.0") & "%)"
    End If

End Function


Public Function Hex2Hsl( _
    ByVal hexColorString As String) _
As String

    '@Description: This function converts a HEX color value into an HSL color value
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: hexColorString is the hex color
    '@Returns: Returns a string with the HSL value of the color
    '@Example: =Hex2Hsl("084080") -> "(212.0, 88.2%, 26.7%)"
    '@Example: =Hex2Hsl("#084080") -> "(212.0, 88.2%, 26.7%)"

    hexColorString = Replace(hexColorString, "#", "")

    Dim redValue As Integer
    Dim greenValue As Integer
    Dim blueValue As Integer
    
    redValue = CInt(Hex2Dec(Left(hexColorString, 2)))
    greenValue = CInt(Hex2Dec(Mid(hexColorString, 3, 2)))
    blueValue = CInt(Hex2Dec(Right(hexColorString, 2)))

    Hex2Hsl = Rgb2Hsl(redValue, greenValue, blueValue)

End Function


Public Function Hsl2Rgb( _
    ByVal hueValue As Double, _
    ByVal saturationValue As Double, _
    ByVal lightnessValue As Double, _
    Optional ByVal singleColorNumberOrName As Variant = -1) _
As Variant

    '@Description: This function converts an HSL color value into an RGB color value, or optionally a single value from the RGB value.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: hueValue is the hue value
    '@Param: saturationValue is the saturation value
    '@Param: lightnessValue is the lightness value
    '@Param: singleColorNumberOrName is a number or string that specifies which of the single color values to return. If this is set to 0 or "Red", the red value will be returned. If this is set to 1 or "Green", the green value will be returned. If this is set to 2 or "Blue", the blue value will be returned.
    '@Returns: Returns a string with the RGB value of the color or an individual color value
    '@Example: =Hsl2Rgb(212, .882, .267) -> "(8, 64, 128)"
    '@Example: =Hsl2Rgb(212, .882, .267, 0) -> 8
    '@Example: =Hsl2Rgb(212, .882, .267, "Red") -> 8
    '@Example: =Hsl2Rgb(212, .882, .267, 1) -> 64
    '@Example: =Hsl2Rgb(212, .882, .267, "Green") -> 64
    '@Example: =Hsl2Rgb(212, .882, .267, 2) -> 128
    '@Example: =Hsl2Rgb(212, .882, .267, "Blue") -> 128

    Dim cValue As Double
    Dim xValue As Double
    Dim mValue As Double
    
    cValue = (1 - Abs(2 * lightnessValue - 1)) * saturationValue
    xValue = cValue * (1 - Abs(ModFloat((hueValue / 60), 2) - 1))
    mValue = lightnessValue - cValue / 2
    
    Dim redValue As Double
    Dim greenValue As Double
    Dim blueValue As Double
    
    If hueValue >= 0 And hueValue < 60 Then
        redValue = cValue
        greenValue = xValue
        blueValue = 0
    ElseIf hueValue >= 60 And hueValue < 120 Then
        redValue = xValue
        greenValue = cValue
        blueValue = 0
    ElseIf hueValue >= 120 And hueValue < 180 Then
        redValue = 0
        greenValue = cValue
        blueValue = xValue
    ElseIf hueValue >= 180 And hueValue < 240 Then
        redValue = 0
        greenValue = xValue
        blueValue = cValue
    ElseIf hueValue >= 240 And hueValue < 300 Then
        redValue = xValue
        greenValue = 0
        blueValue = cValue
    ElseIf hueValue >= 300 And hueValue < 360 Then
        redValue = cValue
        greenValue = 0
        blueValue = xValue
    End If
    
    redValue = (redValue + mValue) * 255
    greenValue = (greenValue + mValue) * 255
    blueValue = (blueValue + mValue) * 255
    
    If singleColorNumberOrName = 0 Or singleColorNumberOrName = "Red" Then
        Hsl2Rgb = Round(redValue, 0)
    ElseIf singleColorNumberOrName = 1 Or singleColorNumberOrName = "Green" Then
        Hsl2Rgb = Round(greenValue, 0)
    ElseIf singleColorNumberOrName = 2 Or singleColorNumberOrName = "Blue" Then
        Hsl2Rgb = Round(blueValue, 0)
    Else
        Hsl2Rgb = "(" & Round(redValue, 0) & ", " & Round(greenValue, 0) & ", " & Round(blueValue, 0) & ")"
    End If

End Function


Public Function Hsl2Hex( _
    ByVal hueValue As Double, _
    ByVal saturationValue As Double, _
    ByVal lightnessValue As Double) _
As Variant

    '@Description: This function converts an HSL color value into a HEX color value.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Code the formula directly instead of using an additional conversion to speed up the function
    '@Param: hueValue is the hue value
    '@Param: saturationValue is the saturation value
    '@Param: lightnessValue is the lightness value
    '@Returns: Returns a string with the HEX value of the color
    '@Example: =Hsl2Rgb(212, .882, .267) -> "(8, 64, 128)"

    Dim redValue As Integer
    Dim greenValue As Integer
    Dim blueValue As Integer

    redValue = Hsl2Rgb(hueValue, saturationValue, lightnessValue, 0)
    greenValue = Hsl2Rgb(hueValue, saturationValue, lightnessValue, 1)
    blueValue = Hsl2Rgb(hueValue, saturationValue, lightnessValue, 2)

    Hsl2Hex = Rgb2Hex(redValue, greenValue, blueValue)

End Function


Public Function Rgb2Hsv( _
    ByVal redColorInteger As Integer, _
    ByVal greenColorInteger As Integer, _
    ByVal blueColorInteger As Integer, _
    Optional ByVal singleColorNumberOrName As Variant = -1) _
As Variant

    '@Description: This function converts an RGB color value into an HSV color value, or optionally a single value from the HSV value.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: redColorInteger is the red value
    '@Param: greenColorInteger is the green value
    '@Param: blueColorInteger is the blue value
    '@Param: singleColorNumberOrName is a number or string that specifies which of the single color values to return. If this is set to 0 or "Hue", the hue value will be returned. If this is set to 1 or "Saturation", the saturation value will be returned. If this is set to 2 or "Value", the value value will be returned.
    '@Returns: Returns a string with the RGB value of the color or an individual color value
    '@Example: =Rgb2Hsv(8, 64, 128) -> "(212.0, 93.8%, 50.2%)"
    '@Example: =Rgb2Hsv(8, 64, 128, 0) -> 212
    '@Example: =Rgb2Hsv(8, 64, 128, "Red") -> 212
    '@Example: =Rgb2Hsv(8, 64, 128, 1) -> .938
    '@Example: =Rgb2Hsv(8, 64, 128, "Green") -> .938
    '@Example: =Rgb2Hsv(8, 64, 128, 2) -> .502
    '@Example: =Rgb2Hsv(8, 64, 128, "Blue") -> .502

    ' Calculating values needed to calculate HSV
    Dim redPrime As Double
    Dim greenPrime As Double
    Dim bluePrime As Double
    
    redPrime = redColorInteger / 255
    greenPrime = greenColorInteger / 255
    bluePrime = blueColorInteger / 255
    
    Dim colorMax As Double
    Dim colorMin As Double
    
    colorMax = Max(redPrime, greenPrime, bluePrime)
    colorMin = Min(redPrime, greenPrime, bluePrime)
    
    Dim deltaValue As Double
    
    deltaValue = colorMax - colorMin
    
    Dim hueValue As Double
    Dim saturationValue As Double
    Dim valueValue As Double

    ' Calculating Hue
    If deltaValue = 0 Then
        hueValue = 0
    Else
        Select Case colorMax
            Case redPrime
                hueValue = 60 * (((greenPrime - bluePrime) / deltaValue) Mod 6)
            Case greenPrime
                hueValue = 60 * (((bluePrime - redPrime) / deltaValue) + 2)
            Case bluePrime
                hueValue = 60 * (((redPrime - greenPrime) / deltaValue) + 4)
        End Select
    End If
    
    
    ' Calculating Saturation
    If colorMax = 0 Then
        saturationValue = 0
    Else
        saturationValue = deltaValue / colorMax
    End If
    
    
    ' Calculating Value
    valueValue = colorMax
    

    If singleColorNumberOrName = 0 Or singleColorNumberOrName = "Hue" Then
        Rgb2Hsv = hueValue
    ElseIf singleColorNumberOrName = 1 Or singleColorNumberOrName = "Saturation" Then
        Rgb2Hsv = saturationValue
    ElseIf singleColorNumberOrName = 2 Or singleColorNumberOrName = "Value" Then
        Rgb2Hsv = valueValue
    Else
        Rgb2Hsv = "(" & Format(hueValue, "#.0") & ", " & Format(saturationValue * 100, "#.0") & "%, " & Format(valueValue * 100, "#.0") & "%)"
    End If
    
End Function



'@Module: This module contains a set of functions for working with dates and times.



Public Function WeekdayName2( _
    Optional ByVal dayNumber As Byte) _
As String
    
    '@Description: This function takes a weekday number and returns the name of the day of the week.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: dayNumber is a number that should be between 1 and 7, with 1 being Sunday and 7 being Saturday. If no dayNumber is given, the value will default to the current day of the week.
    '@Returns: Returns the day of the week as a string
    '@Example: =WeekdayName2(4) -> Wednesday
    '@Example: To get today's weekday name: =WeekdayName2()

    If dayNumber = 0 Then
        WeekdayName2 = WeekdayName(Weekday(Now()))
    Else
        WeekdayName2 = WeekdayName(dayNumber)
    End If

End Function


Public Function MonthName2( _
    Optional ByVal monthNumber As Byte) _
As String

    '@Description: This function takes a month number and returns the name of the month.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: monthNumber is a number that should be between 1 and 12, with 1 being January and 12 being December. If no monthNumber is given, the value will default to the current month.
    '@Returns: Returns the month name as a string
    '@Example: =MonthName2(1) -> "January"
    '@Example: =MonthName2(3) -> "March"
    '@Example: To get today's month name: =MonthName2()

    If monthNumber = 0 Then
        MonthName2 = MonthName(Month(Now()))
    Else
        MonthName2 = MonthName(monthNumber)
    End If

End Function


Public Function Quarter( _
    Optional ByVal monthNumberOrName As Variant) _
As Byte
    
    '@Description: This function takes a month as a number and returns the Quarter of the year the month resides.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Look further into DatePart function and see if its a better choice for generating the Quarter of the year. Also look into adding the month name as well as an option for this function
    '@Param: monthNumberOrName is a number that should be between 1 and 12, with 1 being January and 12 being December, or the name of a Month, such as "January" or "March".
    '@Returns: Returns the Quarter of the month as a number
    '@Example: =Quarter(4) -> 2
    '@Example: =Quarter("April") -> 2
    '@Example: =Quarter(12) -> 4
    '@Example: =Quarter("December") -> 4
    '@Example: To get today's Quarter: =Quarter()
    
    If IsMissing(monthNumberOrName) Then
       monthNumberOrName = MonthName(Month(Now()))
    End If
    
    If IsNumeric(monthNumberOrName) Then
        monthNumberOrName = MonthName(monthNumberOrName)
    End If
    
    
    If monthNumberOrName = MonthName(1) Or monthNumberOrName = MonthName(2) Or monthNumberOrName = MonthName(3) Then
        Quarter = 1
    End If
    
    If monthNumberOrName = MonthName(4) Or monthNumberOrName = MonthName(5) Or monthNumberOrName = MonthName(6) Then
        Quarter = 2
    End If
    
    If monthNumberOrName = MonthName(7) Or monthNumberOrName = MonthName(8) Or monthNumberOrName = MonthName(9) Then
        Quarter = 3
    End If
    
    If monthNumberOrName = MonthName(10) Or monthNumberOrName = MonthName(11) Or monthNumberOrName = MonthName(12) Then
        Quarter = 4
    End If

End Function


Public Function TimeConverter( _
    ByVal date1 As Date, _
    Optional ByVal secondsInteger As Integer, _
    Optional ByVal minutesInteger As Integer, _
    Optional ByVal hoursInteger As Integer, _
    Optional ByVal daysInteger As Integer, _
    Optional ByVal monthsInteger As Integer, _
    Optional ByVal yearsInteger As Integer) _
As Date
    
    '@Description: This function takes a date, and then a series of optional arguments for a number of seconds, minutes, hours, days, and years, and then converts the date given to a new date adding in the other date argument values.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: date1 is the original date that will be converted into a new date
    '@Param: secondsInteger is the number of seconds that will be added
    '@Param: minutesInteger is the number of minutes that will be added
    '@Param: hoursInteger is the number of hours that will be added
    '@Param: daysInteger is the number of days that will be added
    '@Param: monthsInteger is the number of months that will be added
    '@Param: yearsInteger is the number of years that will be added
    '@Returns: Returns a new date with all the date arguments added to it
    '@Note: You can skip earlier date arguments in the function by putting a 0 in place. For example, if we only wanted to change the month, which is the 5th argument, we can do =TimeConverter(A1,0,0,0,2) which will add 2 months to the date chosen
    '@Example: =TimeConverter(A1,60) -> 1/1/2000 1:01; Where A1 contains the date 1/1/2000 1:00
    '@Example: =TimeConverter(A1,0,5) -> 1/1/2000 1:05; Where A1 contains the date 1/1/2000 1:00
    '@Example: =TimeConverter(A1,0,0,2) -> 1/1/2000 3:00; Where A1 contains the date 1/1/2000 1:00
    '@Example: =TimeConverter(A1,0,0,0,4) -> 1/5/2000 1:00; Where A1 contains the date 1/1/2000 1:00
    '@Example: =TimeConverter(A1,0,0,0,0,1) -> 2/1/2000 1:00; Where A1 contains the date 1/1/2000 1:00
    '@Example: =TimeConverter(A1,0,0,0,0,0,5) -> 1/1/2005 1:00; Where A1 contains the date 1/1/2000 1:00
    '@Example: =TimeConverter(A1,60,5,3,10,5,15) -> 6/11/2015 4:06; Where A1 contains the date 1/1/2000 1:00
    
    secondsInteger = Second(date1) + secondsInteger
    minutesInteger = Minute(date1) + minutesInteger
    hoursInteger = Hour(date1) + hoursInteger
    daysInteger = Day(date1) + daysInteger
    monthsInteger = Month(date1) + monthsInteger
    yearsInteger = Year(date1) + yearsInteger
    
    TimeConverter = DateSerial(yearsInteger, monthsInteger, daysInteger) + TimeSerial(hoursInteger, minutesInteger, secondsInteger)

End Function


Public Function DaysOfMonth( _
    Optional ByVal monthNumberOrName As Variant, _
    Optional ByVal yearNumber As Integer) _
As Variant

    '@Description: This function takes a month number or month name and returns the number of days in the month. Optionally, a year number can be specified. If no year number is provided, the current year will be used. Finally, note that the month name or number argument is optional and if omitted will use the current month.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: monthNumberOrName is a number that should be between 1 and 12, with 1 being January and 12 being December, or the name of a Month, such as "January" or "March". If omitted the current month will be used.
    '@Param: yearNumber is the year that will be used. If omitted, the current year will be used.
    '@Returns: Returns the number of days in the month and year specified
    '@Example: =DaysOfMonth() -> 31; Where the current month is January
    '@Example: =DaysOfMonth(1) -> 31
    '@Example: =DaysOfMonth("January") -> 31
    '@Example: =DaysOfMonth(2, 2019) -> 28
    '@Example: =DaysOfMonth(2, 2020) -> 29

    If IsMissing(monthNumberOrName) Then
        monthNumberOrName = Month(Now())
    End If

    If yearNumber = 0 Then
        yearNumber = Year(Now())
    End If

    If monthNumberOrName = 1 Or monthNumberOrName = MonthName(1) Then
        DaysOfMonth = 31
    ElseIf monthNumberOrName = 2 Or monthNumberOrName = MonthName(2) Then
        If yearNumber Mod 4 <> 0 Then
            DaysOfMonth = 28
        Else
            DaysOfMonth = 29
        End If
    ElseIf monthNumberOrName = 3 Or monthNumberOrName = MonthName(3) Then
        DaysOfMonth = 31
    ElseIf monthNumberOrName = 4 Or monthNumberOrName = MonthName(4) Then
        DaysOfMonth = 30
    ElseIf monthNumberOrName = 5 Or monthNumberOrName = MonthName(5) Then
        DaysOfMonth = 31
    ElseIf monthNumberOrName = 6 Or monthNumberOrName = MonthName(6) Then
        DaysOfMonth = 30
    ElseIf monthNumberOrName = 7 Or monthNumberOrName = MonthName(7) Then
        DaysOfMonth = 31
    ElseIf monthNumberOrName = 8 Or monthNumberOrName = MonthName(8) Then
        DaysOfMonth = 31
    ElseIf monthNumberOrName = 9 Or monthNumberOrName = MonthName(9) Then
        DaysOfMonth = 30
    ElseIf monthNumberOrName = 10 Or monthNumberOrName = MonthName(10) Then
        DaysOfMonth = 31
    ElseIf monthNumberOrName = 11 Or monthNumberOrName = MonthName(11) Then
        DaysOfMonth = 30
    ElseIf monthNumberOrName = 12 Or monthNumberOrName = MonthName(12) Then
        DaysOfMonth = 31
    Else
        DaysOfMonth = "#NotAValidMonthNumberOrName"
    End If

End Function


Public Function WeekOfMonth( _
    Optional ByVal date1 As Date) _
As Byte

    '@Description: This function takes a date and returns the number of the week of the month for that date. If no date is given, the current date is used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: date1 is a date whose week number will be found
    '@Returns: Returns the number of week in the month
    '@Example: =WeekOfMonth() -> 5; Where the current date is 1/29/2020
    '@Example: =WeekOfMonth(1/29/2020) -> 5
    '@Example: =WeekOfMonth(1/28/2020) -> 5
    '@Example: =WeekOfMonth(1/27/2020) -> 5
    '@Example: =WeekOfMonth(1/26/2020) -> 5
    '@Example: =WeekOfMonth(1/25/2020) -> 4
    '@Example: =WeekOfMonth(1/24/2020) -> 4
    '@Example: =WeekOfMonth(1/1/2020) -> 1
    

    Dim weekNumber As Byte
    Dim currentDay As Byte
    Dim currentWeekday As Byte
    
    weekNumber = 1
    
    ' When year is 1899, no year was given as an input
    If Year(date1) = 1899 Then
        currentDay = Day(Now())
        currentWeekday = Weekday(Now())
    Else
        currentDay = Day(date1)
        currentWeekday = Weekday(date1)
    End If
    
    While currentDay <> 0
        If currentWeekday = 0 Then
            weekNumber = weekNumber + 1
            currentWeekday = 7
        End If
        
        currentDay = currentDay - 1
        currentWeekday = currentWeekday - 1
    Wend
    
    WeekOfMonth = weekNumber

End Function

'@Module: This module contains a set of functions for gathering information on the environment that Excel is being run on, such as the UserName of the computer, the OS Excel is being run on, and other Environment Variable values.



Public Function OS() As String

    '@Description: This function returns the Operating System name. Currently it will return either "Windows" or "Mac" depending on the OS used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns the name of the Operating System
    '@Example: =OS() -> "Windows"; When running this function on Windows
    '@Example: =OS() -> "Mac"; When running this function on MacOS

    #If Mac Then
        OS = "Mac"
    #Else
        OS = "Windows"
    #End If

End Function


Public Function UserName() As String

    '@Description: This function takes no arguments and returns a string of the USERNAME of the computer
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Returns: Returns a string of the username
    '@Example: =UserName() -> "Anthony"
    
    #If Mac Then
        UserName = Environ("USER")
    #Else
        UserName = Environ("USERNAME")
    #End If

End Function


Public Function UserDomain() As String

    '@Description: This function takes no arguments and returns a string of the USERDOMAIN of the computer
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Returns: Returns a string of the user domain of the computer
    '@Example: =UserDomain() -> "DESKTOP-XYZ1234"
    
    #If Mac Then
        UserDomain = Environ("HOST")
    #Else
        UserDomain = Environ("USERDOMAIN")
    #End If

End Function


Public Function ComputerName() As String

    '@Description: This function takes no arguments and returns a string of the COMPUTERNAME of the computer
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Returns: Returns a string of the computer name of the computer
    '@Example: =ComputerName() -> "DESKTOP-XYZ1234"

    ComputerName = Environ("COMPUTERNAME")

End Function

'@Module: This module contains a set of functions for gathering info on files. It includes functions for gathering file info on the current workbook presentation, document, or database, as well as functions for reading and writing to files, and functions for manipulating file path strings.



Public Function GetActivePathAndName() As String

    '@Description: This function returns the path of the file of the office program that is calling this function. It currently supports Excel, Word, PowerPoint, Access, and Publisher.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current path
    '@Example: =GetActivePathAndName() -> "C:\Users\UserName\Documents\XLib.xlsm"

    If Application.Name = "Microsoft Excel" Then
        GetActivePathAndName = GetActivePathAndNameExcel()
        
    ElseIf Application.Name = "Microsoft Word" Then
        GetActivePathAndName = GetActivePathAndNameWord()
        
    ElseIf Application.Name = "Microsoft PowerPoint" Then
        GetActivePathAndName = GetActivePathAndNamePowerPoint()
        
    ElseIf Application.Name = "Microsoft Access" Then
        GetActivePathAndName = GetActivePathAndNameAccess()
        
    ElseIf Application.Name = "Microsoft Publisher" Then
        GetActivePathAndName = GetActivePathAndNamePublisher()
        
    End If

End Function


Private Function GetActivePathAndNameExcel() As String

    '@Description: This function returns the path of the workbook calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current workbook path
    
    #If Mac Then
        GetActivePathAndNameExcel = ThisWorkbook.Path & "/" & ThisWorkbook.Name
    #Else
        GetActivePathAndNameExcel = ThisWorkbook.Path & "\" & ThisWorkbook.Name
    #End If

End Function


Private Function GetActivePathAndNameWord() As String

    '@Description: This function returns the path of the document calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current document path
    
    #If Mac Then
        GetActivePathAndNameWord = ThisDocument.Path & "/" & ThisDocument.Name
    #Else
        GetActivePathAndNameWord = ThisDocument.Path & "\" & ThisDocument.Name
    #End If

End Function


Private Function GetActivePathAndNamePowerPoint() As String

    '@Description: This function returns the path of the presentation calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current presentation path
    
    #If Mac Then
        GetActivePathAndNamePowerPoint = ActivePresentation.Path & "/" & ActivePresentation.Name
    #Else
        GetActivePathAndNamePowerPoint = ActivePresentation.Path & "\" & ActivePresentation.Name
    #End If

End Function


Private Function GetActivePathAndNameAccess() As String

    '@Description: This function returns the path of the database calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current database path
    
    #If Mac Then
        GetActivePathAndNameAccess = CurrentProject.Path & "/" & CurrentProject.Name
    #Else
        GetActivePathAndNameAccess = CurrentProject.Path & "\" & CurrentProject.Name
    #End If

End Function


Private Function GetActivePathAndNamePublisher() As String

    '@Description: This function returns the path of the publisher file calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current publisher file path
    
    #If Mac Then
        GetActivePathAndNamePublisher = ThisDocument.Path & "/" & ThisDocument.Name
    #Else
        GetActivePathAndNamePublisher = ThisDocument.Path & "\" & ThisDocument.Name
    #End If

End Function


Public Function GetActivePath() As String

    '@Description: This function returns the path of the folder of the office program that is calling this function. It currently supports Excel, Word, PowerPoint, Access, and Publisher.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current folder path
    '@Example: =GetActivePath() -> "C:\Users\UserName\Documents\"; Where the file resides in the Documents folder

    If Application.Name = "Microsoft Excel" Then
        GetActivePath = GetActivePathExcel()
        
    ElseIf Application.Name = "Microsoft Word" Then
        GetActivePath = GetActivePathWord()
        
    ElseIf Application.Name = "Microsoft PowerPoint" Then
        GetActivePath = GetActivePathPowerPoint()
        
    ElseIf Application.Name = "Microsoft Access" Then
        GetActivePath = GetActivePathAccess()
        
    ElseIf Application.Name = "Microsoft Publisher" Then
        GetActivePath = GetActivePathPublisher()
        
    End If

End Function


Private Function GetActivePathExcel() As String

    '@Description: This function returns the folder path of the workbook calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current workbook folder path
    
    #If Mac Then
        GetActivePathExcel = ThisWorkbook.Path & "/"
    #Else
        GetActivePathExcel = ThisWorkbook.Path & "\"
    #End If

End Function


Private Function GetActivePathWord() As String

    '@Description: This function returns the folder path of the document calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current document folder path
    
    #If Mac Then
        GetActivePathWord = ThisDocument.Path & "/"
    #Else
        GetActivePathWord = ThisDocument.Path & "\"
    #End If

End Function


Private Function GetActivePathPowerPoint() As String

    '@Description: This function returns the folder path of the presentation calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current presentation folder path
    
    #If Mac Then
        GetActivePathPowerPoint = ActivePresentation.Path & "/"
    #Else
        GetActivePathPowerPoint = ActivePresentation.Path & "\"
    #End If

End Function


Private Function GetActivePathAccess() As String

    '@Description: This function returns the folder path of the database calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current database folder path
    
    #If Mac Then
        GetActivePathAccess = CurrentProject.Path & "/"
    #Else
        GetActivePathAccess = CurrentProject.Path & "\"
    #End If

End Function


Private Function GetActivePathPublisher() As String

    '@Description: This function returns the folder path of the publisher file calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current publisher folder path
    
    #If Mac Then
        GetActivePathPublisher = ThisDocument.Path & "/"
    #Else
        GetActivePathPublisher = ThisDocument.Path & "\"
    #End If

End Function


Public Function FileCreationTime( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the file creation time of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the file creation time of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileCreationTime() -> "1/1/2020 1:23:45 PM"
    '@Example: =FileCreationTime("C:\hello\world.txt") -> "1/1/2020 5:55:55 PM"
    '@Example: =FileCreationTime("vba.txt") -> "12/25/2000 1:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        FileCreationTime = FSO.GetFile(GetActivePathAndName()).DateCreated
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileCreationTime = FSO.GetFile(GetActivePath() & filePath).DateCreated
        Else
            FileCreationTime = FSO.GetFile(filePath).DateCreated
        End If
    End If

End Function


Public Function FileLastModifiedTime( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the file last modified time of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the file last modified time of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileLastModifiedTime() -> "1/1/2020 2:23:45 PM"
    '@Example: =FileLastModifiedTime("C:\hello\world.txt") -> "1/1/2020 7:55:55 PM"
    '@Example: =FileLastModifiedTime("vba.txt") -> "12/25/2000 3:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        FileLastModifiedTime = FSO.GetFile(GetActivePathAndName()).DateLastModified
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileLastModifiedTime = FSO.GetFile(GetActivePath() & filePath).DateLastModified
        Else
            FileLastModifiedTime = FSO.GetFile(filePath).DateLastModified
        End If
    End If

End Function


Public Function FileDrive( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the drive of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the file drive of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileDrive() -> "A:"; Where the current workbook resides on the A: drive
    '@Example: =FileDrive("C:\hello\world.txt") -> "C:"
    '@Example: =FileDrive("vba.txt") -> "B:"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in, and where the workbook resides in the B: drive

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        FileDrive = FSO.GetFile(GetActivePathAndName()).Drive
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileDrive = FSO.GetFile(GetActivePath() & filePath).Drive
        Else
            FileDrive = FSO.GetFile(filePath).Drive
        End If
    End If

End Function


Public Function FileName( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the name of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the name of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileName() -> "MyWorkbook.xlsm"
    '@Example: =FileName("C:\hello\world.txt") -> "world.txt"
    '@Example: =FileName("vba.txt") -> "vba.txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        FileName = FSO.GetFile(GetActivePathAndName()).Name
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileName = FSO.GetFile(GetActivePath() & filePath).Name
        Else
            FileName = FSO.GetFile(filePath).Name
        End If
    End If

End Function


Public Function FileFolder( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the path of the folder of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the path of the folder where the file resides in as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileFolder() -> "C:\my_excel_files"
    '@Example: =FileFolder("C:\hello\world.txt") -> "C:\hello"
    '@Example: =FileFolder("vba.txt") -> "C:\my_excel_files"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        FileFolder = FSO.GetFile(GetActivePathAndName()).ParentFolder
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileFolder = FSO.GetFile(GetActivePath() & filePath).ParentFolder
        Else
            FileFolder = FSO.GetFile(filePath).ParentFolder
        End If
    End If

End Function


Public Function CurrentFilePath( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the path of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the path of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =CurrentFilePath() -> "C:\my_excel_files\MyWorkbook.xlsx"
    '@Example: =CurrentFilePath("C:\hello\world.txt") -> "C:\hello\world.txt"
    '@Example: =CurrentFilePath("vba.txt") -> "C:\hello\world.txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        CurrentFilePath = FSO.GetFile(GetActivePathAndName()).Path
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            CurrentFilePath = FSO.GetFile(GetActivePath() & filePath).Path
        Else
            CurrentFilePath = FSO.GetFile(filePath).Path
        End If
    End If

End Function


Public Function FileSize( _
    Optional ByVal filePath As String, _
    Optional ByVal byteSize As String) _
As Double

    '@Description: This function returns the file size of the file specified in the file path argument, with the option to set if the file size is returned in Bytes, Kilobytes, Megabytes, or Gigabytes. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Param: byteSize is a string of value "KB", "MB", or "GB"
    '@Returns: Returns the size of the file as a Double
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileSize() -> 1024
    '@Example: =FileSize(,"KB") -> 1
    '@Example: =FileSize("vba.txt", "KB") -> 0.25; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim totalBytes As Double
    
    If filePath = "" Then
        totalBytes = FSO.GetFile(GetActivePathAndName()).Size
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            totalBytes = FSO.GetFile(GetActivePath() & filePath).Size
        Else
            totalBytes = FSO.GetFile(filePath).Size
        End If
    End If
    
    Select Case LCase(byteSize)
        Case "kb"
            totalBytes = totalBytes / (2 ^ 10)
        Case "mb"
            totalBytes = totalBytes / (2 ^ 20)
        Case "gb"
            totalBytes = totalBytes / (2 ^ 30)
    End Select

    FileSize = totalBytes

End Function


Public Function FileType( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the file type of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the file type of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: FileType() -> "Microsoft Excel Macro-Enabled Worksheet"
    '@Example: FileType("C:\hello\world.txt") -> "Text Document"
    '@Example: FileType("vba.txt") -> "Text Document"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    If filePath = "" Then
        FileType = FSO.GetFile(GetActivePathAndName()).Type
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileType = FSO.GetFile(GetActivePath() & filePath).Type
        Else
            FileType = FSO.GetFile(filePath).Type
        End If
    End If

End Function


Public Function FileExtension( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the extension of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the extension of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileExtension() = "xlsx"
    '@Example: =FileExtension("C:\hello\world.txt") -> "txt"
    '@Example: =FileExtension("vba.txt") -> "txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim FileName As String
    If filePath = "" Then
        FileName = FSO.GetFile(GetActivePathAndName()).Name
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileName = FSO.GetFile(GetActivePath() & filePath).Name
        Else
            FileName = FSO.GetFile(filePath).Name
        End If
    End If
    
    FileExtension = Right(FileName, Len(FileName) - InStrRev(FileName, "."))

End Function


Public Function ReadFile( _
    ByVal filePath As String, _
    Optional ByVal lineNumber As Integer) _
As String

    '@Description: This function reads the file specified in the file path argument and returns it's contents. Optionally, a line number can be specified so that only a single line is read. If a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Param: lineNumber is the number of the line that will be read, and if left blank all the file contents will be read. Note that the first line starts at line number 1.
    '@Returns: Returns the contents of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function may run very slowly when running it on large files. Also, for files that are not in text format (such as compressed zip files) this file contents returned will not be in a usable format.
    '@Example: =ReadFile("C:\hello\world.txt") -> "Hello" World
    '@Example: =ReadFile("vba.txt") -> "This is my VBA text file"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in
    '@Example: =ReadFile("multline.txt", 1) -> "This is line 1";
    '@Example: =ReadFile("multline.txt", 2) -> "This is line 2";

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim FileName As String
    Dim fileStream As Object
    
    ' Checking if the file exists in the current directory, and then if it
    ' exists in the path specified, and if it doesn't exist in either, returns
    ' a "#FileDoesntExist!"
    If FSO.FileExists(GetActivePath() & filePath) Then
        filePath = GetActivePath() & filePath
    ElseIf FSO.FileExists(filePath) Then
        filePath = filePath
    Else
        ReadFile = "#FileDoesntExist!"
    End If
    
    Set fileStream = FSO.GetFile(filePath)
    Set fileStream = fileStream.OpenAsTextStream(1, -2)
    
    
    ' If lineNumber is positive, read a line, else read the whole contents
    If lineNumber > 0 Then
        Dim fileLinesArray() As String
        
        fileLinesArray = Split(fileStream.ReadAll(), vbCrLf)
        ReadFile = fileLinesArray(lineNumber)
    Else
        ReadFile = fileStream.ReadAll()
    End If

End Function


Public Function WriteFile( _
    ByVal filePath As String, _
    ByVal fileText As String, _
    Optional ByVal appendModeFlag As Boolean) _
As Boolean

    '@Description: This function creates and writes to the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Refactor a bit shorten the code a bit, such as the area where the file is written to.
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Param: fileText is the text that will be written to the file
    '@Param: appendModeFlag is a Boolean value that if set to TRUE will append to the existing file instead of creating a new file and writing over the contents.
    '@Returns: Returns a message stating the file written to successfully
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: Be careful when writing files, as misuse of this function can results in files being overwritten accidently as well as creating large numbers of files accidently.
    '@Example: =WriteFile("C:\MyWorkbookFolder\hello.txt", "Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"
    '@Example: =WriteFile("hello.txt", "Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"; Where the Workbook resides in "C:\MyWorkbookFolder\"

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Checking if the folder exists if the path is an absolute path
    If InStr(filePath, "\") = 0 Then
        If InStr(filePath, "/") = 0 Then
            filePath = GetActivePath() & filePath
        End If
    ElseIf Right(filePath, 1) = "\" Or Right(filePath, 1) = "/" Then
        If Not FSO.FolderExists(Left(filePath, InStrRev(filePath, "\"))) Then
            WriteFile = False
            Exit Function
        End If
    ElseIf Not FSO.FolderExists(filePath) Then
        WriteFile = False
        Exit Function
    End If
    
    
    ' Writing to the file
    Dim fileStream As Object
    
    If appendModeFlag = False Then
        Set fileStream = FSO.CreateTextFile(filePath, True)
        fileStream.Write fileText
        
    Else
        Dim fileObject As Object
        
        Set fileObject = FSO.GetFile(filePath)
        Set fileStream = fileObject.OpenAsTextStream(8)
        fileStream.Write fileText
    End If
    
    WriteFile = True

End Function


Public Function PathSeparator() As String

    '@Description: This function returns the path separator character of the OS running this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Example: =PathSeparator() -> "\"; When running this code on Windows
    '@Example: =PathSeparator() -> "/"; When running this code on Mac
    
    #If Mac Then
        PathSeparator = "/"
    #Else
        PathSeparator = "\"
    #End If

End Function


Public Function PathJoin( _
    ParamArray pathArray() As Variant) _
As String

    '@Description: This function combines multiple strings into a file path by placing the path separator character between the arguments
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: pathArray is an array of strings that will be combined into a path
    '@Returns: Returns a string with the combined file path
    '@Example: =PathJoin("C:", "hello", "world.txt") -> "C:\hello\world.txt"; On Windows
    '@Example: =PathJoin("hello", "world.txt") -> "/hello/world.txt"; On Mac

    Dim individualPath As Variant
    Dim combinedPath As String
    Dim individualValue As Variant

    For Each individualPath In pathArray
        If IsArray(individualPath) Then
            For Each individualValue In individualPath
                combinedPath = combinedPath & individualValue & PathSeparator()
            Next
        Else
            combinedPath = combinedPath & CStr(individualPath) & PathSeparator()
        End If
    Next
    
    combinedPath = Left(combinedPath, Len(combinedPath) - 1)
    
    PathJoin = combinedPath
    
End Function


Public Function CountFiles( _
    Optional ByVal filePath As String) _
As Integer

    '@Description: This function returns the number of files at the specified folder path. If no path is given, the current workbook path will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the number of files in the folder
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function includes the count for hidden files as well. For example, when a workbook is open, a hidden file for the workbook is created, so if you run this function in the same folder as the workbook and notice the file count is one higher than expected, it is likely due to the hidden file.
    '@Example: =CountFiles() -> 6
    '@Example: =CountFiles("C:\hello") -> 10

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        CountFiles = FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).Files.Count
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            CountFiles = FSO.GetFolder(GetActivePath() & filePath).Files.Count
        Else
            CountFiles = FSO.GetFolder(filePath).Files.Count
        End If
    End If

End Function


Public Function CountFolders( _
    Optional ByVal filePath As String) _
As Integer

    '@Description: This function returns the number of folders at the specified folder path. If no path is given, the current workbook path will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the number of folders in the folder
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function includes the count for hidden folders as well. Hidden folders are often prefixed with a . character at the beginning
    '@Example: =CountFolders() -> 2
    '@Example: =CountFolders("C:\hello") -> 20

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        CountFolders = FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).SubFolders.Count
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            CountFolders = FSO.GetFolder(GetActivePath() & filePath).SubFolders.Count
        Else
            CountFolders = FSO.GetFolder(filePath).SubFolders.Count
        End If
    End If

End Function


Public Function CountFilesAndFolders( _
    Optional ByVal filePath As String) _
As Integer

    '@Description: This function returns the number of files and folders at the specified folder path. If no path is given, the current workbook path will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the number of files and folders in the folder
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function includes the count for hidden files and folders as well
    '@Example: =CountFilesAndFolders() -> 8
    '@Example: =CountFilesAndFolders("C:\hello") -> 30

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        CountFilesAndFolders = FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).Files.Count + FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).SubFolders.Count
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            CountFilesAndFolders = FSO.GetFolder(GetActivePath() & filePath).Files.Count + FSO.GetFolder(GetActivePath() & filePath).SubFolders.Count
        Else
            CountFilesAndFolders = FSO.GetFolder(filePath).Files.Count + FSO.GetFolder(filePath).SubFolders.Count
        End If
    End If

End Function


Public Function GetFileNameByNumber( _
    Optional ByVal filePath As String, _
    Optional ByVal fileNumber As Integer = -1) _
As String

    '@Description: This function returns the name of a file in a folder given the number of the file in the list of all files
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Param: fileNumber is the number of the file in the folder. For example, if there are 3 files in a folder, this should be a number between 1 and 3
    '@Returns: Returns the name of the specified file
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function includes hidden files as well. For example, when a workbook is open, a hidden file for the workbook is created, so if you run this function in the same folder as the workbook and notice the file count is one higher than expected, it is likely due to the hidden file.
    '@Example: =GetFileName(,1) -> "hello.txt"
    '@Example: =GetFileName(,1) -> "world.txt"
    '@Example: =GetFileName("C:\hello", 1) -> "one.txt"
    '@Example: =GetFileName("C:\hello", 1) -> "two.txt"
    '@Example: =GetFileName("C:\hello", 1) -> "three.txt"

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim fileCounter As Integer
    Dim individualFile As Object
    Dim fileCollection As Object
    
    If filePath = "" Then
        Set fileCollection = FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).Files
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            Set fileCollection = FSO.GetFolder(GetActivePath() & filePath).Files
        Else
            Set fileCollection = FSO.GetFolder(filePath).Files
        End If
    End If
    
    For Each individualFile In fileCollection
        fileCounter = fileCounter + 1
        If fileNumber = -1 Then
            GetFileNameByNumber = individualFile.Name
            Exit Function
        ElseIf fileCounter = fileNumber Then
            GetFileNameByNumber = individualFile.Name
            Exit Function
        End If
    Next

End Function

'@Module: This module contains a set of basic mathematical functions where those functions don't already exist as base Excel functions.



Public Function Ceil( _
    ByVal number As Double) _
As Long

    '@Description: This function takes a number and rounds it up to the nearest whole integer
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: number is the number that will be rounded up
    '@Returns: Returns the number rounded up to the nearest integer
    '@Example: =Ceil(1.5) -> 2
    '@Example: =Ceil(1.0001) -> 2
    '@Example: =Ceil(1.0) -> 1
    '@Example: =Ceil(1) -> 1

    If number = Fix(number) Then
        Ceil = number
    Else
        Ceil = Fix(number + 1)
    End If

End Function


Public Function Floor( _
    ByVal number As Double) _
As Long

    '@Description: This function takes a number and rounds it down to the nearest whole integer
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: number is the number that will be rounded down
    '@Returns: Returns the number rounded down to the nearest integer
    '@Example: =Floor(1.9) -> 1
    '@Example: =Floor(1.1) -> 1
    '@Example: =Floor(1.0) -> 1
    '@Example: =Floor(1) -> 1

    Floor = Fix(number)

End Function


Public Function InterpolateNumber( _
    ByVal startingNumber As Double, _
    ByVal endingNumber As Double, _
    ByVal interpolationPercentage As Double) _
As Double

    '@Description: This function takes three numbers, a starting number, an ending number, and an interpolation percent, and linearly interpolates the number at the given percentage between the starting and ending number.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: startingNumber is the beginning number of the interpolation
    '@Param: endingNumber is the ending number of the interpolation
    '@Param: interpolationPercentage is the percentage that will be interpolated linearly between the startingNumber and the endingNumber
    '@Returns: Returns the linearly interpolated number between the two points
    '@Example: =InterpolateNumber(10, 20, 0.5) -> 15; Where 0.5 would be 50% between 10 and 20
    '@Example: =InterpolateNumber(16, 124, 0.64) -> 85.12; Where 0.64 would be 64% between 16 and 124

    InterpolateNumber = startingNumber + ((endingNumber - startingNumber) * interpolationPercentage)

End Function


Public Function InterpolatePercent( _
    ByVal startingNumber As Double, _
    ByVal endingNumber As Double, _
    ByVal interpolationNumber As Double) _
As Double

    '@Description: This function takes three numbers, a starting number, an ending number, and an interpolation number, and linearly interpolates the percentage location of the interpolated number between the starting and ending number.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: startingNumber is the beginning number of the interpolation
    '@Param: endingNumber is the ending number of the interpolation
    '@Param: interpolationNumber is the number that will be interpolated linearly between the startingNumber and the endingNumber to calculate a percentage
    '@Returns: Returns the linearly interpolated percent between the two points given the interpolation number
    '@Example: =InterpolatePercent(10, 18, 12) -> 0.25; As 12 is 25% of the way from 10 to 18
    '@Example: =InterpolatePercent(10, 20, 15) -> 0.5; As 15 is 50% of the way from 10 to 20

    InterpolatePercent = (interpolationNumber - startingNumber) / (endingNumber - startingNumber)

End Function


Public Function Max( _
    ParamArray numbers() As Variant) _
As Double

    '@Description: This function takes multiple numbers or multiple arrays of numbers and returns the max number. This function also accounts for numbers that are formatted as strings by converting them into numbers
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: numbers is a single number, multiple numbers, or multiple arrays of numbers
    '@Returns: Returns the max number
    '@Example: =Max(1, 2, 3) -> 3
    '@Example: =Max(4.4, 5, "6") -> 6
    '@Example: =Max(x) -> 3; Where x is an array with these values [1, 2.2, "3"]
    '@Example: =Max(x, y, 10) -> 15; Where x = [1, 2.2, "3"] and y = [5, 15, -100]

    Dim individualParamArrayValue As Variant
    Dim individualValue As Variant
    Dim maxValue As Variant
    
    maxValue = Empty
    
    For Each individualParamArrayValue In numbers
        If IsArray(individualParamArrayValue) Then
            For Each individualValue In individualParamArrayValue
                If TypeName(individualValue) = "String" Then
                    individualValue = CDbl(individualValue)
                End If
            
                If IsEmpty(maxValue) Then
                    maxValue = individualValue
                ElseIf individualValue > maxValue Then
                    maxValue = individualValue
                End If
            Next
        Else
            If TypeName(individualParamArrayValue) = "String" Then
                individualParamArrayValue = CDbl(individualParamArrayValue)
            End If
        
            If IsEmpty(maxValue) Then
                maxValue = individualParamArrayValue
            ElseIf individualParamArrayValue > maxValue Then
                maxValue = individualParamArrayValue
            End If
        End If
    Next
    
    Max = maxValue

End Function


Public Function Min( _
    ParamArray numbers() As Variant) _
As Double

    '@Description: This function takes multiple numbers or multiple arrays of numbers and returns the min number. This function also accounts for numbers that are formatted as strings by converting them into numbers
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: numbers is a single number, multiple numbers, or multiple arrays of numbers
    '@Returns: Returns the min number
    '@Example: =Min(1, 2, 3) -> 1
    '@Example: =Min(4.4, 5, "6") -> 4.4
    '@Example: =Min(-1, -2, -3) -> -3
    '@Example: =Min(x) -> 1; Where x is an array with these values [1, 2.2, "3"]
    '@Example: =Min(x, y, 10) -> -100; Where x = [1, 2.2, "3"] and y = [5, 15, -100]

    Dim individualParamArrayValue As Variant
    Dim individualValue As Variant
    Dim minValue As Variant
    
    minValue = Empty
    
    For Each individualParamArrayValue In numbers
        If IsArray(individualParamArrayValue) Then
            For Each individualValue In individualParamArrayValue
                If TypeName(individualValue) = "String" Then
                    individualValue = CDbl(individualValue)
                End If
            
                If IsEmpty(minValue) Then
                    minValue = individualValue
                ElseIf individualValue < minValue Then
                    minValue = individualValue
                End If
            Next
        Else
            If TypeName(individualParamArrayValue) = "String" Then
                individualParamArrayValue = CDbl(individualParamArrayValue)
            End If
        
            If IsEmpty(minValue) Then
                minValue = individualParamArrayValue
            ElseIf individualParamArrayValue < minValue Then
                minValue = individualParamArrayValue
            End If
        End If
    Next
    
    Min = minValue

End Function


Public Function ModFloat( _
    numerator As Double, _
    denominator As Double) _
As Double

    '@Description: This function performs modulus operations with floats as the Mod operator in VBA does not support floats.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Find out if numerator and denominator are the correct names for Modulo operation
    '@Param: numerator is the left value of the Mod
    '@Param: denominator is the right value of the Mod
    '@Returns: Returns a double with ModFloat operator performed on it
    '@Example: =ModFloat(3.55, 2) -> 1.55

    Dim modValue As Double

    modValue = numerator - Fix(numerator / denominator) * denominator

    If modValue >= -2 ^ -52 Then
        If modValue <= 2 ^ -52 Then
            modValue = 0
        End If
    End If
    
    ModFloat = modValue
    
End Function

'@Module: This module contains a set of functions that return information on the Xlib library, such as the version number, credits, and a link to the documentation.



Public Function XlibVersion() As String

    '@Description: This function returns the version number of XPlus
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns the XPlus version number
    '@Example: =XlibVersion() -> "1.0.0"; Where the version of XPlus you are using is 1.0.0

    XlibVersion = "1.0.0"

End Function


Public Function XlibCredits() As String

    '@Description: This function returns credits for the XPlus library
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns the XPlus credits
    '@Example: =XlibCredits() -> "Copyright (c) 2020 Anthony Mancini. XLib is Licensed under an MIT License."

    XlibCredits = "Copyright (c) 2020 Anthony Mancini. XLib is Licensed under an MIT License."

End Function


Public Function XlibDocumentation() As String

    '@Description: This function returns a link to the Documentation for XPlus
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns the XPlus Documentation link
    '@Example: =XlibDocumentation() -> "https://x-vba.com/xlib"

    XlibDocumentation = "https://x-vba.com/xlib"

End Function


'@Module: This module contains a set of functions for performing networking tasks such as performing HTTP requests and parsing HTML.



Public Function Http( _
    ByVal url As String, _
    Optional ByVal httpMethod As String = "GET", _
    Optional ByVal headers As Variant, _
    Optional ByVal postData As Variant = "", _
    Optional ByVal asyncFlag As Boolean, _
    Optional ByVal statusErrorHandlerFlag As Boolean, _
    Optional ByVal parseArguments As Variant) _
As String

    '@Description: This function performs an HTTP request to the web and returns the response as a string. It provides many options to change the http method, provide data for a POST request, change the headers, handle errors for non-successful requests, and parse out text from a request using a light parsing language.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: url is a string of the URL of the website you want to fetch data from
    '@Param: httpMethod is a string with the http method, with the default being a GET request. For POST requests, use "POST", for PUT use "PUT", and for DELETE use "DELETE"
    '@Param: headers is either an array or a Scripting Dictionary of headers that will be used in the request. For an array, the 1st, 3rd, 5th... will be used as the key and the 2nd, 4th, 6th... will be used as the values. For a Scripting Dictionary, the dictionary keys will be used as header keys, and the values as values. Finally, in the case when no headers are set, the User-Agent will be set to "XPlus" as a courtesy to the web server.
    '@Param: postData is a string that will contain data for a POST request
    '@Param: asyncFlag is a Boolean value that if set to TRUE will make the request asynchronous. By default requests will be synchronous, which will lock Excel while fetching but will also prevent errors when performing calculations based on fetched data.
    '@Param: statusErrorHandlerFlag is a Boolean value that if set to TRUE will result in a User-Defined Error String being returned for all non 200 requests that tells the user the status code that occured. This flag is useful in cases where requests need to be successful and if not errors should be thrown.
    '@Param: parseArguments is an array of arguments that perform string parsing on the response. It uses a light scripting language that includes commands similar to the Excel Built-in LEFT(), RIGHT(), and MID() that allow you to parse the request before it gets returned. See the Note on the scripting language, and the Warning on why this argument should be used.
    '@Returns: Returns the parsed HTTP response as a string
    '@Note: The parseArguments parameter uses a light scripting language to perform string manipulations on the HTTP response text that allows you to parse out the relevant information to you. The language contains 5 commands that can be used for parsing. Please check out the examples as well below for a better understanding of how to use the parsing language:<br><br> {"ID", "idOfAnElement"} -> HTML inside of the element with the specified ID <br> {"TAG", "div", 2} -> HTML inside of the second div tag found <br> {"LEFT", 100} -> The 100 leftmost characters <br> {"LEFT", "Hello World"} -> All characters left of the first "Hello World" found in the HTML <br> {"RIGHT", 100} -> The 100 rightmost characters <br> {"RIGHT", "Hello World"} -> All characters right of the last "Hello World" found in the HTML <br> {"MID", 100} -> All character to the right of the 100th character in the string <br> {"MID", "Hello World"} -> All characters right of the first "Hello World" found in the HTML
    '@Warning: Excel has a limit on the number of characters that can be placed within a cell. This limit is a max of 32767 characters. If the request returns any more than this, a #VALUE! error will be returned. Most webpages surpass this number of characters, which makes the Excel Built-in function WEBSERVICE() not very useful. However, internally VBA can handle around 2,000,000,000 characters, which more characters that found on virtually every single webpage. As a result, parsing arguments should be used with this function so that you can parse out the relevant information for a request without this function failing. See the Note on the syntax of the light parsing language.
    '@Example: =Http("https://httpbin.org/uuid") -> "{"uuid: "41416bcf-ef11-4256-9490-63853d14e4e8"}"
    '@Example: =Http("https://httpbin.org/user-agent", "GET", {"User-Agent","MicrosoftExcel"}) -> "{"user-agent": "MicrosoftExcel"}"
    '@Example: =Http("https://httpbin.org/status/404",,,,,TRUE) -> "#RequestFailedStatusCode404!"; Since the status error handler flag is set and since this URL returns a 404 status code. Also note that this formula is easier to construct using the Excel Formula Builder
    '@Example: =Http("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000}) -> Returning a string with the leftmost 3000 characters found within the element with the ID "mw-content-text" (we are trying to get the release date of VBA from the VBA wikipedia page, but we need to do more parsing first)
    '@Example: =Http("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000,"MID","appeared"}) -> Returns the prior string, but now with all characters right of the first occurance of the word "appeared" in the HTML (getting closer to parsing the VBA creation date)
    '@Example: =Http("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>"}) -> From the prior result, now returning everything after the first occurance of the "<TD>" in the prior string
    '@Example: =Http("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>","LEFT","<span"}) -> "1993"; Finally this is all the parsing needed to be able to return the date 1993 that we were looking for

    Dim WinHttpRequest As Object
    Set WinHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    WinHttpRequest.Open httpMethod, url, asyncFlag
    
    ' Setting the request headers
    ' Case where headers come in the form of an Array
    If IsArray(headers) Then
        Dim i As Integer
        
        For i = 0 To UBound(headers) - LBound(headers) Step 2
            WinHttpRequest.SetRequestHeader headers(i), headers(i + 1)
        Next
        
    ' Case where headers come in the form of a Dictionary
    ElseIf TypeName(headers) = "Dictionary" Then
        Dim dictKey As Variant
        
        For Each dictKey In headers.Keys()
            WinHttpRequest.SetRequestHeader dictKey, headers(dictKey)
        Next
        
    ' In cases where no headers are given by the user, set a base User-Agent to
    ' "XPlus" as a courtesy to the webserver
    Else
        WinHttpRequest.SetRequestHeader "User-Agent", "XLib"
    End If
    
    ' Sending the HTTP request
    If postData = "" Then
        WinHttpRequest.Send
    Else
        WinHttpRequest.Send postData
    End If
    
    ' If the status error handler flag is set to True, then enable error returns
    ' in cases where the status code is not a 200
    If statusErrorHandlerFlag Then
        If WinHttpRequest.Status = 200 Then
            Http = WinHttpRequest.ResponseText
        Else
            Http = "#RequestFailedStatusCode" & WinHttpRequest.Status & "!"
        End If
    
    ' Case when the status code error handler is not used
    Else
        Http = WinHttpRequest.ResponseText
    End If
    
    ' Parsing Html Response
    If IsArray(parseArguments) Then
        Dim reorderedParseArguments() As Variant
        i = UBound(parseArguments) - LBound(parseArguments)
        ReDim reorderedParseArguments(i)
        
        ' Reordering here, as possibly had some name collision with the name parseArguments somewhere

        For i = 0 To UBound(parseArguments) - LBound(parseArguments)
            reorderedParseArguments(i) = parseArguments(i)
        Next
        
        Http = ParseHtmlString(Http, reorderedParseArguments)
    
    End If

End Function


Public Function SimpleHttp( _
    ByVal url As String, _
    ParamArray parseArguments() As Variant) _
As String

    '@Description: This function performs an HTTP request to the web and returns the response as a string, similar to the HTTP() function, except that only requires one parameter, the URL, and then takes an infinite number of strings after it as the parsing arguments instead of requiring an Array to use. Essentially, this function is a little cleaner to set up when performing very basic GET requests.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: url is a string of the URL of the website you want to fetch data from
    '@Param: parseArguments is an array of arguments that perform string parsing on the response. It uses a light scripting language that includes commands similar to the Excel Built-in LEFT(), RIGHT(), and MID() that allow you to parse the request before it gets returned. See the Note on the HTTP() function, and the Warning on the HTTP() function on why this argument should be used.
    '@Returns: Returns the parsed HTTP response as a string
    '@Example: =SimpleHttp("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications","ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>","LEFT","<span") -> "1993"; See the examples in the HTTP() function, as this example has the same result as the example in the HTTP() function. You can see that this function is cleaner and easier to set up than the corresponding HTTP() function.

    ' Case where parse arguments are provided
    If UBound(parseArguments) > 0 Then
        ' Need to reorder the arguments of the Array since when the caller is a
        ' Range, the Array is 1-based, where as when the caller is another VBA function,
        ' the Array is 0-based
        Dim i As Integer
        Dim reorderedParseArguments() As Variant
        i = UBound(parseArguments) - LBound(parseArguments)
        ReDim reorderedParseArguments(i)
        
        ' Reordering for Range
        For i = 0 To UBound(parseArguments) - LBound(parseArguments)
            reorderedParseArguments(i) = parseArguments(i)
        Next
        
        SimpleHttp = ParseHtmlString(Http(url), reorderedParseArguments)
    
    ' In case of no parse arguments, simply perform an HTTP request
    Else
        SimpleHttp = Http(url)
    End If

End Function


Public Function ParseHtmlString( _
    ByVal htmlString As String, _
    ByVal parseArguments As Variant) _
As Variant

    '@Description: This function parses an HTML string using the same parsing language that the HTTP() function uses. See the HTTP() function for more information on how to use this function.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: htmlString is a string of the HTML
    '@Param: parseArguments is an array of arguments that perform string parsing on the response. It uses a light scripting language that includes commands similar to the Excel Built-in LEFT(), RIGHT(), and MID() that allow you to parse the request before it gets returned. See the Note on the HTTP() function, and the Warning on the HTTP() function on why this argument should be used.
    '@Returns: Returns the parsed HTTP response as a string
    '@Example: =ParseHtmlString("HTML String from the webpage: https://en.wikipedia.org/wiki/Visual_Basic_for_Applications","ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>","LEFT","<span") -> "1993"

    Dim partialHtml As String
    Dim html As Object
    Set html = CreateObject("HtmlFile")
    
    ' Setting the HTML Document
    html.body.innerHTML = htmlString
    
    ' Parsing out info from the HTML Document
    Dim i As Integer
    
    For i = LBound(parseArguments) To UBound(parseArguments)
        ' Note that id and tag will truncate poorly formatted HTML
        ' Works with late bindings
        If LCase(parseArguments(i)) = "id" Then
            If partialHtml <> "" Then
                html.body.innerHTML = partialHtml
            End If
            partialHtml = html.getElementById(parseArguments(i + 1)).innerHTML
            html.body.innerHTML = partialHtml
            i = i + 1
            
        ' Requires early bindings. Don't include in final code, but potentially consider for future updates
        'ElseIf LCase(parseArguments(i)) = "class" Then
        '    partialHtml = html.getElementsByClassName(parseArguments(i + 1))(i + 2).innerHTML
        '    i = i + 2
        
        ' Works with late bindings
        ElseIf LCase(parseArguments(i)) = "tag" Then
            If partialHtml <> "" Then
                html.body.innerHTML = partialHtml
            End If
            partialHtml = html.getElementsByTagName(parseArguments(i + 1))(i + 2).innerHTML
            html.body.innerHTML = partialHtml
            i = i + 2
            
        ' Left string manipulation
        ElseIf LCase(parseArguments(i)) = "left" Then
            If IsNumeric(parseArguments(i + 1)) And TypeName(parseArguments(i + 1)) <> "String" Then
                partialHtml = Left(partialHtml, parseArguments(i + 1))
            Else
                partialHtml = Left(partialHtml, InStr(1, partialHtml, CStr(parseArguments(i + 1)), vbTextCompare) - 1)
            End If
            i = i + 1
            
        ' Right string manipulation
        ElseIf LCase(parseArguments(i)) = "right" Then
            If IsNumeric(parseArguments(i + 1)) And TypeName(parseArguments(i + 1)) <> "String" Then
                partialHtml = Right(partialHtml, parseArguments(i + 1))
            Else
                partialHtml = Right(partialHtml, Len(partialHtml) - Len(parseArguments(i + 1)) + 1 - InStrRev(partialHtml, CStr(parseArguments(i + 1)), Compare:=vbTextCompare))
            End If
            i = i + 1
            
        ' Mid string manipulation. Possibly update this to allow Mid length argument
        ElseIf LCase(parseArguments(i)) = "mid" Then
            If IsNumeric(parseArguments(i + 1)) And TypeName(parseArguments(i + 1)) <> "String" Then
                partialHtml = Mid(partialHtml, parseArguments(i + 1))
            Else
                partialHtml = Mid(partialHtml, Len(parseArguments(i + 1)) + InStr(1, partialHtml, CStr(parseArguments(i + 1)), vbTextCompare))
            End If
            i = i + 1
        End If
    Next
    
    ParseHtmlString = partialHtml

End Function



'@Module: This module contains a set of functions for generating and sampling random data.



Public Function RandBetween( _
    ByVal minNumber As Long, _
    ByVal maxNumber As Long) _
As Variant

    '@Description: This function returns a random number between the min and max numbers
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: minNumber is the minimum number in the range
    '@Param: maxNumber is the maximum number in the range
    '@Returns: Returns a random number between the range
    '@Example: =RandBetween(1, 20) -> 5
    '@Example: =RandBetween(1, 20) -> 9
    '@Example: =RandBetween(1, 20) -> 13
    '@Example: =RandBetween(1, 20) -> 2
    '@Example: =RandBetween(1, 20) -> 20
    '@Example: =RandBetween(1, 20) -> 6

    RandBetween = Fix(Rnd * (maxNumber - minNumber + 1) + minNumber)

End Function


Public Function BigRandBetween( _
    ByVal minNumber As Variant, _
    ByVal maxNumber As Variant) _
As Variant

    '@Description: This function is an implementation of RandBetween that allows for 14-byte integers to be used
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: minNumber is the minimum number in the range
    '@Param: maxNumber is the maximum number in the range
    '@Returns: Returns a random number between the range
    '@Example: =RandBetween(0, 3000000000) -> Error; as RandBetween only works with 4-byte and less integers
    '@Example: =BigRandBetween(0, 3000000000) -> 2116642535; as BigRandBetween supports up to 14-byte integers

    BigRandBetween = Fix(Rnd * (maxNumber - minNumber + 1) + minNumber)

End Function


Public Function RandomSample( _
    ByRef variantArray As Variant) _
As Variant

    '@Description: This function takes an array of cells and returns a random value from the cells chosen
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: variantArray a single cell or multiple cells where the sample will be pulled from
    '@Returns: Returns a random cell value from the array of cells chosen
    '@Example: =RandomSample(A1:A5) -> "Hello"; where "Hello" is the value in cell A3, and where A3 was the chosen random cell
    '@Example: =RandomSample(A1:A5) -> "World"; where "World" is the value in cell A2, and where A2 was the chosen random cell

    Dim randomNumber As Long
    
    randomNumber = RandBetween(1, UBound(variantArray) - LBound(variantArray) + 1)
    
    RandomSample = variantArray(randomNumber - 1)

End Function


Public Function RandomRange( _
    ByVal startNumber As Long, _
    ByVal stopNumber As Long, _
    ByVal stepNumber As Long) _
As Long

    '@Description: This function takes 3 numbers, a start number, a stop number, and a step number, and returns a random number between the start number and stop number that is an interval of the step number.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: startNumber is the beginning value of the range
    '@Param: stopNumber is the end value of the range
    '@Param: stepNumber is the step of the range
    '@Returns: Returns a random number between the start and stop that is a multiple of the step
    '@Example: =RandomRange(50, 100, 10) -> 60
    '@Example: =RandomRange(50, 100, 10) -> 50
    '@Example: =RandomRange(50, 100, 10) -> 90
    '@Example: =RandomRange(0, 10, 2) -> 8
    '@Example: =RandomRange(0, 10, 2) -> 0
    '@Example: =RandomRange(0, 10, 2) -> 4
    '@Example: =RandomRange(0, 10, 2) -> 10

    Dim randomNumber As Long
    
    randomNumber = RandBetween(startNumber / stepNumber, stopNumber / stepNumber) * stepNumber
    
    RandomRange = randomNumber

End Function


Public Function RandBool() As Boolean

    '@Description: This function generates a random Boolean (TRUE or FALSE) value
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns either TRUE or FALSE based on the random value choosen
    '@Example: =RandBool() -> TRUE
    '@Example: =RandBool() -> FALSE
    '@Example: =RandBool() -> TRUE
    '@Example: =RandBool() -> TRUE
    '@Example: =RandBool() -> FALSE
    '@Example: =RandBool() -> FALSE

    RandBool = CBool(RandBetween(0, 1))

End Function


Public Function RandBetweens( _
    ParamArray startOrEndNumberArray() As Variant) _
As Variant

    '@Description: This function is similar to RANDBETWEEN, except that it allows multiple ranges from which to pick a random number. One of the ranges from which to generate a random number between is chosen at an equal probably.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns either TRUE or FALSE based on the random value choosen
    '@Note: This function always requires an even number of inputs. Essentially, when using multiple numbers, the 1st and 2nd will make up a range from which to pull a random number between, the 3rd and 4th will make a different range, and so on. If an even number is used, this function will return a User-Defined Error. See the ISERRORALL() function for how to handle these numbers.
    '@Example: =RandBetweens(1, 10, 5000, 5010) -> 6
    '@Example: =RandBetweens(1, 10, 5000, 5010) -> 5002
    '@Example: =RandBetweens(1, 10, 5000, 5010) -> 8
    '@Example: =RandBetweens(1, 10, 5000, 5010) -> 3
    '@Example: =RandBetweens(1, 10, 5000, 5010) -> 5010
    '@Example: =RandBetweens(1, 10, 5000, 5010) -> 2
    '@Example: =RandBetweens(5, 10, 15, 20, 25, 30, 35, 40) -> 32

    Dim pickNumber As Byte

    ' Checking for ParamArray length, as it needs to be even or it won't be
    ' possible to generate and min and max number.
    If (UBound(startOrEndNumberArray) - LBound(startOrEndNumberArray) + 1) Mod 2 = 1 Then
        RandBetweens = "#NotAnEvenNumberOfParameters!"
    End If

    pickNumber = Ceil(RandBetween(1, (UBound(startOrEndNumberArray) - LBound(startOrEndNumberArray) + 1)) / 2) * 2
    
    RandBetweens = RandBetween(startOrEndNumberArray(pickNumber - 2), startOrEndNumberArray(pickNumber - 1))

End Function


'@Module: This module contains a set of functions for performing Regular Expressions, which are a type of string pattern matching. For more info on Regular Expression Pattern matching, please check "https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference"



Public Function RegexSearch( _
    ByVal string1 As String, _
    ByVal stringPattern As String, _
    Optional ByVal globalFlag As Boolean, _
    Optional ByVal ignoreCaseFlag As Boolean, _
    Optional ByVal multilineFlag As Boolean) _
As String

    '@Description: This function takes a string that we will perform the Regular Expression on and a Regular Expression string pattern, and returns the first value of the matched string. This function also contains optional arguments for various Regular Expression flags.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that the regex will be performed on
    '@Param: stringPattern is the regex pattern
    '@Param: globalFlag is a boolean value that if set TRUE will perform a global search
    '@Param: ignoreCaseFlag is a boolean value that if set TRUE will perform a case insensitive search
    '@Param: multilineFlag is a boolean value that if set TRUE will perform a mulitline search
    '@Returns: Returns a string of the regex value that is found
    '@Example: =RegexSearch("Hello World","[a-z]{2}\s[W]") -> "lo W";

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
    Dim searchResults As Object
    
    With Regex
        .Global = globalFlag
        .IgnoreCase = ignoreCaseFlag
        .MultiLine = multilineFlag
        .Pattern = stringPattern
    End With
    
    Set searchResults = Regex.Execute(string1)
    
    RegexSearch = searchResults(0).Value

End Function


Public Function RegexTest( _
    ByVal string1 As String, _
    ByVal stringPattern As String, _
    Optional ByVal globalFlag As Boolean, _
    Optional ByVal ignoreCaseFlag As Boolean, _
    Optional ByVal multilineFlag As Boolean) _
As Boolean

    '@Description: This function takes a string that we will perform the Regular Expression on and a Regular Expression string pattern, and returns TRUE if the pattern is found in the string. This function also contains optional arguments for various Regular Expression flags.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that the regex will be performed on
    '@Param: stringPattern is the regex pattern
    '@Param: globalFlag is a boolean value that if set TRUE will perform a global search
    '@Param: ignoreCaseFlag is a boolean value that if set TRUE will perform a case insensitive search
    '@Param: multilineFlag is a boolean value that if set TRUE will perform a mulitline search
    '@Returns: Returns TRUE if the regex value that is found, or FALSE if it isn't
    '@Example: =RegexTest("Hello World","[a-z]{2}\s[W]") -> TRUE;

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
    
    With Regex
        .Global = globalFlag
        .IgnoreCase = ignoreCaseFlag
        .MultiLine = multilineFlag
        .Pattern = stringPattern
    End With
    
    RegexTest = Regex.Test(string1)

End Function


Public Function RegexReplace( _
    ByVal string1 As String, _
    ByVal stringPattern As String, _
    ByVal replacementString As String, _
    Optional ByVal globalFlag As Boolean, _
    Optional ByVal ignoreCaseFlag As Boolean, _
    Optional ByVal multilineFlag As Boolean) _
As String

    '@Description: This function takes a string that we will perform the Regular Expression on, a Regular Expression string pattern, and a string that we will replace if the pattern is found, and returns a new string with the replacement string in place of the pattern. This function also contains optional arguments for various Regular Expression flags.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that the regex will be performed on
    '@Param: stringPattern is the regex pattern
    '@Param: replacementString is a string that will be replaced if the pattern is found
    '@Param: globalFlag is a boolean value that if set TRUE will perform a global search
    '@Param: ignoreCaseFlag is a boolean value that if set TRUE will perform a case insensitive search
    '@Param: multilineFlag is a boolean value that if set TRUE will perform a mulitline search
    '@Returns: Returns a new string with the replaced string values
    '@Example: =RegexReplace("Hello World","[W][a-z]{4}", "VBA") -> "Hello VBA"

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
    
    With Regex
        .Global = globalFlag
        .IgnoreCase = ignoreCaseFlag
        .MultiLine = multilineFlag
        .Pattern = stringPattern
    End With
    
    RegexReplace = Regex.Replace(string1, replacementString)

End Function


'@Module: This module contains a set of basic functions for manipulating strings.



Public Function Capitalize( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and returns the same string with the first character capitalized and all other characters lowercased
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that the capitalization will be performed on
    '@Returns: Returns a new string with the first character capitalized and all others lowercased
    '@Example: =Capitalize("hello World") -> "Hello world"

    Capitalize = UCase(Left(string1, 1)) & LCase(Mid(string1, 2))
    
End Function


Public Function LeftFind( _
    ByVal string1 As String, _
    ByVal searchString As String) _
As String

    '@Description: This function takes a string and a search string, and returns a string with all characters to the left of the first search string found within string1. Similar to Excel's built-in =SEARCH() function, this function is case-sensitive. For a case-insensitive version of this function, see =LeftSearch().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be searched
    '@Param: searchString is the string that will be used to search within string1
    '@Returns: Returns a new string with all characters to the left of the first search string within string1
    '@Example: =LeftFind("Hello World", "r") -> "Hello Wo"
    '@Example: =LeftFind("Hello World", "R") -> "#VALUE!"; Since string1 does not contain "R" in it.

    LeftFind = Left(string1, InStr(1, string1, searchString) - 1)

End Function


Public Function RightFind( _
    ByVal string1 As String, _
    ByVal searchString As String) _
As String

    '@Description: This function takes a string and a search string, and returns a string with all characters to the right of the last search string found within string 1. Similar to Excel's built-in =SEARCH() function, this function is case-sensitive. For a case-insensitive version of this function, see =RightSearch().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be searched
    '@Param: searchString is the string that will be used to search within string1
    '@Returns: Returns a new string with all characters to the right of the last search string within string1
    '@Example: =RightFind("Hello World", "o") -> "rld"
    '@Example: =RightFind("Hello World", "O") -> "#VALUE!"; Since string1 does not contain "O" in it.

    RightFind = Right(string1, Len(string1) - InStrRev(string1, searchString))

End Function


Public Function LeftSearch( _
    ByVal string1 As String, _
    ByVal searchString As String) _
As String

    '@Description: This function takes a string and a search string, and returns a string with all characters to the left of the first search string found within string1. Similar to Excel's built-in =FIND() function, this function is NOT case-sensitive (it's case-insensitive). For a case-sensitive version of this function, see =LeftFind().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be searched
    '@Param: searchString is the string that will be used to search within string1
    '@Returns: Returns a new string with all characters to the left of the first search string within string1
    '@Example: =LeftSearch("Hello World", "r") -> "Hello Wo"
    '@Example: =LeftSearch("Hello World", "R") -> "Hello Wo"

    LeftSearch = Left(string1, InStr(1, string1, searchString, vbTextCompare) - 1)

End Function


Public Function RightSearch( _
    ByVal string1 As String, _
    ByVal searchString As String) _
As String

    '@Description: This function takes a string and a search string, and returns a string with all characters to the right of the last search string found within string 1. Similar to Excel's built-in =FIND() function, this function is NOT case-sensitive (it's case-insensitive). For a case-sensitive version of this function, see =RightFind().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be searched
    '@Param: searchString is the string that will be used to search within string1
    '@Returns: Returns a new string with all characters to the right of the last search string within string1
    '@Example: =RightSearch("Hello World", "o") -> "rld"
    '@Example: =RightSearch("Hello World", "O") -> "rld"

    RightSearch = Right(string1, Len(string1) - InStrRev(string1, searchString, Compare:=vbTextCompare))

End Function


Public Function Substr( _
    ByVal string1 As String, _
    ByVal startCharacterNumber As Integer, _
    ByVal endCharacterNumber As Integer) _
As String

    '@Description: This function takes a string and a starting character number and ending character number, and returns the substring between these two numbers. The total number of characters returned will be endCharacterNumber - startCharacterNumber.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that we will get a substring from
    '@Param: startCharacterNumber is the character number of the start of the substring, with 1 being the first character in the string
    '@Param: endCharacterNumber is the character number of the end of the substring
    '@Returns: Returns a substring between the two numbers.
    '@Example: =Substr("Hello World", 2, 6) -> "ello"

    Substr = Mid(string1, startCharacterNumber, endCharacterNumber - startCharacterNumber)

End Function


Public Function SubstrFind( _
    ByVal string1 As String, _
    ByVal RightFindString As String, _
    ByVal rightSearchString As String, _
    Optional ByVal noninclusiveFlag As Boolean) _
As String

    '@Description: This function takes a string and a left string and right string, and returns a substring between those two strings. The left string will find the first matching string starting from the left, and the right string will find the first matching string starting from the right. Finally, and optional final parameter can be set to TRUE to make the substring noninclusive of the two searched strings. SubstrFind is case-sensitive. For case-insensitive version, see SubstrSearch
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that we will get a substring from
    '@Param: RightFindString is the string that will be searched from the left
    '@Param: rightSearchString is the string that will be searched from the right
    '@Param: noninclusiveFlag is an optional parameter that if set to TRUE will result in the substring not including the left and right searched characters
    '@Returns: Returns a substring between the two strings.
    '@Example: =SubstrFind("Hello World", "e", "o") -> "ello Wo"
    '@Example: =SubstrFind("Hello World", "e", "o", TRUE) -> "llo W"
    '@Example: =SubstrFind("One Two Three", "ne ", " Thr") -> "ne Two Thr"
    '@Example: =SubstrFind("One Two Three", "NE ", " THR") -> "#VALUE!"; Since SubstrFind() is case-sensitive
    '@Example: =SubstrFind("One Two Three", "ne ", " Thr", TRUE) -> "Two"
    '@Example: =SubstrFind("Country Code: +51; Area Code: 315; Phone Number: 762-5929;", "Area Code: ", "; Phone", TRUE) -> 315
    '@Example: =SubstrFind("Country Code: +313; Area Code: 423; Phone Number: 284-2468;", "Area Code: ", "; Phone", TRUE) -> 423
    '@Example: =SubstrFind("Country Code: +171; Area Code: 629; Phone Number: 731-5456;", "Area Code: ", "; Phone", TRUE) -> 629

    Dim leftCharacterNumber As Integer
    Dim rightCharacterNumber As Integer
    
    leftCharacterNumber = InStr(1, string1, RightFindString)
    rightCharacterNumber = InStrRev(string1, rightSearchString)
    
    If noninclusiveFlag = True Then
        leftCharacterNumber = leftCharacterNumber + Len(RightFindString)
        rightCharacterNumber = rightCharacterNumber - Len(rightSearchString)
    End If
    
    SubstrFind = Mid(string1, leftCharacterNumber, rightCharacterNumber - leftCharacterNumber + Len(rightSearchString))

End Function


Public Function SubstrSearch( _
    ByVal string1 As String, _
    ByVal RightFindString As String, _
    ByVal rightSearchString As String, _
    Optional ByVal noninclusiveFlag As Boolean) _
As String

    '@Description: This function takes a string and a left string and right string, and returns a substring between those two strings. The left string will find the first matching string starting from the left, and the right string will find the first matching string starting from the right. Finally, and optional final parameter can be set to TRUE to make the substring noninclusive of the two searched strings. SubstrSearch is case-insensitive. For case-sensitive version, see SubstrFind
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Param: string1 is the string that we will get a substring from
    '@Param: RightFindString is the string that will be searched from the left
    '@Param: rightSearchString is the string that will be searched from the right
    '@Param: noninclusiveFlag is an optional parameter that if set to TRUE will result in the substring not including the left and right searched characters
    '@Returns: Returns a substring between the two strings.
    '@Example: =SubstrSearch("Hello World", "e", "o") -> "ello Wo"
    '@Example: =SubstrSearch("Hello World", "e", "o", TRUE) -> "llo W"
    '@Example: =SubstrSearch("One Two Three", "ne ", " Thr") -> "ne Two Thr"
    '@Example: =SubstrSearch("One Two Three", "NE ", " THR") -> "ne Two Thr"; No error, since SubstrSearch is case-insensitive
    '@Example: =SubstrSearch("One Two Three", "ne ", " Thr", TRUE) -> "Two"
    '@Example: =SubstrSearch("Country Code: +51; Area Code: 315; Phone Number: 762-5929;", "Area Code: ", "; Phone", TRUE) -> 315
    '@Example: =SubstrSearch("Country Code: +313; Area Code: 423; Phone Number: 284-2468;", "Area Code: ", "; Phone", TRUE) -> 423
    '@Example: =SubstrSearch("Country Code: +171; Area Code: 629; Phone Number: 731-5456;", "Area Code: ", "; Phone", TRUE) -> 629

    Dim leftCharacterNumber As Integer
    Dim rightCharacterNumber As Integer
    
    leftCharacterNumber = InStr(1, string1, RightFindString, vbTextCompare)
    rightCharacterNumber = InStrRev(string1, rightSearchString, Compare:=vbTextCompare)
    
    If noninclusiveFlag = True Then
        leftCharacterNumber = leftCharacterNumber + Len(RightFindString)
        rightCharacterNumber = rightCharacterNumber - Len(rightSearchString)
    End If
    
    SubstrSearch = Mid(string1, leftCharacterNumber, rightCharacterNumber - leftCharacterNumber + Len(rightSearchString))

End Function

    
Public Function Repeat( _
    ByVal string1 As String, _
    ByVal numberOfRepeats As Integer) _
As String

    '@Description: This function repeats string1 based on the number of repeats specified in the second argument
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be repeated
    '@Param: numberOfRepeats is the number of times string1 will be repeated
    '@Returns: Returns a string repeated multiple times based on the numberOfRepeats
    '@Example: =Repeat("Hello", 2) -> HelloHello"
    '@Example: =Repeat("=", 10) -> "=========="

    Dim i As Integer
    Dim combinedString As String

    For i = 1 To numberOfRepeats
        combinedString = combinedString & string1
    Next

    Repeat = combinedString

End Function


Public Function Formatter( _
    ByVal formatString As String, _
    ParamArray textArray() As Variant) _
As String

    '@Description: This function takes a Formatter string and then an array of ranges or strings, and replaces the format placeholders with the values in the range or strings. The format syntax is "{1} - {2}" where the "{1}" and "{2}" will be replaced with the values given in the text array.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: formatString is the string that will be used as the format and which will be replaced with the individual strings
    '@Param: textArray are the ranges or strings that will be placed within the slots of the format string
    '@Returns: Returns a new string with the individual strings in the placeholder slots of the format string
    '@Example: =Formatter("Hello {1}", "World") -> "Hello World"
    '@Example: =Formatter("{1} {2}", "Hello", "World") -> "Hello World"
    '@Example: =Formatter("{1}.{2}@{3}", "FirstName", "LastName", "email.com") -> "FirstName.LastName@email.com"
    '@Example: =Formatter("{1}.{2}@{3}", A1:A3) -> "FirstName.LastName@email.com"; where A1="FirstName", A2="LastName", and A3="email.com"
    '@Example: =Formatter("{1}.{2}@{3}", A1, A2, A3) -> "FirstName.LastName@email.com"; where A1="FirstName", A2="LastName", and A3="email.com"

    Dim i As Byte
    Dim individualTextItem As Variant
    Dim individualValue As Variant
    
    i = 0
    
    For Each individualTextItem In textArray
        If IsArray(individualTextItem) Then
            For Each individualValue In individualTextItem
                i = i + 1
                
                formatString = Replace(formatString, "{" & i & "}", individualValue)
            Next
        Else
            i = i + 1
            
            formatString = Replace(formatString, "{" & i & "}", individualTextItem)
        End If
    Next

    Formatter = formatString

End Function


Public Function Zfill( _
    ByVal string1 As String, _
    ByVal fillLength As Byte, _
    Optional ByVal fillCharacter As String = "0", _
    Optional ByVal rightToLeftFlag As Boolean) _
As String

    '@Description: This function pads zeros to the left of a string until the string is at least the length of the fill length. Optional parameters can be used to pad with a different character than 0, and to pad from right to left instead of from the default left to right.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be filled
    '@Param: fillLength is the length that string1 will be padded to. In cases where string1 is of greater length than this argument, no padding will occur.
    '@Param: fillCharacter is an optional string that will change the character that will be padded with
    '@Param: rightToLeftFlag is a Boolean parameter that if set to TRUE will result in padding from right to leftt instead of left to right
    '@Returns: Returns a new padded string of the length of specified by fillLength at minimum
    '@Example: =Zfill(123, 5) -> "00123"
    '@Example: =Zfill(5678, 5) -> "05678"
    '@Example: =Zfill(12345678, 5) -> "12345678"
    '@Example: =Zfill(123, 5, "X") -> "XX123"
    '@Example: =Zfill(123, 5, "X", TRUE) -> "123XX"
    
    While Len(string1) < fillLength
        If rightToLeftFlag = False Then
            string1 = fillCharacter + string1
        Else
            string1 = string1 + fillCharacter
        End If
    Wend
    
    Zfill = string1

End Function


Public Function SplitText( _
    ByVal string1 As String, _
    ByVal substringNumber As Integer, _
    Optional ByVal delimiterString As String = " ") _
As String
    
    '@Description: This function takes a string and a number, splits the string by the space characters, and returns the substring in the position of the number specified in the second argument.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be split and a substring returned
    '@Param: substringNumber is the number of the substring that will be chosen
    '@Param: delimiterString is an optional parameter that can be used to specify a different delimiter
    '@Returns: Returns a substring of the split text in the location specified
    '@Example: =SplitText("Hello World", 1) -> "Hello"
    '@Example: =SplitText("Hello World", 2) -> "World"
    '@Example: =SplitText("One Two Three", 2) -> "Two"
    '@Example: =SplitText("One-Two-Three", 2, "-") -> "Two"
    
    SplitText = Split(string1, delimiterString)(substringNumber - 1)

End Function


Public Function CountWords( _
    ByVal string1 As String, _
    Optional ByVal delimiterString As String = " ") _
As Integer

    '@Description: This function takes a string and returns the number of words in the string
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Note: If the number given is higher than the number of words, its possible that the string contains excess whitespace. Try using the =TRIM() function first to remove the excess whitespace
    '@Param: string1 is the string whose number of words will be counted
    '@Param: delimiterString is an optional parameter that can be used to specify a different delimiter
    '@Returns: Returns the number of words in the string
    '@Example: =CountWords("Hello World") -> 2
    '@Example: =CountWords("One Two Three") -> 3
    '@Example: =CountWords("One-Two-Three", "-") -> 3

    Dim stringArray() As String

    stringArray = Split(string1, delimiterString)
    
    CountWords = UBound(stringArray) - LBound(stringArray) + 1

End Function


Public Function CamelCase( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and returns the same string in camel case, removing all the spaces.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be camel cased
    '@Returns: Returns a new string in camel case, where the first character of the first word is lowercase, and uppercased for all other words
    '@Example: =CamelCase("Hello World") -> "helloWorld"
    '@Example: =CamelCase("One Two Three") -> "oneTwoThree"

    Dim i As Integer
    Dim stringArray() As String
    
    stringArray = Split(string1, " ")
    stringArray(0) = LCase(stringArray(0))
    
    For i = 1 To (UBound(stringArray) - LBound(stringArray))
        stringArray(i) = UCase(Left(stringArray(i), 1)) & LCase(Mid(stringArray(i), 2))
    Next
    
    CamelCase = Join(stringArray, "")

End Function


Public Function KebabCase( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and returns the same string in kebab case.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be kebab cased
    '@Returns: Returns a new string in kebab case, where all letters are lowercase and seperated by a "-" character
    '@Example: =KebabCase("Hello World") -> "hello-world"
    '@Example: =KebabCase("One Two Three") -> "one-two-three"

    KebabCase = LCase(Join(Split(string1, " "), "-"))

End Function


Public Function RemoveCharacters( _
    ByVal string1 As String, _
    ParamArray removedCharacters() As Variant) _
As String

    '@Description: This function takes a string and either another string or multiple strings and removes all characters from the first string that are in the second string.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Consider adding a Boolean flag that will make non-case sensitive replacements
    '@Note: This function is case sensitive. If you want to remove the "H" from "Hello World" you would need to use "H" as a removed character, not "h".
    '@Param: string1 is the string that will have characters removed
    '@Param: removedCharacters is an array of strings that will be removed from string1
    '@Returns: Returns the origional string with characters removed
    '@Example: =RemoveCharacters("Hello World", "l") -> "Heo Word"
    '@Example: =RemoveCharacters("Hello World", "lo") -> "He Wrd"
    '@Example: =RemoveCharacters("Hello World", "l", "o") -> "He Wrd"
    '@Example: =RemoveCharacters("Hello World", "lod") -> "He Wr"
    '@Example: =RemoveCharacters("Two Three Four", "f", "t") -> "Two Three Four"; Nothing is replaced since this function is case sensitive
    '@Example: =RemoveCharacters("Two Three Four", "F", "T") -> "wo hree our"

    Dim i As Integer
    Dim individualCharacter As Variant
    
    For Each individualCharacter In removedCharacters
        If Len(individualCharacter) > 1 Then
            For i = 1 To Len(individualCharacter)
                string1 = Replace(string1, Mid(individualCharacter, i, 1), "")
            Next
        Else
            string1 = Replace(string1, individualCharacter, "")
        End If
    Next
    
    RemoveCharacters = string1

End Function


Private Function NumberOfUppercaseLetters( _
    ByVal string1 As String) _
As Integer

    '@Description: This function returns the number of uppercase letter found within a string based on the ASCII character code range for uppercase letters
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose uppercase letters will be counted
    '@Returns: Returns the number of uppercase letters

    Dim i As Integer
    Dim numberOfUppercase As Integer
    
    For i = 1 To Len(string1)
        If Asc(Mid(string1, i, 1)) >= 65 Then
            If Asc(Mid(string1, i, 1)) <= 90 Then
                numberOfUppercase = numberOfUppercase + 1
            End If
        End If
    Next
    
    NumberOfUppercaseLetters = numberOfUppercase

End Function


Public Function CompanyCase( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and uses an algorithm to return the string in Company Case. The standard =PROPER() function in Excel will not capitalize company names properly, as it only capitalizes based on space characters, so a name like "j.p. morgan" will be incorrectly formatted as "J.p. Morgan" instead of the correct "J.P. Morgan". Additionally =PROPER() may incorrectly lowercase company abbreviations, such as the last "H" in "GmbH", as =PROPER() returns "Gmbh" instead of the correct "GmbH". This function attempts to adjust for these issues when a string is a company name.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Warning: There is no perfect algorithm for correctly formatting company names, and while this function can give better performance for correct formatting when compared to =PROPER(), if the performance of this function isn't as accurate as one needs, another solution would be to try Partial Lookup functions in the String Metrics Module and compare that to a known list of well formatted company strings.
    '@Param: string1 is the string that will be formatted
    '@Returns: Returns the origional string in a Company Case format
    '@Example: =CompanyCase("hello world") -> "Hello World"
    '@Example: =CompanyCase("x.y.z company & co.") -> "X.Y.Z Company & Co."
    '@Example: =CompanyCase("x.y.z plc") -> "X.Y.Z PLC"
    '@Example: =CompanyCase("one company gmbh") -> "One Company GmbH"
    '@Example: =CompanyCase("three company s. en n.c.") -> "Three Company S. en N.C."
    '@Example: =CompanyCase("FOUR COMPANY SPOL S.R.O.") -> "Four Company spol s.r.o."
    '@Example: =CompanyCase("five company bvba") -> "Five Company BVBA"

    Dim i As Integer
    Dim k As Integer
    Dim origionalString As String
    Dim stringArray() As String
    Dim splitCharacters As String
    
    origionalString = string1
    string1 = LCase(string1)
    splitCharacters = " ./()-_,*&1234567890"
    
    For k = 1 To Len(splitCharacters)
        stringArray = Split(string1, Mid(splitCharacters, k, 1))
        For i = 0 To UBound(stringArray) - LBound(stringArray)
            If NumberOfUppercaseLetters(Split(origionalString, Mid(splitCharacters, k, 1))(i)) <= 1 Then
                stringArray(i) = UCase(Left(stringArray(i), 1)) & Mid(stringArray(i), 2)
            Else
                If UCase(Join(stringArray, Mid(splitCharacters, k, 1))) = origionalString Then
                    stringArray(i) = UCase(Left(stringArray(i), 1)) & Mid(stringArray(i), 2)
                Else
                    stringArray(i) = Split(origionalString, Mid(splitCharacters, k, 1))(i)
                End If
            End If
            
        Next
        string1 = Join(stringArray, Mid(splitCharacters, k, 1))
    Next
    
    
    ' Checking the final words in the string to see if they are one of the
    ' company abbreviation strings, and if it is, replace the ending with
    ' the correct cases of the company abbreviation
    Dim companyAbbreviationArray() As String
    companyAbbreviationArray = Split("AB|AG|GmbH|LLC|LLP|NV|PLC|SA|A. en P.|ACE|AD|AE|AL|AmbA|ANS|ApS|AS|ASA|AVV|BVBA|CA|CVA|d.d.|d.n.o.|d.o.o.|DA|e.V.|EE|EEG|EIRL|ELP|EOOD|EPE|EURL|GbR|GCV|GesmbH|GIE|HB|hf|IBC|j.t.d.|k.d.|k.d.d.|k.s.|KA/S|KB|KD|KDA|KG|KGaA|KK|Kol. SrK|Kom. SrK|LDC|Ltï¿½e.|NT|OE|OHG|Oy|OYJ|Oï¿½|PC Ltd|PMA|PMDN|PrC|PT|RAS|S. de R.L.|S. en N.C.|SA de CV|SAFI|SAS|SC|SCA|SCP|SCS|SENC|SGPS|SK|SNC|SOPARFI|sp|Sp. z.o.o.|SpA|spol s.r.o.|SPRL|TD|TLS|v.o.s.|VEB|VOF|BYSHR", "|")

    Dim stringArrayLength As Integer

    stringArray = Split(string1, " ")
    stringArrayLength = UBound(stringArray) - LBound(stringArray)

    Dim companyAbbreviationString As Variant
    
    For Each companyAbbreviationString In companyAbbreviationArray
        If InStrRev(LCase(string1), " " & LCase(companyAbbreviationString)) = (Len(string1) - Len(companyAbbreviationString)) Then
            If InStrRev(LCase(string1), " " & LCase(companyAbbreviationString)) <> 0 Then
                CompanyCase = Left(string1, InStrRev(LCase(string1), LCase(companyAbbreviationString)) - 1) & companyAbbreviationString
                Exit Function
            End If
        End If
    Next

    CompanyCase = string1

End Function


Public Function ReverseText( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and reverses all the characters in it so that the returned string is backwards
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be reversed
    '@Returns: Returns the origional string in reverse
    '@Example: =ReverseText("Hello World") -> "dlroW olleH"

    Dim i As Integer
    Dim reversedString As String
    
    For i = 1 To Len(string1)
        reversedString = reversedString & Mid(string1, Len(string1) - i + 1, 1)
    Next
    
    ReverseText = reversedString

End Function


Public Function ReverseWords( _
    ByVal string1 As String, _
    Optional ByVal delimiterCharacter As String = " ") _
As String

    '@Description: This function takes a string and reverses all the words in it so that the returned string's words are backwards. By default, this function uses the space character as a delimiter, but you can optionally specify a different delimiter.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose words will be reversed
    '@Param: delimiterCharacter is the delimiter that will be used, with the default being " "
    '@Returns: Returns the origional string with it's words reversed
    '@Example: =ReverseWords("Hello World") -> "World Hello"
    '@Example: =ReverseWords("One Two Three") -> "Three Two One"
    '@Example: =ReverseWords("One-Two-Three", "-") -> "Three-Two-One"

    Dim i As Integer
    Dim stringArray() As String
    Dim stringArrayLength As Integer
    Dim reversedStringArray() As String
    
    stringArray = Split(string1, delimiterCharacter)
    stringArrayLength = (UBound(stringArray) - LBound(stringArray))
    
    ReDim reversedStringArray(stringArrayLength)
    
    For i = 0 To stringArrayLength
        reversedStringArray(i) = stringArray(stringArrayLength - i)
    Next
    
    ReverseWords = Join(reversedStringArray, delimiterCharacter)

End Function


Public Function IndentText( _
    ByVal string1 As String, _
    Optional ByVal indentAmount As Byte = 4) _
As String

    '@Description: This function takes a string and indents all of its lines by a specified number of space characters (or 4 space characters if left blank)
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be indented
    '@Param: indentAmount is the amount of " " characters that will be indented to the left of string1
    '@Returns: Returns the origional string indented by a specified number of space characters
    '@Example: =IndentText("Hello") -> "    Hello"
    '@Example: =IndentText("Hello", 4) -> "    Hello"
    '@Example: =IndentText("Hello", 3) -> "   Hello"
    '@Example: =IndentText("Hello", 2) -> "  Hello"
    '@Example: =IndentText("Hello", 1) -> " Hello"

    Dim i As Integer
    Dim stringArray() As String

    stringArray = Split(string1, Chr(10))
    
    string1 = ""
    For i = 1 To indentAmount
        string1 = string1 & " "
    Next
    
    For i = 0 To (UBound(stringArray) - LBound(stringArray))
        stringArray(i) = string1 & stringArray(i)
    Next

    IndentText = Join(stringArray, Chr(10))

End Function


Public Function DedentText( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and dedents all of its lines so that there are no space characters to the left or right of each line
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be dedented
    '@Returns: Returns the origional string dedented on each line
    '@Note: Unlike the Excel built-in TRIM() function, this function will dedent every single line, so for strings that span multiple lines in a cell, this will dedent all lines.
    '@Example: =DedentText("    Hello") -> "Hello"

    Dim i As Integer
    Dim stringArray() As String

    stringArray = Split(string1, Chr(10))
    
    For i = 0 To (UBound(stringArray) - LBound(stringArray))
        stringArray(i) = Trim(stringArray(i))
    Next

    DedentText = Join(stringArray, Chr(10))

End Function


Public Function ShortenText( _
    ByVal string1 As String, _
    Optional ByVal shortenWidth As Integer = 80, _
    Optional ByVal placeholderText As String = "[...]", _
    Optional ByVal delimiterCharacter As String = " ") _
As String

    '@Description: This function takes a string and shortens it with placeholder text so that it is no longer in length than the specified width.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be shortened
    '@Param: shortenWidth is the max width of the string. By default this is set to 80
    '@Param: placeholderText is the text that will be placed at the end of the string if it is longer than the shortenWidth. By default this placeholder string is "[...]
    '@Param: delimiterCharacter is the character that will be used as the word delimiter. By default this is the space character " "
    '@Returns: Returns a shortened string with placeholder text if it is longer than the shorten width
    '@Example: =ShortenText("Hello World One Two Three", 20) -> "Hello World [...]"; Only the first two words and the placeholder will result in a string that is less than or equal to 20 in length
    '@Example: =ShortenText("Hello World One Two Three", 15) -> "Hello [...]"; Only the first word and the placeholder will result in a string that is less than or equal to 15 in length
    '@Example: =ShortenText("Hello World One Two Three") -> "Hello World One Two Three"; Since this string is shorter than the default 80 shorten width value, no placeholder will be used and the string wont be shortened
    '@Example: =ShortenText("Hello World One Two Three", 15, "-->") -> "Hello World -->"; A new placeholder is used
    '@Example: =ShortenText("Hello_World_One_Two_Three", 15, "-->", "_") -> "Hello_World_-->"; A new placeholder andd delimiter is used

    Dim shortenedString As String
    Dim individualString As Variant
    Dim stringArray() As String
    
    ' In cases where the origional string is less than the threshold needed to
    ' shorten the string, simply return the origional string
    If Len(string1) <= (shortenWidth - Len(placeholderText) - Len(delimiterCharacter)) Then
        ShortenText = string1
        Exit Function
    End If
    
    stringArray = Split(string1, delimiterCharacter)

    For Each individualString In stringArray
        If Len(shortenedString & individualString) > (shortenWidth - Len(placeholderText) - Len(delimiterCharacter)) Then
            shortenedString = shortenedString & placeholderText
            Exit For
        Else
            shortenedString = shortenedString & individualString & delimiterCharacter
        End If
    Next

    ShortenText = shortenedString

End Function


Public Function InSplit( _
    ByVal string1 As String, _
    ByVal splitString As String, _
    Optional ByVal delimiterCharacter As String = " ") _
As Boolean

    '@Description: This function takes a search string and checks if it exists within a larger string that is split by a delimiter character.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be checked if it exists within the splitString after the split
    '@Param: splitString is the string that will be split and of which string1 will be searched in
    '@Param: delimiterCharacter is the character that will be used as the delimiter for the split. By default this is the space character " "
    '@Returns: Returns TRUE if string1 is found in splitString after the split occurs
    '@Example: =InSplit("Hello", "Hello World One Two Three") -> TRUE; Since "Hello" is found within the searchString after being split
    '@Example: =InSplit("NotInString", "Hello World One Two Three") -> FALSE; Since "NotInString" is not found within the searchString after being split
    '@Example: =InSplit("Hello", "Hello-World-One-Two-Three", "-") -> TRUE; Since "Hello" is found and since the delimiter is set to "-"

    Dim individualString As Variant
    
    For Each individualString In Split(splitString, delimiterCharacter)
        If string1 = individualString Then
            InSplit = True
            Exit Function
        End If
    Next
    
    InSplit = False

End Function


Public Function EliteCase( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and returns the string with characters replaced by similar in appearance numbers
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will have characters replaced
    '@Returns: Returns the string with characters replaced with similar in appearance numbers
    '@Example: =EliteCase("Hello World") -> "H3110 W0r1d"

    string1 = Replace(string1, "o", "0", Compare:=vbTextCompare)
    string1 = Replace(string1, "l", "1", Compare:=vbTextCompare)
    string1 = Replace(string1, "z", "2", Compare:=vbTextCompare)
    string1 = Replace(string1, "e", "3", Compare:=vbTextCompare)
    string1 = Replace(string1, "a", "4", Compare:=vbTextCompare)
    string1 = Replace(string1, "s", "5", Compare:=vbTextCompare)
    string1 = Replace(string1, "t", "7", Compare:=vbTextCompare)

    EliteCase = string1

End Function


Public Function ScrambleCase( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string scrambles the case on each character in the string
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose character's cases will be scrambled
    '@Returns: Returns the origional string with cases scrambled
    '@Example: =ScrambleCase("Hello World") -> "helLo WORlD"
    '@Example: =ScrambleCase("Hello World") -> "HElLo WorLD"
    '@Example: =ScrambleCase("Hello World") -> "hELlo WOrLd"

    Dim i As Integer

    For i = 1 To Len(string1)
        If RandBetween(0, 1) = 1 Then
            Mid(string1, i, 1) = UCase(Mid(string1, i, 1))
        Else
            Mid(string1, i, 1) = LCase(Mid(string1, i, 1))
        End If
    Next
    
    ScrambleCase = string1

End Function


Public Function LeftSplit( _
    ByVal string1 As String, _
    ByVal numberOfSplit As Integer, _
    Optional ByVal delimiterCharacter As String = " ") _
As String

    '@Description: This function takes a string, splits it based on a delimiter, and returns all characters to the left of the specified position of the split.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be split to get a substring
    '@Param: numberOfSplit is the number of the location within the split that we will get all characters to the left of
    '@Param: delimiterCharacter is the delimiter that will be used for the split. By default, the delimiter will be the space character " "
    '@Returns: Returns all characters to the left of the number of the split
    '@Example: =LeftSplit("Hello World One Two Three", 1) -> "Hello"
    '@Example: =LeftSplit("Hello World One Two Three", 2) -> "Hello World"
    '@Example: =LeftSplit("Hello World One Two Three", 3) -> "Hello World One"
    '@Example: =LeftSplit("Hello World One Two Three", 10) -> "Hello World One Two Three"
    '@Example: =LeftSplit("Hello-World-One-Two-Three", 2, "-") -> "Hello-World"

    Dim i As Integer
    Dim newString As String
    Dim stringArray() As String
    Dim stringArrayLength As Integer
    
    numberOfSplit = numberOfSplit - 1
    stringArray = Split(string1, delimiterCharacter)
    stringArrayLength = (UBound(stringArray) - LBound(stringArray) + 1)
    
    ' Checking if the number of split is greater than the length of the split
    ' array, and if so returns the origional string
    If numberOfSplit >= stringArrayLength Then
        LeftSplit = string1
        Exit Function
    End If
    
    For i = 0 To numberOfSplit
        If i = numberOfSplit Then
            newString = newString & stringArray(i)
        Else
            newString = newString & stringArray(i) & delimiterCharacter
        End If
    Next
    
    LeftSplit = newString

End Function


Public Function RightSplit( _
    ByVal string1 As String, _
    ByVal numberOfSplit As Integer, _
    Optional ByVal delimiterCharacter As String = " ") _
As String

    '@Description: This function takes a string, splits it based on a delimiter, and returns all characters to the right of the specified position of the split.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be split to get a substring
    '@Param: numberOfSplit is the number of the location within the split that we will get all characters to the right of
    '@Param: delimiterCharacter is the delimiter that will be used for the split. By default, the delimiter will be the space character " "
    '@Returns: Returns all characters to the right of the number of the split
    '@Example: =RightSplit("Hello World One Two Three", 1) -> "Three"
    '@Example: =RightSplit("Hello World One Two Three", 2) -> "Two Three"
    '@Example: =RightSplit("Hello World One Two Three", 3) -> "One Two Three"
    '@Example: =RightSplit("Hello World One Two Three", 10) -> "Hello World One Two Three"
    '@Example: =RightSplit("Hello-World-One-Two-Three", 2, "-") -> "Two-Three"

    Dim i As Integer
    Dim newString As String
    Dim stringArray() As String
    Dim stringArrayLength As Integer
    
    numberOfSplit = numberOfSplit - 1
    stringArray = Split(string1, delimiterCharacter)
    stringArrayLength = (UBound(stringArray) - LBound(stringArray) + 1)
    
    ' Checking if the number of split is greater than the length of the split
    ' array, and if so returns the origional string
    If numberOfSplit >= stringArrayLength Then
        RightSplit = string1
        Exit Function
    End If
    
    For i = 0 To numberOfSplit
        If i = numberOfSplit Then
            newString = newString & stringArray(stringArrayLength - (numberOfSplit - i) - 1)
        Else
            newString = newString & stringArray(stringArrayLength - (numberOfSplit - i) - 1) & delimiterCharacter
        End If
    Next
    
    RightSplit = newString

End Function


Public Function TrimChar( _
    ByVal string1 As String, _
    Optional ByVal trimCharacter As String = " ") _
As String

    '@Description: This function takes a string trims characters to the left and right of the string, similar to Excel's Built-in TRIM() function, except that an optional different character can be used for the trim.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Allow more than 1 character to be used for trimming
    '@Param: string1 is the string that will be trimmed
    '@Param: trimCharacter is an optional character that will be trimmed from the string. By default, this character will be the space character " "
    '@Returns: Returns the origional string with characters trimmed from the left and right
    '@Note: This function currently supports only single characters for trimming
    '@Example: =TrimChar("   Hello World   ") -> "Hello World"
    '@Example: =TrimChar("---Hello World---", "-") -> "Hello World"

    While Left(string1, 1) = trimCharacter
        Mid(string1, 1) = Chr(1)
        string1 = Replace(string1, Chr(1), "")
    Wend
    
    While Right(string1, 1) = trimCharacter
        Mid(string1, Len(string1)) = Chr(1)
        string1 = Replace(string1, Chr(1), "")
    Wend
    
    TrimChar = string1

End Function


Public Function TrimLeft( _
    ByVal string1 As String, _
    Optional ByVal trimCharacter As String = " ") _
As String

    '@Description: This function takes a string trims characters to the left of the string, similar to Excel's Built-in TRIM() function, except that an optional different character can be used for the trim.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Allow more than 1 character to be used for trimming
    '@Param: string1 is the string that will be trimmed
    '@Param: trimCharacter is an optional character that will be trimmed from the string. By default, this character will be the space character " "
    '@Returns: Returns the origional string with characters trimmed from the left only
    '@Note: This function currently supports only single characters for trimming
    '@Example: =TrimLeft("   Hello World   ") -> "Hello World   "
    '@Example: =TrimLeft("---Hello World---", "-") -> "Hello World---"

    While Left(string1, 1) = trimCharacter
        Mid(string1, 1) = Chr(1)
        string1 = Replace(string1, Chr(1), "")
    Wend
    
    TrimLeft = string1

End Function


Public Function TrimRight( _
    ByVal string1 As String, _
    Optional ByVal trimCharacter As String = " ") _
As String

    '@Description: This function takes a string trims characters to the right of the string, similar to Excel's Built-in TRIM() function, except that an optional different character can be used for the trim.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Allow more than 1 character to be used for trimming
    '@Param: string1 is the string that will be trimmed
    '@Param: trimCharacter is an optional character that will be trimmed from the string. By default, this character will be the space character " "
    '@Returns: Returns the origional string with characters trimmed from the right only
    '@Note: This function currently supports only single characters for trimming
    '@Example: =TrimRight("   Hello World   ") -> "   Hello World"
    '@Example: =TrimRight("---Hello World---", "-") -> "---Hello World"
    
    While Right(string1, 1) = trimCharacter
        Mid(string1, Len(string1)) = Chr(1)
        string1 = Replace(string1, Chr(1), "")
    Wend
    
    TrimRight = string1

End Function


Public Function CountUppercaseCharacters( _
    ByVal string1 As String) _
As Integer

    '@Description: This function takes a string and counts the number of uppercase characters in it
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose characters will be counted
    '@Returns: Returns the number of uppercase characters in the string
    '@Example: =CountUppercaseCharacters("Hello World") -> 2; As the "H" and the "E" are the only 2 uppercase characters in the string

    Dim i As Integer
    Dim characterAsciiCode As Byte
    Dim uppercaseCounter As Integer
    
    For i = 1 To Len(string1)
        characterAsciiCode = Asc(Mid(string1, i, 1))
        If characterAsciiCode >= 65 And characterAsciiCode <= 90 Then
            uppercaseCounter = uppercaseCounter + 1
        End If
    Next
    
    CountUppercaseCharacters = uppercaseCounter

End Function


Public Function CountLowercaseCharacters( _
    ByVal string1 As String) _
As Integer

    '@Description: This function takes a string and counts the number of lowercase characters in it
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose characters will be counted
    '@Returns: Returns the number of lowercase characters in the string
    '@Example: =CountLowercaseCharacters("Hello World") -> 8; As the "ello" and the "orld" are lowercase

    Dim i As Integer
    Dim characterAsciiCode As Byte
    Dim lowercaseCounter As Integer
    
    For i = 1 To Len(string1)
        characterAsciiCode = Asc(Mid(string1, i, 1))
        If characterAsciiCode >= 97 And characterAsciiCode <= 122 Then
            lowercaseCounter = lowercaseCounter + 1
        End If
    Next
    
    CountLowercaseCharacters = lowercaseCounter

End Function


Public Function TextJoin( _
    ByVal stringArray As Variant, _
    Optional ByVal delimiterCharacter As String, _
    Optional ByVal ignoreEmptyCellsFlag As Boolean) _
As String

    '@Description: This function takes a range of cells and combines all the text together, optionally allowing a character delimiter between all the combined strings, and optionally allowing blank cells to be ignored when combining the text. Finally note that this function is very similar to the TEXTJOIN function available in Excel 2019, and thus is a polyfill for that function for earlier versions of Excel.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: stringArray is the range with all the strings we want to combine
    '@Param: delimiterCharacter is an optional character that will be used as the delimiter between the combined text. By default, no delimiter character will be used.
    '@Param: ignoreEmptyCellsFlag if set to TRUE will skip combining empty cells into the combined string, and is useful when specifying a delimiter so that the delimiter does not repeat for empty cells.
    '@Returns: Returns a new combined string containing the strings in the range delimited by the delimiter character.
    '@Example: =TextJoin(A1:A3) -> "123"; Where A1:A3 contains ["1", "2", "3"]
    '@Example: =TextJoin(A1:A3, "--") -> "1--2--3"; Where A1:A3 contains ["1", "2", "3"]
    '@Example: =TextJoin(A1:A3, "--") -> "1----3"; Where A1:A3 contains ["1", "", "3"]
    '@Example: =TextJoin(A1:A3, "-") -> "1--3"; Where A1:A3 contains ["1", "", "3"]
    '@Example: =TextJoin(A1:A3, "-", TRUE) -> "1-3"; Where A1:A3 contains ["1", "", "3"]

    Dim individualString As Variant
    Dim combinedString As String
    
    For Each individualString In stringArray
        individualString = CStr(individualString)
        If ignoreEmptyCellsFlag Then
            If Not (IsEmpty(individualString) Or individualString = "") Then
                combinedString = combinedString & individualString & delimiterCharacter
            End If
        Else
            combinedString = combinedString & individualString & delimiterCharacter
        End If
    Next
    
    If delimiterCharacter <> "" Then
        combinedString = Left(combinedString, InStrRev(combinedString, delimiterCharacter) - 1)
    End If
    
    TextJoin = combinedString

End Function

'@Module: This module contains a set of functions for performing fuzzy string matches. It can be useful when you have 2 columns containing text that is close but not 100% the same. However, since the functions in this module only perform fuzzy matches, there is no guarantee that there will be 100% accuracy in the matches. However, for small groups of string where each string is very different than the other (such as a small group of fairly dissimilar names), these functions can be highly accurate. Finally, some of the functions in this Module will take a long time to calculate for large numbers of cells, as the number of calculations for some functions will grow exponentially, but for small sets of data (such as 100 strings to compare), these functions perform fairly quickly.



'========================================
'  Hamming Distance
'========================================

Public Function Hamming( _
    string1 As String, _
    string2 As String) _
As Integer

    '@Description: This function takes two strings of the same length and calculates the Hamming Distance between them. Hamming Distance measures how close two strings are by checking how many Substitutions are needed to turn one string into the other. Lower numbers mean the strings are closer than high numbers.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the first string
    '@Param: string2 is the second string that will be compared to the first string
    '@Returns: Returns an integer of the Hamming Distance between two string
    '@Example: =Hamming("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
    '@Example: =Hamming("Cat", "Bag") -> 2; 2 changes are needed, changing the "B" to "C" and the "g" to "t"
    '@Example: =Hamming("Cat", "Dog") -> 3; Every single character needs to be substituted in this case

    If Len(string1) <> Len(string2) Then
        Hamming = CVErr(2015)
    End If
    
    Dim totalDistance As Integer
    totalDistance = 0
    
    Dim i As Integer
    
    For i = 1 To Len(string1)
        If Mid(string1, i, 1) <> Mid(string2, i, 1) Then
            totalDistance = totalDistance + 1
        End If
    Next
    
    Hamming = totalDistance
    
End Function



'========================================
'  Levenshtein Distance
'========================================

Public Function Levenshtein( _
    string1 As String, _
    string2 As String) _
As Long

    '@Description: This function takes two strings of any length and calculates the Levenshtein Distance between them. Levenshtein Distance measures how close two strings are by checking how many Insertions, Deletions, or Substitutions are needed to turn one string into the other. Lower numbers mean the strings are closer than high numbers. Unlike Hamming Distance, Levenshtein Distance works for strings of any length and includes 2 more operations. However, calculation time will be slower than Hamming Distance for same length strings, so if you know the two strings are the same length, its preferred to use Hamming Distance.
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Param: string1 is the first string
    '@Param: string2 is the second string that will be compared to the first string
    '@Returns: Returns an integer of the Levenshtein Distance between two string
    '@Example: =Levenshtein("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
    '@Example: =Levenshtein("Cat", "Ca") -> 1; Since only one Insertion needs to occur (adding a "t" at the end)
    '@Example: =Levenshtein("Cat", "Cta") -> 2; Since the "t" in "Cta" needs to be substituted into an "a", and the final character "a" needs to be substituted into a "t"

    ' **Error Checking**
    ' Quick returns for common errors
    If string1 = string2 Then
        Levenshtein = 0
        Exit Function
    ElseIf string1 = Empty Then
        Levenshtein = Len(string2)
        Exit Function
    ElseIf string2 = Empty Then
        Levenshtein = Len(string1)
        Exit Function
    End If
    

    ' **Algorithm Code**
    ' Creating the distance metrix and filling it with values
    Dim numberOfRows As Integer
    Dim numberOfColumns As Integer
    
    numberOfRows = Len(string1)
    numberOfColumns = Len(string2)
    
    Dim distanceArray() As Integer
    ReDim distanceArray(numberOfRows, numberOfColumns)
    
    Dim r As Integer
    Dim c As Integer
    
    For r = 0 To numberOfRows
        For c = 0 To numberOfColumns
            distanceArray(r, c) = 0
        Next
    Next
    
    For r = 1 To numberOfRows
        distanceArray(r, 0) = r
    Next
    
    For c = 1 To numberOfColumns
        distanceArray(0, c) = c
    Next
    
    ' Non-recursive Levenshtein Distance matrix walk
    Dim operationCost As Integer
    
    For c = 1 To numberOfColumns
        For r = 1 To numberOfRows
            If Mid(string1, r, 1) = Mid(string2, c, 1) Then
                operationCost = 0
            Else
                operationCost = 1
            End If
                                                           
            distanceArray(r, c) = Min(distanceArray(r - 1, c) + 1, distanceArray(r, c - 1) + 1, distanceArray(r - 1, c - 1) + operationCost)
        Next
    Next
    
    Levenshtein = distanceArray(numberOfRows, numberOfColumns)

End Function



'========================================
'  Damerau-Levenshtein Distance
'========================================

Public Function Damerau( _
    string1 As String, _
    string2 As String) _
As Integer

    '@Description: This function takes two strings of any length and calculates the Damerau-Levenshtein Distance between them. Damerau-Levenshtein Distance differs from Levenshtein Distance in that it includes an additional operation, called Transpositions, which occurs when two adjacent characters are swapped. Thus, Damerau-Levenshtein Distance calculates the number of Insertions, Deletions, Substitutions, and Transpositons needed to convert string1 into string2. As a result, this function is good when it is likely that spelling errors have occured between two string where the error is simply a transposition of 2 adjacent characters.
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Param: string1 is the first string
    '@Param: string2 is the second string that will be compared to the first string
    '@Returns: Returns an integer of the Damerau-Levenshtein Distance between two string
    '@Example: =Damerau("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
    '@Example: =Damerau("Cat", "Ca") -> 1; Since only one Insertion needs to occur (adding a "t" at the end)
    '@Example: =Damerau("Cat", "Cta") -> 1; Since the "t" and "a" can be transposed as they are adjacent to each other. Notice how LEVENSHTEIN("Cat","Cta")=2 but DAMERAU("Cat","Cta")=1

    ' **Error Checking**
    ' Quick returns for common errors
    If string1 = string2 Then
        Damerau = 0
    ElseIf string1 = Empty Then
        Damerau = Len(string2)
    ElseIf string2 = Empty Then
        Damerau = Len(string1)
    End If
    
    Dim inf As Long
    Dim da As Object
    inf = Len(string1) + Len(string2)
    Set da = CreateObject("Scripting.Dictionary")
    
    ' 35 - 38 = filling the dictionary
    Dim i As Integer
    For i = 1 To Len(string1)
        If da.exists(Mid(string1, i, 1)) = False Then
            da.Add Mid(string1, i, 1), "0"
        End If
    Next
    
    For i = 1 To Len(string2)
        If da.exists(Mid(string2, i, 1)) = False Then
            da.Add Mid(string2, i, 1), "0"
        End If
    Next
    
    ' 39 = creating h matrix
    Dim H() As Long
    ReDim H(Len(string1) + 1, Len(string2) + 1)
    
    Dim k As Integer
    For i = 0 To (Len(string1) + 1)
        For k = 0 To (Len(string2) + 1)
            H(i, k) = 0
        Next
    Next
    
    ' 40 - 45 = updating the matrix
    For i = 0 To Len(string1)
        H(i + 1, 0) = inf
        H(i + 1, 1) = i
    Next
    For k = 0 To Len(string2)
        H(0, k + 1) = inf
        H(1, k + 1) = k
    Next
    

    ' 46 - 60 = running the array
    Dim db As Long
    Dim i1 As Long
    Dim k1 As Long
    Dim cost As Long
    
    For i = 1 To Len(string1)
        db = 0
        For k = 1 To Len(string2)
            i1 = CInt(da(Mid(string2, k, 1)))
            k1 = db
            cost = 1
            
            If Mid(string1, i, 1) = Mid(string2, k, 1) Then
                cost = 0
                db = k
            End If
            
            H(i + 1, k + 1) = Min(H(i, k) + cost, _
                                  H(i + 1, k) + 1, _
                                  H(i, k + 1) + 1, _
                                  H(i1, k1) + (i - i1 - 1) + 1 + (k - k1 - 1))
                           
            
        Next
        
        If da.exists(Mid(string1, i, 1)) Then
            da.Remove Mid(string1, i, 1)
            da.Add Mid(string1, i, 1), CStr(i)
        Else
            da.Add Mid(string1, i, 1), CStr(i)
        End If
        
    Next

    Damerau = H(Len(string1) + 1, Len(string2) + 1)

End Function


'@Module: This module contains a set of basic miscellaneous utility functions



Public Function Jsonify( _
    ByVal indentLevel As Byte, _
    ParamArray stringArray() As Variant) _
As String

    '@Description: This function takes an array of strings and numbers and returns the array as a JSON string. This function takes into account formatting for numbers, and supports specifying the indentation level.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: indentLevel is an optional number that specifying the indentation level. Leaving this argument out will result in no indentation
    '@Param: stringArray() is an array of strings and number in the following format: {"Hello", "World"}
    '@Returns: Returns a JSON valid string of all elements in the array
    '@Example: =Jsonify(0, "Hello", "World", "1", "2", 3, 4.5) -> "["Hello","World",1,2,3,4.5]"
    '@Example: =Jsonify(0, {"Hello", "World", "1", "2", 3, 4.5}, 10) -> "["Hello","World",1,2,3,4.5]"

    Dim i As Byte
    Dim jsonString As String
    Dim individualTextItem As Variant
    Dim individualValue As Variant
    Dim indentString As String
    
    ' Setting up some base JSON features and the indenting
    jsonString = "["
    
    For i = 1 To indentLevel
        indentString = indentString & " "
    Next
    
    If indentLevel > 0 Then
        jsonString = jsonString & Chr(10)
    End If
    
    
    ' Creating the contents of the JSON string
    For Each individualTextItem In stringArray
    
        ' In cases of ranges
        If IsArray(individualTextItem) Then
            For Each individualValue In individualTextItem
                jsonString = jsonString & indentString
                
                If IsNumeric(individualValue) Then
                    jsonString = jsonString & individualValue & ","
                Else
                    jsonString = jsonString & Chr(34) & individualValue & Chr(34) & ","
                End If
                
                If indentLevel > 0 Then
                    jsonString = jsonString & Chr(10)
                End If
            Next
            
        ' In cases of text
        Else
            jsonString = jsonString & indentString
            
            If IsNumeric(individualTextItem) Then
                jsonString = jsonString & individualTextItem & ","
            Else
                jsonString = jsonString & Chr(34) & individualTextItem & Chr(34) & ","
            End If
            
            If indentLevel > 0 Then
                jsonString = jsonString & Chr(10)
            End If
        End If

    Next
    
    jsonString = Left(jsonString, InStrRev(jsonString, ",") - 1)
    
    If indentLevel > 0 Then
        jsonString = jsonString & Chr(10)
    End If
    
    jsonString = jsonString & "]"
    
    Jsonify = jsonString

End Function


Public Function UuidFour() As String

    '@Description: This function generates a unique ID based on the UUID V4 specification. This function is useful for generating unique IDs of a fixed character length.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string unique ID based on UUID V4. The format of the string will always be in the form of "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx" where each x is a hex digit, and y is either 8, 9, A, or B.
    '@Example: =UuidFour() -> "3B4BDC26-E76E-4D6C-9E05-7AE7D2D68304"
    '@Example: =UuidFour() -> "D5761256-8385-4FDA-AD56-6AEF0AD6B9A5"
    '@Example: =UuidFour() -> "CDCAE2F5-B52F-4C90-A38A-42BD58BCED4B"

    Dim firstGroup As String
    Dim secondGroup As String
    Dim thirdGroup As String
    Dim fourthGroup As String
    Dim fifthGroup As String
    Dim sixthGroup As String

    firstGroup = BigDec2Hex(BigRandBetween(0, 4294967295#), 8) & "-"
    secondGroup = Dec2Hex(RandBetween(0, 65535), 4) & "-"
    thirdGroup = Dec2Hex(RandBetween(16384, 20479), 4) & "-"
    fourthGroup = Dec2Hex(RandBetween(32768, 49151), 4) & "-"
    fifthGroup = Dec2Hex(RandBetween(0, 65535), 4)
    sixthGroup = BigDec2Hex(BigRandBetween(0, 4294967295#), 8)

    UuidFour = firstGroup & secondGroup & thirdGroup & fourthGroup & fifthGroup & sixthGroup

End Function


Public Function HideText( _
    ByVal string1 As String, _
    ByVal hiddenFlag As Boolean, _
    Optional ByVal hideString As String) _
As String

    '@Description: This function takes the value in a cell and visibly hides it if the HideText flag set to TRUE. If TRUE, the value will appear as "********", with the option to set the HideText characters to a different set of text.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be HideText
    '@Param: hiddenFlag if set to TRUE will hide string1
    '@Param: hideString is an optional string that if set will be used instead of "********"
    '@Returns: Returns a string to hide string1 if hideFlag is TRUE
    '@Example: =HideText("Hello World", FALSE) -> "Hello World"
    '@Example: =HideText("Hello World", TRUE) -> "********"
    '@Example: =HideText("Hello World", TRUE, "[Hidden Text]") -> "[Hidden Text]"
    '@Example: =HideText("Hello World", UserName()="Anthony") -> "********"

    If hiddenFlag Then
        If hideString = "" Then
            HideText = "********"
        Else
            HideText = hideString
        End If
    Else
        HideText = string1
    End If

End Function


Public Function JavaScript( _
    ByVal jsFuncCode As String, _
    ByVal jsFuncName As String, _
    Optional ByVal argument1 As Variant, _
    Optional ByVal argument2 As Variant, _
    Optional ByVal argument3 As Variant, _
    Optional ByVal argument4 As Variant, _
    Optional ByVal argument5 As Variant, _
    Optional ByVal argument6 As Variant, _
    Optional ByVal argument7 As Variant, _
    Optional ByVal argument8 As Variant, _
    Optional ByVal argument9 As Variant, _
    Optional ByVal argument10 As Variant, _
    Optional ByVal argument11 As Variant, _
    Optional ByVal argument12 As Variant, _
    Optional ByVal argument13 As Variant, _
    Optional ByVal argument14 As Variant, _
    Optional ByVal argument15 As Variant, _
    Optional ByVal argument16 As Variant) _
As Variant

    '@Description: This function executes JavaScript code using Microsoft's JScript scripting language. It takes 3 arguments, the JavaScript code that will be executed, the name of the JavaScript function that will be executed, and up to 16 optional arguments to be used in the JavaScript function that is called. One thing to note is that ES5 syntax should be used in the JavaScript code, as ES6 features are unlikely to be supported.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: jsFuncCode is a string of the JavaScript source code that will be executed
    '@Param: jsFuncName is the name of the JavaScript function that will be called
    '@Param: argument1 - argument16 are optional arguments used in the JScript function call
    '@Returns: Returns the result of the JavaScript function that is called
    '@Example: =JavaScript("function helloFunc(){return 'Hello World!'}", "helloFunc") -> "Hello World!"
    '@Example: =JavaScript("function addTwo(a, b){return a + b}","addTwo",12,24) -> 36

    Dim ScriptContoller As Object
    Set ScriptContoller = CreateObject("ScriptControl")
    
    ScriptContoller.Language = "JScript"
    ScriptContoller.addCode jsFuncCode

    JavaScript = ScriptContoller.Run(jsFuncName, _
        argument1, argument2, argument3, argument4, _
        argument5, argument6, argument7, argument8, _
        argument9, argument10, argument11, argument12, _
        argument13, argument14, argument15, argument16)

End Function

Public Function HtmlEscape( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and escapes the HTML characters in it. For example, the character ">" will be escaped into "%gt;"
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will have its characters HTML escaped
    '@Returns: Returns an HTML escaped string
    '@Example: =HtmlEscape("<p>Hello World</p>") -> "&lt;p&gt;Hello World&lt;/p&gt;"

    string1 = Replace(string1, "&", "&amp;")
    string1 = Replace(string1, Chr(34), "&quot;")
    string1 = Replace(string1, "'", "&apos;")
    string1 = Replace(string1, "<", "&lt;")
    string1 = Replace(string1, ">", "&gt;")
    
    HtmlEscape = string1

End Function


Public Function HtmlUnescape( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and unescapes the HTML characters in it. For example, the character "%gt;" will be escaped into ">"
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will have its characters HTML unescaped
    '@Returns: Returns an HTML unescaped string
    '@Example: =HtmlUnescape("&lt;p&gt;Hello World&lt;/p&gt;") -> "<p>Hello World</p>"

    string1 = Replace(string1, "&amp;", "&")
    string1 = Replace(string1, "&quot;", Chr(34))
    string1 = Replace(string1, "&apos;", "'")
    string1 = Replace(string1, "&lt;", "<")
    string1 = Replace(string1, "&gt;", ">")

    HtmlUnescape = string1

End Function


Private Sub CallTextToSpeech(combinedString)

    '@Description: This subroutine simply calls the text-to-speech API
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: combinedString is the string that will be spoken

    Application.Speech.Speak combinedString, True

End Sub


Public Function SpeakText( _
    ParamArray textArray() As Variant) _
As String

    '@Description: This function takes the range of the cell that this function resides, and then an array of text, and when this function is recalculated manually by the user (for example when pressing the F2 key while on the cell) this function will use Microsoft's text-to-speech to speak out the text through the speakers or microphone.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: textArray() is an array of ranges, strings, or number that will be displayed
    '@Note: Note that text-to-speech is only available on Microsoft Excel. This function will still return the combined string from the text array, but will only result in speech through the speakers in Microsoft Excel
    '@Returns: Returns all the strings in the text array combined as well as displays all the text in the text array
    '@Example: =SpeakText("Hello", "World") -> "Hello World" and the text will be spoken through the speaker

    Dim combinedString As String
    Dim individualTextItem As Variant
    
    For Each individualTextItem In textArray
        combinedString = combinedString & individualTextItem & " "
    Next
    
    If Application.Name = "Microsoft Excel" Then
        CallTextToSpeech combinedString
    End If

    SpeakText = Trim(combinedString)

End Function


Public Function Dec2Hex( _
    ByVal number As Long, _
    Optional ByVal zeroFillAmount As Integer) _
As String

    '@Description: This function takes an integer and converts it to a hex string, with the option to specify the number of leading zeros for the hex string returned
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: number is the integer that will be converted to a hex string
    '@Returns: Returns the number rounded down to the nearest integer
    '@Example: =Dec2Hex(5) -> "5"
    '@Example: =Dec2Hex(5, 2) -> "05"
    '@Example: =Dec2Hex(255, 2) -> "FF"
    '@Example: =Dec2Hex(255, 8) -> "000000FF"

    Dim i As Integer
    Dim hexString As String
    
    hexString = Hex(number)
    
    If zeroFillAmount > 0 Then
        While Len(hexString) < zeroFillAmount
            hexString = "0" & hexString
        Wend
    End If
    
    Dec2Hex = hexString

End Function


Public Function BigDec2Hex( _
    ByVal number As Variant, _
    Optional ByVal zeroFillAmount As Integer) _
As String

    '@Description: This function is an implementation of Dec2Hex that allows big integers up to 14-byte to be used
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: number is the integer that will be converted to a hex string
    '@Returns: Returns the number rounded down to the nearest integer
    '@Example: =Dec2Hex(255, 8) -> "000000FF"
    '@Example: =Dec2Hex(3000000000, 16) -> Error; As Dec2Hex does not support integers this large
    '@Example: =BigDec2Hex(3000000000, 16) -> "00000000B2D05E00"

    Dim i As Integer
    Dim hexString As String
    
    hexString = BigHex(number)
    
    If zeroFillAmount > 0 Then
        While Len(hexString) < zeroFillAmount
            hexString = "0" & hexString
        Wend
    End If
    
    BigDec2Hex = hexString

End Function


Public Function BigHex( _
    ByVal number As Variant) _
As String

    '@Description: This function is an implementation of the Hex() function that allows for 14-byte integers to be used
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: number is the number that will be converted to hex
    '@Returns: Returns a string of the number converted to hex
    '@Example: =BigHex(255) -> "FF"
    '@Example: =Hex(3000000000) -> Error; As hex does not support big integers
    '@Example: =BigHex(3000000000) -> "B2D05E00"

    Dim integerString As String
    Dim decimalString As String
    Dim hexString As String

    While number > 0
        number = number / 16
        If InStr(1, CStr(number), ".") > 0 Then
            integerString = Split(CStr(number), ".")(0)
            decimalString = Split(CStr(number), ".")(1)
        Else
            integerString = CStr(number)
            decimalString = "0"
        End If
        
        Select Case decimalString
            Case "0"
                hexString = "0" & hexString
            Case "0625"
                hexString = "1" & hexString
            Case "125"
                hexString = "2" & hexString
            Case "1875"
                hexString = "3" & hexString
            Case "25"
                hexString = "4" & hexString
            Case "3125"
                hexString = "5" & hexString
            Case "375"
                hexString = "6" & hexString
            Case "4375"
                hexString = "7" & hexString
            Case "5"
                hexString = "8" & hexString
            Case "5625"
                hexString = "9" & hexString
            Case "625"
                hexString = "A" & hexString
            Case "6875"
                hexString = "B" & hexString
            Case "75"
                hexString = "C" & hexString
            Case "8125"
                hexString = "D" & hexString
            Case "875"
                hexString = "E" & hexString
            Case "9375"
                hexString = "F" & hexString
        End Select
        
        number = Fix(number)
    Wend

    BigHex = hexString

End Function

Public Function Hex2Dec( _
    ByVal hexNumber As String) _
As Long

    '@Description: This function takes a hex number as a string and converts it to a decimal long
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: hexNumber is the hex number that will be converted to a long
    '@Returns: Returns a decimal base number converted from the hex number
    '@Example: =Hex2Dec("FF") -> 255
    '@Example: =Hex2Dec("FFFF") -> 65535

    Hex2Dec = CLng("&H" & hexNumber)

End Function


Public Function Len2( _
    ByVal val As Variant) _
As Integer

    '@Description: This function is an extension on the Len() function by returning the length of strings, arrays, numbers, and many other objects in Excel, Word, PowerPoint, and Access, including Objects such as Dictionaries. Internally, any Object that implements a .Count property will have a length returned by this function. Also, any number used within this function will be converted to a string and then its length returned.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: val is the value you want the length from
    '@Returns: Returns an integer of the length of the value specified
    '@Example: =Len2("Hello") -> 5; As the string is 5 characters long
    '@Example: =Len2(arr) -> 3; Where arr is an array with {1, 2, 3} in it, and the array has 3 values in it
    '@Example: =Len2("100") -> 3; As the string is 3 characters long
    '@Example: =Len2(100) -> 3; As the integer is 3 characters long when converted to a string
    '@Example: =Len2(Range("A1:A3")) -> 3; As the Excel Range has 3
    '@Example: =Len2(col) -> 5; Where col is a Collection with 5 items in it
    '@Example: =Len2(dict) -> 2; Where dict is a Dictionary with 2 key/value pairs in it
    '@Example: =Len2(Application.Documents) -> 3; Where we currently have 3 documents open
    '@Example: =Len2(Application.ActivePresentation.Slides) -> 10; Where the active PowerPoint Presentation has 10 slides

    If IsArray(val) And Right(TypeName(val), 2) = "()" Then
        Len2 = UBound(val) - LBound(val) + 1
    ElseIf TypeName(val) = "String" Then
        Len2 = Len(val)
    ElseIf IsNumeric(val) Then
        Len2 = Len(CStr(val))
    Else
        Len2 = val.Count
    End If

End Function


'@Module: This module contains a set of functions for validating some commonly used string, such as validators for email addresses and phone numbers.



Public Function IsEmail( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid email address.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Improve regex robustness
    '@Param: string1 is the string we are checking if its a valid email
    '@Returns: Returns TRUE if the string is a valid email, and FALSE if its invalid
    '@Example: =IsEmail("JohnDoe@testmail.com") -> TRUE
    '@Example: =IsEmail("JohnDoe@test/mail.com") -> FALSE
    '@Example: =IsEmail("not_an_email_address") -> FALSE

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
        
    With Regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "^[a-zA-Z0-9_.]*?[@][a-zA-Z0-9.]*?[.][a-zA-Z]{2,15}$"
    End With

    IsEmail = Regex.Test(string1)

End Function


Public Function IsPhone( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a phone number is valid.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Improve regex robustness
    '@Todo: Add a second argument that lets the user add a country code and uses a different regex for phone number formats for that country. Also make the regx more robust so it can include more common formats.
    '@Param: string1 is the string we are checking if its a valid phone number
    '@Returns: Returns TRUE if the string is a valid phone number, and FALSE if its invalid
    '@Example: =IsPhone("123 456 7890") -> TRUE
    '@Example: =IsPhone("1234567890") -> TRUE
    '@Example: =IsPhone("1-234-567-890") -> FALSE; Not enough digits
    '@Example: =IsPhone("1-234-567-8905") -> TRUE
    '@Example: =IsPhone("+1-234-567-890") -> FALSE; Not enough digits
    '@Example: =IsPhone("+1-234-567-8905") -> TRUE
    '@Example: =IsPhone("+1-(234)-567-8905") -> TRUE
    '@Example: =IsPhone("+1 (234) 567 8905") -> TRUE
    '@Example: =IsPhone("1(234)5678905") -> TRUE
    '@Example: =IsPhone("123-456-789") -> FALSE; Not enough digits
    '@Example: =IsPhone("Hello World") -> FALSE; Not a phone number

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
        
    With Regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "^\s*[+]{0,1}[0-9]{0,1}[\s-]{0,1}\({0,1}([0-9]{3})\){0,1}[\s-]{0,1}([0-9]{3})[\s-]{0,1}([0-9]{4})$"
    End With

    IsPhone = Regex.Test(string1)

End Function


Public Function IsCreditCard( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid credit card from one of the major card issuing companies.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string we are checking if its a valid credit card number
    '@Returns: Returns TRUE if the string is a valid credit card number, and FALSE if its invalid. Currently supports these cards: Visa, MasterCard, Discover, Amex, Diners, JCB
    '@Example: =IsCreditCard("5111567856785678") -> TRUE; This is a valid Mastercard number
    '@Example: =IsCreditCard("511156785678567") -> FALSE; Not enough digits
    '@Example: =IsCreditCard("9999999999999999") -> FALSE; Enough digits, but not a valid card number
    '@Example: =IsCreditCard("Hello World") -> FALSE

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
        
    Dim regexPattern As String
    
    ' Regex for Amex
    regexPattern = regexPattern & "(3[47][0-9]{13})|"
    
    ' Regex for Diners
    regexPattern = regexPattern & "(3(0[0-5]|[68][0-9])?[0-9]{11})|"
    
    ' Regex for Discover
    regexPattern = regexPattern & "(6(011|5[0-9]{2})[0-9]{12})|"
    
    ' Regex for JCB
    regexPattern = regexPattern & "((2131|1800|35[0-9]{3})[0-9]{11})|"
    
    ' Regex for MasterCard
    regexPattern = regexPattern & "(5[1-5][0-9]{14})|"
    
    ' Regex for Visa
    regexPattern = regexPattern & "(4[0-9]{12}([0-9]{3})?)"
    
    With Regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = regexPattern
    End With

    IsCreditCard = Regex.Test(string1)

End Function


Public Function IsUrl( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid URL address.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Improve regex robustness
    '@Param: string1 is the string we are checking if its a valid URL
    '@Returns: Returns TRUE if the string is a valid URL, and FALSE if its invalid
    '@Example: =IsUrl("https://www.wikipedia.org/") -> TRUE
    '@Example: =IsUrl("http://www.wikipedia.org/") -> TRUE
    '@Example: =IsUrl("hello_world") -> FALSE

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
        
    With Regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "http(s){0,1}://www.[a-zA-Z0-9_.]*?[.][a-zA-Z]{2,15}"
    End With

    IsUrl = Regex.Test(string1)

End Function


Public Function IsIPFour( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid IPv4 address.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Improve regex robustness
    '@Param: string1 is the string we are checking if its a valid IPv4 address
    '@Returns: Returns TRUE if the string is a valid IPv4, and FALSE if its invalid
    '@Example: =IsIPFour("0.0.0.0") -> TRUE
    '@Example: =IsIPFour("100.100.100.100") -> TRUE
    '@Example: =IsIPFour("255.255.255.255") -> TRUE
    '@Example: =IsIPFour("255.255.255.256") -> FALSE; as the final 256 makes the address outside of the bounds of IPv4
    '@Example: =IsIPFour("0.0.0") -> FALSE; as the fourth octet is missing

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
        
    With Regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "^((2[0-4]\d|25[0-5]|1\d\d|\d{1,2})[.]){3}(2[0-4]\d|25[0-5]|1\d\d|\d{1,2})$"
    End With

    IsIPFour = Regex.Test(string1)

End Function


Public Function IsMacAddress( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid 48-bit Mac Address.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string we are checking if its a valid 48-bit Mac Address
    '@Returns: Returns TRUE if the string is a valid 48-bit Mac Address, and FALSE if its invalid
    '@Example: =IsMacAddress("00:25:96:12:34:56") -> TRUE
    '@Example: =IsMacAddress("FF:FF:FF:FF:FF:FF") -> TRUE
    '@Example: =IsMacAddress("00-25-96-12-34-56") -> TRUE
    '@Example: =IsMacAddress("123.789.abc.DEF") -> TRUE
    '@Example: =IsMacAddress("Not A Mac Address") -> FALSE
    '@Example: =IsMacAddress("FF:FF:FF:FF:FF:FH") -> FALSE; the H at the end is not a valid Hex number

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
        
    With Regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = "^(([a-fA-F0-9]{2}([:]|[-])){5}[a-fA-F0-9]{2}|([a-fA-F0-9]{3}[.]){3}[a-fA-F0-9]{3})$"
    End With

    IsMacAddress = Regex.Test(string1)

End Function


Public Function CreditCardName( _
    ByVal string1 As String) _
As String

    '@Description: This function checks if a string is a valid credit card from one of the major card issuing companies, and then returns the name of the credit card name. This function assumes no spaces or hyphens (if you have card numbers with spaces or hyphens you can remove these using =SUBSTITUTE("-", "") function.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the credit card string
    '@Returns: Returns the name of the credit card. Currently supports these cards: Visa, MasterCard, Discover, Amex, Diners, JCB
    '@Example: =CreditCardName("5111567856785678") -> "MasterCard"; This is a valid Mastercard number
    '@Example: =CreditCardName("not_a_card_number") -> #VALUE!

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
    
    Regex.Global = True
    Regex.IgnoreCase = True
    Regex.MultiLine = True

    ' Regex for Amex
    Regex.Pattern = "(3[47][0-9]{13})"
    If Regex.Test(string1) Then
        CreditCardName = "Amex"
        Exit Function
    End If
    
    ' Regex for Diners
    Regex.Pattern = "(3(0[0-5]|[68][0-9])?[0-9]{11})"
    If Regex.Test(string1) Then
        CreditCardName = "Diners"
        Exit Function
    End If
    
    ' Regex for Discover
    Regex.Pattern = "(6(011|5[0-9]{2})[0-9]{12})"
    If Regex.Test(string1) Then
        CreditCardName = "Discover"
        Exit Function
    End If
    
    ' Regex for JCB
    Regex.Pattern = "((2131|1800|35[0-9]{3})[0-9]{11})"
    If Regex.Test(string1) Then
        CreditCardName = "JCB"
        Exit Function
    End If
    
    ' Regex for MasterCard
    Regex.Pattern = "(5[1-5][0-9]{14})"
    If Regex.Test(string1) Then
        CreditCardName = "MasterCard"
        Exit Function
    End If
    
    ' Regex for Visa
    Regex.Pattern = "(4[0-9]{12}([0-9]{3})?)"
    If Regex.Test(string1) Then
        CreditCardName = "Visa"
        Exit Function
    End If
    
    CreditCardName = "#NotAValidCreditCardNumber!"

End Function


Public Function FormatCreditCard( _
    ByVal string1 As String) _
As String

    '@Description: This function checks if a string is a valid credit card, and if it is formats it in a more readable way. The format used is XXXX-XXXX-XXXX-XXXX.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is credit card number
    '@Returns: Returns a string formatted as a more readable credit card number
    '@Example: =FormatCreditCard("5111567856785678") -> "5111-5678-5678-5678"

    If IsCreditCard(string1) Then
        FormatCreditCard = Left(string1, 4) & "-" & Mid(string1, 5, 4) & "-" & Mid(string1, 9, 4) & "-" & Mid(string1, 13)
    Else
        FormatCreditCard = "#NotAValidCreditCardNumber!"
    End If

End Function
