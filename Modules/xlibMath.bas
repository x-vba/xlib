Attribute VB_Name = "xlibMath"
'@Module: This module contains a set of basic mathematical functions where those functions don't already exist as base Excel functions.

Option Explicit


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
