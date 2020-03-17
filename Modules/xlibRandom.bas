Attribute VB_Name = "xlibRandom"
'@Module: This module contains a set of functions for generating and sampling random data.

Option Explicit


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

