Attribute VB_Name = "xlibArray"
'@Module: This module contains a set of functions for manipulating and working with arrays.

Option Explicit


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


