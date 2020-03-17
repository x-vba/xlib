Attribute VB_Name = "xlibArrayTests"
Option Explicit

Public Function AllXlibArrayTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not CountUniqueTest() Then
        Debug.Print "Failed: CountUniqueTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CountUniqueTest"
    End If
    
    If Not SortTest() Then
        Debug.Print "Failed: SortTest"
        TestStatus = False
    Else
        Debug.Print "Passed: SortTest"
    End If
    
    If Not ReverseTest() Then
        Debug.Print "Failed: ReverseTest"
        TestStatus = False
    Else
        Debug.Print "Passed: ReverseTest"
    End If
    
    If Not SumHighTest() Then
        Debug.Print "Failed: SumHighTest"
        TestStatus = False
    Else
        Debug.Print "Passed: SumHighTest"
    End If
    
    If Not SumLowTest() Then
        Debug.Print "Failed: SumLowTest"
        TestStatus = False
    Else
        Debug.Print "Passed: SumLowTest"
    End If
    
    If Not AverageHighTest() Then
        Debug.Print "Failed: AverageHighTest"
        TestStatus = False
    Else
        Debug.Print "Passed: AverageHighTest"
    End If
    
    If Not AverageLowTest() Then
        Debug.Print "Failed: AverageLowTest"
        TestStatus = False
    Else
        Debug.Print "Passed: AverageLowTest"
    End If
    
    If Not LargeTest() Then
        Debug.Print "Failed: LargeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: LargeTest"
    End If
    
    If Not SmallTest() Then
        Debug.Print "Failed: SmallTest"
        TestStatus = False
    Else
        Debug.Print "Passed: SmallTest"
    End If
    
    If Not IsInArrayTest() Then
        Debug.Print "Failed: IsInArrayTest"
        TestStatus = False
    Else
        Debug.Print "Passed: IsInArrayTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibArrayTests = TestStatus
    
End Function



Private Function CountUniqueTest() As Boolean
    
    '@Example: =CountUnique(1, 2, 2, 3) -> 3;
    '@Example: =CountUnique("a", "a", "a") -> 1;
    '@Example: =CountUnique(arr) -> 3; Where arr = [1, 2, 4, 4, 1]
    
    CountUniqueTest = True

    CountUniqueTest = CountUniqueTest And CountUnique(1, 2, 2, 3) = 3
    CountUniqueTest = CountUniqueTest And CountUnique("a", "a", "a") = 1
    CountUniqueTest = CountUniqueTest And CountUnique(Array(1, 2, 4, 4, 1)) = 3
    
End Function


Private Function SortTest() As Boolean

    '@Example: =Sort({1,3,2}) -> {1,2,3}
    '@Example: =Sort({1,3,2}, True) -> {3,2,1}

    SortTest = True

    SortTest = SortTest And Sort(Array(10, 20, 30))(0) = 10
    SortTest = SortTest And Sort(Array(10, 20, 30))(1) = 20
    SortTest = SortTest And Sort(Array(10, 20, 30))(2) = 30
    SortTest = SortTest And Sort(Array(10, 20, 30), True)(0) = 30
    SortTest = SortTest And Sort(Array(10, 20, 30), True)(1) = 20
    SortTest = SortTest And Sort(Array(10, 20, 30), True)(2) = 10
    
End Function


Private Function ReverseTest() As Boolean

    '@Example: =Reverse({1,2,3}) -> {3,2,1}

    ReverseTest = True

    ReverseTest = ReverseTest And Reverse(Array(10, 20, 30))(0) = 30
    ReverseTest = ReverseTest And Reverse(Array(10, 20, 30))(1) = 20
    ReverseTest = ReverseTest And Reverse(Array(10, 20, 30))(2) = 10

End Function


Private Function SumHighTest() As Boolean

    '@Example: =SumHigh({1,2,3,4}, 2) -> 7; as 3 and 4 will be summed
    '@Example: =SumHigh({1,2,3,4}, 3) -> 9; as 2, 3, and 4 will be summed

    SumHighTest = True

    SumHighTest = SumHighTest And SumHigh(Array(1, 2, 3, 4), 2) = 7
    SumHighTest = SumHighTest And SumHigh(Array(1, 2, 3, 4), 3) = 9

End Function


Private Function SumLowTest() As Boolean

    '@Example: =SumLow({1,2,3,4}, 2) -> 3; as 1 and 2 will be summed
    '@Example: =SumLow({1,2,3,4}, 3) -> 6; as 1, 2, and 3 will be summed

    SumLowTest = True

    SumLowTest = SumLowTest And SumLow(Array(1, 2, 3, 4), 2) = 3
    SumLowTest = SumLowTest And SumLow(Array(1, 2, 3, 4), 3) = 6

End Function


Private Function AverageHighTest() As Boolean

    '@Example: =AverageHigh({1,2,3,4}, 2) -> 3.5; as 3 and 4 will be averaged
    '@Example: =AverageHigh({1,2,3,4}, 3) -> 3; as 2, 3, and 4 will be averaged

    AverageHighTest = True

    AverageHighTest = AverageHighTest And AverageHigh(Array(1, 2, 3, 4), 2) = 3.5
    AverageHighTest = AverageHighTest And AverageHigh(Array(1, 2, 3, 4), 3) = 3

End Function


Private Function AverageLowTest() As Boolean

    '@Example: =AverageLow({1,2,3,4}, 2) -> 1.5; as 1 and 2 will be averaged
    '@Example: =AverageLow({1,2,3,4}, 3) -> 2; as 1, 2, and 3 will be averaged

    AverageLowTest = True

    AverageLowTest = AverageLowTest And AverageLow(Array(1, 2, 3, 4), 2) = 1.5
    AverageLowTest = AverageLowTest And AverageLow(Array(1, 2, 3, 4), 3) = 2

End Function


Private Function LargeTest() As Boolean

    '@Example: =Large({1,2,3,4}, 1) -> 4
    '@Example: =Large({1,2,3,4}, 2) -> 3

    LargeTest = True

    LargeTest = LargeTest And Large(Array(1, 2, 3, 4), 1) = 4
    LargeTest = LargeTest And Large(Array(1, 2, 3, 4), 2) = 3

End Function


Private Function SmallTest() As Boolean

    '@Example: =Small({1,2,3,4}, 1) -> 1
    '@Example: =Small({1,2,3,4}, 2) -> 2

    SmallTest = True

    SmallTest = SmallTest And Small(Array(1, 2, 3, 4), 1) = 1
    SmallTest = SmallTest And Small(Array(1, 2, 3, 4), 2) = 2

End Function


Private Function IsInArrayTest() As Boolean

    '@Example: =IsInArray("hello", {"one", 2, "hello"}) -> True
    '@Example: =IsInArray("hello", {1, "two", "three"}) -> False

    IsInArrayTest = True

    IsInArrayTest = IsInArrayTest And IsInArray("hello", Array("one", 2, "hello")) = True
    IsInArrayTest = IsInArrayTest And IsInArray("hello", Array(1, "two", "three")) = False

End Function


