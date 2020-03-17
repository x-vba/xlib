Attribute VB_Name = "xlibMathTests"
Option Explicit

Public Function AllXlibMathTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not CeilTest() Then
        Debug.Print "Failed: CeilTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CeilTest"
    End If
    
    If Not FloorTest() Then
        Debug.Print "Failed: FloorTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FloorTest"
    End If
    
    If Not InterpolateNumberTest() Then
        Debug.Print "Failed: InterpolateNumberTest"
        TestStatus = False
    Else
        Debug.Print "Passed: InterpolateNumberTest"
    End If
    
    If Not InterpolatePercentTest() Then
        Debug.Print "Failed: InterpolatePercentTest"
        TestStatus = False
    Else
        Debug.Print "Passed: InterpolatePercentTest"
    End If
    
    If Not MaxTest() Then
        Debug.Print "Failed: MaxTest"
        TestStatus = False
    Else
        Debug.Print "Passed: MaxTest"
    End If
    
    If Not MinTest() Then
        Debug.Print "Failed: MinTest"
        TestStatus = False
    Else
        Debug.Print "Passed: MinTest"
    End If
    
    If Not ModFloatTest() Then
        Debug.Print "Failed: ModFloatTest"
        TestStatus = False
    Else
        Debug.Print "Passed: ModFloatTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibMathTests = TestStatus
    
End Function



Private Function CeilTest() As Boolean

    '@Example: =Ceil(1.5) -> 2
    '@Example: =Ceil(1.0001) -> 2
    '@Example: =Ceil(1.0) -> 1
    '@Example: =Ceil(1) -> 1

    CeilTest = True

    CeilTest = CeilTest And Ceil(1.5) = 2
    CeilTest = CeilTest And Ceil(1.0001) = 2
    CeilTest = CeilTest And Ceil(1) = 1

End Function


Private Function FloorTest() As Boolean

    '@Example: =Floor(1.9) -> 1
    '@Example: =Floor(1.1) -> 1
    '@Example: =Floor(1.0) -> 1
    '@Example: =Floor(1) -> 1

    FloorTest = True

    FloorTest = FloorTest And Floor(1.9) = 1
    FloorTest = FloorTest And Floor(1.1) = 1
    FloorTest = FloorTest And Floor(1) = 1

End Function


Private Function InterpolateNumberTest() As Boolean

    '@Example: =InterpolateNumber(10, 20, 0.5) -> 15; Where 0.5 would be 50% between 10 and 20
    '@Example: =InterpolateNumber(16, 124, 0.64) -> 85.12; Where 0.64 would be 64% between 16 and 124

    InterpolateNumberTest = True

    InterpolateNumberTest = InterpolateNumberTest And InterpolateNumber(10, 20, 0.5) = 15
    InterpolateNumberTest = InterpolateNumberTest And Round(InterpolateNumber(16, 124, 0.64), 2) = 85.12

End Function


Private Function InterpolatePercentTest() As Boolean

    '@Example: =InterpolatePercent(10, 18, 12) -> 0.25; As 12 is 25% of the way from 10 to 18
    '@Example: =InterpolatePercent(10, 20, 15) -> 0.5; As 15 is 50% of the way from 10 to 20

    InterpolatePercentTest = True

    InterpolatePercentTest = InterpolatePercentTest And InterpolatePercent(10, 18, 12) = 0.25
    InterpolatePercentTest = InterpolatePercentTest And InterpolatePercent(10, 20, 15) = 0.5

End Function


Private Function MaxTest() As Boolean

    '@Example: =Max(1, 2, 3) -> 3
    '@Example: =Max(4.4, 5, "6") -> 6
    '@Example: =Max(x) -> 3; Where x is an array with these values [1, 2.2, "3"]
    '@Example: =Max(x, y, 10) -> 15; Where x = [1, 2.2, "3"] and y = [5, 15, -100]

    MaxTest = True

    MaxTest = MaxTest And Max(1, 2, 3) = 3
    MaxTest = MaxTest And Max(4.4, 5, "6") = 6
    MaxTest = MaxTest And Max(Array(1, 2.2, "3")) = 3
    MaxTest = MaxTest And Max(Array(1, 2.2, "3"), Array(5, 15, -100), 10) = 15

End Function


Private Function MinTest() As Boolean

    '@Example: =Min(1, 2, 3) -> 1
    '@Example: =Min(4.4, 5, "6") -> 4.4
    '@Example: =Min(-1, -2, -3) -> -3
    '@Example: =Min(x) -> 1; Where x is an array with these values [1, 2.2, "3"]
    '@Example: =Min(x, y, 10) -> -100; Where x = [1, 2.2, "3"] and y = [5, 15, -100]

    MinTest = True

    MinTest = MinTest And Min(1, 2, 3) = 1
    MinTest = MinTest And Min(4.4, 5, "6") = 4.4
    MinTest = MinTest And Min(-1, -2, -3) = -3
    MinTest = MinTest And Min(Array(1, 2.2, "3")) = 1
    MinTest = MinTest And Min(Array(1, 2.2, "3"), Array(5, 15, -100), 10) = -100

End Function


Private Function ModFloatTest() As Boolean

    '@Example: =ModFloat(3.55, 2) -> 1.55

    ModFloatTest = True

    ModFloatTest = ModFloatTest And Round(ModFloat(3.55, 2), 2) = 1.55
    
End Function

