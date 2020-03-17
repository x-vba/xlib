Attribute VB_Name = "xlibRandomTests"
Option Explicit

Public Function AllXlibRandomTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not RandBetweenTest() Then
        Debug.Print "Failed: RandBetweenTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RandBetweenTest"
    End If
    
    If Not BigRandBetweenTest() Then
        Debug.Print "Failed: BigRandBetweenTest"
        TestStatus = False
    Else
        Debug.Print "Passed: BigRandBetweenTest"
    End If
    
    If Not RandomSampleTest() Then
        Debug.Print "Failed: RandomSampleTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RandomSampleTest"
    End If
    
    If Not RandomRangeTest() Then
        Debug.Print "Failed: RandomRangeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RandomRangeTest"
    End If
    
    If Not RandBoolTest() Then
        Debug.Print "Failed: RandBoolTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RandBoolTest"
    End If
    
    If Not RandBetweensTest() Then
        Debug.Print "Failed: RandBetweensTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RandBetweensTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibRandomTests = TestStatus
    
End Function



Private Function RandBetweenTest() As Boolean

    '@Example: =RandBetween(1, 20) -> 5
    '@Example: =RandBetween(1, 20) -> 9
    '@Example: =RandBetween(1, 20) -> 13
    '@Example: =RandBetween(1, 20) -> 2
    '@Example: =RandBetween(1, 20) -> 20
    '@Example: =RandBetween(1, 20) -> 6
    
    Dim randomNumber As Integer
    randomNumber = RandBetween(1, 20)
    
    RandBetweenTest = (randomNumber >= 1 And randomNumber <= 20)

End Function


Private Function BigRandBetweenTest() As Boolean

    '@Example: =RandBetween(0, 3000000000) -> Error; as RandBetween only works with 4-byte and less integers
    '@Example: =BigRandBetween(0, 3000000000) -> 2116642535; as BigRandBetween supports up to 14-byte integers

    Dim randomNumber As Integer
    randomNumber = BigRandBetween(1, 20)
    
    BigRandBetweenTest = (randomNumber >= 1 And randomNumber <= 20)

End Function


Private Function RandomSampleTest() As Boolean

    '@Example: =RandomSample(A1:A5) -> "Hello"; where "Hello" is the value in cell A3, and where A3 was the chosen random cell
    '@Example: =RandomSample(A1:A5) -> "World"; where "World" is the value in cell A2, and where A2 was the chosen random cell

    Dim randomNumber As Integer
    randomNumber = RandomSample(Array(1, 2, 3))
    
    RandomSampleTest = (randomNumber = 1 Or randomNumber = 2 Or randomNumber = 3)

End Function


Private Function RandomRangeTest() As Boolean

    '@Example: =RandomRange(50, 100, 10) -> 60
    '@Example: =RandomRange(50, 100, 10) -> 50
    '@Example: =RandomRange(50, 100, 10) -> 90
    '@Example: =RandomRange(0, 10, 2) -> 8
    '@Example: =RandomRange(0, 10, 2) -> 0
    '@Example: =RandomRange(0, 10, 2) -> 4
    '@Example: =RandomRange(0, 10, 2) -> 10
    
    Dim randomNumber As Integer
    randomNumber = RandomRange(50, 60, 10)
    
    RandomRangeTest = (randomNumber = 50 Or randomNumber = 60)

End Function


Private Function RandBoolTest() As Boolean

    '@Example: =RANDBOOL() -> TRUE
    '@Example: =RANDBOOL() -> FALSE
    '@Example: =RANDBOOL() -> TRUE
    '@Example: =RANDBOOL() -> TRUE
    '@Example: =RANDBOOL() -> FALSE
    '@Example: =RANDBOOL() -> FALSE
    
    Dim randomBoolean As Boolean
    randomBoolean = RandBool()
    
    RandBoolTest = (randomBoolean = True Or randomBoolean = False)

End Function


Private Function RandBetweensTest() As Boolean

    '@Example: =RANDBETWEENS(1, 10, 5000, 5010) -> 6
    '@Example: =RANDBETWEENS(1, 10, 5000, 5010) -> 5002
    '@Example: =RANDBETWEENS(1, 10, 5000, 5010) -> 8
    '@Example: =RANDBETWEENS(1, 10, 5000, 5010) -> 3
    '@Example: =RANDBETWEENS(1, 10, 5000, 5010) -> 5010
    '@Example: =RANDBETWEENS(1, 10, 5000, 5010) -> 2
    '@Example: =RANDBETWEENS(5, 10, 15, 20, 25, 30, 35, 40) -> 32

    Dim randomNumber As Integer
    randomNumber = RandBetweens(1, 2, 51, 52)
    
    RandBetweensTest = (randomNumber = 1 Or randomNumber = 2 Or randomNumber = 51 Or randomNumber = 52)

End Function

