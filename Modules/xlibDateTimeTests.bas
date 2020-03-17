Attribute VB_Name = "xlibDateTimeTests"
Option Explicit

Public Function AllXlibDateTimeTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not QuarterTest() Then
        Debug.Print "Failed: QuarterTest"
        TestStatus = False
    Else
        Debug.Print "Passed: QuarterTest"
    End If
    
    If Not DaysOfMonthTest() Then
        Debug.Print "Failed: DaysOfMonthTest"
        TestStatus = False
    Else
        Debug.Print "Passed: DaysOfMonthTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibDateTimeTests = TestStatus
    
End Function



Private Function QuarterTest() As Boolean

    '@Example: =Quarter(4) -> 2
    '@Example: =Quarter("April") -> 2
    '@Example: =Quarter(12) -> 4
    '@Example: =Quarter("December") -> 4
    '@Example: To get today's Quarter: =Quarter()

    QuarterTest = True

    QuarterTest = QuarterTest And Quarter(4) = 2
    QuarterTest = QuarterTest And Quarter(12) = 4

End Function


Private Function DaysOfMonthTest() As Boolean

    '@Example: =DaysOfMonth() -> 31; Where the current month is January
    '@Example: =DaysOfMonth(1) -> 31
    '@Example: =DaysOfMonth("January") -> 31
    '@Example: =DaysOfMonth(2, 2019) -> 28
    '@Example: =DaysOfMonth(2, 2020) -> 29

    DaysOfMonthTest = True

    DaysOfMonthTest = DaysOfMonthTest And DaysOfMonth(1) = 31
    DaysOfMonthTest = DaysOfMonthTest And DaysOfMonth(2, 2019) = 28
    DaysOfMonthTest = DaysOfMonthTest And DaysOfMonth(2, 2020) = 29

End Function


