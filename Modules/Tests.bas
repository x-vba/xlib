Attribute VB_Name = "Tests"
Public Sub XlibTests()

    Dim TestStatus As Boolean
    TestStatus = True

    TestStatus = TestStatus And AllXlibArrayTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibColorTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibDateTimeTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibEnvironmentTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibFileTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibMathTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibMetaTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibNetworkTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibRandomTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibRegexTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibStringManipulationTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibStringMetricsTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibUtilitiesTests
    Debug.Print ""
    
    TestStatus = TestStatus And AllXlibValidatorsTests
    Debug.Print ""
    
    Debug.Print "========================================"
    If TestStatus Then
        Debug.Print "Status of All Tests: Passed"
    Else
        Debug.Print "Status of All Tests: !!! FAILED !!!"
    End If
    Debug.Print "========================================"

End Sub
