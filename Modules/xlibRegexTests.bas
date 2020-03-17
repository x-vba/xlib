Attribute VB_Name = "xlibRegexTests"
Option Explicit

Public Function AllXlibRegexTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not RegexSearchTest() Then
        Debug.Print "Failed: RegexSearchTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RegexSearchTest"
    End If
    
    If Not RegexTestTest() Then
        Debug.Print "Failed: RegexTestTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RegexTestTest"
    End If
    
    If Not RegexReplaceTest() Then
        Debug.Print "Failed: RegexReplaceTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RegexReplaceTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibRegexTests = TestStatus
    
End Function



Private Function RegexSearchTest() As Boolean

    '@Example: =RegexSearch("Hello World","[a-z]{2}\s[W]") -> "lo W";

    RegexSearchTest = True

    RegexSearchTest = RegexSearchTest And RegexSearch("Hello World", "[a-z]{2}\s[W]") = "lo W"

End Function


Private Function RegexTestTest() As Boolean

    '@Example: =RegexTest("Hello World","[a-z]{2}\s[W]") -> TRUE;

    RegexTestTest = True

    RegexTestTest = RegexTestTest And RegexTest("Hello World", "[a-z]{2}\s[W]") = True

End Function


Private Function RegexReplaceTest() As Boolean

    '@Example: =RegexReplace("Hello World","[W][a-z]{4}", "VBA") -> "Hello VBA"

    RegexReplaceTest = True

    RegexReplaceTest = RegexReplaceTest And RegexReplace("Hello World", "[W][a-z]{4}", "VBA") = "Hello VBA"

End Function


