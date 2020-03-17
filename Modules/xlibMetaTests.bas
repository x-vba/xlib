Attribute VB_Name = "xlibMetaTests"
Option Explicit

Public Function AllXlibMetaTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not XlibVersionTest() Then
        Debug.Print "Failed: XlibVersionTest"
        TestStatus = False
    Else
        Debug.Print "Passed: XlibVersionTest"
    End If
    
    If Not XlibCreditsTest() Then
        Debug.Print "Failed: XlibCreditsTest"
        TestStatus = False
    Else
        Debug.Print "Passed: XlibCreditsTest"
    End If
    
    If Not XlibDocumentationTest() Then
        Debug.Print "Failed: XlibDocumentationTest"
        TestStatus = False
    Else
        Debug.Print "Passed: XlibDocumentationTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibMetaTests = TestStatus
    
End Function



Private Function XlibVersionTest() As Boolean

    '@Example: =XlibVersion() -> "1.0.0"; Where the version of XPlus you are using is 1.0.0

    If IsNumeric(Split(XlibVersion(), ".")(0)) Then
        If IsNumeric(Split(XlibVersion(), ".")(1)) Then
            If IsNumeric(Split(XlibVersion(), ".")(2)) Then
                XlibVersionTest = True
            End If
        End If
    End If

End Function


Private Function XlibCreditsTest() As Boolean

    '@Example: =XlibCredits() -> "Copyright (c) 2020 Anthony Mancini. XPlus is Licensed under an MIT License."

    If InStr(1, XlibCredits(), "XLib") > 0 Then
        XlibCreditsTest = True
    End If

End Function


Private Function XlibDocumentationTest() As Boolean

    '@Example: =XlibDocumentation() -> "https://x-vba.com/xlib"

    If Left(XlibDocumentation(), 4) = "http" Then
        XlibDocumentationTest = True
    End If

End Function

