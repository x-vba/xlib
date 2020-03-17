Attribute VB_Name = "xlibStringMetricsTests"
Option Explicit

Public Function AllXlibStringMetricsTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not HammingTest() Then
        Debug.Print "Failed: HammingTest"
        TestStatus = False
    Else
        Debug.Print "Passed: HammingTest"
    End If
    
    If Not LevenshteinTest() Then
        Debug.Print "Failed: LevenshteinTest"
        TestStatus = False
    Else
        Debug.Print "Passed: LevenshteinTest"
    End If
    
    If Not DamerauTest() Then
        Debug.Print "Failed: DamerauTest"
        TestStatus = False
    Else
        Debug.Print "Passed: DamerauTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibStringMetricsTests = TestStatus
    
End Function



Private Function HammingTest() As Boolean

    '@Example: =Hamming("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
    '@Example: =Hamming("Cat", "Bag") -> 2; 2 changes are needed, changing the "B" to "C" and the "g" to "t"
    '@Example: =Hamming("Cat", "Dog") -> 3; Every single character needs to be substituted in this case

    HammingTest = True

    HammingTest = HammingTest And Hamming("Cat", "Bat") = 1
    HammingTest = HammingTest And Hamming("Cat", "Bag") = 2
    HammingTest = HammingTest And Hamming("Cat", "Dog") = 3
    
End Function


Private Function LevenshteinTest() As Boolean

    '@Example: =Levenshtein("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
    '@Example: =Levenshtein("Cat", "Ca") -> 1; Since only one Insertion needs to occur (adding a "t" at the end)
    '@Example: =Levenshtein("Cat", "Cta") -> 2; Since the "t" in "Cta" needs to be substituted into an "a", and the final character "a" needs to be substituted into a "t"

    LevenshteinTest = True

    LevenshteinTest = LevenshteinTest And Levenshtein("Cat", "Bat") = 1
    LevenshteinTest = LevenshteinTest And Levenshtein("Cat", "Ca") = 1
    LevenshteinTest = LevenshteinTest And Levenshtein("Cat", "Cta") = 2

End Function


Private Function DamerauTest() As Boolean

    '@Example: =Damerau("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
    '@Example: =Damerau("Cat", "Ca") -> 1; Since only one Insertion needs to occur (adding a "t" at the end)
    '@Example: =Damerau("Cat", "Cta") -> 1; Since the "t" and "a" can be transposed as they are adjacent to each other. Notice how LEVENSHTEIN("Cat","Cta")=2 but DAMERAU("Cat","Cta")=1

    DamerauTest = True

    DamerauTest = DamerauTest And Damerau("Cat", "Bat") = 1
    DamerauTest = DamerauTest And Damerau("Cat", "Ca") = 1
    DamerauTest = DamerauTest And Damerau("Cat", "Cta") = 1

End Function


