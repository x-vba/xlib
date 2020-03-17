Attribute VB_Name = "xlibUtilitiesTests"
Option Explicit

Public Function AllXlibUtilitiesTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not JsonifyTest() Then
        Debug.Print "Failed: JsonifyTest"
        TestStatus = False
    Else
        Debug.Print "Passed: JsonifyTest"
    End If
    
    If Not UuidFourTest() Then
        Debug.Print "Failed: UuidFourTest"
        TestStatus = False
    Else
        Debug.Print "Passed: UuidFourTest"
    End If
    
    If Not HideTextTest() Then
        Debug.Print "Failed: HideTextTest"
        TestStatus = False
    Else
        Debug.Print "Passed: HideTextTest"
    End If
    
    If Not JavaScriptTest() Then
        Debug.Print "Failed: JavaScriptTest"
        TestStatus = False
    Else
        Debug.Print "Passed: JavaScriptTest"
    End If
    
    If Not HtmlEscapeTest() Then
        Debug.Print "Failed: HtmlEscapeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: HtmlEscapeTest"
    End If
    
    If Not HtmlUnescapeTest() Then
        Debug.Print "Failed: HtmlUnescapeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: HtmlUnescapeTest"
    End If
    
    If Not SpeakTextTest() Then
        Debug.Print "Failed: SpeakTextTest"
        TestStatus = False
    Else
        Debug.Print "Passed: SpeakTextTest"
    End If
    
    If Not Dec2HexTest() Then
        Debug.Print "Failed: Dec2HexTest"
        TestStatus = False
    Else
        Debug.Print "Passed: Dec2HexTest"
    End If
    
    If Not BigDec2HexTest() Then
        Debug.Print "Failed: BigDec2HexTest"
        TestStatus = False
    Else
        Debug.Print "Passed: BigDec2HexTest"
    End If
    
    If Not BigHexTest() Then
        Debug.Print "Failed: BigHexTest"
        TestStatus = False
    Else
        Debug.Print "Passed: BigHexTest"
    End If
    
    If Not Hex2DecTest() Then
        Debug.Print "Failed: Hex2DecTest"
        TestStatus = False
    Else
        Debug.Print "Passed: Hex2DecTest"
    End If
    
    If Not Len2Test() Then
        Debug.Print "Failed: Len2Test"
        TestStatus = False
    Else
        Debug.Print "Passed: Len2Test"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibUtilitiesTests = TestStatus
    
End Function



Private Function JsonifyTest() As Boolean

    '@Example: =Jsonify(0, "Hello", "World", "1", "2", 3, 4.5) -> "["Hello","World",1,2,3,4.5]"
    '@Example: =Jsonify(0, {"Hello", "World", "1", "2", 3, 4.5}, 10) -> "["Hello","World",1,2,3,4.5]"

    JsonifyTest = True

    JsonifyTest = JsonifyTest And Jsonify(0, "Hello", "World", "1", "2", 3, 4.5) = "[" & Chr(34) & "Hello" & Chr(34) & "," & Chr(34) & "World" & Chr(34) & ",1,2,3,4.5]"
    JsonifyTest = JsonifyTest And Jsonify(0, Array("Hello", "World", "1", "2", 3, 4.5)) = "[" & Chr(34) & "Hello" & Chr(34) & "," & Chr(34) & "World" & Chr(34) & ",1,2,3,4.5]"

End Function


Private Function UuidFourTest() As Boolean

    '@Example: =UuidFour() -> "3B4BDC26-E76E-4D6C-9E05-7AE7D2D68304"
    '@Example: =UuidFour() -> "D5761256-8385-4FDA-AD56-6AEF0AD6B9A5"
    '@Example: =UuidFour() -> "CDCAE2F5-B52F-4C90-A38A-42BD58BCED4B"

    Dim uuidGroups As Variant
    uuidGroups = Split(UuidFour(), "-")
    
    If Len(uuidGroups(0)) = 8 _
    And Len(uuidGroups(1)) = 4 _
    And Len(uuidGroups(2)) = 4 _
    And Len(uuidGroups(3)) = 4 _
    And Len(uuidGroups(4)) = 12 Then

        UuidFourTest = True

    End If

End Function


Private Function HideTextTest() As Boolean

    '@Example: =HideText("Hello World", FALSE) -> "Hello World"
    '@Example: =HideText("Hello World", TRUE) -> "********"
    '@Example: =HideText("Hello World", TRUE, "[HideText Text]") -> "[HideText Text]"
    '@Example: =HideText("Hello World", UserName()="Anthony") -> "********"

    HideTextTest = True

    HideTextTest = HideTextTest And HideText("Hello World", False) = "Hello World"
    HideTextTest = HideTextTest And HideText("Hello World", True) = "********"
    HideTextTest = HideTextTest And HideText("Hello World", True, "[Hidden Text]") = "[Hidden Text]"

End Function


Private Function JavaScriptTest() As Boolean

    '@Example: =JavaScript("function helloFunc(){return 'Hello World!'}", "helloFunc") -> "Hello World!"
    '@Example: =JavaScript("function addTwo(a, b){return a + b}","addTwo",12,24) -> 36

    JavaScriptTest = True

    JavaScriptTest = JavaScriptTest And JavaScript("function helloFunc(){return 'Hello World!'}", "helloFunc") = "Hello World!"
    JavaScriptTest = JavaScriptTest And JavaScript("function addTwo(a, b){return a + b}", "addTwo", 12, 24) = 36

End Function


Private Function HtmlEscapeTest() As Boolean

    '@Example: =HtmlEscape("<p>Hello World</p>") -> "&lt;p&gt;Hello World&lt;/p&gt;"

    HtmlEscapeTest = True

    HtmlEscapeTest = HtmlEscapeTest And HtmlEscape("<p>Hello World</p>") = "&lt;p&gt;Hello World&lt;/p&gt;"

End Function


Private Function HtmlUnescapeTest() As Boolean
    
    '@Example: =HtmlUnescape("&lt;p&gt;Hello World&lt;/p&gt;") -> "<p>Hello World</p>"

    HtmlUnescapeTest = True

    HtmlUnescapeTest = HtmlUnescapeTest And HtmlUnescape("&lt;p&gt;Hello World&lt;/p&gt;") = "<p>Hello World</p>"

End Function


Private Function SpeakTextTest() As Boolean

    '@Example: =SpeakText("Hello", "World") -> "Hello World" and the text will be spoken through the speaker

    SpeakTextTest = True

    SpeakTextTest = SpeakTextTest And SpeakText("Hello", "World") = "Hello World"

End Function


Private Function Dec2HexTest() As Boolean

    '@Example: =Dec2Hex(5) -> "5"
    '@Example: =Dec2Hex(5, 2) -> "05"
    '@Example: =Dec2Hex(255, 2) -> "FF"
    '@Example: =Dec2Hex(255, 8) -> "000000FF"

    Dec2HexTest = True

    Dec2HexTest = Dec2HexTest And Dec2Hex(5) = "5"
    Dec2HexTest = Dec2HexTest And Dec2Hex(5, 2) = "05"
    Dec2HexTest = Dec2HexTest And Dec2Hex(255, 2) = "FF"
    Dec2HexTest = Dec2HexTest And Dec2Hex(255, 8) = "000000FF"

End Function


Private Function BigDec2HexTest() As Boolean

    '@Example: =Dec2Hex(255, 8) -> "000000FF"
    '@Example: =Dec2Hex(3000000000, 16) -> Error; As Dec2Hex does not support integers this large
    '@Example: =BigDec2Hex(3000000000, 16) -> "00000000B2D05E00"

    BigDec2HexTest = True

    BigDec2HexTest = BigDec2HexTest And BigDec2Hex(3000000000#, 16) = "00000000B2D05E00"

End Function


Private Function BigHexTest() As Boolean

    '@Example: =BigHex(255) -> "FF"
    '@Example: =Hex(3000000000) -> Error; As hex does not support big integers
    '@Example: =BigHex(3000000000) -> "B2D05E00"

    BigHexTest = True

    BigHexTest = BigHexTest And BigHex(3000000000#) = "B2D05E00"

End Function


Private Function Hex2DecTest() As Boolean

    '@Example: =Hex2Dec("FF") -> 255
    '@Example: =Hex2Dec("FFFF") -> 65535

    Hex2DecTest = True

    Hex2DecTest = Hex2DecTest And Hex2Dec("FF") = 255
    Hex2DecTest = Hex2DecTest And Hex2Dec("FFFF") = 65535

End Function


Private Function Len2Test() As Boolean

    '@Example: =Len2("Hello") -> 5; As the string is 5 characters long
    '@Example: =Len2(arr) -> 3; Where arr is an array with {1, 2, 3} in it, and the array has 3 values in it
    '@Example: =Len2("100") -> 3; As the string is 3 characters long
    '@Example: =Len2(100) -> 3; As the integer is 3 characters long when converted to a string
    '@Example: =Len2(Range("A1:A3")) -> 3; As the Excel Range has 3
    '@Example: =Len2(col) -> 5; Where col is a Collection with 5 items in it
    '@Example: =Len2(dict) -> 2; Where dict is a Dictionary with 2 key/value pairs in it
    '@Example: =Len2(Application.Documents) -> 3; Where we currently have 3 documents open
    '@Example: =Len2(Application.ActivePresentation.Slides) -> 10; Where the active PowerPoint Presentation has 10 slides

    Len2Test = True

    Len2Test = Len2Test And Len2("Hello") = 5
    Len2Test = Len2Test And Len2(Array(1, 2, 3)) = 3
    Len2Test = Len2Test And Len2("100") = 3
    Len2Test = Len2Test And Len2(100) = 3
    
    Dim col As Collection
    Set col = New Collection
    col.Add 1
    col.Add 2
    col.Add 3
    col.Add 4
    col.Add 5
    Len2Test = Len2Test And Len2(col) = 5
    
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.Add "Hello", 1
    dict.Add "World", 2
    Len2Test = Len2Test And Len2(dict) = 2

End Function

