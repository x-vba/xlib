Attribute VB_Name = "xlibStringManipulationTests"
Option Explicit

Public Function AllXlibStringManipulationTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not CapitalizeTest() Then
        Debug.Print "Failed: CapitalizeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CapitalizeTest"
    End If
    
    If Not LeftFindTest() Then
        Debug.Print "Failed: LeftFindTest"
        TestStatus = False
    Else
        Debug.Print "Passed: LeftFindTest"
    End If
    
    If Not RightFindTest() Then
        Debug.Print "Failed: RightFindTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RightFindTest"
    End If
    
    If Not LeftSearchTest() Then
        Debug.Print "Failed: LeftSearchTest"
        TestStatus = False
    Else
        Debug.Print "Passed: LeftSearchTest"
    End If
    
    If Not RightSearchTest() Then
        Debug.Print "Failed: RightSearchTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RightSearchTest"
    End If
    
    If Not SubstrTest() Then
        Debug.Print "Failed: SubstrTest"
        TestStatus = False
    Else
        Debug.Print "Passed: SubstrTest"
    End If
    
    If Not SubstrFindTest() Then
        Debug.Print "Failed: SubstrFindTest"
        TestStatus = False
    Else
        Debug.Print "Passed: SubstrFindTest"
    End If
    
    If Not SubstrSearchTest() Then
        Debug.Print "Failed: SubstrSearchTest"
        TestStatus = False
    Else
        Debug.Print "Passed: SubstrSearchTest"
    End If
    
    If Not RepeatTest() Then
        Debug.Print "Failed: RepeatTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RepeatTest"
    End If
    
    If Not FormatterTest() Then
        Debug.Print "Failed: FormatterTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FormatterTest"
    End If
    
    If Not ZfillTest() Then
        Debug.Print "Failed: ZfillTest"
        TestStatus = False
    Else
        Debug.Print "Passed: ZfillTest"
    End If
    
    If Not SplitTextTest() Then
        Debug.Print "Failed: SplitTextTest"
        TestStatus = False
    Else
        Debug.Print "Passed: SplitTextTest"
    End If
    
    If Not CountWordsTest() Then
        Debug.Print "Failed: CountWordsTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CountWordsTest"
    End If
    
    If Not CamelCaseTest() Then
        Debug.Print "Failed: CamelCaseTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CamelCaseTest"
    End If
    
    If Not KebabCaseTest() Then
        Debug.Print "Failed: KebabCaseTest"
        TestStatus = False
    Else
        Debug.Print "Passed: KebabCaseTest"
    End If
    
    If Not RemoveCharactersTest() Then
        Debug.Print "Failed: RemoveCharactersTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RemoveCharactersTest"
    End If
    
    If Not CompanyCaseTest() Then
        Debug.Print "Failed: CompanyCaseTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CompanyCaseTest"
    End If
    
    If Not ReverseTextTest() Then
        Debug.Print "Failed: ReverseTextTest"
        TestStatus = False
    Else
        Debug.Print "Passed: ReverseTextTest"
    End If
    
    If Not ReverseWordsTest() Then
        Debug.Print "Failed: ReverseWordsTest"
        TestStatus = False
    Else
        Debug.Print "Passed: ReverseWordsTest"
    End If
    
    If Not IndentTextTest() Then
        Debug.Print "Failed: IndentTextTest"
        TestStatus = False
    Else
        Debug.Print "Passed: IndentTextTest"
    End If
    
    If Not DedentTextTest() Then
        Debug.Print "Failed: DedentTextTest"
        TestStatus = False
    Else
        Debug.Print "Passed: DedentTextTest"
    End If
    
    If Not ShortenTextTest() Then
        Debug.Print "Failed: ShortenTextTest"
        TestStatus = False
    Else
        Debug.Print "Passed: ShortenTextTest"
    End If
    
    If Not InSplitTest() Then
        Debug.Print "Failed: InSplitTest"
        TestStatus = False
    Else
        Debug.Print "Passed: InSplitTest"
    End If
    
    If Not EliteCaseTest() Then
        Debug.Print "Failed: EliteCaseTest"
        TestStatus = False
    Else
        Debug.Print "Passed: EliteCaseTest"
    End If
    
    If Not ScrambleCaseTest() Then
        Debug.Print "Failed: ScrambleCaseTest"
        TestStatus = False
    Else
        Debug.Print "Passed: ScrambleCaseTest"
    End If
    
    If Not LeftSplitTest() Then
        Debug.Print "Failed: LeftSplitTest"
        TestStatus = False
    Else
        Debug.Print "Passed: LeftSplitTest"
    End If
    
    If Not RightSplitTest() Then
        Debug.Print "Failed: RightSplitTest"
        TestStatus = False
    Else
        Debug.Print "Passed: RightSplitTest"
    End If
    
    If Not TrimCharTest() Then
        Debug.Print "Failed: TrimCharTest"
        TestStatus = False
    Else
        Debug.Print "Passed: TrimCharTest"
    End If
    
    If Not TrimLeftTest() Then
        Debug.Print "Failed: TrimLeftTest"
        TestStatus = False
    Else
        Debug.Print "Passed: TrimLeftTest"
    End If
    
    If Not TrimRightTest() Then
        Debug.Print "Failed: TrimRightTest"
        TestStatus = False
    Else
        Debug.Print "Passed: TrimRightTest"
    End If
    
    If Not CountUppercaseCharactersTest() Then
        Debug.Print "Failed: CountUppercaseCharactersTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CountUppercaseCharactersTest"
    End If
    
    If Not CountLowercaseCharactersTest() Then
        Debug.Print "Failed: CountLowercaseCharactersTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CountLowercaseCharactersTest"
    End If
    
    If Not TextJoinTest() Then
        Debug.Print "Failed: TextJoinTest"
        TestStatus = False
    Else
        Debug.Print "Passed: TextJoinTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibStringManipulationTests = TestStatus
    
End Function



Private Function CapitalizeTest() As Boolean

    '@Example: =Capitalize("hello World") -> "Hello world"

    CapitalizeTest = True

    CapitalizeTest = CapitalizeTest And Capitalize("hello World") = "Hello world"
    
End Function


Private Function LeftFindTest() As Boolean

    '@Example: =LeftFind("Hello World", "r") -> "Hello Wo"
    '@Example: =LeftFind("Hello World", "R") -> "#VALUE!"; Since string1 does not contain "R" in it.

    LeftFindTest = True

    LeftFindTest = LeftFindTest And LeftFind("Hello World", "r") = "Hello Wo"

End Function


Private Function RightFindTest() As Boolean

    '@Example: =RightFind("Hello World", "o") -> "rld"
    '@Example: =RightFind("Hello World", "O") -> "#VALUE!"; Since string1 does not contain "O" in it.

    RightFindTest = True

    RightFindTest = RightFindTest And RightFind("Hello World", "o") = "rld"

End Function


Private Function LeftSearchTest() As Boolean

    '@Example: =LeftSearch("Hello World", "r") -> "Hello Wo"
    '@Example: =LeftSearch("Hello World", "R") -> "Hello Wo"

    LeftSearchTest = True

    LeftSearchTest = LeftSearchTest And LeftSearch("Hello World", "r") = "Hello Wo"
    LeftSearchTest = LeftSearchTest And LeftSearch("Hello World", "R") = "Hello Wo"

End Function


Private Function RightSearchTest() As Boolean

    '@Example: =RightSearch("Hello World", "o") -> "rld"
    '@Example: =RightSearch("Hello World", "O") -> "rld"

    RightSearchTest = True

    RightSearchTest = RightSearchTest And RightSearch("Hello World", "o") = "rld"
    RightSearchTest = RightSearchTest And RightSearch("Hello World", "O") = "rld"

End Function


Private Function SubstrTest() As Boolean

    '@Example: =Substr("Hello World", 2, 6) -> "ello"

    SubstrTest = True

    SubstrTest = SubstrTest And Substr("Hello World", 2, 6) = "ello"

End Function


Private Function SubstrFindTest() As Boolean

    '@Example: =SubstrFind("Hello World", "e", "o") -> "ello Wo"
    '@Example: =SubstrFind("Hello World", "e", "o", TRUE) -> "llo W"
    '@Example: =SubstrFind("One Two Three", "ne ", " Thr") -> "ne Two Thr"
    '@Example: =SubstrFind("One Two Three", "NE ", " THR") -> "#VALUE!"; Since SubstrFind() is case-sensitive
    '@Example: =SubstrFind("One Two Three", "ne ", " Thr", TRUE) -> "Two"
    '@Example: =SubstrFind("Country Code: +51; Area Code: 315; Phone Number: 762-5929;", "Area Code: ", "; Phone", TRUE) -> 315
    '@Example: =SubstrFind("Country Code: +313; Area Code: 423; Phone Number: 284-2468;", "Area Code: ", "; Phone", TRUE) -> 423
    '@Example: =SubstrFind("Country Code: +171; Area Code: 629; Phone Number: 731-5456;", "Area Code: ", "; Phone", TRUE) -> 629

    SubstrFindTest = True

    SubstrFindTest = SubstrFindTest And SubstrFind("Hello World", "e", "o") = "ello Wo"
    SubstrFindTest = SubstrFindTest And SubstrFind("Hello World", "e", "o", True) = "llo W"
    SubstrFindTest = SubstrFindTest And SubstrFind("One Two Three", "ne ", " Thr") = "ne Two Thr"
    SubstrFindTest = SubstrFindTest And SubstrFind("One Two Three", "ne ", " Thr", True) = "Two"

End Function


Private Function SubstrSearchTest() As Boolean

    '@Example: =SubstrSearch("Hello World", "e", "o") -> "ello Wo"
    '@Example: =SubstrSearch("Hello World", "e", "o", TRUE) -> "llo W"
    '@Example: =SubstrSearch("One Two Three", "ne ", " Thr") -> "ne Two Thr"
    '@Example: =SubstrSearch("One Two Three", "NE ", " THR") -> "ne Two Thr"; No error, since SubstrSearch is case-insensitive
    '@Example: =SubstrSearch("One Two Three", "ne ", " Thr", TRUE) -> "Two"
    '@Example: =SubstrSearch("Country Code: +51; Area Code: 315; Phone Number: 762-5929;", "Area Code: ", "; Phone", TRUE) -> 315
    '@Example: =SubstrSearch("Country Code: +313; Area Code: 423; Phone Number: 284-2468;", "Area Code: ", "; Phone", TRUE) -> 423
    '@Example: =SubstrSearch("Country Code: +171; Area Code: 629; Phone Number: 731-5456;", "Area Code: ", "; Phone", TRUE) -> 629

    SubstrSearchTest = True

    SubstrSearchTest = SubstrSearchTest And SubstrSearch("Hello World", "e", "o") = "ello Wo"
    SubstrSearchTest = SubstrSearchTest And SubstrSearch("Hello World", "e", "o", True) = "llo W"
    SubstrSearchTest = SubstrSearchTest And SubstrSearch("One Two Three", "ne ", " Thr") = "ne Two Thr"
    SubstrSearchTest = SubstrSearchTest And SubstrSearch("One Two Three", "NE ", " THR") = "ne Two Thr"
    SubstrSearchTest = SubstrSearchTest And SubstrSearch("One Two Three", "ne ", " Thr", True) = "Two"

End Function

    
Private Function RepeatTest() As Boolean

    '@Example: =Repeat("Hello", 2) -> HelloHello"
    '@Example: =Repeat("=", 10) -> "=========="

    RepeatTest = True

    RepeatTest = RepeatTest And Repeat("Hello", 2) = "HelloHello"
    RepeatTest = RepeatTest And Repeat("=", 10) = "=========="

End Function


Private Function FormatterTest() As Boolean

    '@Example: =Formatter("Hello {1}", "World") -> "Hello World"
    '@Example: =Formatter("{1} {2}", "Hello", "World") -> "Hello World"
    '@Example: =Formatter("{1}.{2}@{3}", "FirstName", "LastName", "email.com") -> "FirstName.LastName@email.com"
    '@Example: =Formatter("{1}.{2}@{3}", A1:A3) -> "FirstName.LastName@email.com"; where A1="FirstName", A2="LastName", and A3="email.com"
    '@Example: =Formatter("{1}.{2}@{3}", A1, A2, A3) -> "FirstName.LastName@email.com"; where A1="FirstName", A2="LastName", and A3="email.com"

    FormatterTest = True

    FormatterTest = FormatterTest And Formatter("Hello {1}", "World") = "Hello World"
    FormatterTest = FormatterTest And Formatter("{1} {2}", "Hello", "World") = "Hello World"
    FormatterTest = FormatterTest And Formatter("{1}.{2}@{3}", "FirstName", "LastName", "email.com") = "FirstName.LastName@email.com"
    FormatterTest = FormatterTest And Formatter("{1}.{2}@{3}", Array("FirstName", "LastName", "email.com")) = "FirstName.LastName@email.com"

End Function


Private Function ZfillTest() As Boolean

    '@Example: =Zfill(123, 5) -> "00123"
    '@Example: =Zfill(5678, 5) -> "05678"
    '@Example: =Zfill(12345678, 5) -> "12345678"
    '@Example: =Zfill(123, 5, "X") -> "XX123"
    '@Example: =Zfill(123, 5, "X", TRUE) -> "123XX"
    
    ZfillTest = True

    ZfillTest = ZfillTest And Zfill(123, 5) = "00123"
    ZfillTest = ZfillTest And Zfill(5678, 5) = "05678"
    ZfillTest = ZfillTest And Zfill(12345678, 5) = "12345678"
    ZfillTest = ZfillTest And Zfill(123, 5, "X") = "XX123"
    ZfillTest = ZfillTest And Zfill(123, 5, "X", True) = "123XX"

End Function


Private Function SplitTextTest() As Boolean
    
    '@Example: =SplitText("Hello World", 1) -> "Hello"
    '@Example: =SplitText("Hello World", 2) -> "World"
    '@Example: =SplitText("One Two Three", 2) -> "Two"
    '@Example: =SplitText("One-Two-Three", 2, "-") -> "Two"
    
    SplitTextTest = True

    SplitTextTest = SplitTextTest And SplitText("Hello World", 1) = "Hello"
    SplitTextTest = SplitTextTest And SplitText("Hello World", 2) = "World"
    SplitTextTest = SplitTextTest And SplitText("One Two Three", 2) = "Two"
    SplitTextTest = SplitTextTest And SplitText("One-Two-Three", 2, "-") = "Two"

End Function


Private Function CountWordsTest() As Boolean

    '@Example: =CountWords("Hello World") -> 2
    '@Example: =CountWords("One Two Three") -> 3
    '@Example: =CountWords("One-Two-Three", "-") -> 3

    CountWordsTest = True

    CountWordsTest = CountWordsTest And CountWords("Hello World") = 2
    CountWordsTest = CountWordsTest And CountWords("One Two Three") = 3
    CountWordsTest = CountWordsTest And CountWords("One-Two-Three", "-") = 3

End Function


Private Function CamelCaseTest() As Boolean

    '@Example: =CamelCase("Hello World") -> "helloWorld"
    '@Example: =CamelCase("One Two Three") -> "oneTwoThree"

    CamelCaseTest = True

    CamelCaseTest = CamelCaseTest And CamelCase("Hello World") = "helloWorld"
    CamelCaseTest = CamelCaseTest And CamelCase("One Two Three") = "oneTwoThree"

End Function


Private Function KebabCaseTest() As Boolean

    '@Example: =KebabCase("Hello World") -> "hello-world"
    '@Example: =KebabCase("One Two Three") -> "one-two-three"

    KebabCaseTest = True

    KebabCaseTest = KebabCaseTest And KebabCase("Hello World") = "hello-world"
    KebabCaseTest = KebabCaseTest And KebabCase("One Two Three") = "one-two-three"

End Function


Private Function RemoveCharactersTest() As Boolean

    '@Example: =RemoveCharacters("Hello World", "l") -> "Heo Word"
    '@Example: =RemoveCharacters("Hello World", "lo") -> "He Wrd"
    '@Example: =RemoveCharacters("Hello World", "l", "o") -> "He Wrd"
    '@Example: =RemoveCharacters("Hello World", "lod") -> "He Wr"
    '@Example: =RemoveCharacters("One Two Three", "o", "t") -> "One Two Three"; Nothing is replaced since this function is case sensitive
    '@Example: =RemoveCharacters("One Two Three", "O", "T") -> "ne wo hree"

    RemoveCharactersTest = True

    RemoveCharactersTest = RemoveCharactersTest And RemoveCharacters("Hello World", "l") = "Heo Word"
    RemoveCharactersTest = RemoveCharactersTest And RemoveCharacters("Hello World", "lo") = "He Wrd"
    RemoveCharactersTest = RemoveCharactersTest And RemoveCharacters("Hello World", "l", "o") = "He Wrd"
    RemoveCharactersTest = RemoveCharactersTest And RemoveCharacters("Hello World", "lod") = "He Wr"
    RemoveCharactersTest = RemoveCharactersTest And RemoveCharacters("Two Three Four", "f", "t") = "Two Three Four"
    RemoveCharactersTest = RemoveCharactersTest And RemoveCharacters("Two Three Four", "F", "T") = "wo hree our"

End Function


Private Function CompanyCaseTest() As Boolean

    '@Example: =CompanyCase("hello world") -> "Hello World"
    '@Example: =CompanyCase("x.y.z company & co.") -> "X.Y.Z Company & Co."
    '@Example: =CompanyCase("x.y.z plc") -> "X.Y.Z PLC"
    '@Example: =CompanyCase("one company gmbh") -> "One Company GmbH"
    '@Example: =CompanyCase("three company s. en n.c.") -> "Three Company S. en N.C."
    '@Example: =CompanyCase("FOUR COMPANY SPOL S.R.O.") -> "Four Company spol s.r.o."
    '@Example: =CompanyCase("five company bvba") -> "Five Company BVBA"

    CompanyCaseTest = True

    CompanyCaseTest = CompanyCaseTest And CompanyCase("hello world") = "Hello World"
    CompanyCaseTest = CompanyCaseTest And CompanyCase("x.y.z company & co.") = "X.Y.Z Company & Co."
    CompanyCaseTest = CompanyCaseTest And CompanyCase("x.y.z plc") = "X.Y.Z PLC"
    CompanyCaseTest = CompanyCaseTest And CompanyCase("one company gmbh") = "One Company GmbH"
    CompanyCaseTest = CompanyCaseTest And CompanyCase("three company s. en n.c.") = "Three Company S. en N.C."
    CompanyCaseTest = CompanyCaseTest And CompanyCase("FOUR COMPANY SPOL S.R.O.") = "Four Company spol s.r.o."
    CompanyCaseTest = CompanyCaseTest And CompanyCase("five company bvba") = "Five Company BVBA"

End Function


Private Function ReverseTextTest() As Boolean

    '@Example: =ReverseText("Hello World") -> "dlroW olleH"

    ReverseTextTest = True

    ReverseTextTest = ReverseTextTest And ReverseText("Hello World") = "dlroW olleH"

End Function


Private Function ReverseWordsTest() As Boolean

    '@Example: =ReverseWords("Hello World") -> "World Hello"
    '@Example: =ReverseWords("One Two Three") -> "Three Two One"
    '@Example: =ReverseWords("One-Two-Three", "-") -> "Three-Two-One"

    ReverseWordsTest = True

    ReverseWordsTest = ReverseWordsTest And ReverseWords("Hello World") = "World Hello"
    ReverseWordsTest = ReverseWordsTest And ReverseWords("One Two Three") = "Three Two One"
    ReverseWordsTest = ReverseWordsTest And ReverseWords("One-Two-Three", "-") = "Three-Two-One"

End Function


Private Function IndentTextTest() As Boolean

    '@Example: =IndentText("Hello") -> "    Hello"
    '@Example: =IndentText("Hello", 4) -> "    Hello"
    '@Example: =IndentText("Hello", 3) -> "   Hello"
    '@Example: =IndentText("Hello", 2) -> "  Hello"
    '@Example: =IndentText("Hello", 1) -> " Hello"

    IndentTextTest = True

    IndentTextTest = IndentTextTest And IndentText("Hello") = "    Hello"
    IndentTextTest = IndentTextTest And IndentText("Hello", 4) = "    Hello"
    IndentTextTest = IndentTextTest And IndentText("Hello", 3) = "   Hello"
    IndentTextTest = IndentTextTest And IndentText("Hello", 2) = "  Hello"
    IndentTextTest = IndentTextTest And IndentText("Hello", 1) = " Hello"

End Function


Private Function DedentTextTest() As Boolean

    '@Example: =DedentText("    Hello") -> "Hello"

    DedentTextTest = True

    DedentTextTest = DedentTextTest And DedentText("    Hello") = "Hello"

End Function


Private Function ShortenTextTest() As Boolean

    '@Example: =ShortenText("Hello World One Two Three", 20) -> "Hello World [...]"; Only the first two words and the placeholder will result in a string that is less than or equal to 20 in length
    '@Example: =ShortenText("Hello World One Two Three", 15) -> "Hello [...]"; Only the first word and the placeholder will result in a string that is less than or equal to 15 in length
    '@Example: =ShortenText("Hello World One Two Three") -> "Hello World One Two Three"; Since this string is shorter than the default 80 shorten width value, no placeholder will be used and the string wont be shortened
    '@Example: =ShortenText("Hello World One Two Three", 15, "-->") -> "Hello World -->"; A new placeholder is used
    '@Example: =ShortenText("Hello_World_One_Two_Three", 15, "-->", "_") -> "Hello_World_-->"; A new placeholder andd delimiter is used

    ShortenTextTest = True

    ShortenTextTest = ShortenTextTest And ShortenText("Hello World One Two Three", 20) = "Hello World [...]"
    ShortenTextTest = ShortenTextTest And ShortenText("Hello World One Two Three", 15) = "Hello [...]"
    ShortenTextTest = ShortenTextTest And ShortenText("Hello World One Two Three") = "Hello World One Two Three"
    ShortenTextTest = ShortenTextTest And ShortenText("Hello World One Two Three", 15, "-->") = "Hello World -->"
    ShortenTextTest = ShortenTextTest And ShortenText("Hello_World_One_Two_Three", 15, "-->", "_") = "Hello_World_-->"

End Function


Private Function InSplitTest() As Boolean

    '@Example: =InSplit("Hello", "Hello World One Two Three") -> TRUE; Since "Hello" is found within the searchString after being split
    '@Example: =InSplit("NotInString", "Hello World One Two Three") -> FALSE; Since "NotInString" is not found within the searchString after being split
    '@Example: =InSplit("Hello", "Hello-World-One-Two-Three", "-") -> TRUE; Since "Hello" is found and since the delimiter is set to "-"

    InSplitTest = True

    InSplitTest = InSplitTest And InSplit("Hello", "Hello World One Two Three") = True
    InSplitTest = InSplitTest And InSplit("NotInString", "Hello World One Two Three") = False
    InSplitTest = InSplitTest And InSplit("Hello", "Hello-World-One-Two-Three", "-") = True

End Function


Private Function EliteCaseTest() As Boolean

    '@Example: =EliteCase("Hello World") -> "H3110 W0r1d"

    EliteCaseTest = True

    EliteCaseTest = EliteCaseTest And EliteCase("Hello World") = "H3110 W0r1d"

End Function


Private Function ScrambleCaseTest() As Boolean

    '@Example: =ScrambleCase("Hello World") -> "helLo WORlD"
    '@Example: =ScrambleCase("Hello World") -> "HElLo WorLD"
    '@Example: =ScrambleCase("Hello World") -> "hELlo WOrLd"

    ScrambleCaseTest = True

    Dim testString As String
    testString = "a"
    
    ScrambleCaseTest = ScrambleCaseTest And (testString = "a" Or testString = "A")

End Function


Private Function LeftSplitTest() As Boolean

    '@Example: =LeftSplit("Hello World One Two Three", 1) -> "Hello"
    '@Example: =LeftSplit("Hello World One Two Three", 2) -> "Hello World"
    '@Example: =LeftSplit("Hello World One Two Three", 3) -> "Hello World One"
    '@Example: =LeftSplit("Hello World One Two Three", 10) -> "Hello World One Two Three"
    '@Example: =LeftSplit("Hello-World-One-Two-Three", 2, "-") -> "Hello-World"

    LeftSplitTest = True

    LeftSplitTest = LeftSplitTest And LeftSplit("Hello World One Two Three", 1) = "Hello"
    LeftSplitTest = LeftSplitTest And LeftSplit("Hello World One Two Three", 2) = "Hello World"
    LeftSplitTest = LeftSplitTest And LeftSplit("Hello World One Two Three", 3) = "Hello World One"
    LeftSplitTest = LeftSplitTest And LeftSplit("Hello World One Two Three", 10) = "Hello World One Two Three"
    LeftSplitTest = LeftSplitTest And LeftSplit("Hello-World-One-Two-Three", 2, "-") = "Hello-World"

End Function


Private Function RightSplitTest() As Boolean

    '@Example: =RightSplit("Hello World One Two Three", 1) -> "Three"
    '@Example: =RightSplit("Hello World One Two Three", 2) -> "Two Three"
    '@Example: =RightSplit("Hello World One Two Three", 3) -> "One Two Three"
    '@Example: =RightSplit("Hello World One Two Three", 10) -> "Hello World One Two Three"
    '@Example: =RightSplit("Hello-World-One-Two-Three", 2, "-") -> "Two-Three"

    RightSplitTest = True

    RightSplitTest = RightSplitTest And RightSplit("Hello World One Two Three", 1) = "Three"
    RightSplitTest = RightSplitTest And RightSplit("Hello World One Two Three", 2) = "Two Three"
    RightSplitTest = RightSplitTest And RightSplit("Hello World One Two Three", 3) = "One Two Three"
    RightSplitTest = RightSplitTest And RightSplit("Hello World One Two Three", 10) = "Hello World One Two Three"
    RightSplitTest = RightSplitTest And RightSplit("Hello-World-One-Two-Three", 2, "-") = "Two-Three"

End Function


Private Function TrimCharTest() As Boolean

    '@Example: =TrimChar("   Hello World   ") -> "Hello World"
    '@Example: =TrimChar("---Hello World---", "-") -> "Hello World"

    TrimCharTest = True

    TrimCharTest = TrimCharTest And TrimChar("   Hello World   ") = "Hello World"
    TrimCharTest = TrimCharTest And TrimChar("---Hello World---", "-") = "Hello World"

End Function


Private Function TrimLeftTest() As Boolean

    '@Example: =TrimLeft("   Hello World   ") -> "Hello World   "
    '@Example: =TrimLeft("---Hello World---", "-") -> "Hello World---"

    TrimLeftTest = True

    TrimLeftTest = TrimLeftTest And TrimLeft("   Hello World   ") = "Hello World   "
    TrimLeftTest = TrimLeftTest And TrimLeft("---Hello World---", "-") = "Hello World---"

End Function


Private Function TrimRightTest() As Boolean

    '@Example: =TrimRight("   Hello World   ") -> "   Hello World"
    '@Example: =TrimRight("---Hello World---", "-") -> "---Hello World"
    
    TrimRightTest = True

    TrimRightTest = TrimRightTest And TrimRight("   Hello World   ") = "   Hello World"
    TrimRightTest = TrimRightTest And TrimRight("---Hello World---", "-") = "---Hello World"

End Function


Private Function CountUppercaseCharactersTest() As Boolean

    '@Example: =CountUppercaseCharacters("Hello World") -> 2; As the "H" and the "E" are the only 2 uppercase characters in the string

    CountUppercaseCharactersTest = True

    CountUppercaseCharactersTest = CountUppercaseCharactersTest And CountUppercaseCharacters("Hello World") = 2

End Function


Private Function CountLowercaseCharactersTest() As Boolean

    '@Example: =CountLowercaseCharacters("Hello World") -> 8; As the "ello" and the "orld" are lowercase

    CountLowercaseCharactersTest = True

    CountLowercaseCharactersTest = CountLowercaseCharactersTest And CountLowercaseCharacters("Hello World") = 8

End Function


Private Function TextJoinTest() As Boolean

    '@Example: =TextJoin(A1:A3) -> "123"; Where A1:A3 contains ["1", "2", "3"]
    '@Example: =TextJoin(A1:A3, "--") -> "1--2--3"; Where A1:A3 contains ["1", "2", "3"]
    '@Example: =TextJoin(A1:A3, "--") -> "1----3"; Where A1:A3 contains ["1", "", "3"]
    '@Example: =TextJoin(A1:A3, "-") -> "1--3"; Where A1:A3 contains ["1", "", "3"]
    '@Example: =TextJoin(A1:A3, "-", TRUE) -> "1-3"; Where A1:A3 contains ["1", "", "3"]

    TextJoinTest = True

    TextJoinTest = TextJoinTest And TextJoin(Array("1", "2", "3")) = "123"
    TextJoinTest = TextJoinTest And TextJoin(Array("1", "2", "3"), "--") = "1--2--3"
    TextJoinTest = TextJoinTest And TextJoin(Array("1", "", "3"), "--") = "1----3"
    TextJoinTest = TextJoinTest And TextJoin(Array("1", "", "3"), "-") = "1--3"
    TextJoinTest = TextJoinTest And TextJoin(Array("1", "", "3"), "-", True) = "1-3"

End Function


