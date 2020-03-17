Attribute VB_Name = "XlibTests"
'The MIT License (MIT)
'Copyright © 2020 Anthony Mancini
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Option Explicit
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




Public Function AllXlibColorTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    If Not Rgb2HexTest() Then
        Debug.Print "Failed: Rgb2HexTest"
        TestStatus = False
    Else
        Debug.Print "Passed: Rgb2HexTest"
    End If
    
    If Not Hex2RgbTest() Then
        Debug.Print "Failed: Hex2RgbTest"
        TestStatus = False
    Else
        Debug.Print "Passed: Hex2RgbTest"
    End If
    
    If Not Rgb2HslTest() Then
        Debug.Print "Failed: Rgb2HslTest"
        TestStatus = False
    Else
        Debug.Print "Passed: Rgb2HslTest"
    End If
    
    If Not Hex2HslTest() Then
        Debug.Print "Failed: Hex2HslTest"
        TestStatus = False
    Else
        Debug.Print "Passed: Hex2HslTest"
    End If
    
    If Not Hsl2RgbTest() Then
        Debug.Print "Failed: Hsl2RgbTest"
        TestStatus = False
    Else
        Debug.Print "Passed: Hsl2RgbTest"
    End If
    
    If Not Hsl2HexTest() Then
        Debug.Print "Failed: Hsl2HexTest"
        TestStatus = False
    Else
        Debug.Print "Passed: Hsl2HexTest"
    End If
    
    If Not Rgb2HsvTest() Then
        Debug.Print "Failed: Rgb2HsvTest"
        TestStatus = False
    Else
        Debug.Print "Passed: Rgb2HsvTest"
    End If
    
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibColorTests = TestStatus
    
End Function



Private Function Rgb2HexTest() As Boolean

    '@Example: =Rgb2Hex(255, 255, 255) -> "FFFFFF"
    
    Rgb2HexTest = True

    Rgb2HexTest = Rgb2HexTest And Rgb2Hex(255, 255, 255) = "FFFFFF"
    
End Function


Private Function Hex2RgbTest() As Boolean

    '@Example: =Hex2Rgb("FFFFFF") -> "(255, 255, 255)"
    '@Example: =Hex2Rgb("FF0109", 0) -> 255; The red color
    '@Example: =Hex2Rgb("FF0109", "Red") -> 255; The red color
    '@Example: =Hex2Rgb("FF0109", 1) -> 1; The green color
    '@Example: =Hex2Rgb("FF0109", "Green") -> 1; The green color
    '@Example: =Hex2Rgb("FF0109", 2) -> 9; The blue color
    '@Example: =Hex2Rgb("FF0109", "Blue") -> 9; The blue color

    Hex2RgbTest = True

    Hex2RgbTest = Hex2RgbTest And Hex2Rgb("FFFFFF") = "(255, 255, 255)"
    Hex2RgbTest = Hex2RgbTest And Hex2Rgb("FF0109", 0) = 255
    Hex2RgbTest = Hex2RgbTest And Hex2Rgb("FF0109", "Red") = 255
    Hex2RgbTest = Hex2RgbTest And Hex2Rgb("FF0109", 1) = 1
    Hex2RgbTest = Hex2RgbTest And Hex2Rgb("FF0109", "Green") = 1
    Hex2RgbTest = Hex2RgbTest And Hex2Rgb("FF0109", 2) = 9
    Hex2RgbTest = Hex2RgbTest And Hex2Rgb("FF0109", "Blue") = 9

End Function


Private Function Rgb2HslTest() As Boolean

    '@Example: =Rgb2Hsl(8, 64, 128) -> "(212.0ï¿½, 88.2%, 26.7%)"
    '@Example: =Rgb2Hsl(8, 64, 128, 0) -> 212
    '@Example: =Rgb2Hsl(8, 64, 128, "Hue") -> 212
    '@Example: =Rgb2Hsl(8, 64, 128, 1) -> .882
    '@Example: =Rgb2Hsl(8, 64, 128, "Saturation") -> .882
    '@Example: =Rgb2Hsl(8, 64, 128, 2) -> .267
    '@Example: =Rgb2Hsl(8, 64, 128, "Lightness") -> .267

    Rgb2HslTest = True

    Rgb2HslTest = Rgb2HslTest And Rgb2Hsl(8, 64, 128) = "(212.0, 88.2%, 26.7%)"
    Rgb2HslTest = Rgb2HslTest And Rgb2Hsl(8, 64, 128, 0) = 212
    Rgb2HslTest = Rgb2HslTest And Rgb2Hsl(8, 64, 128, "Hue") = 212
    Rgb2HslTest = Rgb2HslTest And Round(Rgb2Hsl(8, 64, 128, 1), 3) = 0.882
    Rgb2HslTest = Rgb2HslTest And Round(Rgb2Hsl(8, 64, 128, "Saturation"), 3) = 0.882
    Rgb2HslTest = Rgb2HslTest And Round(Rgb2Hsl(8, 64, 128, 2), 3) = 0.267
    Rgb2HslTest = Rgb2HslTest And Round(Rgb2Hsl(8, 64, 128, "Lightness"), 3) = 0.267

End Function


Private Function Hex2HslTest() As Boolean

    '@Example: =Hex2Hsl("084080") -> "(212.0, 88.2%, 26.7%)"
    '@Example: =Hex2Hsl("#084080") -> "(212.0, 88.2%, 26.7%)"

    Hex2HslTest = True

    Hex2HslTest = Hex2HslTest And Hex2Hsl("084080") = "(212.0, 88.2%, 26.7%)"
    Hex2HslTest = Hex2HslTest And Hex2Hsl("#084080") = "(212.0, 88.2%, 26.7%)"

End Function


Private Function Hsl2RgbTest() As Boolean

    '@Example: =Hsl2Rgb(212, .882, .267) -> "(8, 64, 128)"
    '@Example: =Hsl2Rgb(212, .882, .267, 0) -> 8
    '@Example: =Hsl2Rgb(212, .882, .267, "Red") -> 8
    '@Example: =Hsl2Rgb(212, .882, .267, 1) -> 64
    '@Example: =Hsl2Rgb(212, .882, .267, "Green") -> 64
    '@Example: =Hsl2Rgb(212, .882, .267, 2) -> 128
    '@Example: =Hsl2Rgb(212, .882, .267, "Blue") -> 128

    Hsl2RgbTest = True

    Hsl2RgbTest = Hsl2RgbTest And Hsl2Rgb(212, 0.882, 0.267) = "(8, 64, 128)"
    Hsl2RgbTest = Hsl2RgbTest And Hsl2Rgb(212, 0.882, 0.267, 0) = 8
    Hsl2RgbTest = Hsl2RgbTest And Hsl2Rgb(212, 0.882, 0.267, "Red") = 8
    Hsl2RgbTest = Hsl2RgbTest And Hsl2Rgb(212, 0.882, 0.267, 1) = 64
    Hsl2RgbTest = Hsl2RgbTest And Hsl2Rgb(212, 0.882, 0.267, "Green") = 64
    Hsl2RgbTest = Hsl2RgbTest And Hsl2Rgb(212, 0.882, 0.267, 2) = 128
    Hsl2RgbTest = Hsl2RgbTest And Hsl2Rgb(212, 0.882, 0.267, 2) = 128
    Hsl2RgbTest = Hsl2RgbTest And Hsl2Rgb(212, 0.882, 0.267, "Blue") = 128

End Function


Private Function Hsl2HexTest() As Boolean

    '@Example: =Hsl2Hex(212, .882, .267) -> "084080"

    Hsl2HexTest = True

    Hsl2HexTest = Hsl2HexTest And Hsl2Hex(212, 0.882, 0.267) = "084080"

End Function


Private Function Rgb2HsvTest() As Boolean

    '@Example: =Rgb2Hsv(8, 64, 128) -> "(212.0, 93.8%, 50.2%)"
    '@Example: =Rgb2Hsv(8, 64, 128, 0) -> 212
    '@Example: =Rgb2Hsv(8, 64, 128, "Hue") -> 212
    '@Example: =Rgb2Hsv(8, 64, 128, 1) -> .938
    '@Example: =Rgb2Hsv(8, 64, 128, "Saturation") -> .938
    '@Example: =Rgb2Hsv(8, 64, 128, 2) -> .502
    '@Example: =Rgb2Hsv(8, 64, 128, "Value") -> .502

    Rgb2HsvTest = True

    Rgb2HsvTest = Rgb2HsvTest And Rgb2Hsv(8, 64, 128) = "(212.0, 93.8%, 50.2%)"
    Rgb2HsvTest = Rgb2HsvTest And Rgb2Hsv(8, 64, 128, 0) = 212
    Rgb2HsvTest = Rgb2HsvTest And Rgb2Hsv(8, 64, 128, "Hue") = 212
    Rgb2HsvTest = Rgb2HsvTest And Round(Rgb2Hsv(8, 64, 128, 1), 3) = 0.938
    Rgb2HsvTest = Rgb2HsvTest And Round(Rgb2Hsv(8, 64, 128, "Saturation"), 3) = 0.938
    Rgb2HsvTest = Rgb2HsvTest And Round(Rgb2Hsv(8, 64, 128, "Value"), 3) = 0.502
    
End Function


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




Public Function AllXlibEnvironmentTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not OSTest() Then
        Debug.Print "Failed: OSTest"
        TestStatus = False
    Else
        Debug.Print "Passed: OSTest"
    End If
    
    If Not UserNameTest() Then
        Debug.Print "Failed: UserNameTest"
        TestStatus = False
    Else
        Debug.Print "Passed: UserNameTest"
    End If
    
    If Not UserDomainTest() Then
        Debug.Print "Failed: UserDomainTest"
        TestStatus = False
    Else
        Debug.Print "Passed: UserDomainTest"
    End If
    
    If Not ComputerNameTest() Then
        Debug.Print "Failed: ComputerNameTest"
        TestStatus = False
    Else
        Debug.Print "Passed: ComputerNameTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibEnvironmentTests = TestStatus
    
End Function



Private Function OSTest() As Boolean

    '@Example: =OS() -> "Windows"; When running this function on Windows
    '@Example: =OS() -> "Mac"; When running this function on MacOS

    OSTest = True

    #If Mac Then
        OSTest = OSTest And OS() = "Mac"
    #Else
        OSTest = OSTest And OS() = "Windows"
    #End If

End Function


Private Function UserNameTest() As Boolean

    '@Example: =UserName() -> "Anthony"
    
    UserNameTest = True

    #If Mac Then
        UserNameTest = UserNameTest And UserName() = Environ("USER")
    #Else
        UserNameTest = UserNameTest And UserName() = Environ("USERNAME")
    #End If

End Function


Private Function UserDomainTest() As Boolean

    '@Example: =UserDomain() -> "DESKTOP-XYZ1234"
    
    UserDomainTest = True
    
    #If Mac Then
        UserDomainTest = UserDomainTest And UserDomain() = Environ("HOST")
    #Else
        UserDomainTest = UserDomainTest And UserDomain() = Environ("USERDOMAIN")
    #End If

End Function


Private Function ComputerNameTest() As Boolean

    '@Example: =ComputerName() -> "DESKTOP-XYZ1234"

    ComputerNameTest = True
    
    ComputerNameTest = ComputerNameTest And ComputerName() = Environ("COMPUTERNAME")

End Function



Public Function AllXlibFileTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not GetActivePathAndNameTest() Then
        Debug.Print "Failed: GetActivePathAndNameTest"
        TestStatus = False
    Else
        Debug.Print "Passed: GetActivePathAndNameTest"
    End If
    
    If Not GetActivePathTest() Then
        Debug.Print "Failed: GetActivePathTest"
        TestStatus = False
    Else
        Debug.Print "Passed: GetActivePathTest"
    End If
    
    If Not FileCreationTimeTest() Then
        Debug.Print "Failed: FileCreationTimeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileCreationTimeTest"
    End If
    
    If Not FileLastModifiedTimeTest() Then
        Debug.Print "Failed: FileLastModifiedTimeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileLastModifiedTimeTest"
    End If
    
    If Not FileDriveTest() Then
        Debug.Print "Failed: FileDriveTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileDriveTest"
    End If
    
    If Not FileNameTest() Then
        Debug.Print "Failed: FileNameTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileNameTest"
    End If
    
    If Not FileFolderTest() Then
        Debug.Print "Failed: FileFolderTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileFolderTest"
    End If
    
    If Not CurrentFilePathTest() Then
        Debug.Print "Failed: CurrentFilePathTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CurrentFilePathTest"
    End If
    
    If Not FileSizeTest() Then
        Debug.Print "Failed: FileSizeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileSizeTest"
    End If
    
    If Not FileTypeTest() Then
        Debug.Print "Failed: FileTypeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileTypeTest"
    End If
    
    If Not FileExtensionTest() Then
        Debug.Print "Failed: FileExtensionTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileExtensionTest"
    End If
    
    If Not WriteFileTest() Then
        Debug.Print "Failed: WriteFileTest"
        TestStatus = False
    Else
        Debug.Print "Passed: WriteFileTest"
    End If
    
    If Not ReadFileTest() Then
        Debug.Print "Failed: ReadFileTest"
        TestStatus = False
    Else
        Debug.Print "Passed: ReadFileTest"
    End If
    
    If Not PathSeparatorTest() Then
        Debug.Print "Failed: PathSeparatorTest"
        TestStatus = False
    Else
        Debug.Print "Passed: PathSeparatorTest"
    End If
    
    If Not PathJoinTest() Then
        Debug.Print "Failed: PathJoinTest"
        TestStatus = False
    Else
        Debug.Print "Passed: PathJoinTest"
    End If
    
    If Not CountFilesTest() Then
        Debug.Print "Failed: CountFilesTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CountFilesTest"
    End If
    
    If Not CountFilesAndFoldersTest() Then
        Debug.Print "Failed: CountFilesAndFoldersTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CountFilesAndFoldersTest"
    End If
    
    If Not GetFileNameByNumberTest() Then
        Debug.Print "Failed: GetFileNameByNumberTest"
        TestStatus = False
    Else
        Debug.Print "Passed: GetFileNameByNumberTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibFileTests = TestStatus
    
End Function



Private Function GetActivePathAndNameTest() As Boolean

    '@Example: =GetActivePathAndName() -> "C:\Users\UserName\Documents\XLib.xlsm"
    
    GetActivePathAndNameTest = True
    
    If InStr(1, GetActivePathAndName(), ".") > 0 Then
        #If Mac Then
            If InStr(1, GetActivePathAndName(), "/") > 0 Then
                GetActivePathAndNameTest = True
            Else
                GetActivePathAndNameTest = False
            End If
        #Else
            If InStr(1, GetActivePathAndName(), ":\") > 0 Then
                GetActivePathAndNameTest = True
            Else
                GetActivePathAndNameTest = False
            End If
        #End If
    End If

End Function


Private Function GetActivePathTest() As Boolean

    '@Example: =GetActivePath() -> "C:\Users\UserName\Documents\"

    GetActivePathTest = True
        
    #If Mac Then
        If InStr(1, GetActivePath(), "/") > 0 Then
            GetActivePathTest = True
        Else
            GetActivePathTest = False
        End If
    #Else
        If InStr(1, GetActivePath(), ":\") > 0 Then
            GetActivePathTest = True
        Else
            GetActivePathTest = False
        End If
    #End If

End Function


Private Function FileCreationTimeTest() As Boolean

    '@Example: =FileCreationTime() -> "1/1/2020 1:23:45 PM"
    '@Example: =FileCreationTime("C:\hello\world.txt") -> "1/1/2020 5:55:55 PM"
    '@Example: =FileCreationTime("vba.txt") -> "12/25/2000 1:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    FileCreationTimeTest = False
    
    If InStr(1, FileCreationTime(), " ") > 0 Then
        If InStr(1, FileCreationTime(), ":") > 0 Then
            FileCreationTimeTest = True
        End If
    End If

End Function


Private Function FileLastModifiedTimeTest() As Boolean

    '@Example: =FileLastModifiedTime() -> "1/1/2020 2:23:45 PM"
    '@Example: =FileLastModifiedTime("C:\hello\world.txt") -> "1/1/2020 7:55:55 PM"
    '@Example: =FileLastModifiedTime("vba.txt") -> "12/25/2000 3:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    FileLastModifiedTimeTest = False
    
    If InStr(1, FileLastModifiedTime(), " ") > 0 Then
        If InStr(1, FileLastModifiedTime(), ":") > 0 Then
            FileLastModifiedTimeTest = True
        End If
    End If

End Function


Private Function FileDriveTest() As Boolean

    '@Example: =FileDrive() -> "A:"; Where the current workbook resides on the A: drive
    '@Example: =FileDrive("C:\hello\world.txt") -> "C:"
    '@Example: =FileDrive("vba.txt") -> "B:"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in, and where the workbook resides in the B: drive

    If InStr(1, FileDrive(), ":") > 0 Then
        FileDriveTest = True
    End If

End Function


Private Function FileNameTest() As Boolean

    '@Example: =FileName() -> "MyWorkbook.xlsm"
    '@Example: =FileName("C:\hello\world.txt") -> "world.txt"
    '@Example: =FileName("vba.txt") -> "vba.txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    If Len(FileName()) > 0 Then
        FileNameTest = True
    End If

End Function


Private Function FileFolderTest() As Boolean

    '@Example: =FileFolder() -> "C:\my_excel_files"
    '@Example: =FileFolder("C:\hello\world.txt") -> "C:\hello"
    '@Example: =FileFolder("vba.txt") -> "C:\my_excel_files"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in
    
    #If Mac Then
        If InStr(1, FileFolder(), "/") > 0 Then
            FileFolderTest = True
        End If
    #Else
        If InStr(1, FileFolder(), ":\") > 0 Then
            FileFolderTest = True
        End If
    #End If

End Function


Private Function CurrentFilePathTest() As Boolean

    '@Example: =CurrentFilePath() -> "C:\my_excel_files\MyWorkbook.xlsx"
    '@Example: =CurrentFilePath("C:\hello\world.txt") -> "C:\hello\world.txt"
    '@Example: =CurrentFilePath("vba.txt") -> "C:\hello\world.txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    #If Mac Then
        If InStr(1, CurrentFilePath(), "/") > 0 Then
            CurrentFilePathTest = True
        End If
    #Else
        If InStr(1, CurrentFilePath(), ":\") > 0 Then
            CurrentFilePathTest = True
        End If
    #End If

End Function


Private Function FileSizeTest() As Boolean

    '@Example: =FileSize() -> 1024
    '@Example: =FileSize(,"KB") -> 1
    '@Example: =FileSize("vba.txt", "KB") -> 0.25; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in
    
    If FileSize() > 0 Then
        FileSizeTest = True
    End If

End Function


Private Function FileTypeTest() As Boolean

    '@Example: FileType() -> "Microsoft Excel Macro-Enabled Worksheet"
    '@Example: FileType("C:\hello\world.txt") -> "Text Document"
    '@Example: FileType("vba.txt") -> "Text Document"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in
    
    If Len(FileType()) > 0 Then
        FileTypeTest = True
    End If

End Function


Private Function FileExtensionTest() As Boolean

    '@Example: =FileExtension() = "xlsx"
    '@Example: =FileExtension("C:\hello\world.txt") -> "txt"
    '@Example: =FileExtension("vba.txt") -> "txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    If Len(FileExtension()) > 0 Then
        FileExtensionTest = True
    End If

End Function


Private Function WriteFileTest() As Boolean

    '@Example: =WriteFile("C:\MyWorkbookFolder\hello.txt", "Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"
    '@Example: =WriteFile("hello.txt", "Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"; Where the Workbook resides in "C:\MyWorkbookFolder\"
    
    If WriteFile("TempTestFile.txt", "Hello World") Then
        WriteFileTest = True
    End If

End Function


Private Function ReadFileTest() As Boolean

    '@Example: =ReadFile("C:\hello\world.txt") -> "Hello" World
    '@Example: =ReadFile("vba.txt") -> "This is my VBA text file"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in
    '@Example: =ReadFile("multline.txt", 1) -> "This is line 1";
    '@Example: =ReadFile("multline.txt", 2) -> "This is line 2";

    ReadFileTest = IIf(ReadFile("TempTestFile.txt") = "Hello World", True, False)
    Kill (GetActivePath() & "TempTestFile.txt")

End Function


Private Function PathSeparatorTest() As Boolean

    '@Example: =PathSeparator() -> "\"; When running this code on Windows
    '@Example: =PathSeparator() -> "/"; When running this code on Mac
    
    If PathSeparator() = "\" Or PathSeparator() = "/" Then
        PathSeparatorTest = True
    End If

End Function


Private Function PathJoinTest() As Boolean

    '@Example: =PathJoin("C:", "hello", "world.txt") -> "C:\hello\world.txt"; On Windows
    '@Example: =PathJoin("hello", "world.txt") -> "/hello/world.txt"; On Mac

    If PathJoin("hello", "world") = "hello/world" Or PathJoin("hello", "world") = "hello\world" Then
        PathJoinTest = True
    End If
    
End Function


Private Function CountFilesTest() As Boolean

    '@Example: =CountFiles() -> 6
    '@Example: =CountFiles("C:\hello") -> 10

    If CountFiles() > 0 Then
        CountFilesTest = True
    End If

End Function


Private Function CountFilesAndFoldersTest() As Boolean

    '@Example: =CountFilesAndFolders() -> 8
    '@Example: =CountFilesAndFolders("C:\hello") -> 30
    
    If CountFilesAndFolders() > 0 Then
        CountFilesAndFoldersTest = True
    End If

End Function


Private Function GetFileNameByNumberTest() As Boolean

    '@Example: =GetFileName(,1) -> "hello.txt"
    '@Example: =GetFileName(,1) -> "world.txt"
    '@Example: =GetFileName("C:\hello", 1) -> "one.txt"
    '@Example: =GetFileName("C:\hello", 1) -> "two.txt"
    '@Example: =GetFileName("C:\hello", 1) -> "three.txt"

    If Len(GetFileNameByNumber()) > 0 Then
        GetFileNameByNumberTest = True
    End If

End Function




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



Public Function AllXlibNetworkTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not HttpTest() Then
        Debug.Print "Failed: HttpTest"
        TestStatus = False
    Else
        Debug.Print "Passed: HttpTest"
    End If
    
    If Not SimpleHttpTest() Then
        Debug.Print "Failed: SimpleHttpTest"
        TestStatus = False
    Else
        Debug.Print "Passed: SimpleHttpTest"
    End If
    
    If Not ParseHtmlStringTest() Then
        Debug.Print "Failed: ParseHtmlStringTest"
        TestStatus = False
    Else
        Debug.Print "Passed: ParseHtmlStringTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibNetworkTests = TestStatus
    
End Function



Private Function HttpTest() As Boolean

    '@Example: =HTTP("https://httpbin.org/uuid") -> "{"uuid: "41416bcf-ef11-4256-9490-63853d14e4e8"}"
    '@Example: =HTTP("https://httpbin.org/user-agent", "GET", {"User-Agent","MicrosoftExcel"}) -> "{"user-agent": "MicrosoftExcel"}"
    '@Example: =HTTP("https://httpbin.org/status/404",,,,,TRUE) -> "#RequestFailedStatusCode404!"; Since the status error handler flag is set and since this URL returns a 404 status code. Also note that this formula is easier to construct using the Excel Formula Builder
    '@Example: =HTTP("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000}) -> Returning a string with the leftmost 3000 characters found within the element with the ID "mw-content-text" (we are trying to get the release date of VBA from the VBA wikipedia page, but we need to do more parsing first)
    '@Example: =HTTP("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000,"MID","appeared"}) -> Returns the prior string, but now with all characters right of the first occurance of the word "appeared" in the HTML (getting closer to parsing the VBA creation date)
    '@Example: =HTTP("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>"}) -> From the prior result, now returning everything after the first occurance of the "<TD>" in the prior string
    '@Example: =HTTP("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>","LEFT","<span"}) -> "1993"; Finally this is all the parsing needed to be able to return the date 1993 that we were looking for

    If InStr(1, Http("https://httpbin.org/user-agent", "GET", Array("User-Agent", "MicrosoftExcel")), Chr(34) & "user-agent" & Chr(34) & ": " & Chr(34) & "MicrosoftExcel" & Chr(34)) > 0 Then
        HttpTest = True
    End If

End Function


Private Function SimpleHttpTest() As Boolean

    '@Example: =SimpleHttp("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications","ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>","LEFT","<span") -> "1993"; See the examples in the HTTP() function, as this example has the same result as the example in the HTTP() function. You can see that this function is cleaner and easier to set up than the corresponding HTTP() function.

    If InStr(1, SimpleHttp("https://httpbin.org/get?hello=world"), "world") > 0 Then
        SimpleHttpTest = True
    End If

End Function


Private Function ParseHtmlStringTest() As Boolean

    '@Example: =ParseHtmlString("HTML String from the webpage: https://en.wikipedia.org/wiki/Visual_Basic_for_Applications","ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>","LEFT","<span") -> "1993"

    If ParseHtmlString("<div><p id='main'>Hello World</p></div>", Array("id", "main")) = "Hello World" Then
        ParseHtmlStringTest = True
    End If

End Function




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



Public Function AllXlibValidatorsTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not IsEmailTest() Then
        Debug.Print "Failed: IsEmailTest"
        TestStatus = False
    Else
        Debug.Print "Passed: IsEmailTest"
    End If
    
    If Not IsPhoneTest() Then
        Debug.Print "Failed: IsPhoneTest"
        TestStatus = False
    Else
        Debug.Print "Passed: IsPhoneTest"
    End If
    
    If Not IsCreditCardTest() Then
        Debug.Print "Failed: IsCreditCardTest"
        TestStatus = False
    Else
        Debug.Print "Passed: IsCreditCardTest"
    End If
    
    If Not IsUrlTest() Then
        Debug.Print "Failed: IsUrlTest"
        TestStatus = False
    Else
        Debug.Print "Passed: IsUrlTest"
    End If
    
    If Not IsIPFourTest() Then
        Debug.Print "Failed: IsIPFourTest"
        TestStatus = False
    Else
        Debug.Print "Passed: IsIPFourTest"
    End If
    
    If Not IsMacAddressTest() Then
        Debug.Print "Failed: IsMacAddressTest"
        TestStatus = False
    Else
        Debug.Print "Passed: IsMacAddressTest"
    End If
    
    If Not CreditCardNameTest() Then
        Debug.Print "Failed: CreditCardNameTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CreditCardNameTest"
    End If
    
    If Not FormatCreditCardTest() Then
        Debug.Print "Failed: FormatCreditCardTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FormatCreditCardTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibValidatorsTests = TestStatus
    
End Function



Private Function IsEmailTest() As Boolean

    '@Example: =IsEmail("JohnDoe@testmail.com") -> TRUE
    '@Example: =IsEmail("JohnDoe@test/mail.com") -> FALSE
    '@Example: =IsEmail("not_an_email_address") -> FALSE

    IsEmailTest = True

    IsEmailTest = IsEmailTest And IsEmail("JohnDoe@testmail.com") = True
    IsEmailTest = IsEmailTest And IsEmail("JohnDoe@test/mail.com") = False
    IsEmailTest = IsEmailTest And IsEmail("not_an_email_address") = False

End Function


Private Function IsPhoneTest() As Boolean

    '@Example: =IsPhone("123 456 7890") -> TRUE
    '@Example: =IsPhone("1234567890") -> TRUE
    '@Example: =IsPhone("1-234-567-890") -> FALSE; Not enough digits
    '@Example: =IsPhone("1-234-567-8905") -> TRUE
    '@Example: =IsPhone("+1-234-567-890") -> FALSE; Not enough digits
    '@Example: =IsPhone("+1-234-567-8905") -> TRUE
    '@Example: =IsPhone("+1-(234)-567-8905") -> TRUE
    '@Example: =IsPhone("+1 (234) 567 8905") -> TRUE
    '@Example: =IsPhone("1(234)5678905") -> TRUE
    '@Example: =IsPhone("123-456-789") -> FALSE; Not enough digits
    '@Example: =IsPhone("Hello World") -> FALSE; Not a phone number

    IsPhoneTest = True

    IsPhoneTest = IsPhoneTest And IsPhone("123 456 7890") = True
    IsPhoneTest = IsPhoneTest And IsPhone("1234567890") = True
    IsPhoneTest = IsPhoneTest And IsPhone("1-234-567-890") = False
    IsPhoneTest = IsPhoneTest And IsPhone("1-234-567-8905") = True
    IsPhoneTest = IsPhoneTest And IsPhone("+1-234-567-890") = False
    IsPhoneTest = IsPhoneTest And IsPhone("+1-234-567-8905") = True
    IsPhoneTest = IsPhoneTest And IsPhone("+1-(234)-567-8905") = True
    IsPhoneTest = IsPhoneTest And IsPhone("+1 (234) 567 8905") = True
    IsPhoneTest = IsPhoneTest And IsPhone("1(234)5678905") = True
    IsPhoneTest = IsPhoneTest And IsPhone("123-456-789") = False
    IsPhoneTest = IsPhoneTest And IsPhone("Hello World") = False

End Function


Private Function IsCreditCardTest() As Boolean

    '@Example: =IsCreditCard("5111567856785678") -> TRUE; This is a valid Mastercard number
    '@Example: =IsCreditCard("511156785678567") -> FALSE; Not enough digits
    '@Example: =IsCreditCard("9999999999999999") -> FALSE; Enough digits, but not a valid card number
    '@Example: =IsCreditCard("Hello World") -> FALSE

    IsCreditCardTest = True

    IsCreditCardTest = IsCreditCardTest And IsCreditCard("5111567856785678") = True
    IsCreditCardTest = IsCreditCardTest And IsCreditCard("511156785678567") = False
    IsCreditCardTest = IsCreditCardTest And IsCreditCard("9999999999999999") = False
    IsCreditCardTest = IsCreditCardTest And IsCreditCard("Hello World") = False

End Function


Private Function IsUrlTest() As Boolean

    '@Example: =IsUrl("https://www.wikipedia.org/") -> TRUE
    '@Example: =IsUrl("http://www.wikipedia.org/") -> TRUE
    '@Example: =IsUrl("hello_world") -> FALSE

    IsUrlTest = True

    IsUrlTest = IsUrlTest And IsUrl("https://www.wikipedia.org/") = True
    IsUrlTest = IsUrlTest And IsUrl("http://www.wikipedia.org/") = True
    IsUrlTest = IsUrlTest And IsUrl("hello_world") = False

End Function


Private Function IsIPFourTest() As Boolean

    '@Example: =IsIPFour("0.0.0.0") -> TRUE
    '@Example: =IsIPFour("100.100.100.100") -> TRUE
    '@Example: =IsIPFour("255.255.255.255") -> TRUE
    '@Example: =IsIPFour("255.255.255.256") -> FALSE; as the final 256 makes the address outside of the bounds of IPv4
    '@Example: =IsIPFour("0.0.0") -> FALSE; as the fourth octet is missing

    IsIPFourTest = True

    IsIPFourTest = IsIPFourTest And IsIPFour("0.0.0.0") = True
    IsIPFourTest = IsIPFourTest And IsIPFour("100.100.100.100") = True
    IsIPFourTest = IsIPFourTest And IsIPFour("255.255.255.255") = True
    IsIPFourTest = IsIPFourTest And IsIPFour("255.255.255.256") = False
    IsIPFourTest = IsIPFourTest And IsIPFour("0.0.0") = False

End Function


Private Function IsMacAddressTest() As Boolean

    '@Example: =IsMacAddress("00:25:96:12:34:56") -> TRUE
    '@Example: =IsMacAddress("FF:FF:FF:FF:FF:FF") -> TRUE
    '@Example: =IsMacAddress("00-25-96-12-34-56") -> TRUE
    '@Example: =IsMacAddress("123.789.abc.DEF") -> TRUE
    '@Example: =IsMacAddress("Not A Mac Address") -> FALSE
    '@Example: =IsMacAddress("FF:FF:FF:FF:FF:FH") -> FALSE; the H at the end is not a valid Hex number

    IsMacAddressTest = True

    IsMacAddressTest = IsMacAddressTest And IsMacAddress("00:25:96:12:34:56") = True
    IsMacAddressTest = IsMacAddressTest And IsMacAddress("FF:FF:FF:FF:FF:FF") = True
    IsMacAddressTest = IsMacAddressTest And IsMacAddress("00-25-96-12-34-56") = True
    IsMacAddressTest = IsMacAddressTest And IsMacAddress("123.789.abc.DEF") = True
    IsMacAddressTest = IsMacAddressTest And IsMacAddress("Not A Mac Address") = False
    IsMacAddressTest = IsMacAddressTest And IsMacAddress("FF:FF:FF:FF:FF:FH") = False

End Function


Private Function CreditCardNameTest() As Boolean

    '@Example: =CreditCardName("5111567856785678") -> "MasterCard"; This is a valid Mastercard number
    '@Example: =CreditCardName("not_a_card_number") -> #VALUE!

    CreditCardNameTest = True

    CreditCardNameTest = CreditCardNameTest And CreditCardName("5111567856785678") = "MasterCard"

End Function


Private Function FormatCreditCardTest() As Boolean

    '@Example: =FormatCreditCard("5111567856785678") -> "5111-5678-5678-5678"

    FormatCreditCardTest = True

    FormatCreditCardTest = FormatCreditCardTest And FormatCreditCard("5111567856785678") = "5111-5678-5678-5678"

End Function


