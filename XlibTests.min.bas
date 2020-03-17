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
CountUniqueTest = True
CountUniqueTest = CountUniqueTest And CountUnique(1, 2, 2, 3) = 3
CountUniqueTest = CountUniqueTest And CountUnique("a", "a", "a") = 1
CountUniqueTest = CountUniqueTest And CountUnique(Array(1, 2, 4, 4, 1)) = 3
End Function
Private Function SortTest() As Boolean
SortTest = True
SortTest = SortTest And Sort(Array(10, 20, 30))(0) = 10
SortTest = SortTest And Sort(Array(10, 20, 30))(1) = 20
SortTest = SortTest And Sort(Array(10, 20, 30))(2) = 30
SortTest = SortTest And Sort(Array(10, 20, 30), True)(0) = 30
SortTest = SortTest And Sort(Array(10, 20, 30), True)(1) = 20
SortTest = SortTest And Sort(Array(10, 20, 30), True)(2) = 10
End Function
Private Function ReverseTest() As Boolean
ReverseTest = True
ReverseTest = ReverseTest And Reverse(Array(10, 20, 30))(0) = 30
ReverseTest = ReverseTest And Reverse(Array(10, 20, 30))(1) = 20
ReverseTest = ReverseTest And Reverse(Array(10, 20, 30))(2) = 10
End Function
Private Function SumHighTest() As Boolean
SumHighTest = True
SumHighTest = SumHighTest And SumHigh(Array(1, 2, 3, 4), 2) = 7
SumHighTest = SumHighTest And SumHigh(Array(1, 2, 3, 4), 3) = 9
End Function
Private Function SumLowTest() As Boolean
SumLowTest = True
SumLowTest = SumLowTest And SumLow(Array(1, 2, 3, 4), 2) = 3
SumLowTest = SumLowTest And SumLow(Array(1, 2, 3, 4), 3) = 6
End Function
Private Function AverageHighTest() As Boolean
AverageHighTest = True
AverageHighTest = AverageHighTest And AverageHigh(Array(1, 2, 3, 4), 2) = 3.5
AverageHighTest = AverageHighTest And AverageHigh(Array(1, 2, 3, 4), 3) = 3
End Function
Private Function AverageLowTest() As Boolean
AverageLowTest = True
AverageLowTest = AverageLowTest And AverageLow(Array(1, 2, 3, 4), 2) = 1.5
AverageLowTest = AverageLowTest And AverageLow(Array(1, 2, 3, 4), 3) = 2
End Function
Private Function LargeTest() As Boolean
LargeTest = True
LargeTest = LargeTest And Large(Array(1, 2, 3, 4), 1) = 4
LargeTest = LargeTest And Large(Array(1, 2, 3, 4), 2) = 3
End Function
Private Function SmallTest() As Boolean
SmallTest = True
SmallTest = SmallTest And Small(Array(1, 2, 3, 4), 1) = 1
SmallTest = SmallTest And Small(Array(1, 2, 3, 4), 2) = 2
End Function
Private Function IsInArrayTest() As Boolean
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
Rgb2HexTest = True
Rgb2HexTest = Rgb2HexTest And Rgb2Hex(255, 255, 255) = "FFFFFF"
End Function
Private Function Hex2RgbTest() As Boolean
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
Hex2HslTest = True
Hex2HslTest = Hex2HslTest And Hex2Hsl("084080") = "(212.0, 88.2%, 26.7%)"
Hex2HslTest = Hex2HslTest And Hex2Hsl("#084080") = "(212.0, 88.2%, 26.7%)"
End Function
Private Function Hsl2RgbTest() As Boolean
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
Hsl2HexTest = True
Hsl2HexTest = Hsl2HexTest And Hsl2Hex(212, 0.882, 0.267) = "084080"
End Function
Private Function Rgb2HsvTest() As Boolean
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
QuarterTest = True
QuarterTest = QuarterTest And Quarter(4) = 2
QuarterTest = QuarterTest And Quarter(12) = 4
End Function
Private Function DaysOfMonthTest() As Boolean
DaysOfMonthTest = True
DaysOfMonthTest = DaysOfMonthTest And DaysOfMonth(1) = 31
DaysOfMonthTest = DaysOfMonthTest And DaysOfMonth(2, 2019) = 28
DaysOfMonthTest = DaysOfMonthTest And DaysOfMonth(2, 2020) = 29
End Function
Public Function AllXlibEnvironmentTests()
Dim TestStatus As Boolean
TestStatus = True
Debug.Print "========================================"
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
OSTest = True
#If Mac Then
OSTest = OSTest And OS() = "Mac"
#Else
OSTest = OSTest And OS() = "Windows"
#End If
End Function
Private Function UserNameTest() As Boolean
UserNameTest = True
#If Mac Then
UserNameTest = UserNameTest And UserName() = Environ("USER")
#Else
UserNameTest = UserNameTest And UserName() = Environ("USERNAME")
#End If
End Function
Private Function UserDomainTest() As Boolean
UserDomainTest = True
#If Mac Then
UserDomainTest = UserDomainTest And UserDomain() = Environ("HOST")
#Else
UserDomainTest = UserDomainTest And UserDomain() = Environ("USERDOMAIN")
#End If
End Function
Private Function ComputerNameTest() As Boolean
ComputerNameTest = True
ComputerNameTest = ComputerNameTest And ComputerName() = Environ("COMPUTERNAME")
End Function
Public Function AllXlibFileTests()
Dim TestStatus As Boolean
TestStatus = True
Debug.Print "========================================"
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
FileCreationTimeTest = False
If InStr(1, FileCreationTime(), " ") > 0 Then
If InStr(1, FileCreationTime(), ":") > 0 Then
FileCreationTimeTest = True
End If
End If
End Function
Private Function FileLastModifiedTimeTest() As Boolean
FileLastModifiedTimeTest = False
If InStr(1, FileLastModifiedTime(), " ") > 0 Then
If InStr(1, FileLastModifiedTime(), ":") > 0 Then
FileLastModifiedTimeTest = True
End If
End If
End Function
Private Function FileDriveTest() As Boolean
If InStr(1, FileDrive(), ":") > 0 Then
FileDriveTest = True
End If
End Function
Private Function FileNameTest() As Boolean
If Len(FileName()) > 0 Then
FileNameTest = True
End If
End Function
Private Function FileFolderTest() As Boolean
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
If FileSize() > 0 Then
FileSizeTest = True
End If
End Function
Private Function FileTypeTest() As Boolean
If Len(FileType()) > 0 Then
FileTypeTest = True
End If
End Function
Private Function FileExtensionTest() As Boolean
If Len(FileExtension()) > 0 Then
FileExtensionTest = True
End If
End Function
Private Function WriteFileTest() As Boolean
If WriteFile("TempTestFile.txt", "Hello World") Then
WriteFileTest = True
End If
End Function
Private Function ReadFileTest() As Boolean
ReadFileTest = IIf(ReadFile("TempTestFile.txt") = "Hello World", True, False)
Kill (GetActivePath() & "TempTestFile.txt")
End Function
Private Function PathSeparatorTest() As Boolean
If PathSeparator() = "\" Or PathSeparator() = "/" Then
PathSeparatorTest = True
End If
End Function
Private Function PathJoinTest() As Boolean
If PathJoin("hello", "world") = "hello/world" Or PathJoin("hello", "world") = "hello\world" Then
PathJoinTest = True
End If
End Function
Private Function CountFilesTest() As Boolean
If CountFiles() > 0 Then
CountFilesTest = True
End If
End Function
Private Function CountFilesAndFoldersTest() As Boolean
If CountFilesAndFolders() > 0 Then
CountFilesAndFoldersTest = True
End If
End Function
Private Function GetFileNameByNumberTest() As Boolean
If Len(GetFileNameByNumber()) > 0 Then
GetFileNameByNumberTest = True
End If
End Function
Public Function AllXlibMathTests()
Dim TestStatus As Boolean
TestStatus = True
Debug.Print "========================================"
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
CeilTest = True
CeilTest = CeilTest And Ceil(1.5) = 2
CeilTest = CeilTest And Ceil(1.0001) = 2
CeilTest = CeilTest And Ceil(1) = 1
End Function
Private Function FloorTest() As Boolean
FloorTest = True
FloorTest = FloorTest And Floor(1.9) = 1
FloorTest = FloorTest And Floor(1.1) = 1
FloorTest = FloorTest And Floor(1) = 1
End Function
Private Function InterpolateNumberTest() As Boolean
InterpolateNumberTest = True
InterpolateNumberTest = InterpolateNumberTest And InterpolateNumber(10, 20, 0.5) = 15
InterpolateNumberTest = InterpolateNumberTest And Round(InterpolateNumber(16, 124, 0.64), 2) = 85.12
End Function
Private Function InterpolatePercentTest() As Boolean
InterpolatePercentTest = True
InterpolatePercentTest = InterpolatePercentTest And InterpolatePercent(10, 18, 12) = 0.25
InterpolatePercentTest = InterpolatePercentTest And InterpolatePercent(10, 20, 15) = 0.5
End Function
Private Function MaxTest() As Boolean
MaxTest = True
MaxTest = MaxTest And Max(1, 2, 3) = 3
MaxTest = MaxTest And Max(4.4, 5, "6") = 6
MaxTest = MaxTest And Max(Array(1, 2.2, "3")) = 3
MaxTest = MaxTest And Max(Array(1, 2.2, "3"), Array(5, 15, -100), 10) = 15
End Function
Private Function MinTest() As Boolean
MinTest = True
MinTest = MinTest And Min(1, 2, 3) = 1
MinTest = MinTest And Min(4.4, 5, "6") = 4.4
MinTest = MinTest And Min(-1, -2, -3) = -3
MinTest = MinTest And Min(Array(1, 2.2, "3")) = 1
MinTest = MinTest And Min(Array(1, 2.2, "3"), Array(5, 15, -100), 10) = -100
End Function
Private Function ModFloatTest() As Boolean
ModFloatTest = True
ModFloatTest = ModFloatTest And Round(ModFloat(3.55, 2), 2) = 1.55
End Function
Public Function AllXlibMetaTests()
Dim TestStatus As Boolean
TestStatus = True
Debug.Print "========================================"
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
If IsNumeric(Split(XlibVersion(), ".")(0)) Then
If IsNumeric(Split(XlibVersion(), ".")(1)) Then
If IsNumeric(Split(XlibVersion(), ".")(2)) Then
XlibVersionTest = True
End If
End If
End If
End Function
Private Function XlibCreditsTest() As Boolean
If InStr(1, XlibCredits(), "XLib") > 0 Then
XlibCreditsTest = True
End If
End Function
Private Function XlibDocumentationTest() As Boolean
If Left(XlibDocumentation(), 4) = "http" Then
XlibDocumentationTest = True
End If
End Function
Public Function AllXlibNetworkTests()
Dim TestStatus As Boolean
TestStatus = True
Debug.Print "========================================"
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
If InStr(1, Http("https://httpbin.org/user-agent", "GET", Array("User-Agent", "MicrosoftExcel")), Chr(34) & "user-agent" & Chr(34) & ": " & Chr(34) & "MicrosoftExcel" & Chr(34)) > 0 Then
HttpTest = True
End If
End Function
Private Function SimpleHttpTest() As Boolean
If InStr(1, SimpleHttp("https://httpbin.org/get?hello=world"), "world") > 0 Then
SimpleHttpTest = True
End If
End Function
Private Function ParseHtmlStringTest() As Boolean
If ParseHtmlString("<div><p id='main'>Hello World</p></div>", Array("id", "main")) = "Hello World" Then
ParseHtmlStringTest = True
End If
End Function
Public Function AllXlibRandomTests()
Dim TestStatus As Boolean
TestStatus = True
Debug.Print "========================================"
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
Dim randomNumber%
randomNumber = RandBetween(1, 20)
RandBetweenTest = (randomNumber >= 1 And randomNumber <= 20)
End Function
Private Function BigRandBetweenTest() As Boolean
Dim randomNumber%
randomNumber = BigRandBetween(1, 20)
BigRandBetweenTest = (randomNumber >= 1 And randomNumber <= 20)
End Function
Private Function RandomSampleTest() As Boolean
Dim randomNumber%
randomNumber = RandomSample(Array(1, 2, 3))
RandomSampleTest = (randomNumber = 1 Or randomNumber = 2 Or randomNumber = 3)
End Function
Private Function RandomRangeTest() As Boolean
Dim randomNumber%
randomNumber = RandomRange(50, 60, 10)
RandomRangeTest = (randomNumber = 50 Or randomNumber = 60)
End Function
Private Function RandBoolTest() As Boolean
Dim randomBoolean As Boolean
randomBoolean = RandBool()
RandBoolTest = (randomBoolean = True Or randomBoolean = False)
End Function
Private Function RandBetweensTest() As Boolean
Dim randomNumber%
randomNumber = RandBetweens(1, 2, 51, 52)
RandBetweensTest = (randomNumber = 1 Or randomNumber = 2 Or randomNumber = 51 Or randomNumber = 52)
End Function
Public Function AllXlibRegexTests()
Dim TestStatus As Boolean
TestStatus = True
Debug.Print "========================================"
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
RegexSearchTest = True
RegexSearchTest = RegexSearchTest And RegexSearch("Hello World", "[a-z]{2}\s[W]") = "lo W"
End Function
Private Function RegexTestTest() As Boolean
RegexTestTest = True
RegexTestTest = RegexTestTest And RegexTest("Hello World", "[a-z]{2}\s[W]") = True
End Function
Private Function RegexReplaceTest() As Boolean
RegexReplaceTest = True
RegexReplaceTest = RegexReplaceTest And RegexReplace("Hello World", "[W][a-z]{4}", "VBA") = "Hello VBA"
End Function
Public Function AllXlibStringManipulationTests()
Dim TestStatus As Boolean
TestStatus = True
Debug.Print "========================================"
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
CapitalizeTest = True
CapitalizeTest = CapitalizeTest And Capitalize("hello World") = "Hello world"
End Function
Private Function LeftFindTest() As Boolean
LeftFindTest = True
LeftFindTest = LeftFindTest And LeftFind("Hello World", "r") = "Hello Wo"
End Function
Private Function RightFindTest() As Boolean
RightFindTest = True
RightFindTest = RightFindTest And RightFind("Hello World", "o") = "rld"
End Function
Private Function LeftSearchTest() As Boolean
LeftSearchTest = True
LeftSearchTest = LeftSearchTest And LeftSearch("Hello World", "r") = "Hello Wo"
LeftSearchTest = LeftSearchTest And LeftSearch("Hello World", "R") = "Hello Wo"
End Function
Private Function RightSearchTest() As Boolean
RightSearchTest = True
RightSearchTest = RightSearchTest And RightSearch("Hello World", "o") = "rld"
RightSearchTest = RightSearchTest And RightSearch("Hello World", "O") = "rld"
End Function
Private Function SubstrTest() As Boolean
SubstrTest = True
SubstrTest = SubstrTest And Substr("Hello World", 2, 6) = "ello"
End Function
Private Function SubstrFindTest() As Boolean
SubstrFindTest = True
SubstrFindTest = SubstrFindTest And SubstrFind("Hello World", "e", "o") = "ello Wo"
SubstrFindTest = SubstrFindTest And SubstrFind("Hello World", "e", "o", True) = "llo W"
SubstrFindTest = SubstrFindTest And SubstrFind("One Two Three", "ne ", " Thr") = "ne Two Thr"
SubstrFindTest = SubstrFindTest And SubstrFind("One Two Three", "ne ", " Thr", True) = "Two"
End Function
Private Function SubstrSearchTest() As Boolean
SubstrSearchTest = True
SubstrSearchTest = SubstrSearchTest And SubstrSearch("Hello World", "e", "o") = "ello Wo"
SubstrSearchTest = SubstrSearchTest And SubstrSearch("Hello World", "e", "o", True) = "llo W"
SubstrSearchTest = SubstrSearchTest And SubstrSearch("One Two Three", "ne ", " Thr") = "ne Two Thr"
SubstrSearchTest = SubstrSearchTest And SubstrSearch("One Two Three", "NE ", " THR") = "ne Two Thr"
SubstrSearchTest = SubstrSearchTest And SubstrSearch("One Two Three", "ne ", " Thr", True) = "Two"
End Function
Private Function RepeatTest() As Boolean
RepeatTest = True
RepeatTest = RepeatTest And Repeat("Hello", 2) = "HelloHello"
RepeatTest = RepeatTest And Repeat("=", 10) = "=========="
End Function
Private Function FormatterTest() As Boolean
FormatterTest = True
FormatterTest = FormatterTest And Formatter("Hello {1}", "World") = "Hello World"
FormatterTest = FormatterTest And Formatter("{1} {2}", "Hello", "World") = "Hello World"
FormatterTest = FormatterTest And Formatter("{1}.{2}@{3}", "FirstName", "LastName", "email.com") = "FirstName.LastName@email.com"
FormatterTest = FormatterTest And Formatter("{1}.{2}@{3}", Array("FirstName", "LastName", "email.com")) = "FirstName.LastName@email.com"
End Function
Private Function ZfillTest() As Boolean
ZfillTest = True
ZfillTest = ZfillTest And Zfill(123, 5) = "00123"
ZfillTest = ZfillTest And Zfill(5678, 5) = "05678"
ZfillTest = ZfillTest And Zfill(12345678, 5) = "12345678"
ZfillTest = ZfillTest And Zfill(123, 5, "X") = "XX123"
ZfillTest = ZfillTest And Zfill(123, 5, "X", True) = "123XX"
End Function
Private Function SplitTextTest() As Boolean
SplitTextTest = True
SplitTextTest = SplitTextTest And SplitText("Hello World", 1) = "Hello"
SplitTextTest = SplitTextTest And SplitText("Hello World", 2) = "World"
SplitTextTest = SplitTextTest And SplitText("One Two Three", 2) = "Two"
SplitTextTest = SplitTextTest And SplitText("One-Two-Three", 2, "-") = "Two"
End Function
Private Function CountWordsTest() As Boolean
CountWordsTest = True
CountWordsTest = CountWordsTest And CountWords("Hello World") = 2
CountWordsTest = CountWordsTest And CountWords("One Two Three") = 3
CountWordsTest = CountWordsTest And CountWords("One-Two-Three", "-") = 3
End Function
Private Function CamelCaseTest() As Boolean
CamelCaseTest = True
CamelCaseTest = CamelCaseTest And CamelCase("Hello World") = "helloWorld"
CamelCaseTest = CamelCaseTest And CamelCase("One Two Three") = "oneTwoThree"
End Function
Private Function KebabCaseTest() As Boolean
KebabCaseTest = True
KebabCaseTest = KebabCaseTest And KebabCase("Hello World") = "hello-world"
KebabCaseTest = KebabCaseTest And KebabCase("One Two Three") = "one-two-three"
End Function
Private Function RemoveCharactersTest() As Boolean
RemoveCharactersTest = True
RemoveCharactersTest = RemoveCharactersTest And RemoveCharacters("Hello World", "l") = "Heo Word"
RemoveCharactersTest = RemoveCharactersTest And RemoveCharacters("Hello World", "lo") = "He Wrd"
RemoveCharactersTest = RemoveCharactersTest And RemoveCharacters("Hello World", "l", "o") = "He Wrd"
RemoveCharactersTest = RemoveCharactersTest And RemoveCharacters("Hello World", "lod") = "He Wr"
RemoveCharactersTest = RemoveCharactersTest And RemoveCharacters("Two Three Four", "f", "t") = "Two Three Four"
RemoveCharactersTest = RemoveCharactersTest And RemoveCharacters("Two Three Four", "F", "T") = "wo hree our"
End Function
Private Function CompanyCaseTest() As Boolean
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
ReverseTextTest = True
ReverseTextTest = ReverseTextTest And ReverseText("Hello World") = "dlroW olleH"
End Function
Private Function ReverseWordsTest() As Boolean
ReverseWordsTest = True
ReverseWordsTest = ReverseWordsTest And ReverseWords("Hello World") = "World Hello"
ReverseWordsTest = ReverseWordsTest And ReverseWords("One Two Three") = "Three Two One"
ReverseWordsTest = ReverseWordsTest And ReverseWords("One-Two-Three", "-") = "Three-Two-One"
End Function
Private Function IndentTextTest() As Boolean
IndentTextTest = True
IndentTextTest = IndentTextTest And IndentText("Hello") = "    Hello"
IndentTextTest = IndentTextTest And IndentText("Hello", 4) = "    Hello"
IndentTextTest = IndentTextTest And IndentText("Hello", 3) = "   Hello"
IndentTextTest = IndentTextTest And IndentText("Hello", 2) = "  Hello"
IndentTextTest = IndentTextTest And IndentText("Hello", 1) = " Hello"
End Function
Private Function DedentTextTest() As Boolean
DedentTextTest = True
DedentTextTest = DedentTextTest And DedentText("    Hello") = "Hello"
End Function
Private Function ShortenTextTest() As Boolean
ShortenTextTest = True
ShortenTextTest = ShortenTextTest And ShortenText("Hello World One Two Three", 20) = "Hello World [...]"
ShortenTextTest = ShortenTextTest And ShortenText("Hello World One Two Three", 15) = "Hello [...]"
ShortenTextTest = ShortenTextTest And ShortenText("Hello World One Two Three") = "Hello World One Two Three"
ShortenTextTest = ShortenTextTest And ShortenText("Hello World One Two Three", 15, "-->") = "Hello World -->"
ShortenTextTest = ShortenTextTest And ShortenText("Hello_World_One_Two_Three", 15, "-->", "_") = "Hello_World_-->"
End Function
Private Function InSplitTest() As Boolean
InSplitTest = True
InSplitTest = InSplitTest And InSplit("Hello", "Hello World One Two Three") = True
InSplitTest = InSplitTest And InSplit("NotInString", "Hello World One Two Three") = False
InSplitTest = InSplitTest And InSplit("Hello", "Hello-World-One-Two-Three", "-") = True
End Function
Private Function EliteCaseTest() As Boolean
EliteCaseTest = True
EliteCaseTest = EliteCaseTest And EliteCase("Hello World") = "H3110 W0r1d"
End Function
Private Function ScrambleCaseTest() As Boolean
ScrambleCaseTest = True
Dim testString$
testString = "a"
ScrambleCaseTest = ScrambleCaseTest And (testString = "a" Or testString = "A")
End Function
Private Function LeftSplitTest() As Boolean
LeftSplitTest = True
LeftSplitTest = LeftSplitTest And LeftSplit("Hello World One Two Three", 1) = "Hello"
LeftSplitTest = LeftSplitTest And LeftSplit("Hello World One Two Three", 2) = "Hello World"
LeftSplitTest = LeftSplitTest And LeftSplit("Hello World One Two Three", 3) = "Hello World One"
LeftSplitTest = LeftSplitTest And LeftSplit("Hello World One Two Three", 10) = "Hello World One Two Three"
LeftSplitTest = LeftSplitTest And LeftSplit("Hello-World-One-Two-Three", 2, "-") = "Hello-World"
End Function
Private Function RightSplitTest() As Boolean
RightSplitTest = True
RightSplitTest = RightSplitTest And RightSplit("Hello World One Two Three", 1) = "Three"
RightSplitTest = RightSplitTest And RightSplit("Hello World One Two Three", 2) = "Two Three"
RightSplitTest = RightSplitTest And RightSplit("Hello World One Two Three", 3) = "One Two Three"
RightSplitTest = RightSplitTest And RightSplit("Hello World One Two Three", 10) = "Hello World One Two Three"
RightSplitTest = RightSplitTest And RightSplit("Hello-World-One-Two-Three", 2, "-") = "Two-Three"
End Function
Private Function TrimCharTest() As Boolean
TrimCharTest = True
TrimCharTest = TrimCharTest And TrimChar("   Hello World   ") = "Hello World"
TrimCharTest = TrimCharTest And TrimChar("---Hello World---", "-") = "Hello World"
End Function
Private Function TrimLeftTest() As Boolean
TrimLeftTest = True
TrimLeftTest = TrimLeftTest And TrimLeft("   Hello World   ") = "Hello World   "
TrimLeftTest = TrimLeftTest And TrimLeft("---Hello World---", "-") = "Hello World---"
End Function
Private Function TrimRightTest() As Boolean
TrimRightTest = True
TrimRightTest = TrimRightTest And TrimRight("   Hello World   ") = "   Hello World"
TrimRightTest = TrimRightTest And TrimRight("---Hello World---", "-") = "---Hello World"
End Function
Private Function CountUppercaseCharactersTest() As Boolean
CountUppercaseCharactersTest = True
CountUppercaseCharactersTest = CountUppercaseCharactersTest And CountUppercaseCharacters("Hello World") = 2
End Function
Private Function CountLowercaseCharactersTest() As Boolean
CountLowercaseCharactersTest = True
CountLowercaseCharactersTest = CountLowercaseCharactersTest And CountLowercaseCharacters("Hello World") = 8
End Function
Private Function TextJoinTest() As Boolean
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
HammingTest = True
HammingTest = HammingTest And Hamming("Cat", "Bat") = 1
HammingTest = HammingTest And Hamming("Cat", "Bag") = 2
HammingTest = HammingTest And Hamming("Cat", "Dog") = 3
End Function
Private Function LevenshteinTest() As Boolean
LevenshteinTest = True
LevenshteinTest = LevenshteinTest And Levenshtein("Cat", "Bat") = 1
LevenshteinTest = LevenshteinTest And Levenshtein("Cat", "Ca") = 1
LevenshteinTest = LevenshteinTest And Levenshtein("Cat", "Cta") = 2
End Function
Private Function DamerauTest() As Boolean
DamerauTest = True
DamerauTest = DamerauTest And Damerau("Cat", "Bat") = 1
DamerauTest = DamerauTest And Damerau("Cat", "Ca") = 1
DamerauTest = DamerauTest And Damerau("Cat", "Cta") = 1
End Function
Public Function AllXlibUtilitiesTests()
Dim TestStatus As Boolean
TestStatus = True
Debug.Print "========================================"
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
JsonifyTest = True
JsonifyTest = JsonifyTest And Jsonify(0, "Hello", "World", "1", "2", 3, 4.5) = "[" & Chr(34) & "Hello" & Chr(34) & "," & Chr(34) & "World" & Chr(34) & ",1,2,3,4.5]"
JsonifyTest = JsonifyTest And Jsonify(0, Array("Hello", "World", "1", "2", 3, 4.5)) = "[" & Chr(34) & "Hello" & Chr(34) & "," & Chr(34) & "World" & Chr(34) & ",1,2,3,4.5]"
End Function
Private Function UuidFourTest() As Boolean
Dim uuidGroups
uuidGroups = Split(UuidFour(), "-")
If Len(uuidGroups(0)) = 8 And Len(uuidGroups(1)) = 4 And Len(uuidGroups(2)) = 4 And Len(uuidGroups(3)) = 4 And Len(uuidGroups(4)) = 12 Then
UuidFourTest = True
End If
End Function
Private Function HideTextTest() As Boolean
HideTextTest = True
HideTextTest = HideTextTest And HideText("Hello World", False) = "Hello World"
HideTextTest = HideTextTest And HideText("Hello World", True) = "********"
HideTextTest = HideTextTest And HideText("Hello World", True, "[Hidden Text]") = "[Hidden Text]"
End Function
Private Function JavaScriptTest() As Boolean
JavaScriptTest = True
JavaScriptTest = JavaScriptTest And JavaScript("function helloFunc(){return 'Hello World!'}", "helloFunc") = "Hello World!"
JavaScriptTest = JavaScriptTest And JavaScript("function addTwo(a, b){return a + b}", "addTwo", 12, 24) = 36
End Function
Private Function HtmlEscapeTest() As Boolean
HtmlEscapeTest = True
HtmlEscapeTest = HtmlEscapeTest And HtmlEscape("<p>Hello World</p>") = "&lt;p&gt;Hello World&lt;/p&gt;"
End Function
Private Function HtmlUnescapeTest() As Boolean
HtmlUnescapeTest = True
HtmlUnescapeTest = HtmlUnescapeTest And HtmlUnescape("&lt;p&gt;Hello World&lt;/p&gt;") = "<p>Hello World</p>"
End Function
Private Function SpeakTextTest() As Boolean
SpeakTextTest = True
SpeakTextTest = SpeakTextTest And SpeakText("Hello", "World") = "Hello World"
End Function
Private Function Dec2HexTest() As Boolean
Dec2HexTest = True
Dec2HexTest = Dec2HexTest And Dec2Hex(5) = "5"
Dec2HexTest = Dec2HexTest And Dec2Hex(5, 2) = "05"
Dec2HexTest = Dec2HexTest And Dec2Hex(255, 2) = "FF"
Dec2HexTest = Dec2HexTest And Dec2Hex(255, 8) = "000000FF"
End Function
Private Function BigDec2HexTest() As Boolean
BigDec2HexTest = True
BigDec2HexTest = BigDec2HexTest And BigDec2Hex(3000000000#, 16) = "00000000B2D05E00"
End Function
Private Function BigHexTest() As Boolean
BigHexTest = True
BigHexTest = BigHexTest And BigHex(3000000000#) = "B2D05E00"
End Function
Private Function Hex2DecTest() As Boolean
Hex2DecTest = True
Hex2DecTest = Hex2DecTest And Hex2Dec("FF") = 255
Hex2DecTest = Hex2DecTest And Hex2Dec("FFFF") = 65535
End Function
Private Function Len2Test() As Boolean
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
IsEmailTest = True
IsEmailTest = IsEmailTest And IsEmail("JohnDoe@testmail.com") = True
IsEmailTest = IsEmailTest And IsEmail("JohnDoe@test/mail.com") = False
IsEmailTest = IsEmailTest And IsEmail("not_an_email_address") = False
End Function
Private Function IsPhoneTest() As Boolean
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
IsCreditCardTest = True
IsCreditCardTest = IsCreditCardTest And IsCreditCard("5111567856785678") = True
IsCreditCardTest = IsCreditCardTest And IsCreditCard("511156785678567") = False
IsCreditCardTest = IsCreditCardTest And IsCreditCard("9999999999999999") = False
IsCreditCardTest = IsCreditCardTest And IsCreditCard("Hello World") = False
End Function
Private Function IsUrlTest() As Boolean
IsUrlTest = True
IsUrlTest = IsUrlTest And IsUrl("https://www.wikipedia.org/") = True
IsUrlTest = IsUrlTest And IsUrl("http://www.wikipedia.org/") = True
IsUrlTest = IsUrlTest And IsUrl("hello_world") = False
End Function
Private Function IsIPFourTest() As Boolean
IsIPFourTest = True
IsIPFourTest = IsIPFourTest And IsIPFour("0.0.0.0") = True
IsIPFourTest = IsIPFourTest And IsIPFour("100.100.100.100") = True
IsIPFourTest = IsIPFourTest And IsIPFour("255.255.255.255") = True
IsIPFourTest = IsIPFourTest And IsIPFour("255.255.255.256") = False
IsIPFourTest = IsIPFourTest And IsIPFour("0.0.0") = False
End Function
Private Function IsMacAddressTest() As Boolean
IsMacAddressTest = True
IsMacAddressTest = IsMacAddressTest And IsMacAddress("00:25:96:12:34:56") = True
IsMacAddressTest = IsMacAddressTest And IsMacAddress("FF:FF:FF:FF:FF:FF") = True
IsMacAddressTest = IsMacAddressTest And IsMacAddress("00-25-96-12-34-56") = True
IsMacAddressTest = IsMacAddressTest And IsMacAddress("123.789.abc.DEF") = True
IsMacAddressTest = IsMacAddressTest And IsMacAddress("Not A Mac Address") = False
IsMacAddressTest = IsMacAddressTest And IsMacAddress("FF:FF:FF:FF:FF:FH") = False
End Function
Private Function CreditCardNameTest() As Boolean
CreditCardNameTest = True
CreditCardNameTest = CreditCardNameTest And CreditCardName("5111567856785678") = "MasterCard"
End Function
Private Function FormatCreditCardTest() As Boolean
FormatCreditCardTest = True
FormatCreditCardTest = FormatCreditCardTest And FormatCreditCard("5111567856785678") = "5111-5678-5678-5678"
End Function
