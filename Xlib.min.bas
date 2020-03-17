Attribute VB_Name = "Xlib"
'The MIT License (MIT)
'Copyright © 2020 Anthony Mancini
'Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the “Software”), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:
'The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.
'THE SOFTWARE IS PROVIDED “AS IS”, WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
Option Explicit
Option Private Module
Public Function CountUnique(ParamArray array1()) As Integer
Dim individualElement
Dim individualValue
Dim uniqueDictionary As Object
Dim uniqueCount%
Set uniqueDictionary = CreateObject("Scripting.Dictionary")
For Each individualElement In array1
If IsArray(individualElement) Then
For Each individualValue In individualElement
If Not uniqueDictionary.exists(individualValue) Then
uniqueDictionary.Add individualValue, 0
uniqueCount = uniqueCount + 1
End If
Next
Else
If Not uniqueDictionary.exists(individualElement) Then
uniqueDictionary.Add individualElement, 0
uniqueCount = uniqueCount + 1
End If
End If
Next
CountUnique = uniqueCount
End Function
Public Function Sort(ByVal sortableArray, Optional ByVal descendingFlag As Boolean)
Dim i%
Dim swapOccuredBool As Boolean
Dim arrayLength%
arrayLength = UBound(sortableArray) - LBound(sortableArray) + 1
Dim sortedArray()
ReDim sortedArray(arrayLength)
For i = 0 To arrayLength - 1
sortedArray(i) = sortableArray(i)
Next
Dim temporaryValue
Do
swapOccuredBool = False
For i = 0 To arrayLength - 1
If (sortedArray(i)) < sortedArray(i + 1) Then
temporaryValue = sortedArray(i)
sortedArray(i) = sortedArray(i + 1)
sortedArray(i + 1) = temporaryValue
swapOccuredBool = True
End If
Next
Loop While swapOccuredBool
If descendingFlag = True Then
Sort = sortedArray
Else
Dim ascendingArray()
ReDim ascendingArray(arrayLength)
For i = 0 To arrayLength - 1
ascendingArray(i) = sortedArray(arrayLength - i - 1)
Next
Sort = ascendingArray
End If
End Function
Public Function Reverse(ByVal array1)
Dim i%
Dim arrayLength%
Dim reversedArray()
arrayLength = UBound(array1) - LBound(array1)
ReDim reversedArray(arrayLength)
For i = LBound(array1) To UBound(array1)
reversedArray(arrayLength - i) = array1(i)
Next
Reverse = reversedArray
End Function
Public Function SumHigh(ByVal array1, ByVal numberSummed As Integer)
Dim i%
Dim sumValue#
For i = 1 To numberSummed
sumValue = sumValue + Large(array1, i)
Next
SumHigh = sumValue
End Function
Public Function SumLow(ByVal array1, ByVal numberSummed As Integer)
Dim i%
Dim sumValue#
For i = 1 To numberSummed
sumValue = sumValue + Small(array1, i)
Next
SumLow = sumValue
End Function
Public Function AverageHigh(ByVal array1, ByVal numberAveraged As Integer)
Dim i%
Dim sumValue#
For i = 1 To numberAveraged
sumValue = sumValue + Large(array1, i)
Next
AverageHigh = sumValue / numberAveraged
End Function
Public Function AverageLow(ByVal array1, ByVal numberAveraged As Integer)
Dim i%
Dim sumValue#
For i = 1 To numberAveraged
sumValue = sumValue + Small(array1, i)
Next
AverageLow = sumValue / numberAveraged
End Function
Public Function Large(ByVal array1, ByVal nthNumber As Integer)
Large = Sort(array1)(UBound(array1) - (nthNumber - 1))
End Function
Public Function Small(ByVal array1, ByVal nthNumber As Integer)
Small = Sort(array1, True)(UBound(array1) - (nthNumber - 1))
End Function
Public Function IsInArray(ByVal value1, ByVal array1) As Boolean
Dim individualElement
For Each individualElement In array1
If individualElement = value1 Then
IsInArray = True
Exit Function
End If
Next
IsInArray = False
End Function
Public Function Rgb2Hex(ByVal redColorInteger As Integer, ByVal greenColorInteger As Integer, ByVal blueColorInteger As Integer) As String
Rgb2Hex = Dec2Hex(redColorInteger, 2) & Dec2Hex(greenColorInteger, 2) & Dec2Hex(blueColorInteger, 2)
End Function
Public Function Hex2Rgb(ByVal hexColorString As String, Optional ByVal singleColorNumberOrName = -1)
hexColorString = Replace(hexColorString, "#", "")
If singleColorNumberOrName = 0 Or singleColorNumberOrName = "Red" Then
Hex2Rgb = Hex2Dec(Left(hexColorString, 2))
ElseIf singleColorNumberOrName = 1 Or singleColorNumberOrName = "Green" Then
Hex2Rgb = Hex2Dec(Mid(hexColorString, 3, 2))
ElseIf singleColorNumberOrName = 2 Or singleColorNumberOrName = "Blue" Then
Hex2Rgb = Hex2Dec(Right(hexColorString, 2))
Else
Hex2Rgb = "(" & Hex2Dec(Left(hexColorString, 2)) & ", " & Hex2Dec(Mid(hexColorString, 3, 2)) & ", " & Hex2Dec(Right(hexColorString, 2)) & ")"
End If
End Function
Public Function Rgb2Hsl(ByVal redColorInteger As Integer, ByVal greenColorInteger As Integer, ByVal blueColorInteger As Integer, Optional ByVal singleColorNumberOrName = -1)
Dim redPrime#
Dim greenPrime#
Dim bluePrime#
redPrime = redColorInteger / 255
greenPrime = greenColorInteger / 255
bluePrime = blueColorInteger / 255
Dim colorMax#
Dim colorMin#
colorMax = Max(redPrime, greenPrime, bluePrime)
colorMin = Min(redPrime, greenPrime, bluePrime)
Dim deltaValue#
deltaValue = colorMax - colorMin
Dim hueValue#
Dim saturationValue#
Dim lightnessValue#
If deltaValue = 0 Then
hueValue = 0
Else
Select Case colorMax
Case redPrime
hueValue = 60 * (((greenPrime - bluePrime) / deltaValue) Mod 6)
Case greenPrime
hueValue = 60 * (((bluePrime - redPrime) / deltaValue) + 2)
Case bluePrime
hueValue = 60 * (((redPrime - greenPrime) / deltaValue) + 4)
End Select
End If
lightnessValue = (colorMax + colorMin) / 2
If deltaValue = 0 Then
saturationValue = 0
Else
saturationValue = deltaValue / (1 - Abs((2 * lightnessValue - 1)))
End If
If singleColorNumberOrName = 0 Or singleColorNumberOrName = "Hue" Then
Rgb2Hsl = hueValue
ElseIf singleColorNumberOrName = 1 Or singleColorNumberOrName = "Saturation" Then
Rgb2Hsl = saturationValue
ElseIf singleColorNumberOrName = 2 Or singleColorNumberOrName = "Lightness" Then
Rgb2Hsl = lightnessValue
Else
Rgb2Hsl = "(" & Format(hueValue, "#.0") & ", " & Format(saturationValue * 100, "#.0") & "%, " & Format(lightnessValue * 100, "#.0") & "%)"
End If
End Function
Public Function Hex2Hsl(ByVal hexColorString As String) As String
hexColorString = Replace(hexColorString, "#", "")
Dim redValue%
Dim greenValue%
Dim blueValue%
redValue = CInt(Hex2Dec(Left(hexColorString, 2)))
greenValue = CInt(Hex2Dec(Mid(hexColorString, 3, 2)))
blueValue = CInt(Hex2Dec(Right(hexColorString, 2)))
Hex2Hsl = Rgb2Hsl(redValue, greenValue, blueValue)
End Function
Public Function Hsl2Rgb(ByVal hueValue As Double, ByVal saturationValue As Double, ByVal lightnessValue As Double, Optional ByVal singleColorNumberOrName = -1)
Dim cValue#
Dim xValue#
Dim mValue#
cValue = (1 - Abs(2 * lightnessValue - 1)) * saturationValue
xValue = cValue * (1 - Abs(ModFloat((hueValue / 60), 2) - 1))
mValue = lightnessValue - cValue / 2
Dim redValue#
Dim greenValue#
Dim blueValue#
If hueValue >= 0 And hueValue < 60 Then
redValue = cValue
greenValue = xValue
blueValue = 0
ElseIf hueValue >= 60 And hueValue < 120 Then
redValue = xValue
greenValue = cValue
blueValue = 0
ElseIf hueValue >= 120 And hueValue < 180 Then
redValue = 0
greenValue = cValue
blueValue = xValue
ElseIf hueValue >= 180 And hueValue < 240 Then
redValue = 0
greenValue = xValue
blueValue = cValue
ElseIf hueValue >= 240 And hueValue < 300 Then
redValue = xValue
greenValue = 0
blueValue = cValue
ElseIf hueValue >= 300 And hueValue < 360 Then
redValue = cValue
greenValue = 0
blueValue = xValue
End If
redValue = (redValue + mValue) * 255
greenValue = (greenValue + mValue) * 255
blueValue = (blueValue + mValue) * 255
If singleColorNumberOrName = 0 Or singleColorNumberOrName = "Red" Then
Hsl2Rgb = Round(redValue, 0)
ElseIf singleColorNumberOrName = 1 Or singleColorNumberOrName = "Green" Then
Hsl2Rgb = Round(greenValue, 0)
ElseIf singleColorNumberOrName = 2 Or singleColorNumberOrName = "Blue" Then
Hsl2Rgb = Round(blueValue, 0)
Else
Hsl2Rgb = "(" & Round(redValue, 0) & ", " & Round(greenValue, 0) & ", " & Round(blueValue, 0) & ")"
End If
End Function
Public Function Hsl2Hex(ByVal hueValue As Double, ByVal saturationValue As Double, ByVal lightnessValue As Double)
Dim redValue%
Dim greenValue%
Dim blueValue%
redValue = Hsl2Rgb(hueValue, saturationValue, lightnessValue, 0)
greenValue = Hsl2Rgb(hueValue, saturationValue, lightnessValue, 1)
blueValue = Hsl2Rgb(hueValue, saturationValue, lightnessValue, 2)
Hsl2Hex = Rgb2Hex(redValue, greenValue, blueValue)
End Function
Public Function Rgb2Hsv(ByVal redColorInteger As Integer, ByVal greenColorInteger As Integer, ByVal blueColorInteger As Integer, Optional ByVal singleColorNumberOrName = -1)
Dim redPrime#
Dim greenPrime#
Dim bluePrime#
redPrime = redColorInteger / 255
greenPrime = greenColorInteger / 255
bluePrime = blueColorInteger / 255
Dim colorMax#
Dim colorMin#
colorMax = Max(redPrime, greenPrime, bluePrime)
colorMin = Min(redPrime, greenPrime, bluePrime)
Dim deltaValue#
deltaValue = colorMax - colorMin
Dim hueValue#
Dim saturationValue#
Dim valueValue#
If deltaValue = 0 Then
hueValue = 0
Else
Select Case colorMax
Case redPrime
hueValue = 60 * (((greenPrime - bluePrime) / deltaValue) Mod 6)
Case greenPrime
hueValue = 60 * (((bluePrime - redPrime) / deltaValue) + 2)
Case bluePrime
hueValue = 60 * (((redPrime - greenPrime) / deltaValue) + 4)
End Select
End If
If colorMax = 0 Then
saturationValue = 0
Else
saturationValue = deltaValue / colorMax
End If
valueValue = colorMax
If singleColorNumberOrName = 0 Or singleColorNumberOrName = "Hue" Then
Rgb2Hsv = hueValue
ElseIf singleColorNumberOrName = 1 Or singleColorNumberOrName = "Saturation" Then
Rgb2Hsv = saturationValue
ElseIf singleColorNumberOrName = 2 Or singleColorNumberOrName = "Value" Then
Rgb2Hsv = valueValue
Else
Rgb2Hsv = "(" & Format(hueValue, "#.0") & ", " & Format(saturationValue * 100, "#.0") & "%, " & Format(valueValue * 100, "#.0") & "%)"
End If
End Function
Public Function WeekdayName2(Optional ByVal dayNumber As Byte) As String
If dayNumber = 0 Then
WeekdayName2 = WeekdayName(Weekday(Now()))
Else
WeekdayName2 = WeekdayName(dayNumber)
End If
End Function
Public Function MonthName2(Optional ByVal monthNumber As Byte) As String
If monthNumber = 0 Then
MonthName2 = MonthName(Month(Now()))
Else
MonthName2 = MonthName(monthNumber)
End If
End Function
Public Function Quarter(Optional ByVal monthNumberOrName) As Byte
If IsMissing(monthNumberOrName) Then
monthNumberOrName = MonthName(Month(Now()))
End If
If IsNumeric(monthNumberOrName) Then
monthNumberOrName = MonthName(monthNumberOrName)
End If
If monthNumberOrName = MonthName(1) Or monthNumberOrName = MonthName(2) Or monthNumberOrName = MonthName(3) Then
Quarter = 1
End If
If monthNumberOrName = MonthName(4) Or monthNumberOrName = MonthName(5) Or monthNumberOrName = MonthName(6) Then
Quarter = 2
End If
If monthNumberOrName = MonthName(7) Or monthNumberOrName = MonthName(8) Or monthNumberOrName = MonthName(9) Then
Quarter = 3
End If
If monthNumberOrName = MonthName(10) Or monthNumberOrName = MonthName(11) Or monthNumberOrName = MonthName(12) Then
Quarter = 4
End If
End Function
Public Function TimeConverter(ByVal date1 As Date, Optional ByVal secondsInteger As Integer, Optional ByVal minutesInteger As Integer, Optional ByVal hoursInteger As Integer, Optional ByVal daysInteger As Integer, Optional ByVal monthsInteger As Integer, Optional ByVal yearsInteger As Integer) As Date
secondsInteger = Second(date1) + secondsInteger
minutesInteger = Minute(date1) + minutesInteger
hoursInteger = Hour(date1) + hoursInteger
daysInteger = Day(date1) + daysInteger
monthsInteger = Month(date1) + monthsInteger
yearsInteger = Year(date1) + yearsInteger
TimeConverter = DateSerial(yearsInteger, monthsInteger, daysInteger) + TimeSerial(hoursInteger, minutesInteger, secondsInteger)
End Function
Public Function DaysOfMonth(Optional ByVal monthNumberOrName, Optional ByVal yearNumber As Integer)
If IsMissing(monthNumberOrName) Then
monthNumberOrName = Month(Now())
End If
If yearNumber = 0 Then
yearNumber = Year(Now())
End If
If monthNumberOrName = 1 Or monthNumberOrName = MonthName(1) Then
DaysOfMonth = 31
ElseIf monthNumberOrName = 2 Or monthNumberOrName = MonthName(2) Then
If yearNumber Mod 4 <> 0 Then
DaysOfMonth = 28
Else
DaysOfMonth = 29
End If
ElseIf monthNumberOrName = 3 Or monthNumberOrName = MonthName(3) Then
DaysOfMonth = 31
ElseIf monthNumberOrName = 4 Or monthNumberOrName = MonthName(4) Then
DaysOfMonth = 30
ElseIf monthNumberOrName = 5 Or monthNumberOrName = MonthName(5) Then
DaysOfMonth = 31
ElseIf monthNumberOrName = 6 Or monthNumberOrName = MonthName(6) Then
DaysOfMonth = 30
ElseIf monthNumberOrName = 7 Or monthNumberOrName = MonthName(7) Then
DaysOfMonth = 31
ElseIf monthNumberOrName = 8 Or monthNumberOrName = MonthName(8) Then
DaysOfMonth = 31
ElseIf monthNumberOrName = 9 Or monthNumberOrName = MonthName(9) Then
DaysOfMonth = 30
ElseIf monthNumberOrName = 10 Or monthNumberOrName = MonthName(10) Then
DaysOfMonth = 31
ElseIf monthNumberOrName = 11 Or monthNumberOrName = MonthName(11) Then
DaysOfMonth = 30
ElseIf monthNumberOrName = 12 Or monthNumberOrName = MonthName(12) Then
DaysOfMonth = 31
Else
DaysOfMonth = "#NotAValidMonthNumberOrName"
End If
End Function
Public Function WeekOfMonth(Optional ByVal date1 As Date) As Byte
Dim weekNumber As Byte
Dim currentDay As Byte
Dim currentWeekday As Byte
weekNumber = 1
If Year(date1) = 1899 Then
currentDay = Day(Now())
currentWeekday = Weekday(Now())
Else
currentDay = Day(date1)
currentWeekday = Weekday(date1)
End If
While currentDay <> 0
If currentWeekday = 0 Then
weekNumber = weekNumber + 1
currentWeekday = 7
End If
currentDay = currentDay - 1
currentWeekday = currentWeekday - 1
Wend
WeekOfMonth = weekNumber
End Function
Public Function OS() As String
#If Mac Then
OS = "Mac"
#Else
OS = "Windows"
#End If
End Function
Public Function UserName() As String
#If Mac Then
UserName = Environ("USER")
#Else
UserName = Environ("USERNAME")
#End If
End Function
Public Function UserDomain() As String
#If Mac Then
UserDomain = Environ("HOST")
#Else
UserDomain = Environ("USERDOMAIN")
#End If
End Function
Public Function ComputerName() As String
ComputerName = Environ("COMPUTERNAME")
End Function
Public Function GetActivePathAndName() As String
If Application.Name = "Microsoft Excel" Then
GetActivePathAndName = GetActivePathAndNameExcel()
ElseIf Application.Name = "Microsoft Word" Then
GetActivePathAndName = GetActivePathAndNameWord()
ElseIf Application.Name = "Microsoft PowerPoint" Then
GetActivePathAndName = GetActivePathAndNamePowerPoint()
ElseIf Application.Name = "Microsoft Access" Then
GetActivePathAndName = GetActivePathAndNameAccess()
ElseIf Application.Name = "Microsoft Publisher" Then
GetActivePathAndName = GetActivePathAndNamePublisher()
End If
End Function
Private Function GetActivePathAndNameExcel() As String
#If Mac Then
GetActivePathAndNameExcel = ThisWorkbook.Path & "/" & ThisWorkbook.Name
#Else
GetActivePathAndNameExcel = ThisWorkbook.Path & "\" & ThisWorkbook.Name
#End If
End Function
Private Function GetActivePathAndNameWord() As String
#If Mac Then
GetActivePathAndNameWord = ThisDocument.Path & "/" & ThisDocument.Name
#Else
GetActivePathAndNameWord = ThisDocument.Path & "\" & ThisDocument.Name
#End If
End Function
Private Function GetActivePathAndNamePowerPoint() As String
#If Mac Then
GetActivePathAndNamePowerPoint = ActivePresentation.Path & "/" & ActivePresentation.Name
#Else
GetActivePathAndNamePowerPoint = ActivePresentation.Path & "\" & ActivePresentation.Name
#End If
End Function
Private Function GetActivePathAndNameAccess() As String
#If Mac Then
GetActivePathAndNameAccess = CurrentProject.Path & "/" & CurrentProject.Name
#Else
GetActivePathAndNameAccess = CurrentProject.Path & "\" & CurrentProject.Name
#End If
End Function
Private Function GetActivePathAndNamePublisher() As String
#If Mac Then
GetActivePathAndNamePublisher = ThisDocument.Path & "/" & ThisDocument.Name
#Else
GetActivePathAndNamePublisher = ThisDocument.Path & "\" & ThisDocument.Name
#End If
End Function
Public Function GetActivePath() As String
If Application.Name = "Microsoft Excel" Then
GetActivePath = GetActivePathExcel()
ElseIf Application.Name = "Microsoft Word" Then
GetActivePath = GetActivePathWord()
ElseIf Application.Name = "Microsoft PowerPoint" Then
GetActivePath = GetActivePathPowerPoint()
ElseIf Application.Name = "Microsoft Access" Then
GetActivePath = GetActivePathAccess()
ElseIf Application.Name = "Microsoft Publisher" Then
GetActivePath = GetActivePathPublisher()
End If
End Function
Private Function GetActivePathExcel() As String
#If Mac Then
GetActivePathExcel = ThisWorkbook.Path & "/"
#Else
GetActivePathExcel = ThisWorkbook.Path & "\"
#End If
End Function
Private Function GetActivePathWord() As String
#If Mac Then
GetActivePathWord = ThisDocument.Path & "/"
#Else
GetActivePathWord = ThisDocument.Path & "\"
#End If
End Function
Private Function GetActivePathPowerPoint() As String
#If Mac Then
GetActivePathPowerPoint = ActivePresentation.Path & "/"
#Else
GetActivePathPowerPoint = ActivePresentation.Path & "\"
#End If
End Function
Private Function GetActivePathAccess() As String
#If Mac Then
GetActivePathAccess = CurrentProject.Path & "/"
#Else
GetActivePathAccess = CurrentProject.Path & "\"
#End If
End Function
Private Function GetActivePathPublisher() As String
#If Mac Then
GetActivePathPublisher = ThisDocument.Path & "/"
#Else
GetActivePathPublisher = ThisDocument.Path & "\"
#End If
End Function
Public Function FileCreationTime(Optional ByVal filePath As String) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FileCreationTime = FSO.GetFile(GetActivePathAndName()).DateCreated
Else
If FSO.FileExists(GetActivePath() & filePath) Then
FileCreationTime = FSO.GetFile(GetActivePath() & filePath).DateCreated
Else
FileCreationTime = FSO.GetFile(filePath).DateCreated
End If
End If
End Function
Public Function FileLastModifiedTime(Optional ByVal filePath As String) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FileLastModifiedTime = FSO.GetFile(GetActivePathAndName()).DateLastModified
Else
If FSO.FileExists(GetActivePath() & filePath) Then
FileLastModifiedTime = FSO.GetFile(GetActivePath() & filePath).DateLastModified
Else
FileLastModifiedTime = FSO.GetFile(filePath).DateLastModified
End If
End If
End Function
Public Function FileDrive(Optional ByVal filePath As String) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FileDrive = FSO.GetFile(GetActivePathAndName()).Drive
Else
If FSO.FileExists(GetActivePath() & filePath) Then
FileDrive = FSO.GetFile(GetActivePath() & filePath).Drive
Else
FileDrive = FSO.GetFile(filePath).Drive
End If
End If
End Function
Public Function FileName(Optional ByVal filePath As String) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FileName = FSO.GetFile(GetActivePathAndName()).Name
Else
If FSO.FileExists(GetActivePath() & filePath) Then
FileName = FSO.GetFile(GetActivePath() & filePath).Name
Else
FileName = FSO.GetFile(filePath).Name
End If
End If
End Function
Public Function FileFolder(Optional ByVal filePath As String) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FileFolder = FSO.GetFile(GetActivePathAndName()).ParentFolder
Else
If FSO.FileExists(GetActivePath() & filePath) Then
FileFolder = FSO.GetFile(GetActivePath() & filePath).ParentFolder
Else
FileFolder = FSO.GetFile(filePath).ParentFolder
End If
End If
End Function
Public Function CurrentFilePath(Optional ByVal filePath As String) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
CurrentFilePath = FSO.GetFile(GetActivePathAndName()).Path
Else
If FSO.FileExists(GetActivePath() & filePath) Then
CurrentFilePath = FSO.GetFile(GetActivePath() & filePath).Path
Else
CurrentFilePath = FSO.GetFile(filePath).Path
End If
End If
End Function
Public Function FileSize(Optional ByVal filePath As String, Optional ByVal byteSize As String) As Double
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim totalBytes#
If filePath = "" Then
totalBytes = FSO.GetFile(GetActivePathAndName()).Size
Else
If FSO.FileExists(GetActivePath() & filePath) Then
totalBytes = FSO.GetFile(GetActivePath() & filePath).Size
Else
totalBytes = FSO.GetFile(filePath).Size
End If
End If
Select Case LCase(byteSize)
Case "kb"
totalBytes = totalBytes / (2 ^ 10)
Case "mb"
totalBytes = totalBytes / (2 ^ 20)
Case "gb"
totalBytes = totalBytes / (2 ^ 30)
End Select
FileSize = totalBytes
End Function
Public Function FileType(Optional ByVal filePath As String) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
FileType = FSO.GetFile(GetActivePathAndName()).Type
Else
If FSO.FileExists(GetActivePath() & filePath) Then
FileType = FSO.GetFile(GetActivePath() & filePath).Type
Else
FileType = FSO.GetFile(filePath).Type
End If
End If
End Function
Public Function FileExtension(Optional ByVal filePath As String) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim FileName$
If filePath = "" Then
FileName = FSO.GetFile(GetActivePathAndName()).Name
Else
If FSO.FileExists(GetActivePath() & filePath) Then
FileName = FSO.GetFile(GetActivePath() & filePath).Name
Else
FileName = FSO.GetFile(filePath).Name
End If
End If
FileExtension = Right(FileName, Len(FileName) - InStrRev(FileName, "."))
End Function
Public Function ReadFile(ByVal filePath As String, Optional ByVal lineNumber As Integer) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim FileName$
Dim fileStream As Object
If FSO.FileExists(GetActivePath() & filePath) Then
filePath = GetActivePath() & filePath
ElseIf FSO.FileExists(filePath) Then
filePath = filePath
Else
ReadFile = "#FileDoesntExist!"
End If
Set fileStream = FSO.GetFile(filePath)
Set fileStream = fileStream.OpenAsTextStream(1, -2)
If lineNumber > 0 Then
Dim fileLinesArray() As String
fileLinesArray = Split(fileStream.ReadAll(), vbCrLf)
ReadFile = fileLinesArray(lineNumber)
Else
ReadFile = fileStream.ReadAll()
End If
End Function
Public Function WriteFile(ByVal filePath As String, ByVal fileText As String, Optional ByVal appendModeFlag As Boolean) As Boolean
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If InStr(filePath, "\") = 0 Then
If InStr(filePath, "/") = 0 Then
filePath = GetActivePath() & filePath
End If
ElseIf Right(filePath, 1) = "\" Or Right(filePath, 1) = "/" Then
If Not FSO.FolderExists(Left(filePath, InStrRev(filePath, "\"))) Then
WriteFile = False
Exit Function
End If
ElseIf Not FSO.FolderExists(filePath) Then
WriteFile = False
Exit Function
End If
Dim fileStream As Object
If appendModeFlag = False Then
Set fileStream = FSO.CreateTextFile(filePath, True)
fileStream.Write fileText
Else
Dim fileObject As Object
Set fileObject = FSO.GetFile(filePath)
Set fileStream = fileObject.OpenAsTextStream(8)
fileStream.Write fileText
End If
WriteFile = True
End Function
Public Function PathSeparator() As String
#If Mac Then
PathSeparator = "/"
#Else
PathSeparator = "\"
#End If
End Function
Public Function PathJoin(ParamArray pathArray()) As String
Dim individualPath
Dim combinedPath$
Dim individualValue
For Each individualPath In pathArray
If IsArray(individualPath) Then
For Each individualValue In individualPath
combinedPath = combinedPath & individualValue & PathSeparator()
Next
Else
combinedPath = combinedPath & CStr(individualPath) & PathSeparator()
End If
Next
combinedPath = Left(combinedPath, Len(combinedPath) - 1)
PathJoin = combinedPath
End Function
Public Function CountFiles(Optional ByVal filePath As String) As Integer
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
CountFiles = FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).Files.Count
Else
If FSO.FileExists(GetActivePath() & filePath) Then
CountFiles = FSO.GetFolder(GetActivePath() & filePath).Files.Count
Else
CountFiles = FSO.GetFolder(filePath).Files.Count
End If
End If
End Function
Public Function CountFolders(Optional ByVal filePath As String) As Integer
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
CountFolders = FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).SubFolders.Count
Else
If FSO.FileExists(GetActivePath() & filePath) Then
CountFolders = FSO.GetFolder(GetActivePath() & filePath).SubFolders.Count
Else
CountFolders = FSO.GetFolder(filePath).SubFolders.Count
End If
End If
End Function
Public Function CountFilesAndFolders(Optional ByVal filePath As String) As Integer
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
If filePath = "" Then
CountFilesAndFolders = FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).Files.Count + FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).SubFolders.Count
Else
If FSO.FileExists(GetActivePath() & filePath) Then
CountFilesAndFolders = FSO.GetFolder(GetActivePath() & filePath).Files.Count + FSO.GetFolder(GetActivePath() & filePath).SubFolders.Count
Else
CountFilesAndFolders = FSO.GetFolder(filePath).Files.Count + FSO.GetFolder(filePath).SubFolders.Count
End If
End If
End Function
Public Function GetFileNameByNumber(Optional ByVal filePath As String, Optional ByVal fileNumber As Integer = -1) As String
Dim FSO As Object
Set FSO = CreateObject("Scripting.FileSystemObject")
Dim fileCounter%
Dim individualFile As Object
Dim fileCollection As Object
If filePath = "" Then
Set fileCollection = FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).Files
Else
If FSO.FileExists(GetActivePath() & filePath) Then
Set fileCollection = FSO.GetFolder(GetActivePath() & filePath).Files
Else
Set fileCollection = FSO.GetFolder(filePath).Files
End If
End If
For Each individualFile In fileCollection
fileCounter = fileCounter + 1
If fileNumber = -1 Then
GetFileNameByNumber = individualFile.Name
Exit Function
ElseIf fileCounter = fileNumber Then
GetFileNameByNumber = individualFile.Name
Exit Function
End If
Next
End Function
Public Function Ceil(ByVal number As Double) As Long
If number = Fix(number) Then
Ceil = number
Else
Ceil = Fix(number + 1)
End If
End Function
Public Function Floor(ByVal number As Double) As Long
Floor = Fix(number)
End Function
Public Function InterpolateNumber(ByVal startingNumber As Double, ByVal endingNumber As Double, ByVal interpolationPercentage As Double) As Double
InterpolateNumber = startingNumber + ((endingNumber - startingNumber) * interpolationPercentage)
End Function
Public Function InterpolatePercent(ByVal startingNumber As Double, ByVal endingNumber As Double, ByVal interpolationNumber As Double) As Double
InterpolatePercent = (interpolationNumber - startingNumber) / (endingNumber - startingNumber)
End Function
Public Function Max(ParamArray numbers()) As Double
Dim individualParamArrayValue
Dim individualValue
Dim maxValue
maxValue = Empty
For Each individualParamArrayValue In numbers
If IsArray(individualParamArrayValue) Then
For Each individualValue In individualParamArrayValue
If TypeName(individualValue) = "String" Then
individualValue = CDbl(individualValue)
End If
If IsEmpty(maxValue) Then
maxValue = individualValue
ElseIf individualValue > maxValue Then
maxValue = individualValue
End If
Next
Else
If TypeName(individualParamArrayValue) = "String" Then
individualParamArrayValue = CDbl(individualParamArrayValue)
End If
If IsEmpty(maxValue) Then
maxValue = individualParamArrayValue
ElseIf individualParamArrayValue > maxValue Then
maxValue = individualParamArrayValue
End If
End If
Next
Max = maxValue
End Function
Public Function Min(ParamArray numbers()) As Double
Dim individualParamArrayValue
Dim individualValue
Dim minValue
minValue = Empty
For Each individualParamArrayValue In numbers
If IsArray(individualParamArrayValue) Then
For Each individualValue In individualParamArrayValue
If TypeName(individualValue) = "String" Then
individualValue = CDbl(individualValue)
End If
If IsEmpty(minValue) Then
minValue = individualValue
ElseIf individualValue < minValue Then
minValue = individualValue
End If
Next
Else
If TypeName(individualParamArrayValue) = "String" Then
individualParamArrayValue = CDbl(individualParamArrayValue)
End If
If IsEmpty(minValue) Then
minValue = individualParamArrayValue
ElseIf individualParamArrayValue < minValue Then
minValue = individualParamArrayValue
End If
End If
Next
Min = minValue
End Function
Public Function ModFloat(numerator As Double, denominator As Double) As Double
Dim modValue#
modValue = numerator - Fix(numerator / denominator) * denominator
If modValue >= -2 ^ -52 Then
If modValue <= 2 ^ -52 Then
modValue = 0
End If
End If
ModFloat = modValue
End Function
Public Function XlibVersion() As String
XlibVersion = "1.0.0"
End Function
Public Function XlibCredits() As String
XlibCredits = "Copyright (c) 2020 Anthony Mancini. XLib is Licensed under an MIT License."
End Function
Public Function XlibDocumentation() As String
XlibDocumentation = "https://x-vba.com/xlib"
End Function
Public Function Http(ByVal url As String, Optional ByVal httpMethod As String = "GET", Optional ByVal headers, Optional ByVal postData = "", Optional ByVal asyncFlag As Boolean, Optional ByVal statusErrorHandlerFlag As Boolean, Optional ByVal parseArguments) As String
Dim WinHttpRequest As Object
Set WinHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
WinHttpRequest.Open httpMethod, url, asyncFlag
If IsArray(headers) Then
Dim i%
For i = 0 To UBound(headers) - LBound(headers) Step 2
WinHttpRequest.SetRequestHeader headers(i), headers(i + 1)
Next
ElseIf TypeName(headers) = "Dictionary" Then
Dim dictKey
For Each dictKey In headers.Keys()
WinHttpRequest.SetRequestHeader dictKey, headers(dictKey)
Next
Else
WinHttpRequest.SetRequestHeader "User-Agent", "XLib"
End If
If postData = "" Then
WinHttpRequest.Send
Else
WinHttpRequest.Send postData
End If
If statusErrorHandlerFlag Then
If WinHttpRequest.Status = 200 Then
Http = WinHttpRequest.ResponseText
Else
Http = "#RequestFailedStatusCode" & WinHttpRequest.Status & "!"
End If
Else
Http = WinHttpRequest.ResponseText
End If
If IsArray(parseArguments) Then
Dim reorderedParseArguments()
i = UBound(parseArguments) - LBound(parseArguments)
ReDim reorderedParseArguments(i)
For i = 0 To UBound(parseArguments) - LBound(parseArguments)
reorderedParseArguments(i) = parseArguments(i)
Next
Http = ParseHtmlString(Http, reorderedParseArguments)
End If
End Function
Public Function SimpleHttp(ByVal url As String, ParamArray parseArguments()) As String
If UBound(parseArguments) > 0 Then
Dim i%
Dim reorderedParseArguments()
i = UBound(parseArguments) - LBound(parseArguments)
ReDim reorderedParseArguments(i)
For i = 0 To UBound(parseArguments) - LBound(parseArguments)
reorderedParseArguments(i) = parseArguments(i)
Next
SimpleHttp = ParseHtmlString(Http(url), reorderedParseArguments)
Else
SimpleHttp = Http(url)
End If
End Function
Public Function ParseHtmlString(ByVal htmlString As String, ByVal parseArguments)
Dim partialHtml$
Dim html As Object
Set html = CreateObject("HtmlFile")
html.body.innerHTML = htmlString
Dim i%
For i = LBound(parseArguments) To UBound(parseArguments)
If LCase(parseArguments(i)) = "id" Then
If partialHtml <> "" Then
html.body.innerHTML = partialHtml
End If
partialHtml = html.getElementById(parseArguments(i + 1)).innerHTML
html.body.innerHTML = partialHtml
i = i + 1
ElseIf LCase(parseArguments(i)) = "tag" Then
If partialHtml <> "" Then
html.body.innerHTML = partialHtml
End If
partialHtml = html.getElementsByTagName(parseArguments(i + 1))(i + 2).innerHTML
html.body.innerHTML = partialHtml
i = i + 2
ElseIf LCase(parseArguments(i)) = "left" Then
If IsNumeric(parseArguments(i + 1)) And TypeName(parseArguments(i + 1)) <> "String" Then
partialHtml = Left(partialHtml, parseArguments(i + 1))
Else
partialHtml = Left(partialHtml, InStr(1, partialHtml, CStr(parseArguments(i + 1)), vbTextCompare) - 1)
End If
i = i + 1
ElseIf LCase(parseArguments(i)) = "right" Then
If IsNumeric(parseArguments(i + 1)) And TypeName(parseArguments(i + 1)) <> "String" Then
partialHtml = Right(partialHtml, parseArguments(i + 1))
Else
partialHtml = Right(partialHtml, Len(partialHtml) - Len(parseArguments(i + 1)) + 1 - InStrRev(partialHtml, CStr(parseArguments(i + 1)), Compare:=vbTextCompare))
End If
i = i + 1
ElseIf LCase(parseArguments(i)) = "mid" Then
If IsNumeric(parseArguments(i + 1)) And TypeName(parseArguments(i + 1)) <> "String" Then
partialHtml = Mid(partialHtml, parseArguments(i + 1))
Else
partialHtml = Mid(partialHtml, Len(parseArguments(i + 1)) + InStr(1, partialHtml, CStr(parseArguments(i + 1)), vbTextCompare))
End If
i = i + 1
End If
Next
ParseHtmlString = partialHtml
End Function
Public Function RandBetween(ByVal minNumber As Long, ByVal maxNumber As Long)
RandBetween = Fix(Rnd * (maxNumber - minNumber + 1) + minNumber)
End Function
Public Function BigRandBetween(ByVal minNumber, ByVal maxNumber)
BigRandBetween = Fix(Rnd * (maxNumber - minNumber + 1) + minNumber)
End Function
Public Function RandomSample(ByRef variantArray)
Dim randomNumber&
randomNumber = RandBetween(1, UBound(variantArray) - LBound(variantArray) + 1)
RandomSample = variantArray(randomNumber - 1)
End Function
Public Function RandomRange(ByVal startNumber As Long, ByVal stopNumber As Long, ByVal stepNumber As Long) As Long
Dim randomNumber&
randomNumber = RandBetween(startNumber / stepNumber, stopNumber / stepNumber) * stepNumber
RandomRange = randomNumber
End Function
Public Function RandBool() As Boolean
RandBool = CBool(RandBetween(0, 1))
End Function
Public Function RandBetweens(ParamArray startOrEndNumberArray())
Dim pickNumber As Byte
If (UBound(startOrEndNumberArray) - LBound(startOrEndNumberArray) + 1) Mod 2 = 1 Then
RandBetweens = "#NotAnEvenNumberOfParameters!"
End If
pickNumber = Ceil(RandBetween(1, (UBound(startOrEndNumberArray) - LBound(startOrEndNumberArray) + 1)) / 2) * 2
RandBetweens = RandBetween(startOrEndNumberArray(pickNumber - 2), startOrEndNumberArray(pickNumber - 1))
End Function
Public Function RegexSearch(ByVal string1 As String, ByVal stringPattern As String, Optional ByVal globalFlag As Boolean, Optional ByVal ignoreCaseFlag As Boolean, Optional ByVal multilineFlag As Boolean) As String
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
Dim searchResults As Object
With Regex
.Global = globalFlag
.IgnoreCase = ignoreCaseFlag
.MultiLine = multilineFlag
.Pattern = stringPattern
End With
Set searchResults = Regex.Execute(string1)
RegexSearch = searchResults(0).Value
End Function
Public Function RegexTest(ByVal string1 As String, ByVal stringPattern As String, Optional ByVal globalFlag As Boolean, Optional ByVal ignoreCaseFlag As Boolean, Optional ByVal multilineFlag As Boolean) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
With Regex
.Global = globalFlag
.IgnoreCase = ignoreCaseFlag
.MultiLine = multilineFlag
.Pattern = stringPattern
End With
RegexTest = Regex.Test(string1)
End Function
Public Function RegexReplace(ByVal string1 As String, ByVal stringPattern As String, ByVal replacementString As String, Optional ByVal globalFlag As Boolean, Optional ByVal ignoreCaseFlag As Boolean, Optional ByVal multilineFlag As Boolean) As String
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
With Regex
.Global = globalFlag
.IgnoreCase = ignoreCaseFlag
.MultiLine = multilineFlag
.Pattern = stringPattern
End With
RegexReplace = Regex.Replace(string1, replacementString)
End Function
Public Function Capitalize(ByVal string1 As String) As String
Capitalize = UCase(Left(string1, 1)) & LCase(Mid(string1, 2))
End Function
Public Function LeftFind(ByVal string1 As String, ByVal searchString As String) As String
LeftFind = Left(string1, InStr(1, string1, searchString) - 1)
End Function
Public Function RightFind(ByVal string1 As String, ByVal searchString As String) As String
RightFind = Right(string1, Len(string1) - InStrRev(string1, searchString))
End Function
Public Function LeftSearch(ByVal string1 As String, ByVal searchString As String) As String
LeftSearch = Left(string1, InStr(1, string1, searchString, vbTextCompare) - 1)
End Function
Public Function RightSearch(ByVal string1 As String, ByVal searchString As String) As String
RightSearch = Right(string1, Len(string1) - InStrRev(string1, searchString, Compare:=vbTextCompare))
End Function
Public Function Substr(ByVal string1 As String, ByVal startCharacterNumber As Integer, ByVal endCharacterNumber As Integer) As String
Substr = Mid(string1, startCharacterNumber, endCharacterNumber - startCharacterNumber)
End Function
Public Function SubstrFind(ByVal string1 As String, ByVal RightFindString As String, ByVal rightSearchString As String, Optional ByVal noninclusiveFlag As Boolean) As String
Dim leftCharacterNumber%
Dim rightCharacterNumber%
leftCharacterNumber = InStr(1, string1, RightFindString)
rightCharacterNumber = InStrRev(string1, rightSearchString)
If noninclusiveFlag = True Then
leftCharacterNumber = leftCharacterNumber + Len(RightFindString)
rightCharacterNumber = rightCharacterNumber - Len(rightSearchString)
End If
SubstrFind = Mid(string1, leftCharacterNumber, rightCharacterNumber - leftCharacterNumber + Len(rightSearchString))
End Function
Public Function SubstrSearch(ByVal string1 As String, ByVal RightFindString As String, ByVal rightSearchString As String, Optional ByVal noninclusiveFlag As Boolean) As String
Dim leftCharacterNumber%
Dim rightCharacterNumber%
leftCharacterNumber = InStr(1, string1, RightFindString, vbTextCompare)
rightCharacterNumber = InStrRev(string1, rightSearchString, Compare:=vbTextCompare)
If noninclusiveFlag = True Then
leftCharacterNumber = leftCharacterNumber + Len(RightFindString)
rightCharacterNumber = rightCharacterNumber - Len(rightSearchString)
End If
SubstrSearch = Mid(string1, leftCharacterNumber, rightCharacterNumber - leftCharacterNumber + Len(rightSearchString))
End Function
Public Function Repeat(ByVal string1 As String, ByVal numberOfRepeats As Integer) As String
Dim i%
Dim combinedString$
For i = 1 To numberOfRepeats
combinedString = combinedString & string1
Next
Repeat = combinedString
End Function
Public Function Formatter(ByVal formatString As String, ParamArray textArray()) As String
Dim i As Byte
Dim individualTextItem
Dim individualValue
i = 0
For Each individualTextItem In textArray
If IsArray(individualTextItem) Then
For Each individualValue In individualTextItem
i = i + 1
formatString = Replace(formatString, "{" & i & "}", individualValue)
Next
Else
i = i + 1
formatString = Replace(formatString, "{" & i & "}", individualTextItem)
End If
Next
Formatter = formatString
End Function
Public Function Zfill(ByVal string1 As String, ByVal fillLength As Byte, Optional ByVal fillCharacter As String = "0", Optional ByVal rightToLeftFlag As Boolean) As String
While Len(string1) < fillLength
If rightToLeftFlag = False Then
string1 = fillCharacter + string1
Else
string1 = string1 + fillCharacter
End If
Wend
Zfill = string1
End Function
Public Function SplitText(ByVal string1 As String, ByVal substringNumber As Integer, Optional ByVal delimiterString As String = " ") As String
SplitText = Split(string1, delimiterString)(substringNumber - 1)
End Function
Public Function CountWords(ByVal string1 As String, Optional ByVal delimiterString As String = " ") As Integer
Dim stringArray() As String
stringArray = Split(string1, delimiterString)
CountWords = UBound(stringArray) - LBound(stringArray) + 1
End Function
Public Function CamelCase(ByVal string1 As String) As String
Dim i%
Dim stringArray() As String
stringArray = Split(string1, " ")
stringArray(0) = LCase(stringArray(0))
For i = 1 To (UBound(stringArray) - LBound(stringArray))
stringArray(i) = UCase(Left(stringArray(i), 1)) & LCase(Mid(stringArray(i), 2))
Next
CamelCase = Join(stringArray, "")
End Function
Public Function KebabCase(ByVal string1 As String) As String
KebabCase = LCase(Join(Split(string1, " "), "-"))
End Function
Public Function RemoveCharacters(ByVal string1 As String, ParamArray removedCharacters()) As String
Dim i%
Dim individualCharacter
For Each individualCharacter In removedCharacters
If Len(individualCharacter) > 1 Then
For i = 1 To Len(individualCharacter)
string1 = Replace(string1, Mid(individualCharacter, i, 1), "")
Next
Else
string1 = Replace(string1, individualCharacter, "")
End If
Next
RemoveCharacters = string1
End Function
Private Function NumberOfUppercaseLetters(ByVal string1 As String) As Integer
Dim i%
Dim numberOfUppercase%
For i = 1 To Len(string1)
If Asc(Mid(string1, i, 1)) >= 65 Then
If Asc(Mid(string1, i, 1)) <= 90 Then
numberOfUppercase = numberOfUppercase + 1
End If
End If
Next
NumberOfUppercaseLetters = numberOfUppercase
End Function
Public Function CompanyCase(ByVal string1 As String) As String
Dim i%
Dim k%
Dim origionalString$
Dim stringArray() As String
Dim splitCharacters$
origionalString = string1
string1 = LCase(string1)
splitCharacters = " ./()-_,*&1234567890"
For k = 1 To Len(splitCharacters)
stringArray = Split(string1, Mid(splitCharacters, k, 1))
For i = 0 To UBound(stringArray) - LBound(stringArray)
If NumberOfUppercaseLetters(Split(origionalString, Mid(splitCharacters, k, 1))(i)) <= 1 Then
stringArray(i) = UCase(Left(stringArray(i), 1)) & Mid(stringArray(i), 2)
Else
If UCase(Join(stringArray, Mid(splitCharacters, k, 1))) = origionalString Then
stringArray(i) = UCase(Left(stringArray(i), 1)) & Mid(stringArray(i), 2)
Else
stringArray(i) = Split(origionalString, Mid(splitCharacters, k, 1))(i)
End If
End If
Next
string1 = Join(stringArray, Mid(splitCharacters, k, 1))
Next
Dim companyAbbreviationArray() As String
companyAbbreviationArray = Split("AB|AG|GmbH|LLC|LLP|NV|PLC|SA|A. en P.|ACE|AD|AE|AL|AmbA|ANS|ApS|AS|ASA|AVV|BVBA|CA|CVA|d.d.|d.n.o.|d.o.o.|DA|e.V.|EE|EEG|EIRL|ELP|EOOD|EPE|EURL|GbR|GCV|GesmbH|GIE|HB|hf|IBC|j.t.d.|k.d.|k.d.d.|k.s.|KA/S|KB|KD|KDA|KG|KGaA|KK|Kol. SrK|Kom. SrK|LDC|Ltï¿½e.|NT|OE|OHG|Oy|OYJ|Oï¿½|PC Ltd|PMA|PMDN|PrC|PT|RAS|S. de R.L.|S. en N.C.|SA de CV|SAFI|SAS|SC|SCA|SCP|SCS|SENC|SGPS|SK|SNC|SOPARFI|sp|Sp. z.o.o.|SpA|spol s.r.o.|SPRL|TD|TLS|v.o.s.|VEB|VOF|BYSHR", "|")
Dim stringArrayLength%
stringArray = Split(string1, " ")
stringArrayLength = UBound(stringArray) - LBound(stringArray)
Dim companyAbbreviationString
For Each companyAbbreviationString In companyAbbreviationArray
If InStrRev(LCase(string1), " " & LCase(companyAbbreviationString)) = (Len(string1) - Len(companyAbbreviationString)) Then
If InStrRev(LCase(string1), " " & LCase(companyAbbreviationString)) <> 0 Then
CompanyCase = Left(string1, InStrRev(LCase(string1), LCase(companyAbbreviationString)) - 1) & companyAbbreviationString
Exit Function
End If
End If
Next
CompanyCase = string1
End Function
Public Function ReverseText(ByVal string1 As String) As String
Dim i%
Dim reversedString$
For i = 1 To Len(string1)
reversedString = reversedString & Mid(string1, Len(string1) - i + 1, 1)
Next
ReverseText = reversedString
End Function
Public Function ReverseWords(ByVal string1 As String, Optional ByVal delimiterCharacter As String = " ") As String
Dim i%
Dim stringArray() As String
Dim stringArrayLength%
Dim reversedStringArray() As String
stringArray = Split(string1, delimiterCharacter)
stringArrayLength = (UBound(stringArray) - LBound(stringArray))
ReDim reversedStringArray(stringArrayLength)
For i = 0 To stringArrayLength
reversedStringArray(i) = stringArray(stringArrayLength - i)
Next
ReverseWords = Join(reversedStringArray, delimiterCharacter)
End Function
Public Function IndentText(ByVal string1 As String, Optional ByVal indentAmount As Byte = 4) As String
Dim i%
Dim stringArray() As String
stringArray = Split(string1, Chr(10))
string1 = ""
For i = 1 To indentAmount
string1 = string1 & " "
Next
For i = 0 To (UBound(stringArray) - LBound(stringArray))
stringArray(i) = string1 & stringArray(i)
Next
IndentText = Join(stringArray, Chr(10))
End Function
Public Function DedentText(ByVal string1 As String) As String
Dim i%
Dim stringArray() As String
stringArray = Split(string1, Chr(10))
For i = 0 To (UBound(stringArray) - LBound(stringArray))
stringArray(i) = Trim(stringArray(i))
Next
DedentText = Join(stringArray, Chr(10))
End Function
Public Function ShortenText(ByVal string1 As String, Optional ByVal shortenWidth As Integer = 80, Optional ByVal placeholderText As String = "[...]", Optional ByVal delimiterCharacter As String = " ") As String
Dim shortenedString$
Dim individualString
Dim stringArray() As String
If Len(string1) <= (shortenWidth - Len(placeholderText) - Len(delimiterCharacter)) Then
ShortenText = string1
Exit Function
End If
stringArray = Split(string1, delimiterCharacter)
For Each individualString In stringArray
If Len(shortenedString & individualString) > (shortenWidth - Len(placeholderText) - Len(delimiterCharacter)) Then
shortenedString = shortenedString & placeholderText
Exit For
Else
shortenedString = shortenedString & individualString & delimiterCharacter
End If
Next
ShortenText = shortenedString
End Function
Public Function InSplit(ByVal string1 As String, ByVal splitString As String, Optional ByVal delimiterCharacter As String = " ") As Boolean
Dim individualString
For Each individualString In Split(splitString, delimiterCharacter)
If string1 = individualString Then
InSplit = True
Exit Function
End If
Next
InSplit = False
End Function
Public Function EliteCase(ByVal string1 As String) As String
string1 = Replace(string1, "o", "0", Compare:=vbTextCompare)
string1 = Replace(string1, "l", "1", Compare:=vbTextCompare)
string1 = Replace(string1, "z", "2", Compare:=vbTextCompare)
string1 = Replace(string1, "e", "3", Compare:=vbTextCompare)
string1 = Replace(string1, "a", "4", Compare:=vbTextCompare)
string1 = Replace(string1, "s", "5", Compare:=vbTextCompare)
string1 = Replace(string1, "t", "7", Compare:=vbTextCompare)
EliteCase = string1
End Function
Public Function ScrambleCase(ByVal string1 As String) As String
Dim i%
For i = 1 To Len(string1)
If RandBetween(0, 1) = 1 Then
Mid(string1, i, 1) = UCase(Mid(string1, i, 1))
Else
Mid(string1, i, 1) = LCase(Mid(string1, i, 1))
End If
Next
ScrambleCase = string1
End Function
Public Function LeftSplit(ByVal string1 As String, ByVal numberOfSplit As Integer, Optional ByVal delimiterCharacter As String = " ") As String
Dim i%
Dim newString$
Dim stringArray() As String
Dim stringArrayLength%
numberOfSplit = numberOfSplit - 1
stringArray = Split(string1, delimiterCharacter)
stringArrayLength = (UBound(stringArray) - LBound(stringArray) + 1)
If numberOfSplit >= stringArrayLength Then
LeftSplit = string1
Exit Function
End If
For i = 0 To numberOfSplit
If i = numberOfSplit Then
newString = newString & stringArray(i)
Else
newString = newString & stringArray(i) & delimiterCharacter
End If
Next
LeftSplit = newString
End Function
Public Function RightSplit(ByVal string1 As String, ByVal numberOfSplit As Integer, Optional ByVal delimiterCharacter As String = " ") As String
Dim i%
Dim newString$
Dim stringArray() As String
Dim stringArrayLength%
numberOfSplit = numberOfSplit - 1
stringArray = Split(string1, delimiterCharacter)
stringArrayLength = (UBound(stringArray) - LBound(stringArray) + 1)
If numberOfSplit >= stringArrayLength Then
RightSplit = string1
Exit Function
End If
For i = 0 To numberOfSplit
If i = numberOfSplit Then
newString = newString & stringArray(stringArrayLength - (numberOfSplit - i) - 1)
Else
newString = newString & stringArray(stringArrayLength - (numberOfSplit - i) - 1) & delimiterCharacter
End If
Next
RightSplit = newString
End Function
Public Function TrimChar(ByVal string1 As String, Optional ByVal trimCharacter As String = " ") As String
While Left(string1, 1) = trimCharacter
Mid(string1, 1) = Chr(1)
string1 = Replace(string1, Chr(1), "")
Wend
While Right(string1, 1) = trimCharacter
Mid(string1, Len(string1)) = Chr(1)
string1 = Replace(string1, Chr(1), "")
Wend
TrimChar = string1
End Function
Public Function TrimLeft(ByVal string1 As String, Optional ByVal trimCharacter As String = " ") As String
While Left(string1, 1) = trimCharacter
Mid(string1, 1) = Chr(1)
string1 = Replace(string1, Chr(1), "")
Wend
TrimLeft = string1
End Function
Public Function TrimRight(ByVal string1 As String, Optional ByVal trimCharacter As String = " ") As String
While Right(string1, 1) = trimCharacter
Mid(string1, Len(string1)) = Chr(1)
string1 = Replace(string1, Chr(1), "")
Wend
TrimRight = string1
End Function
Public Function CountUppercaseCharacters(ByVal string1 As String) As Integer
Dim i%
Dim characterAsciiCode As Byte
Dim uppercaseCounter%
For i = 1 To Len(string1)
characterAsciiCode = Asc(Mid(string1, i, 1))
If characterAsciiCode >= 65 And characterAsciiCode <= 90 Then
uppercaseCounter = uppercaseCounter + 1
End If
Next
CountUppercaseCharacters = uppercaseCounter
End Function
Public Function CountLowercaseCharacters(ByVal string1 As String) As Integer
Dim i%
Dim characterAsciiCode As Byte
Dim lowercaseCounter%
For i = 1 To Len(string1)
characterAsciiCode = Asc(Mid(string1, i, 1))
If characterAsciiCode >= 97 And characterAsciiCode <= 122 Then
lowercaseCounter = lowercaseCounter + 1
End If
Next
CountLowercaseCharacters = lowercaseCounter
End Function
Public Function TextJoin(ByVal stringArray, Optional ByVal delimiterCharacter As String, Optional ByVal ignoreEmptyCellsFlag As Boolean) As String
Dim individualString
Dim combinedString$
For Each individualString In stringArray
individualString = CStr(individualString)
If ignoreEmptyCellsFlag Then
If Not (IsEmpty(individualString) Or individualString = "") Then
combinedString = combinedString & individualString & delimiterCharacter
End If
Else
combinedString = combinedString & individualString & delimiterCharacter
End If
Next
If delimiterCharacter <> "" Then
combinedString = Left(combinedString, InStrRev(combinedString, delimiterCharacter) - 1)
End If
TextJoin = combinedString
End Function
Public Function Hamming(string1 As String, string2 As String) As Integer
If Len(string1) <> Len(string2) Then
Hamming = CVErr(2015)
End If
Dim totalDistance%
totalDistance = 0
Dim i%
For i = 1 To Len(string1)
If Mid(string1, i, 1) <> Mid(string2, i, 1) Then
totalDistance = totalDistance + 1
End If
Next
Hamming = totalDistance
End Function
Public Function Levenshtein(string1 As String, string2 As String) As Long
If string1 = string2 Then
Levenshtein = 0
Exit Function
ElseIf string1 = Empty Then
Levenshtein = Len(string2)
Exit Function
ElseIf string2 = Empty Then
Levenshtein = Len(string1)
Exit Function
End If
Dim numberOfRows%
Dim numberOfColumns%
numberOfRows = Len(string1)
numberOfColumns = Len(string2)
Dim distanceArray() As Integer
ReDim distanceArray(numberOfRows, numberOfColumns)
Dim r%
Dim c%
For r = 0 To numberOfRows
For c = 0 To numberOfColumns
distanceArray(r, c) = 0
Next
Next
For r = 1 To numberOfRows
distanceArray(r, 0) = r
Next
For c = 1 To numberOfColumns
distanceArray(0, c) = c
Next
Dim operationCost%
For c = 1 To numberOfColumns
For r = 1 To numberOfRows
If Mid(string1, r, 1) = Mid(string2, c, 1) Then
operationCost = 0
Else
operationCost = 1
End If
distanceArray(r, c) = Min(distanceArray(r - 1, c) + 1, distanceArray(r, c - 1) + 1, distanceArray(r - 1, c - 1) + operationCost)
Next
Next
Levenshtein = distanceArray(numberOfRows, numberOfColumns)
End Function
Public Function Damerau(string1 As String, string2 As String) As Integer
If string1 = string2 Then
Damerau = 0
ElseIf string1 = Empty Then
Damerau = Len(string2)
ElseIf string2 = Empty Then
Damerau = Len(string1)
End If
Dim inf&
Dim da As Object
inf = Len(string1) + Len(string2)
Set da = CreateObject("Scripting.Dictionary")
Dim i%
For i = 1 To Len(string1)
If da.exists(Mid(string1, i, 1)) = False Then
da.Add Mid(string1, i, 1), "0"
End If
Next
For i = 1 To Len(string2)
If da.exists(Mid(string2, i, 1)) = False Then
da.Add Mid(string2, i, 1), "0"
End If
Next
Dim H() As Long
ReDim H(Len(string1) + 1, Len(string2) + 1)
Dim k%
For i = 0 To (Len(string1) + 1)
For k = 0 To (Len(string2) + 1)
H(i, k) = 0
Next
Next
For i = 0 To Len(string1)
H(i + 1, 0) = inf
H(i + 1, 1) = i
Next
For k = 0 To Len(string2)
H(0, k + 1) = inf
H(1, k + 1) = k
Next
Dim db&
Dim i1&
Dim k1&
Dim cost&
For i = 1 To Len(string1)
db = 0
For k = 1 To Len(string2)
i1 = CInt(da(Mid(string2, k, 1)))
k1 = db
cost = 1
If Mid(string1, i, 1) = Mid(string2, k, 1) Then
cost = 0
db = k
End If
H(i + 1, k + 1) = Min(H(i, k) + cost, H(i + 1, k) + 1, H(i, k + 1) + 1, H(i1, k1) + (i - i1 - 1) + 1 + (k - k1 - 1))
Next
If da.exists(Mid(string1, i, 1)) Then
da.Remove Mid(string1, i, 1)
da.Add Mid(string1, i, 1), CStr(i)
Else
da.Add Mid(string1, i, 1), CStr(i)
End If
Next
Damerau = H(Len(string1) + 1, Len(string2) + 1)
End Function
Public Function Jsonify(ByVal indentLevel As Byte, ParamArray stringArray()) As String
Dim i As Byte
Dim jsonString$
Dim individualTextItem
Dim individualValue
Dim indentString$
jsonString = "["
For i = 1 To indentLevel
indentString = indentString & " "
Next
If indentLevel > 0 Then
jsonString = jsonString & Chr(10)
End If
For Each individualTextItem In stringArray
If IsArray(individualTextItem) Then
For Each individualValue In individualTextItem
jsonString = jsonString & indentString
If IsNumeric(individualValue) Then
jsonString = jsonString & individualValue & ","
Else
jsonString = jsonString & Chr(34) & individualValue & Chr(34) & ","
End If
If indentLevel > 0 Then
jsonString = jsonString & Chr(10)
End If
Next
Else
jsonString = jsonString & indentString
If IsNumeric(individualTextItem) Then
jsonString = jsonString & individualTextItem & ","
Else
jsonString = jsonString & Chr(34) & individualTextItem & Chr(34) & ","
End If
If indentLevel > 0 Then
jsonString = jsonString & Chr(10)
End If
End If
Next
jsonString = Left(jsonString, InStrRev(jsonString, ",") - 1)
If indentLevel > 0 Then
jsonString = jsonString & Chr(10)
End If
jsonString = jsonString & "]"
Jsonify = jsonString
End Function
Public Function UuidFour() As String
Dim firstGroup$
Dim secondGroup$
Dim thirdGroup$
Dim fourthGroup$
Dim fifthGroup$
Dim sixthGroup$
firstGroup = BigDec2Hex(BigRandBetween(0, 4294967295#), 8) & "-"
secondGroup = Dec2Hex(RandBetween(0, 65535), 4) & "-"
thirdGroup = Dec2Hex(RandBetween(16384, 20479), 4) & "-"
fourthGroup = Dec2Hex(RandBetween(32768, 49151), 4) & "-"
fifthGroup = Dec2Hex(RandBetween(0, 65535), 4)
sixthGroup = BigDec2Hex(BigRandBetween(0, 4294967295#), 8)
UuidFour = firstGroup & secondGroup & thirdGroup & fourthGroup & fifthGroup & sixthGroup
End Function
Public Function HideText(ByVal string1 As String, ByVal hiddenFlag As Boolean, Optional ByVal hideString As String) As String
If hiddenFlag Then
If hideString = "" Then
HideText = "********"
Else
HideText = hideString
End If
Else
HideText = string1
End If
End Function
Public Function JavaScript(ByVal jsFuncCode As String, ByVal jsFuncName As String, Optional ByVal argument1, Optional ByVal argument2, Optional ByVal argument3, Optional ByVal argument4, Optional ByVal argument5, Optional ByVal argument6, Optional ByVal argument7, Optional ByVal argument8, Optional ByVal argument9, Optional ByVal argument10, Optional ByVal argument11, Optional ByVal argument12, Optional ByVal argument13, Optional ByVal argument14, Optional ByVal argument15, Optional ByVal argument16)
Dim ScriptContoller As Object
Set ScriptContoller = CreateObject("ScriptControl")
ScriptContoller.Language = "JScript"
ScriptContoller.addCode jsFuncCode
JavaScript = ScriptContoller.Run(jsFuncName, argument1, argument2, argument3, argument4, argument5, argument6, argument7, argument8, argument9, argument10, argument11, argument12, argument13, argument14, argument15, argument16)
End Function
Public Function HtmlEscape(ByVal string1 As String) As String
string1 = Replace(string1, "&", "&amp;")
string1 = Replace(string1, Chr(34), "&quot;")
string1 = Replace(string1, "'", "&apos;")
string1 = Replace(string1, "<", "&lt;")
string1 = Replace(string1, ">", "&gt;")
HtmlEscape = string1
End Function
Public Function HtmlUnescape(ByVal string1 As String) As String
string1 = Replace(string1, "&amp;", "&")
string1 = Replace(string1, "&quot;", Chr(34))
string1 = Replace(string1, "&apos;", "'")
string1 = Replace(string1, "&lt;", "<")
string1 = Replace(string1, "&gt;", ">")
HtmlUnescape = string1
End Function
Private Sub CallTextToSpeech(combinedString)
Application.Speech.Speak combinedString, True
End Sub
Public Function SpeakText(ParamArray textArray()) As String
Dim combinedString$
Dim individualTextItem
For Each individualTextItem In textArray
combinedString = combinedString & individualTextItem & " "
Next
If Application.Name = "Microsoft Excel" Then
CallTextToSpeech combinedString
End If
SpeakText = Trim(combinedString)
End Function
Public Function Dec2Hex(ByVal number As Long, Optional ByVal zeroFillAmount As Integer) As String
Dim i%
Dim hexString$
hexString = Hex(number)
If zeroFillAmount > 0 Then
While Len(hexString) < zeroFillAmount
hexString = "0" & hexString
Wend
End If
Dec2Hex = hexString
End Function
Public Function BigDec2Hex(ByVal number, Optional ByVal zeroFillAmount As Integer) As String
Dim i%
Dim hexString$
hexString = BigHex(number)
If zeroFillAmount > 0 Then
While Len(hexString) < zeroFillAmount
hexString = "0" & hexString
Wend
End If
BigDec2Hex = hexString
End Function
Public Function BigHex(ByVal number) As String
Dim integerString$
Dim decimalString$
Dim hexString$
While number > 0
number = number / 16
If InStr(1, CStr(number), ".") > 0 Then
integerString = Split(CStr(number), ".")(0)
decimalString = Split(CStr(number), ".")(1)
Else
integerString = CStr(number)
decimalString = "0"
End If
Select Case decimalString
Case "0"
hexString = "0" & hexString
Case "0625"
hexString = "1" & hexString
Case "125"
hexString = "2" & hexString
Case "1875"
hexString = "3" & hexString
Case "25"
hexString = "4" & hexString
Case "3125"
hexString = "5" & hexString
Case "375"
hexString = "6" & hexString
Case "4375"
hexString = "7" & hexString
Case "5"
hexString = "8" & hexString
Case "5625"
hexString = "9" & hexString
Case "625"
hexString = "A" & hexString
Case "6875"
hexString = "B" & hexString
Case "75"
hexString = "C" & hexString
Case "8125"
hexString = "D" & hexString
Case "875"
hexString = "E" & hexString
Case "9375"
hexString = "F" & hexString
End Select
number = Fix(number)
Wend
BigHex = hexString
End Function
Public Function Hex2Dec(ByVal hexNumber As String) As Long
Hex2Dec = CLng("&H" & hexNumber)
End Function
Public Function Len2(ByVal val) As Integer
If IsArray(val) And Right(TypeName(val), 2) = "()" Then
Len2 = UBound(val) - LBound(val) + 1
ElseIf TypeName(val) = "String" Then
Len2 = Len(val)
ElseIf IsNumeric(val) Then
Len2 = Len(CStr(val))
Else
Len2 = val.Count
End If
End Function
Public Function IsEmail(ByVal string1 As String) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
With Regex
.Global = True
.IgnoreCase = True
.MultiLine = True
.Pattern = "^[a-zA-Z0-9_.]*?[@][a-zA-Z0-9.]*?[.][a-zA-Z]{2,15}$"
End With
IsEmail = Regex.Test(string1)
End Function
Public Function IsPhone(ByVal string1 As String) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
With Regex
.Global = True
.IgnoreCase = True
.MultiLine = True
.Pattern = "^\s*[+]{0,1}[0-9]{0,1}[\s-]{0,1}\({0,1}([0-9]{3})\){0,1}[\s-]{0,1}([0-9]{3})[\s-]{0,1}([0-9]{4})$"
End With
IsPhone = Regex.Test(string1)
End Function
Public Function IsCreditCard(ByVal string1 As String) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
Dim regexPattern$
regexPattern = regexPattern & "(3[47][0-9]{13})|"
regexPattern = regexPattern & "(3(0[0-5]|[68][0-9])?[0-9]{11})|"
regexPattern = regexPattern & "(6(011|5[0-9]{2})[0-9]{12})|"
regexPattern = regexPattern & "((2131|1800|35[0-9]{3})[0-9]{11})|"
regexPattern = regexPattern & "(5[1-5][0-9]{14})|"
regexPattern = regexPattern & "(4[0-9]{12}([0-9]{3})?)"
With Regex
.Global = True
.IgnoreCase = True
.MultiLine = True
.Pattern = regexPattern
End With
IsCreditCard = Regex.Test(string1)
End Function
Public Function IsUrl(ByVal string1 As String) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
With Regex
.Global = True
.IgnoreCase = True
.MultiLine = True
.Pattern = "http(s){0,1}://www.[a-zA-Z0-9_.]*?[.][a-zA-Z]{2,15}"
End With
IsUrl = Regex.Test(string1)
End Function
Public Function IsIPFour(ByVal string1 As String) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
With Regex
.Global = True
.IgnoreCase = True
.MultiLine = True
.Pattern = "^((2[0-4]\d|25[0-5]|1\d\d|\d{1,2})[.]){3}(2[0-4]\d|25[0-5]|1\d\d|\d{1,2})$"
End With
IsIPFour = Regex.Test(string1)
End Function
Public Function IsMacAddress(ByVal string1 As String) As Boolean
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
With Regex
.Global = True
.IgnoreCase = True
.MultiLine = True
.Pattern = "^(([a-fA-F0-9]{2}([:]|[-])){5}[a-fA-F0-9]{2}|([a-fA-F0-9]{3}[.]){3}[a-fA-F0-9]{3})$"
End With
IsMacAddress = Regex.Test(string1)
End Function
Public Function CreditCardName(ByVal string1 As String) As String
Dim Regex As Object
Set Regex = CreateObject("VBScript.RegExp")
Regex.Global = True
Regex.IgnoreCase = True
Regex.MultiLine = True
Regex.Pattern = "(3[47][0-9]{13})"
If Regex.Test(string1) Then
CreditCardName = "Amex"
Exit Function
End If
Regex.Pattern = "(3(0[0-5]|[68][0-9])?[0-9]{11})"
If Regex.Test(string1) Then
CreditCardName = "Diners"
Exit Function
End If
Regex.Pattern = "(6(011|5[0-9]{2})[0-9]{12})"
If Regex.Test(string1) Then
CreditCardName = "Discover"
Exit Function
End If
Regex.Pattern = "((2131|1800|35[0-9]{3})[0-9]{11})"
If Regex.Test(string1) Then
CreditCardName = "JCB"
Exit Function
End If
Regex.Pattern = "(5[1-5][0-9]{14})"
If Regex.Test(string1) Then
CreditCardName = "MasterCard"
Exit Function
End If
Regex.Pattern = "(4[0-9]{12}([0-9]{3})?)"
If Regex.Test(string1) Then
CreditCardName = "Visa"
Exit Function
End If
CreditCardName = "#NotAValidCreditCardNumber!"
End Function
Public Function FormatCreditCard(ByVal string1 As String) As String
If IsCreditCard(string1) Then
FormatCreditCard = Left(string1, 4) & "-" & Mid(string1, 5, 4) & "-" & Mid(string1, 9, 4) & "-" & Mid(string1, 13)
Else
FormatCreditCard = "#NotAValidCreditCardNumber!"
End If
End Function
