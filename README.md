# XLib

XLib is very small (~60 KB), open source, free, MIT Licensed, VBA function
library that adds around 120 additional functions to VBA, making it much
easier to develop VBA application for the Microsoft Office programs. 

Some additional characteristics of Xlib include:
* MIT Licensed
* 0 dependencies
* Extremely portable
* Uses late bindings for all functions
	- Late bindings make it easy to install and include within your own projects 
	  without having to configure your office programs to use the library. 
	  Additionally, when you ship an Office file to another user, they won't 
	  need to install or configure anything to use the XLib functions, making 
	  it very portable.
* Well tested
	- I have written tests for virtually every single function in XLib, which
	  can be found in the github reposity, so you can quickly see if the functions
	  work on your machine and with your version of the Office programs.
* Works with Excel, Word, PowerPoint, Outlook, Access, and Publisher
	- When needed, XLib checks the Office program that is calling a function
	  and runs different code to ensure that the same API works with many of
	  the different Office programs.
* Many modules are cross platform with Mac
	- Conditional compilation is used so that many of the XLib functions work
	  on Mac as well.


# Sample of XLib functions

* **Len2** -> Returns length of Strings, Arrays, Dictionaries, Collections, and
  any other objects that implement the property .Count, including Workbooks,
  Sheets, Worksheets, Ranges, Documents, Presentations, Slides, and many other
  Office Objects
* **Sort** -> Sorts an array in ascedning or descending order
* **Large/Small** -> Same as the Large() and Small() functions in Excel, but can
   be used in Word, PowerPoint, etc.
* **SubstrFind** -> Returns all characters between two substring
* **IsInArray** -> Returns True if the value is found in an array
* **Quarter** -> Returns the quarter of the year
* *RandBetween* -> Same as Excel RandBetween(), but can be used in Word, 
  PowerPoint, etc.
* **RegexTest** -> Tests if the regex is found in a string
* **Jsonify** -> Converts arrays into JSON format
* **Http** -> Performs a web request and returns the response, with options to
  set headers, send post data, etc.
* **ReadFile** -> To easily read files
* **WriteFile** -> To easily write files


# Installation

Xlib is written in pure VBA code and uses late bindings, so installation is as
simple as importing the Xlib.min.bas module, or alternatively you can simply
copy and paste the source code from the Xlib.min.bas module on the github
repository.


# Table Of Contents

Below are a list of all Modules in XLib and all functions within those modules:

* **Array**
	- AverageHigh
	- AverageLow
	- CountUnique
	- IsInArray
	- Large
	- Reverse
	- Small
	- Sort
	- SumHigh
	- SumLow

* **Color**
	- Hex2Hsl
	- Hex2Rgb
	- Hsl2Hex
	- Hsl2Rgb
	- Rgb2Hex
	- Rgb2Hsl
	- Rgb2Hsv

* **DateTime**
	- DaysOfMonth
	- MonthName2
	- Quarter
	- TimeConverter
	- WeekOfMonth
	- WeekdayName2

* **Environment**
	- ComputerName
	- OS
	- UserDomain
	- UserName

* **File**
	- CountFiles
	- CountFilesAndFolders
	- CountFolders
	- CurrentFilePath
	- FileCreationTime
	- FileDrive
	- FileExtension
	- FileFolder
	- FileLastModifiedTime
	- FileName
	- FileSize
	- FileType
	- GetActivePath
	- GetActivePathAccess
	- GetActivePathAndName
	- GetActivePathAndNameAccess
	- GetActivePathAndNameExcel
	- GetActivePathAndNamePowerPoint
	- GetActivePathAndNamePublisher
	- GetActivePathAndNameWord
	- GetActivePathExcel
	- GetActivePathPowerPoint
	- GetActivePathPublisher
	- GetActivePathWord
	- GetFileNameByNumber
	- PathJoin
	- PathSeparator
	- ReadFile
	- WriteFile

* **Math**
	- Ceil
	- Floor
	- InterpolateNumber
	- InterpolatePercent
	- Max
	- Min
	- ModFloat

* **Meta**
	- XlibCredits
	- XlibDocumentation
	- XlibVersion

* **Network**
	- Http
	- ParseHtmlString
	- SimpleHttp

* **Random**
	- BigRandBetween
	- RandBetween
	- RandBetweens
	- RandBool
	- RandomRange
	- RandomSample

* **Regex**
	- RegexReplace
	- RegexSearch
	- RegexTest

* **StringManipulation**
	- CamelCase
	- Capitalize
	- CompanyCase
	- CountLowercaseCharacters
	- CountUppercaseCharacters
	- CountWords
	- DedentText
	- EliteCase
	- Formatter
	- InSplit
	- IndentText
	- KebabCase
	- LeftFind
	- LeftSearch
	- LeftSplit
	- NumberOfUppercaseLetters
	- RemoveCharacters
	- Repeat
	- ReverseText
	- ReverseWords
	- RightFind
	- RightSearch
	- RightSplit
	- ScrambleCase
	- ShortenText
	- SplitText
	- Substr
	- SubstrFind
	- SubstrSearch
	- TextJoin
	- TrimChar
	- TrimLeft
	- TrimRight
	- Zfill

* **StringMetrics**
	- Damerau
	- Hamming
	- Levenshtein

* **Utilities**
	- BigDec2Hex
	- BigHex
	- CallTextToSpeech
	- Dec2Hex
	- Hex2Dec
	- HideText
	- HtmlEscape
	- HtmlUnescape
	- JavaScript
	- Jsonify
	- Len2
	- SpeakText
	- UuidFour

* **Validators**
	- CreditCardName
	- FormatCreditCard
	- IsCreditCard
	- IsEmail
	- IsIPFour
	- IsMacAddress
	- IsPhone
	- IsUrl



# Documentation

## Array Module

This module contains a set of functions for manipulating and working with arrays.

========================================

### AverageHigh

This function returns the average of the top values of the number specified in the second argument. For example, if the second argument is 3, only the top 3 values will be averaged

**Arguments**
 * **array1** {Variant}: is the range that will be averaged
 * **numberAveraged** {Integer}: is the number of the top values that will be averaged

**Returns**
 * {Variant}: Returns the average of the top numbers specified

**Examples**
 *  =AverageHigh({1,2,3,4}, 2) -> 3.5; as 3 and 4 will be averaged
 *  =AverageHigh({1,2,3,4}, 3) -> 3; as 2, 3, and 4 will be averaged


========================================

### AverageLow

This function returns the average of the bottom values of the number specified in the second argument. For example, if the second argument is 3, only the bottom 3 values will be averaged

**Arguments**
 * **array1** {Variant}: is the range that will be averaged
 * **numberAveraged** {Integer}: is the number of the bottom values that will be averaged

**Returns**
 * {Variant}: Returns the average of the bottom numbers specified

**Examples**
 *  =AverageLow({1,2,3,4}, 2) -> 1.5; as 1 and 2 will be averaged
 *  =AverageLow({1,2,3,4}, 3) -> 2; as 1, 2, and 3 will be averaged


========================================

### CountUnique

This function counts the number of unique occurances of values within a range or multiple ranges

**Arguments**
 * **array1** {Variant}: is the group of cells we are counting the unique values of

**Returns**
 * {Variant}: Returns the number of unique values

**Examples**
 *  =CountUnique(1, 2, 2, 3) -> 3;
 *  =CountUnique("a", "a", "a") -> 1;
 *  =CountUnique(arr) -> 3; Where arr = [1, 2, 4, 4, 1]


========================================

### IsInArray

This function checks if a value is in an array

**Arguments**
 * **value1** {Variant}: is the value that will be checked if its in the array
 * **array1** {Variant}: is the array

**Returns**
 * {Boolean}: Returns boolean True if the value is in the array, and false otherwise

**Examples**
 *  =IsInArray("hello", {"one", 2, "hello"}) -> True
 *  =IsInArray("hello", {1, "two", "three"}) -> False


========================================

### Large

This function returns the nth highest number an in array, similar to Excel's LARGE function.

**Arguments**
 * **array1** {Variant}: is the array that the number will be pulled from
 * **nthNumber** {Integer}: is the number of the top value that will be chosen. For example, a nthNumber of 1 results in the 1st highest value being chosen, when a number of 2 results in the 2nd, etc.

**Returns**
 * {Variant}: Returns the nth highest number

**Examples**
 *  =Large({1,2,3,4}, 1) -> 4
 *  =Large({1,2,3,4}, 2) -> 3


========================================

### Reverse

This function takes an array and reverses all its elements

**Arguments**
 * **array1** {Variant}: is the array that will be reversed

**Returns**
 * {Variant}: Returns the a reversed array

**Examples**
 *  =Reverse({1,2,3}) -> {3,2,1}


========================================

### Small

This function returns the nth lowest number an in array, similar to Excel's SMALL function.

**Arguments**
 * **array1** {Variant}: is the array that the number will be pulled from
 * **nthNumber** {Integer}: is the number of the bottom value that will be chosen. For example, a nthNumber of 1 results in the 1st smallest value being chosen, when a number of 2 results in the 2nd, etc.

**Returns**
 * {Variant}: Returns the nth smallest number

**Examples**
 *  =Small({1,2,3,4}, 1) -> 1
 *  =Small({1,2,3,4}, 2) -> 2


========================================

### Sort

This function is an implementation of Bubble Sort, allowing the user to sort an array, optionally allowing the user to specify the array to be sorted in descending order

**Arguments**
 * **sortableArray** {Variant}: is the array that will be sorted
 * (Optional) _[**descendingFlag** {Boolean}]_: changes the sort to descending

**Returns**
 * {Variant}: Returns the a sorted array

**Examples**
 *  =Sort({1,3,2}) -> {1,2,3}
 *  =Sort({1,3,2}, True) -> {3,2,1}


========================================

### SumHigh

This function returns the sum of the top values of the number specified in the second argument. For example, if the second argument is 3, only the top 3 values will be summed

**Arguments**
 * **array1** {Variant}: is the range that will be summed
 * **numberSummed** {Integer}: is the number of the top values that will be summed

**Returns**
 * {Variant}: Returns the sum of the top numbers specified

**Examples**
 *  =SumHigh({1,2,3,4}, 2) -> 7; as 3 and 4 will be summed
 *  =SumHigh({1,2,3,4}, 3) -> 9; as 2, 3, and 4 will be summed


========================================

### SumLow

This function returns the sum of the bottom values of the number specified in the second argument. For example, if the second argument is 3, only the bottom 3 values will be summed

**Arguments**
 * **array1** {Variant}: is the range that will be summed
 * **numberSummed** {Integer}: is the number of the bottom values that will be summed

**Returns**
 * {Variant}: Returns the sum of the bottom numbers specified

**Examples**
 *  =SumLow({1,2,3,4}, 2) -> 3; as 1 and 2 will be summed
 *  =SumLow({1,2,3,4}, 3) -> 6; as 1, 2, and 3 will be summed


---

## Color Module

This module contains a set of functions for working with colors

========================================

### Hex2Hsl

This function converts a HEX color value into an HSL color value

**Arguments**
 * **hexColorString** {String}: is the hex color

**Returns**
 * {String}: Returns a string with the HSL value of the color

**Examples**
 *  =Hex2Hsl("084080") -> "(212.0, 88.2%, 26.7%)"
 *  =Hex2Hsl("#084080") -> "(212.0, 88.2%, 26.7%)"


========================================

### Hex2Rgb

This function converts a HEX color value into an RGB color value, or optionally a single value from the RGB value.

**Arguments**
 * **hexColorString** {String}: is the color in HEX format
 * (Optional) _[**singleColorNumberOrName** {Variant = -1}]_: is a number or string that specifies which of the single color values to return. If this is set to 0 or "Red", the red value will be returned. If this is set to 1 or "Green", the green value will be returned. If this is set to 2 or "Blue", the blue value will be returned.

**Returns**
 * {Variant}: Returns a string with the RGB value of the color or the number of the individual color chosen

**Examples**
 *  =Hex2Rgb("FFFFFF") -> "(255, 255, 255)"
 *  =Hex2Rgb("FF0109", 0) -> 255; The red color
 *  =Hex2Rgb("FF0109", "Red") -> 255; The red color
 *  =Hex2Rgb("FF0109", 1) -> 1; The green color
 *  =Hex2Rgb("FF0109", "Green") -> 1; The green color
 *  =Hex2Rgb("FF0109", 2) -> 9; The blue color
 *  =Hex2Rgb("FF0109", "Blue") -> 9; The blue color


========================================

### Hsl2Hex

This function converts an HSL color value into a HEX color value.

**Arguments**
 * **hueValue** {Double}: is the hue value
 * **saturationValue** {Double}: is the saturation value
 * **lightnessValue** {Double}: is the lightness value

**Returns**
 * {Variant}: Returns a string with the HEX value of the color

**Examples**
 *  =Hsl2Rgb(212, .882, .267) -> "(8, 64, 128)"


========================================

### Hsl2Rgb

This function converts an HSL color value into an RGB color value, or optionally a single value from the RGB value.

**Arguments**
 * **hueValue** {Double}: is the hue value
 * **saturationValue** {Double}: is the saturation value
 * **lightnessValue** {Double}: is the lightness value
 * (Optional) _[**singleColorNumberOrName** {Variant = -1}]_: is a number or string that specifies which of the single color values to return. If this is set to 0 or "Red", the red value will be returned. If this is set to 1 or "Green", the green value will be returned. If this is set to 2 or "Blue", the blue value will be returned.

**Returns**
 * {Variant}: Returns a string with the RGB value of the color or an individual color value

**Examples**
 *  =Hsl2Rgb(212, .882, .267) -> "(8, 64, 128)"
 *  =Hsl2Rgb(212, .882, .267, 0) -> 8
 *  =Hsl2Rgb(212, .882, .267, "Red") -> 8
 *  =Hsl2Rgb(212, .882, .267, 1) -> 64
 *  =Hsl2Rgb(212, .882, .267, "Green") -> 64
 *  =Hsl2Rgb(212, .882, .267, 2) -> 128
 *  =Hsl2Rgb(212, .882, .267, "Blue") -> 128


========================================

### Rgb2Hex

This function converts an RGB color value into a HEX color value

**Arguments**
 * **redColorInteger** {Integer}: is the red value
 * **greenColorInteger** {Integer}: is the green value
 * **blueColorInteger** {Integer}: is the blue value

**Returns**
 * {String}: Returns a string with the HEX value of the color

**Examples**
 *  =Rgb2Hex(255, 255, 255) -> "FFFFFF"


========================================

### Rgb2Hsl

This function converts an RGB color value into an HSL color value and returns a string of the HSL value, or optionally a single value from the HSL value.

**Arguments**
 * **redColorInteger** {Integer}: is the red value
 * **greenColorInteger** {Integer}: is the green value
 * **blueColorInteger** {Integer}: is the blue value
 * (Optional) _[**singleColorNumberOrName** {Variant = -1}]_: is a number or string that specifies which of the single color values to return. If this is set to 0 or "Hue", the hue value will be returned. If this is set to 1 or "Saturation", the saturation value will be returned. If this is set to 2 or "Lightness", the lightness value will be returned.

**Returns**
 * {Variant}: Returns a string with the HSL value of the color

**Examples**
 *  =Rgb2Hsl(8, 64, 128) -> "(212.0�, 88.2%, 26.7%)"
 *  =Rgb2Hsl(8, 64, 128, 0) -> 212
 *  =Rgb2Hsl(8, 64, 128, "Hue") -> 212
 *  =Rgb2Hsl(8, 64, 128, 1) -> .882
 *  =Rgb2Hsl(8, 64, 128, "Saturation") -> .882
 *  =Rgb2Hsl(8, 64, 128, 2) -> .267
 *  =Rgb2Hsl(8, 64, 128, "Lightness") -> .267


========================================

### Rgb2Hsv

This function converts an RGB color value into an HSV color value, or optionally a single value from the HSV value.

**Arguments**
 * **redColorInteger** {Integer}: is the red value
 * **greenColorInteger** {Integer}: is the green value
 * **blueColorInteger** {Integer}: is the blue value
 * (Optional) _[**singleColorNumberOrName** {Variant = -1}]_: is a number or string that specifies which of the single color values to return. If this is set to 0 or "Hue", the hue value will be returned. If this is set to 1 or "Saturation", the saturation value will be returned. If this is set to 2 or "Value", the value value will be returned.

**Returns**
 * {Variant}: Returns a string with the RGB value of the color or an individual color value

**Examples**
 *  =Rgb2Hsv(8, 64, 128) -> "(212.0, 93.8%, 50.2%)"
 *  =Rgb2Hsv(8, 64, 128, 0) -> 212
 *  =Rgb2Hsv(8, 64, 128, "Red") -> 212
 *  =Rgb2Hsv(8, 64, 128, 1) -> .938
 *  =Rgb2Hsv(8, 64, 128, "Green") -> .938
 *  =Rgb2Hsv(8, 64, 128, 2) -> .502
 *  =Rgb2Hsv(8, 64, 128, "Blue") -> .502


---

## DateTime Module

This module contains a set of functions for working with dates and times.

========================================

### DaysOfMonth

This function takes a month number or month name and returns the number of days in the month. Optionally, a year number can be specified. If no year number is provided, the current year will be used. Finally, note that the month name or number argument is optional and if omitted will use the current month.

**Arguments**
 * (Optional) _[**monthNumberOrName** {Variant}]_: is a number that should be between 1 and 12, with 1 being January and 12 being December, or the name of a Month, such as "January" or "March". If omitted the current month will be used.
 * (Optional) _[**yearNumber** {Integer}]_: is the year that will be used. If omitted, the current year will be used.

**Returns**
 * {Variant}: Returns the number of days in the month and year specified

**Examples**
 *  =DaysOfMonth() -> 31; Where the current month is January
 *  =DaysOfMonth(1) -> 31
 *  =DaysOfMonth("January") -> 31
 *  =DaysOfMonth(2, 2019) -> 28
 *  =DaysOfMonth(2, 2020) -> 29


========================================

### MonthName2

This function takes a month number and returns the name of the month.

**Arguments**
 * (Optional) _[**monthNumber** {Byte}]_: is a number that should be between 1 and 12, with 1 being January and 12 being December. If no monthNumber is given, the value will default to the current month.

**Returns**
 * {String}: Returns the month name as a string

**Examples**
 *  =MonthName2(1) -> "January"
 *  =MonthName2(3) -> "March"
 *  To get today's month name: =MonthName2()


========================================

### Quarter

This function takes a month as a number and returns the Quarter of the year the month resides.

**Arguments**
 * (Optional) _[**monthNumberOrName** {Variant}]_: is a number that should be between 1 and 12, with 1 being January and 12 being December, or the name of a Month, such as "January" or "March".

**Returns**
 * {Byte}: Returns the Quarter of the month as a number

**Examples**
 *  =Quarter(4) -> 2
 *  =Quarter("April") -> 2
 *  =Quarter(12) -> 4
 *  =Quarter("December") -> 4
 *  To get today's Quarter: =Quarter()


========================================

### TimeConverter

This function takes a date, and then a series of optional arguments for a number of seconds, minutes, hours, days, and years, and then converts the date given to a new date adding in the other date argument values.

**Arguments**
 * **date1** {Date}: is the original date that will be converted into a new date
 * (Optional) _[**secondsInteger** {Integer}]_: is the number of seconds that will be added
 * (Optional) _[**minutesInteger** {Integer}]_: is the number of minutes that will be added
 * (Optional) _[**hoursInteger** {Integer}]_: is the number of hours that will be added
 * (Optional) _[**daysInteger** {Integer}]_: is the number of days that will be added
 * (Optional) _[**monthsInteger** {Integer}]_: is the number of months that will be added
 * (Optional) _[**yearsInteger** {Integer}]_: is the number of years that will be added

**Returns**
 * {Date}: Returns a new date with all the date arguments added to it

**Examples**
 *  =TimeConverter(A1,60) -> 1/1/2000 1:01; Where A1 contains the date 1/1/2000 1:00
 *  =TimeConverter(A1,0,5) -> 1/1/2000 1:05; Where A1 contains the date 1/1/2000 1:00
 *  =TimeConverter(A1,0,0,2) -> 1/1/2000 3:00; Where A1 contains the date 1/1/2000 1:00
 *  =TimeConverter(A1,0,0,0,4) -> 1/5/2000 1:00; Where A1 contains the date 1/1/2000 1:00
 *  =TimeConverter(A1,0,0,0,0,1) -> 2/1/2000 1:00; Where A1 contains the date 1/1/2000 1:00
 *  =TimeConverter(A1,0,0,0,0,0,5) -> 1/1/2005 1:00; Where A1 contains the date 1/1/2000 1:00
 *  =TimeConverter(A1,60,5,3,10,5,15) -> 6/11/2015 4:06; Where A1 contains the date 1/1/2000 1:00


========================================

### WeekOfMonth

This function takes a date and returns the number of the week of the month for that date. If no date is given, the current date is used.

**Arguments**
 * (Optional) _[**date1** {Date}]_: is a date whose week number will be found

**Returns**
 * {Byte}: Returns the number of week in the month

**Examples**
 *  =WeekOfMonth() -> 5; Where the current date is 1/29/2020
 *  =WeekOfMonth(1/29/2020) -> 5
 *  =WeekOfMonth(1/28/2020) -> 5
 *  =WeekOfMonth(1/27/2020) -> 5
 *  =WeekOfMonth(1/26/2020) -> 5
 *  =WeekOfMonth(1/25/2020) -> 4
 *  =WeekOfMonth(1/24/2020) -> 4
 *  =WeekOfMonth(1/1/2020) -> 1


========================================

### WeekdayName2

This function takes a weekday number and returns the name of the day of the week.

**Arguments**
 * (Optional) _[**dayNumber** {Byte}]_: is a number that should be between 1 and 7, with 1 being Sunday and 7 being Saturday. If no dayNumber is given, the value will default to the current day of the week.

**Returns**
 * {String}: Returns the day of the week as a string

**Examples**
 *  =WeekdayName2(4) -> Wednesday
 *  To get today's weekday name: =WeekdayName2()


---

## Environment Module

This module contains a set of functions for gathering information on the environment that Excel is being run on, such as the UserName of the computer, the OS Excel is being run on, and other Environment Variable values.

========================================

### ComputerName

This function takes no arguments and returns a string of the COMPUTERNAME of the computer

**Arguments**
 * None

**Returns**
 * {String}: Returns a string of the computer name of the computer

**Examples**
 *  =ComputerName() -> "DESKTOP-XYZ1234"


========================================

### OS

This function returns the Operating System name. Currently it will return either "Windows" or "Mac" depending on the OS used.

**Arguments**
 * None

**Returns**
 * {String}: Returns the name of the Operating System

**Examples**
 *  =OS() -> "Windows"; When running this function on Windows
 *  =OS() -> "Mac"; When running this function on MacOS


========================================

### UserDomain

This function takes no arguments and returns a string of the USERDOMAIN of the computer

**Arguments**
 * None

**Returns**
 * {String}: Returns a string of the user domain of the computer

**Examples**
 *  =UserDomain() -> "DESKTOP-XYZ1234"


========================================

### UserName

This function takes no arguments and returns a string of the USERNAME of the computer

**Arguments**
 * None

**Returns**
 * {String}: Returns a string of the username

**Examples**
 *  =UserName() -> "Anthony"


---

## File Module

This module contains a set of functions for gathering info on files. It includes functions for gathering file info on the current workbook presentation, document, or database, as well as functions for reading and writing to files, and functions for manipulating file path strings.

========================================

### CountFiles

This function returns the number of files at the specified folder path. If no path is given, the current workbook path will be used.

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path to the file on the system, such as "C:\hello\world.txt"

**Returns**
 * {Integer}: Returns the number of files in the folder

**Examples**
 *  =CountFiles() -> 6
 *  =CountFiles("C:\hello") -> 10


========================================

### CountFilesAndFolders

This function returns the number of files and folders at the specified folder path. If no path is given, the current workbook path will be used.

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path to the file on the system, such as "C:\hello\world.txt"

**Returns**
 * {Integer}: Returns the number of files and folders in the folder

**Examples**
 *  =CountFilesAndFolders() -> 8
 *  =CountFilesAndFolders("C:\hello") -> 30


========================================

### CountFolders

This function returns the number of folders at the specified folder path. If no path is given, the current workbook path will be used.

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path to the file on the system, such as "C:\hello\world.txt"

**Returns**
 * {Integer}: Returns the number of folders in the folder

**Examples**
 *  =CountFolders() -> 2
 *  =CountFolders("C:\hello") -> 20


========================================

### CurrentFilePath

This function returns the path of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path of the file on the system, such as "C:\hello\world.txt"

**Returns**
 * {String}: Returns the path of the file as a string

**Examples**
 *  =CurrentFilePath() -> "C:\my\_excel\_files\MyWorkbook.xlsx"
 *  =CurrentFilePath("C:\hello\world.txt") -> "C:\hello\world.txt"
 *  =CurrentFilePath("vba.txt") -> "C:\hello\world.txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in


========================================

### FileCreationTime

This function returns the file creation time of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path to the file on the system, such as "C:\hello\world.txt"

**Returns**
 * {String}: Returns the file creation time of the file as a string

**Examples**
 *  =FileCreationTime() -> "1/1/2020 1:23:45 PM"
 *  =FileCreationTime("C:\hello\world.txt") -> "1/1/2020 5:55:55 PM"
 *  =FileCreationTime("vba.txt") -> "12/25/2000 1:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in


========================================

### FileDrive

This function returns the drive of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path to the file on the system, such as "C:\hello\world.txt"

**Returns**
 * {String}: Returns the file drive of the file as a string

**Examples**
 *  =FileDrive() -> "A:"; Where the current workbook resides on the A: drive
 *  =FileDrive("C:\hello\world.txt") -> "C:"
 *  =FileDrive("vba.txt") -> "B:"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in, and where the workbook resides in the B: drive


========================================

### FileExtension

This function returns the extension of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path of the file on the system, such as "C:\hello\world.txt"

**Returns**
 * {String}: Returns the extension of the file as a string

**Examples**
 *  =FileExtension() = "xlsx"
 *  =FileExtension("C:\hello\world.txt") -> "txt"
 *  =FileExtension("vba.txt") -> "txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in


========================================

### FileFolder

This function returns the path of the folder of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path to the file on the system, such as "C:\hello\world.txt"

**Returns**
 * {String}: Returns the path of the folder where the file resides in as a string

**Examples**
 *  =FileFolder() -> "C:\my\_excel\_files"
 *  =FileFolder("C:\hello\world.txt") -> "C:\hello"
 *  =FileFolder("vba.txt") -> "C:\my\_excel\_files"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in


========================================

### FileLastModifiedTime

This function returns the file last modified time of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path to the file on the system, such as "C:\hello\world.txt"

**Returns**
 * {String}: Returns the file last modified time of the file as a string

**Examples**
 *  =FileLastModifiedTime() -> "1/1/2020 2:23:45 PM"
 *  =FileLastModifiedTime("C:\hello\world.txt") -> "1/1/2020 7:55:55 PM"
 *  =FileLastModifiedTime("vba.txt") -> "12/25/2000 3:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in


========================================

### FileName

This function returns the name of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path to the file on the system, such as "C:\hello\world.txt"

**Returns**
 * {String}: Returns the name of the file as a string

**Examples**
 *  =FileName() -> "MyWorkbook.xlsm"
 *  =FileName("C:\hello\world.txt") -> "world.txt"
 *  =FileName("vba.txt") -> "vba.txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in


========================================

### FileSize

This function returns the file size of the file specified in the file path argument, with the option to set if the file size is returned in Bytes, Kilobytes, Megabytes, or Gigabytes. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path of the file on the system, such as "C:\hello\world.txt"
 * (Optional) _[**byteSize** {String}]_: is a string of value "KB", "MB", or "GB"

**Returns**
 * {Double}: Returns the size of the file as a Double

**Examples**
 *  =FileSize() -> 1024
 *  =FileSize(,"KB") -> 1
 *  =FileSize("vba.txt", "KB") -> 0.25; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in


========================================

### FileType

This function returns the file type of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used.

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path to the file on the system, such as "C:\hello\world.txt"

**Returns**
 * {String}: Returns the file type of the file as a string

**Examples**
 *  FileType() -> "Microsoft Excel Macro-Enabled Worksheet"
 *  FileType("C:\hello\world.txt") -> "Text Document"
 *  FileType("vba.txt") -> "Text Document"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in


========================================

### GetActivePath

This function returns the path of the folder of the office program that is calling this function. It currently supports Excel, Word, PowerPoint, Access, and Publisher.

**Arguments**
 * None

**Returns**
 * {String}: Returns a string of the current folder path

**Examples**
 *  =GetActivePath() -> "C:\Users\UserName\Documents\"; Where the file resides in the Documents folder


========================================

### GetActivePathAndName

This function returns the path of the file of the office program that is calling this function. It currently supports Excel, Word, PowerPoint, Access, and Publisher.

**Arguments**
 * None

**Returns**
 * {String}: Returns a string of the current path

**Examples**
 *  =GetActivePathAndName() -> "C:\Users\UserName\Documents\XLib.xlsm"


========================================

### GetFileNameByNumber

This function returns the name of a file in a folder given the number of the file in the list of all files

**Arguments**
 * (Optional) _[**filePath** {String}]_: is a string path to the file on the system, such as "C:\hello\world.txt"
 * (Optional) _[**fileNumber** {Integer = -1}]_: is the number of the file in the folder. For example, if there are 3 files in a folder, this should be a number between 1 and 3

**Returns**
 * {String}: Returns the name of the specified file

**Examples**
 *  =GetFileName(,1) -> "hello.txt"
 *  =GetFileName(,1) -> "world.txt"
 *  =GetFileName("C:\hello", 1) -> "one.txt"
 *  =GetFileName("C:\hello", 1) -> "two.txt"
 *  =GetFileName("C:\hello", 1) -> "three.txt"


========================================

### PathJoin

This function combines multiple strings into a file path by placing the path separator character between the arguments

**Arguments**
 * **pathArray** {Variant}: is an array of strings that will be combined into a path

**Returns**
 * {Variant}: Returns a string with the combined file path

**Examples**
 *  =PathJoin("C:", "hello", "world.txt") -> "C:\hello\world.txt"; On Windows
 *  =PathJoin("hello", "world.txt") -> "/hello/world.txt"; On Mac


========================================

### PathSeparator

This function returns the path separator character of the OS running this function

**Arguments**
 * None

**Returns**
 * {String}: undefined

**Examples**
 *  =PathSeparator() -> "\"; When running this code on Windows
 *  =PathSeparator() -> "/"; When running this code on Mac


========================================

### ReadFile

This function reads the file specified in the file path argument and returns it's contents. Optionally, a line number can be specified so that only a single line is read. If a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.

**Arguments**
 * **filePath** {String}: is a string path of the file on the system, such as "C:\hello\world.txt"
 * (Optional) _[**lineNumber** {Integer}]_: is the number of the line that will be read, and if left blank all the file contents will be read. Note that the first line starts at line number 1.

**Returns**
 * {String}: Returns the contents of the file as a string

**Examples**
 *  =ReadFile("C:\hello\world.txt") -> "Hello" World
 *  =ReadFile("vba.txt") -> "This is my VBA text file"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in
 *  =ReadFile("multline.txt", 1) -> "This is line 1";
 *  =ReadFile("multline.txt", 2) -> "This is line 2";


========================================

### WriteFile

This function creates and writes to the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.

**Arguments**
 * **filePath** {String}: is a string path of the file on the system, such as "C:\hello\world.txt"
 * **fileText** {String}: is the text that will be written to the file
 * (Optional) _[**appendModeFlag** {Boolean}]_: is a Boolean value that if set to TRUE will append to the existing file instead of creating a new file and writing over the contents.

**Returns**
 * {Boolean}: Returns a message stating the file written to successfully

**Examples**
 *  =WriteFile("C:\MyWorkbookFolder\hello.txt", "Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"
 *  =WriteFile("hello.txt", "Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"; Where the Workbook resides in "C:\MyWorkbookFolder\"


---

## Math Module

This module contains a set of basic mathematical functions where those functions don't already exist as base Excel functions.

========================================

### Ceil

This function takes a number and rounds it up to the nearest whole integer

**Arguments**
 * **number** {Double}: is the number that will be rounded up

**Returns**
 * {Long}: Returns the number rounded up to the nearest integer

**Examples**
 *  =Ceil(1.5) -> 2
 *  =Ceil(1.0001) -> 2
 *  =Ceil(1.0) -> 1
 *  =Ceil(1) -> 1


========================================

### Floor

This function takes a number and rounds it down to the nearest whole integer

**Arguments**
 * **number** {Double}: is the number that will be rounded down

**Returns**
 * {Long}: Returns the number rounded down to the nearest integer

**Examples**
 *  =Floor(1.9) -> 1
 *  =Floor(1.1) -> 1
 *  =Floor(1.0) -> 1
 *  =Floor(1) -> 1


========================================

### InterpolateNumber

This function takes three numbers, a starting number, an ending number, and an interpolation percent, and linearly interpolates the number at the given percentage between the starting and ending number.

**Arguments**
 * **startingNumber** {Double}: is the beginning number of the interpolation
 * **endingNumber** {Double}: is the ending number of the interpolation
 * **interpolationPercentage** {Double}: is the percentage that will be interpolated linearly between the startingNumber and the endingNumber

**Returns**
 * {Double}: Returns the linearly interpolated number between the two points

**Examples**
 *  =InterpolateNumber(10, 20, 0.5) -> 15; Where 0.5 would be 50% between 10 and 20
 *  =InterpolateNumber(16, 124, 0.64) -> 85.12; Where 0.64 would be 64% between 16 and 124


========================================

### InterpolatePercent

This function takes three numbers, a starting number, an ending number, and an interpolation number, and linearly interpolates the percentage location of the interpolated number between the starting and ending number.

**Arguments**
 * **startingNumber** {Double}: is the beginning number of the interpolation
 * **endingNumber** {Double}: is the ending number of the interpolation
 * **interpolationNumber** {Double}: is the number that will be interpolated linearly between the startingNumber and the endingNumber to calculate a percentage

**Returns**
 * {Double}: Returns the linearly interpolated percent between the two points given the interpolation number

**Examples**
 *  =InterpolatePercent(10, 18, 12) -> 0.25; As 12 is 25% of the way from 10 to 18
 *  =InterpolatePercent(10, 20, 15) -> 0.5; As 15 is 50% of the way from 10 to 20


========================================

### Max

This function takes multiple numbers or multiple arrays of numbers and returns the max number. This function also accounts for numbers that are formatted as strings by converting them into numbers

**Arguments**
 * **numbers** {Variant}: is a single number, multiple numbers, or multiple arrays of numbers

**Returns**
 * {Variant}: Returns the max number

**Examples**
 *  =Max(1, 2, 3) -> 3
 *  =Max(4.4, 5, "6") -> 6
 *  =Max(x) -> 3; Where x is an array with these values [1, 2.2, "3"]
 *  =Max(x, y, 10) -> 15; Where x = [1, 2.2, "3"] and y = [5, 15, -100]


========================================

### Min

This function takes multiple numbers or multiple arrays of numbers and returns the min number. This function also accounts for numbers that are formatted as strings by converting them into numbers

**Arguments**
 * **numbers** {Variant}: is a single number, multiple numbers, or multiple arrays of numbers

**Returns**
 * {Variant}: Returns the min number

**Examples**
 *  =Min(1, 2, 3) -> 1
 *  =Min(4.4, 5, "6") -> 4.4
 *  =Min(-1, -2, -3) -> -3
 *  =Min(x) -> 1; Where x is an array with these values [1, 2.2, "3"]
 *  =Min(x, y, 10) -> -100; Where x = [1, 2.2, "3"] and y = [5, 15, -100]


========================================

### ModFloat

This function performs modulus operations with floats as the Mod operator in VBA does not support floats.

**Arguments**
 * **numerator** {Double}: is the left value of the Mod
 * **denominator** {Double}: is the right value of the Mod

**Returns**
 * {Double}: Returns a double with ModFloat operator performed on it

**Examples**
 *  =ModFloat(3.55, 2) -> 1.55


---

## Meta Module

This module contains a set of functions that return information on the Xlib library, such as the version number, credits, and a link to the documentation.

========================================

### XlibCredits

This function returns credits for the XPlus library

**Arguments**
 * None

**Returns**
 * {String}: Returns the XPlus credits

**Examples**
 *  =XlibCredits() -> "Copyright (c) 2020 Anthony Mancini. XLib is Licensed under an MIT License."


========================================

### XlibDocumentation

This function returns a link to the Documentation for XPlus

**Arguments**
 * None

**Returns**
 * {String}: Returns the XPlus Documentation link

**Examples**
 *  =XlibDocumentation() -> "https://x-vba.com/xlib"


========================================

### XlibVersion

This function returns the version number of XPlus

**Arguments**
 * None

**Returns**
 * {String}: Returns the XPlus version number

**Examples**
 *  =XlibVersion() -> "1.0.0"; Where the version of XPlus you are using is 1.0.0


---

## Network Module

This module contains a set of functions for performing networking tasks such as performing HTTP requests and parsing HTML.

========================================

### Http

This function performs an HTTP request to the web and returns the response as a string. It provides many options to change the http method, provide data for a POST request, change the headers, handle errors for non-successful requests, and parse out text from a request using a light parsing language.

**Arguments**
 * **url** {String}: is a string of the URL of the website you want to fetch data from
 * (Optional) _[**httpMethod** {String = "GET"}]_: is a string with the http method, with the default being a GET request. For POST requests, use "POST", for PUT use "PUT", and for DELETE use "DELETE"
 * (Optional) _[**headers** {Variant}]_: is either an array or a Scripting Dictionary of headers that will be used in the request. For an array, the 1st, 3rd, 5th... will be used as the key and the 2nd, 4th, 6th... will be used as the values. For a Scripting Dictionary, the dictionary keys will be used as header keys, and the values as values. Finally, in the case when no headers are set, the User-Agent will be set to "XPlus" as a courtesy to the web server.
 * (Optional) _[**postData** {Variant = ""}]_: is a string that will contain data for a POST request
 * (Optional) _[**asyncFlag** {Boolean}]_: is a Boolean value that if set to TRUE will make the request asynchronous. By default requests will be synchronous, which will lock Excel while fetching but will also prevent errors when performing calculations based on fetched data.
 * (Optional) _[**statusErrorHandlerFlag** {Boolean}]_: is a Boolean value that if set to TRUE will result in a User-Defined Error String being returned for all non 200 requests that tells the user the status code that occured. This flag is useful in cases where requests need to be successful and if not errors should be thrown.
 * (Optional) _[**parseArguments** {Variant}]_: is an array of arguments that perform string parsing on the response. It uses a light scripting language that includes commands similar to the Excel Built-in LEFT(), RIGHT(), and MID() that allow you to parse the request before it gets returned. See the Note on the scripting language, and the Warning on why this argument should be used.

**Returns**
 * {String}: Returns the parsed HTTP response as a string

**Examples**
 *  =Http("https://httpbin.org/uuid") -> "{"uuid: "41416bcf-ef11-4256-9490-63853d14e4e8"}"
 *  =Http("https://httpbin.org/user-agent", "GET", {"User-Agent","MicrosoftExcel"}) -> "{"user-agent": "MicrosoftExcel"}"
 *  =Http("https://httpbin.org/status/404",,,,,TRUE) -> "#RequestFailedStatusCode404!"; Since the status error handler flag is set and since this URL returns a 404 status code. Also note that this formula is easier to construct using the Excel Formula Builder
 *  =Http("https://en.wikipedia.org/wiki/Visual\_Basic\_for\_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000}) -> Returning a string with the leftmost 3000 characters found within the element with the ID "mw-content-text" (we are trying to get the release date of VBA from the VBA wikipedia page, but we need to do more parsing first)
 *  =Http("https://en.wikipedia.org/wiki/Visual\_Basic\_for\_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000,"MID","appeared"}) -> Returns the prior string, but now with all characters right of the first occurance of the word "appeared" in the HTML (getting closer to parsing the VBA creation date)
 *  =Http("https://en.wikipedia.org/wiki/Visual\_Basic\_for\_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>"}) -> From the prior result, now returning everything after the first occurance of the "<TD>" in the prior string
 *  =Http("https://en.wikipedia.org/wiki/Visual\_Basic\_for\_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>","LEFT","<span"}) -> "1993"; Finally this is all the parsing needed to be able to return the date 1993 that we were looking for


========================================

### ParseHtmlString

This function parses an HTML string using the same parsing language that the HTTP() function uses. See the HTTP() function for more information on how to use this function.

**Arguments**
 * **htmlString** {String}: is a string of the HTML
 * **parseArguments** {Variant}: is an array of arguments that perform string parsing on the response. It uses a light scripting language that includes commands similar to the Excel Built-in LEFT(), RIGHT(), and MID() that allow you to parse the request before it gets returned. See the Note on the HTTP() function, and the Warning on the HTTP() function on why this argument should be used.

**Returns**
 * {Variant}: Returns the parsed HTTP response as a string

**Examples**
 *  =ParseHtmlString("HTML String from the webpage: https://en.wikipedia.org/wiki/Visual\_Basic\_for\_Applications","ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>","LEFT","<span") -> "1993"


========================================

### SimpleHttp

This function performs an HTTP request to the web and returns the response as a string, similar to the HTTP() function, except that only requires one parameter, the URL, and then takes an infinite number of strings after it as the parsing arguments instead of requiring an Array to use. Essentially, this function is a little cleaner to set up when performing very basic GET requests.

**Arguments**
 * **url** {String}: is a string of the URL of the website you want to fetch data from
 * **parseArguments** {Variant}: is an array of arguments that perform string parsing on the response. It uses a light scripting language that includes commands similar to the Excel Built-in LEFT(), RIGHT(), and MID() that allow you to parse the request before it gets returned. See the Note on the HTTP() function, and the Warning on the HTTP() function on why this argument should be used.

**Returns**
 * {Variant}: Returns the parsed HTTP response as a string

**Examples**
 *  =SimpleHttp("https://en.wikipedia.org/wiki/Visual\_Basic\_for\_Applications","ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>","LEFT","<span") -> "1993"; See the examples in the HTTP() function, as this example has the same result as the example in the HTTP() function. You can see that this function is cleaner and easier to set up than the corresponding HTTP() function.


---

## Random Module

This module contains a set of functions for generating and sampling random data.

========================================

### BigRandBetween

This function is an implementation of RandBetween that allows for 14-byte integers to be used

**Arguments**
 * **minNumber** {Variant}: is the minimum number in the range
 * **maxNumber** {Variant}: is the maximum number in the range

**Returns**
 * {Variant}: Returns a random number between the range

**Examples**
 *  =RandBetween(0, 3000000000) -> Error; as RandBetween only works with 4-byte and less integers
 *  =BigRandBetween(0, 3000000000) -> 2116642535; as BigRandBetween supports up to 14-byte integers


========================================

### RandBetween

This function returns a random number between the min and max numbers

**Arguments**
 * **minNumber** {Long}: is the minimum number in the range
 * **maxNumber** {Long}: is the maximum number in the range

**Returns**
 * {Variant}: Returns a random number between the range

**Examples**
 *  =RandBetween(1, 20) -> 5
 *  =RandBetween(1, 20) -> 9
 *  =RandBetween(1, 20) -> 13
 *  =RandBetween(1, 20) -> 2
 *  =RandBetween(1, 20) -> 20
 *  =RandBetween(1, 20) -> 6


========================================

### RandBetweens

This function is similar to RANDBETWEEN, except that it allows multiple ranges from which to pick a random number. One of the ranges from which to generate a random number between is chosen at an equal probably.

**Arguments**
 * **startOrEndNumberArray** {Variant}: undefined

**Returns**
 * {Variant}: Returns either TRUE or FALSE based on the random value choosen

**Examples**
 *  =RandBetweens(1, 10, 5000, 5010) -> 6
 *  =RandBetweens(1, 10, 5000, 5010) -> 5002
 *  =RandBetweens(1, 10, 5000, 5010) -> 8
 *  =RandBetweens(1, 10, 5000, 5010) -> 3
 *  =RandBetweens(1, 10, 5000, 5010) -> 5010
 *  =RandBetweens(1, 10, 5000, 5010) -> 2
 *  =RandBetweens(5, 10, 15, 20, 25, 30, 35, 40) -> 32


========================================

### RandBool

This function generates a random Boolean (TRUE or FALSE) value

**Arguments**
 * None

**Returns**
 * {Boolean}: Returns either TRUE or FALSE based on the random value choosen

**Examples**
 *  =RandBool() -> TRUE
 *  =RandBool() -> FALSE
 *  =RandBool() -> TRUE
 *  =RandBool() -> TRUE
 *  =RandBool() -> FALSE
 *  =RandBool() -> FALSE


========================================

### RandomRange

This function takes 3 numbers, a start number, a stop number, and a step number, and returns a random number between the start number and stop number that is an interval of the step number.

**Arguments**
 * **startNumber** {Long}: is the beginning value of the range
 * **stopNumber** {Long}: is the end value of the range
 * **stepNumber** {Long}: is the step of the range

**Returns**
 * {Long}: Returns a random number between the start and stop that is a multiple of the step

**Examples**
 *  =RandomRange(50, 100, 10) -> 60
 *  =RandomRange(50, 100, 10) -> 50
 *  =RandomRange(50, 100, 10) -> 90
 *  =RandomRange(0, 10, 2) -> 8
 *  =RandomRange(0, 10, 2) -> 0
 *  =RandomRange(0, 10, 2) -> 4
 *  =RandomRange(0, 10, 2) -> 10


========================================

### RandomSample

This function takes an array of cells and returns a random value from the cells chosen

**Arguments**
 * **variantArray** {Variant}: a single cell or multiple cells where the sample will be pulled from

**Returns**
 * {Variant}: Returns a random cell value from the array of cells chosen

**Examples**
 *  =RandomSample(A1:A5) -> "Hello"; where "Hello" is the value in cell A3, and where A3 was the chosen random cell
 *  =RandomSample(A1:A5) -> "World"; where "World" is the value in cell A2, and where A2 was the chosen random cell


---

## Regex Module

This module contains a set of functions for performing Regular Expressions, which are a type of string pattern matching. For more info on Regular Expression Pattern matching, please check "https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference"

========================================

### RegexReplace

This function takes a string that we will perform the Regular Expression on, a Regular Expression string pattern, and a string that we will replace if the pattern is found, and returns a new string with the replacement string in place of the pattern. This function also contains optional arguments for various Regular Expression flags.

**Arguments**
 * **string1** {String}: is the string that the regex will be performed on
 * **stringPattern** {String}: is the regex pattern
 * **replacementString** {String}: is a string that will be replaced if the pattern is found
 * (Optional) _[**globalFlag** {Boolean}]_: is a boolean value that if set TRUE will perform a global search
 * (Optional) _[**ignoreCaseFlag** {Boolean}]_: is a boolean value that if set TRUE will perform a case insensitive search
 * (Optional) _[**multilineFlag** {Boolean}]_: is a boolean value that if set TRUE will perform a mulitline search

**Returns**
 * {String}: Returns a new string with the replaced string values

**Examples**
 *  =RegexReplace("Hello World","[W][a-z]{4}", "VBA") -> "Hello VBA"


========================================

### RegexSearch

This function takes a string that we will perform the Regular Expression on and a Regular Expression string pattern, and returns the first value of the matched string. This function also contains optional arguments for various Regular Expression flags.

**Arguments**
 * **string1** {String}: is the string that the regex will be performed on
 * **stringPattern** {String}: is the regex pattern
 * (Optional) _[**globalFlag** {Boolean}]_: is a boolean value that if set TRUE will perform a global search
 * (Optional) _[**ignoreCaseFlag** {Boolean}]_: is a boolean value that if set TRUE will perform a case insensitive search
 * (Optional) _[**multilineFlag** {Boolean}]_: is a boolean value that if set TRUE will perform a mulitline search

**Returns**
 * {String}: Returns a string of the regex value that is found

**Examples**
 *  =RegexSearch("Hello World","[a-z]{2}\s[W]") -> "lo W";


========================================

### RegexTest

This function takes a string that we will perform the Regular Expression on and a Regular Expression string pattern, and returns TRUE if the pattern is found in the string. This function also contains optional arguments for various Regular Expression flags.

**Arguments**
 * **string1** {String}: is the string that the regex will be performed on
 * **stringPattern** {String}: is the regex pattern
 * (Optional) _[**globalFlag** {Boolean}]_: is a boolean value that if set TRUE will perform a global search
 * (Optional) _[**ignoreCaseFlag** {Boolean}]_: is a boolean value that if set TRUE will perform a case insensitive search
 * (Optional) _[**multilineFlag** {Boolean}]_: is a boolean value that if set TRUE will perform a mulitline search

**Returns**
 * {Boolean}: Returns TRUE if the regex value that is found, or FALSE if it isn't

**Examples**
 *  =RegexTest("Hello World","[a-z]{2}\s[W]") -> TRUE;


---

## StringManipulation Module

This module contains a set of basic functions for manipulating strings.

========================================

### CamelCase

This function takes a string and returns the same string in camel case, removing all the spaces.

**Arguments**
 * **string1** {String}: is the string that will be camel cased

**Returns**
 * {String}: Returns a new string in camel case, where the first character of the first word is lowercase, and uppercased for all other words

**Examples**
 *  =CamelCase("Hello World") -> "helloWorld"
 *  =CamelCase("One Two Three") -> "oneTwoThree"


========================================

### Capitalize

This function takes a string and returns the same string with the first character capitalized and all other characters lowercased

**Arguments**
 * **string1** {String}: is the string that the capitalization will be performed on

**Returns**
 * {String}: Returns a new string with the first character capitalized and all others lowercased

**Examples**
 *  =Capitalize("hello World") -> "Hello world"


========================================

### CompanyCase

This function takes a string and uses an algorithm to return the string in Company Case. The standard =PROPER() function in Excel will not capitalize company names properly, as it only capitalizes based on space characters, so a name like "j.p. morgan" will be incorrectly formatted as "J.p. Morgan" instead of the correct "J.P. Morgan". Additionally =PROPER() may incorrectly lowercase company abbreviations, such as the last "H" in "GmbH", as =PROPER() returns "Gmbh" instead of the correct "GmbH". This function attempts to adjust for these issues when a string is a company name.

**Arguments**
 * **string1** {String}: is the string that will be formatted

**Returns**
 * {String}: Returns the origional string in a Company Case format

**Examples**
 *  =CompanyCase("hello world") -> "Hello World"
 *  =CompanyCase("x.y.z company & co.") -> "X.Y.Z Company & Co."
 *  =CompanyCase("x.y.z plc") -> "X.Y.Z PLC"
 *  =CompanyCase("one company gmbh") -> "One Company GmbH"
 *  =CompanyCase("three company s. en n.c.") -> "Three Company S. en N.C."
 *  =CompanyCase("FOUR COMPANY SPOL S.R.O.") -> "Four Company spol s.r.o."
 *  =CompanyCase("five company bvba") -> "Five Company BVBA"


========================================

### CountLowercaseCharacters

This function takes a string and counts the number of lowercase characters in it

**Arguments**
 * **string1** {String}: is the string whose characters will be counted

**Returns**
 * {Integer}: Returns the number of lowercase characters in the string

**Examples**
 *  =CountLowercaseCharacters("Hello World") -> 8; As the "ello" and the "orld" are lowercase


========================================

### CountUppercaseCharacters

This function takes a string and counts the number of uppercase characters in it

**Arguments**
 * **string1** {String}: is the string whose characters will be counted

**Returns**
 * {Integer}: Returns the number of uppercase characters in the string

**Examples**
 *  =CountUppercaseCharacters("Hello World") -> 2; As the "H" and the "E" are the only 2 uppercase characters in the string


========================================

### CountWords

This function takes a string and returns the number of words in the string

**Arguments**
 * **string1** {String}: is the string whose number of words will be counted
 * (Optional) _[**delimiterString** {String = " "}]_: is an optional parameter that can be used to specify a different delimiter

**Returns**
 * {Integer}: Returns the number of words in the string

**Examples**
 *  =CountWords("Hello World") -> 2
 *  =CountWords("One Two Three") -> 3
 *  =CountWords("One-Two-Three", "-") -> 3


========================================

### DedentText

This function takes a string and dedents all of its lines so that there are no space characters to the left or right of each line

**Arguments**
 * **string1** {String}: is the string that will be dedented

**Returns**
 * {String}: Returns the origional string dedented on each line

**Examples**
 *  =DedentText("    Hello") -> "Hello"


========================================

### EliteCase

This function takes a string and returns the string with characters replaced by similar in appearance numbers

**Arguments**
 * **string1** {String}: is the string that will have characters replaced

**Returns**
 * {String}: Returns the string with characters replaced with similar in appearance numbers

**Examples**
 *  =EliteCase("Hello World") -> "H3110 W0r1d"


========================================

### Formatter

This function takes a Formatter string and then an array of ranges or strings, and replaces the format placeholders with the values in the range or strings. The format syntax is "{1} - {2}" where the "{1}" and "{2}" will be replaced with the values given in the text array.

**Arguments**
 * **formatString** {String}: is the string that will be used as the format and which will be replaced with the individual strings
 * **textArray** {Variant}: are the ranges or strings that will be placed within the slots of the format string

**Returns**
 * {Variant}: Returns a new string with the individual strings in the placeholder slots of the format string

**Examples**
 *  =Formatter("Hello {1}", "World") -> "Hello World"
 *  =Formatter("{1} {2}", "Hello", "World") -> "Hello World"
 *  =Formatter("{1}.{2}@{3}", "FirstName", "LastName", "email.com") -> "FirstName.LastName@email.com"
 *  =Formatter("{1}.{2}@{3}", A1:A3) -> "FirstName.LastName@email.com"; where A1="FirstName", A2="LastName", and A3="email.com"
 *  =Formatter("{1}.{2}@{3}", A1, A2, A3) -> "FirstName.LastName@email.com"; where A1="FirstName", A2="LastName", and A3="email.com"


========================================

### InSplit

This function takes a search string and checks if it exists within a larger string that is split by a delimiter character.

**Arguments**
 * **string1** {String}: is the string that will be checked if it exists within the splitString after the split
 * **splitString** {String}: is the string that will be split and of which string1 will be searched in
 * (Optional) _[**delimiterCharacter** {String = " "}]_: is the character that will be used as the delimiter for the split. By default this is the space character " "

**Returns**
 * {Boolean}: Returns TRUE if string1 is found in splitString after the split occurs

**Examples**
 *  =InSplit("Hello", "Hello World One Two Three") -> TRUE; Since "Hello" is found within the searchString after being split
 *  =InSplit("NotInString", "Hello World One Two Three") -> FALSE; Since "NotInString" is not found within the searchString after being split
 *  =InSplit("Hello", "Hello-World-One-Two-Three", "-") -> TRUE; Since "Hello" is found and since the delimiter is set to "-"


========================================

### IndentText

This function takes a string and indents all of its lines by a specified number of space characters (or 4 space characters if left blank)

**Arguments**
 * **string1** {String}: is the string that will be indented
 * (Optional) _[**indentAmount** {Byte = 4}]_: is the amount of " " characters that will be indented to the left of string1

**Returns**
 * {String}: Returns the origional string indented by a specified number of space characters

**Examples**
 *  =IndentText("Hello") -> "    Hello"
 *  =IndentText("Hello", 4) -> "    Hello"
 *  =IndentText("Hello", 3) -> "   Hello"
 *  =IndentText("Hello", 2) -> "  Hello"
 *  =IndentText("Hello", 1) -> " Hello"


========================================

### KebabCase

This function takes a string and returns the same string in kebab case.

**Arguments**
 * **string1** {String}: is the string that will be kebab cased

**Returns**
 * {String}: Returns a new string in kebab case, where all letters are lowercase and seperated by a "-" character

**Examples**
 *  =KebabCase("Hello World") -> "hello-world"
 *  =KebabCase("One Two Three") -> "one-two-three"


========================================

### LeftFind

This function takes a string and a search string, and returns a string with all characters to the left of the first search string found within string1. Similar to Excel's built-in =SEARCH() function, this function is case-sensitive. For a case-insensitive version of this function, see =LeftSearch().

**Arguments**
 * **string1** {String}: is the string that will be searched
 * **searchString** {String}: is the string that will be used to search within string1

**Returns**
 * {String}: Returns a new string with all characters to the left of the first search string within string1

**Examples**
 *  =LeftFind("Hello World", "r") -> "Hello Wo"
 *  =LeftFind("Hello World", "R") -> "#VALUE!"; Since string1 does not contain "R" in it.


========================================

### LeftSearch

This function takes a string and a search string, and returns a string with all characters to the left of the first search string found within string1. Similar to Excel's built-in =FIND() function, this function is NOT case-sensitive (it's case-insensitive). For a case-sensitive version of this function, see =LeftFind().

**Arguments**
 * **string1** {String}: is the string that will be searched
 * **searchString** {String}: is the string that will be used to search within string1

**Returns**
 * {String}: Returns a new string with all characters to the left of the first search string within string1

**Examples**
 *  =LeftSearch("Hello World", "r") -> "Hello Wo"
 *  =LeftSearch("Hello World", "R") -> "Hello Wo"


========================================

### LeftSplit

This function takes a string, splits it based on a delimiter, and returns all characters to the left of the specified position of the split.

**Arguments**
 * **string1** {String}: is the string that will be split to get a substring
 * **numberOfSplit** {Integer}: is the number of the location within the split that we will get all characters to the left of
 * (Optional) _[**delimiterCharacter** {String = " "}]_: is the delimiter that will be used for the split. By default, the delimiter will be the space character " "

**Returns**
 * {String}: Returns all characters to the left of the number of the split

**Examples**
 *  =LeftSplit("Hello World One Two Three", 1) -> "Hello"
 *  =LeftSplit("Hello World One Two Three", 2) -> "Hello World"
 *  =LeftSplit("Hello World One Two Three", 3) -> "Hello World One"
 *  =LeftSplit("Hello World One Two Three", 10) -> "Hello World One Two Three"
 *  =LeftSplit("Hello-World-One-Two-Three", 2, "-") -> "Hello-World"


========================================

### RemoveCharacters

This function takes a string and either another string or multiple strings and removes all characters from the first string that are in the second string.

**Arguments**
 * **string1** {String}: is the string that will have characters removed
 * **removedCharacters** {Variant}: is an array of strings that will be removed from string1

**Returns**
 * {Variant}: Returns the origional string with characters removed

**Examples**
 *  =RemoveCharacters("Hello World", "l") -> "Heo Word"
 *  =RemoveCharacters("Hello World", "lo") -> "He Wrd"
 *  =RemoveCharacters("Hello World", "l", "o") -> "He Wrd"
 *  =RemoveCharacters("Hello World", "lod") -> "He Wr"
 *  =RemoveCharacters("Two Three Four", "f", "t") -> "Two Three Four"; Nothing is replaced since this function is case sensitive
 *  =RemoveCharacters("Two Three Four", "F", "T") -> "wo hree our"


========================================

### Repeat

This function repeats string1 based on the number of repeats specified in the second argument

**Arguments**
 * **string1** {String}: is the string that will be repeated
 * **numberOfRepeats** {Integer}: is the number of times string1 will be repeated

**Returns**
 * {String}: Returns a string repeated multiple times based on the numberOfRepeats

**Examples**
 *  =Repeat("Hello", 2) -> HelloHello"
 *  =Repeat("=", 10) -> "=========="


========================================

### ReverseText

This function takes a string and reverses all the characters in it so that the returned string is backwards

**Arguments**
 * **string1** {String}: is the string that will be reversed

**Returns**
 * {String}: Returns the origional string in reverse

**Examples**
 *  =ReverseText("Hello World") -> "dlroW olleH"


========================================

### ReverseWords

This function takes a string and reverses all the words in it so that the returned string's words are backwards. By default, this function uses the space character as a delimiter, but you can optionally specify a different delimiter.

**Arguments**
 * **string1** {String}: is the string whose words will be reversed
 * (Optional) _[**delimiterCharacter** {String = " "}]_: is the delimiter that will be used, with the default being " "

**Returns**
 * {String}: Returns the origional string with it's words reversed

**Examples**
 *  =ReverseWords("Hello World") -> "World Hello"
 *  =ReverseWords("One Two Three") -> "Three Two One"
 *  =ReverseWords("One-Two-Three", "-") -> "Three-Two-One"


========================================

### RightFind

This function takes a string and a search string, and returns a string with all characters to the right of the last search string found within string 1. Similar to Excel's built-in =SEARCH() function, this function is case-sensitive. For a case-insensitive version of this function, see =RightSearch().

**Arguments**
 * **string1** {String}: is the string that will be searched
 * **searchString** {String}: is the string that will be used to search within string1

**Returns**
 * {String}: Returns a new string with all characters to the right of the last search string within string1

**Examples**
 *  =RightFind("Hello World", "o") -> "rld"
 *  =RightFind("Hello World", "O") -> "#VALUE!"; Since string1 does not contain "O" in it.


========================================

### RightSearch

This function takes a string and a search string, and returns a string with all characters to the right of the last search string found within string 1. Similar to Excel's built-in =FIND() function, this function is NOT case-sensitive (it's case-insensitive). For a case-sensitive version of this function, see =RightFind().

**Arguments**
 * **string1** {String}: is the string that will be searched
 * **searchString** {String}: is the string that will be used to search within string1

**Returns**
 * {String}: Returns a new string with all characters to the right of the last search string within string1

**Examples**
 *  =RightSearch("Hello World", "o") -> "rld"
 *  =RightSearch("Hello World", "O") -> "rld"


========================================

### RightSplit

This function takes a string, splits it based on a delimiter, and returns all characters to the right of the specified position of the split.

**Arguments**
 * **string1** {String}: is the string that will be split to get a substring
 * **numberOfSplit** {Integer}: is the number of the location within the split that we will get all characters to the right of
 * (Optional) _[**delimiterCharacter** {String = " "}]_: is the delimiter that will be used for the split. By default, the delimiter will be the space character " "

**Returns**
 * {String}: Returns all characters to the right of the number of the split

**Examples**
 *  =RightSplit("Hello World One Two Three", 1) -> "Three"
 *  =RightSplit("Hello World One Two Three", 2) -> "Two Three"
 *  =RightSplit("Hello World One Two Three", 3) -> "One Two Three"
 *  =RightSplit("Hello World One Two Three", 10) -> "Hello World One Two Three"
 *  =RightSplit("Hello-World-One-Two-Three", 2, "-") -> "Two-Three"


========================================

### ScrambleCase

This function takes a string scrambles the case on each character in the string

**Arguments**
 * **string1** {String}: is the string whose character's cases will be scrambled

**Returns**
 * {String}: Returns the origional string with cases scrambled

**Examples**
 *  =ScrambleCase("Hello World") -> "helLo WORlD"
 *  =ScrambleCase("Hello World") -> "HElLo WorLD"
 *  =ScrambleCase("Hello World") -> "hELlo WOrLd"


========================================

### ShortenText

This function takes a string and shortens it with placeholder text so that it is no longer in length than the specified width.

**Arguments**
 * **string1** {String}: is the string that will be shortened
 * (Optional) _[**shortenWidth** {Integer = 80}]_: is the max width of the string. By default this is set to 80
 * (Optional) _[**placeholderText** {String = "[...]"}]_: is the text that will be placed at the end of the string if it is longer than the shortenWidth. By default this placeholder string is "[...]
 * (Optional) _[**delimiterCharacter** {String = " "}]_: is the character that will be used as the word delimiter. By default this is the space character " "

**Returns**
 * {String}: Returns a shortened string with placeholder text if it is longer than the shorten width

**Examples**
 *  =ShortenText("Hello World One Two Three", 20) -> "Hello World [...]"; Only the first two words and the placeholder will result in a string that is less than or equal to 20 in length
 *  =ShortenText("Hello World One Two Three", 15) -> "Hello [...]"; Only the first word and the placeholder will result in a string that is less than or equal to 15 in length
 *  =ShortenText("Hello World One Two Three") -> "Hello World One Two Three"; Since this string is shorter than the default 80 shorten width value, no placeholder will be used and the string wont be shortened
 *  =ShortenText("Hello World One Two Three", 15, "-->") -> "Hello World -->"; A new placeholder is used
 *  =ShortenText("Hello\_World\_One\_Two\_Three", 15, "-->", "\_") -> "Hello\_World\_-->"; A new placeholder andd delimiter is used


========================================

### SplitText

This function takes a string and a number, splits the string by the space characters, and returns the substring in the position of the number specified in the second argument.

**Arguments**
 * **string1** {String}: is the string that will be split and a substring returned
 * **substringNumber** {Integer}: is the number of the substring that will be chosen
 * (Optional) _[**delimiterString** {String = " "}]_: is an optional parameter that can be used to specify a different delimiter

**Returns**
 * {String}: Returns a substring of the split text in the location specified

**Examples**
 *  =SplitText("Hello World", 1) -> "Hello"
 *  =SplitText("Hello World", 2) -> "World"
 *  =SplitText("One Two Three", 2) -> "Two"
 *  =SplitText("One-Two-Three", 2, "-") -> "Two"


========================================

### Substr

This function takes a string and a starting character number and ending character number, and returns the substring between these two numbers. The total number of characters returned will be endCharacterNumber - startCharacterNumber.

**Arguments**
 * **string1** {String}: is the string that we will get a substring from
 * **startCharacterNumber** {Integer}: is the character number of the start of the substring, with 1 being the first character in the string
 * **endCharacterNumber** {Integer}: is the character number of the end of the substring

**Returns**
 * {String}: Returns a substring between the two numbers.

**Examples**
 *  =Substr("Hello World", 2, 6) -> "ello"


========================================

### SubstrFind

This function takes a string and a left string and right string, and returns a substring between those two strings. The left string will find the first matching string starting from the left, and the right string will find the first matching string starting from the right. Finally, and optional final parameter can be set to TRUE to make the substring noninclusive of the two searched strings. SubstrFind is case-sensitive. For case-insensitive version, see SubstrSearch

**Arguments**
 * **string1** {String}: is the string that we will get a substring from
 * **RightFindString** {String}: is the string that will be searched from the left
 * **rightSearchString** {String}: is the string that will be searched from the right
 * (Optional) _[**noninclusiveFlag** {Boolean}]_: is an optional parameter that if set to TRUE will result in the substring not including the left and right searched characters

**Returns**
 * {String}: Returns a substring between the two strings.

**Examples**
 *  =SubstrFind("Hello World", "e", "o") -> "ello Wo"
 *  =SubstrFind("Hello World", "e", "o", TRUE) -> "llo W"
 *  =SubstrFind("One Two Three", "ne ", " Thr") -> "ne Two Thr"
 *  =SubstrFind("One Two Three", "NE ", " THR") -> "#VALUE!"; Since SubstrFind() is case-sensitive
 *  =SubstrFind("One Two Three", "ne ", " Thr", TRUE) -> "Two"
 *  =SubstrFind("Country Code: +51; Area Code: 315; Phone Number: 762-5929;", "Area Code: ", "; Phone", TRUE) -> 315
 *  =SubstrFind("Country Code: +313; Area Code: 423; Phone Number: 284-2468;", "Area Code: ", "; Phone", TRUE) -> 423
 *  =SubstrFind("Country Code: +171; Area Code: 629; Phone Number: 731-5456;", "Area Code: ", "; Phone", TRUE) -> 629


========================================

### SubstrSearch

This function takes a string and a left string and right string, and returns a substring between those two strings. The left string will find the first matching string starting from the left, and the right string will find the first matching string starting from the right. Finally, and optional final parameter can be set to TRUE to make the substring noninclusive of the two searched strings. SubstrSearch is case-insensitive. For case-sensitive version, see SubstrFind

**Arguments**
 * **string1** {String}: is the string that we will get a substring from
 * **RightFindString** {String}: is the string that will be searched from the left
 * **rightSearchString** {String}: is the string that will be searched from the right
 * (Optional) _[**noninclusiveFlag** {Boolean}]_: is an optional parameter that if set to TRUE will result in the substring not including the left and right searched characters

**Returns**
 * {String}: Returns a substring between the two strings.

**Examples**
 *  =SubstrSearch("Hello World", "e", "o") -> "ello Wo"
 *  =SubstrSearch("Hello World", "e", "o", TRUE) -> "llo W"
 *  =SubstrSearch("One Two Three", "ne ", " Thr") -> "ne Two Thr"
 *  =SubstrSearch("One Two Three", "NE ", " THR") -> "ne Two Thr"; No error, since SubstrSearch is case-insensitive
 *  =SubstrSearch("One Two Three", "ne ", " Thr", TRUE) -> "Two"
 *  =SubstrSearch("Country Code: +51; Area Code: 315; Phone Number: 762-5929;", "Area Code: ", "; Phone", TRUE) -> 315
 *  =SubstrSearch("Country Code: +313; Area Code: 423; Phone Number: 284-2468;", "Area Code: ", "; Phone", TRUE) -> 423
 *  =SubstrSearch("Country Code: +171; Area Code: 629; Phone Number: 731-5456;", "Area Code: ", "; Phone", TRUE) -> 629


========================================

### TextJoin

This function takes a range of cells and combines all the text together, optionally allowing a character delimiter between all the combined strings, and optionally allowing blank cells to be ignored when combining the text. Finally note that this function is very similar to the TEXTJOIN function available in Excel 2019, and thus is a polyfill for that function for earlier versions of Excel.

**Arguments**
 * **stringArray** {Variant}: is the range with all the strings we want to combine
 * (Optional) _[**delimiterCharacter** {String}]_: is an optional character that will be used as the delimiter between the combined text. By default, no delimiter character will be used.
 * (Optional) _[**ignoreEmptyCellsFlag** {Boolean}]_: if set to TRUE will skip combining empty cells into the combined string, and is useful when specifying a delimiter so that the delimiter does not repeat for empty cells.

**Returns**
 * {String}: Returns a new combined string containing the strings in the range delimited by the delimiter character.

**Examples**
 *  =TextJoin(A1:A3) -> "123"; Where A1:A3 contains ["1", "2", "3"]
 *  =TextJoin(A1:A3, "--") -> "1--2--3"; Where A1:A3 contains ["1", "2", "3"]
 *  =TextJoin(A1:A3, "--") -> "1----3"; Where A1:A3 contains ["1", "", "3"]
 *  =TextJoin(A1:A3, "-") -> "1--3"; Where A1:A3 contains ["1", "", "3"]
 *  =TextJoin(A1:A3, "-", TRUE) -> "1-3"; Where A1:A3 contains ["1", "", "3"]


========================================

### TrimChar

This function takes a string trims characters to the left and right of the string, similar to Excel's Built-in TRIM() function, except that an optional different character can be used for the trim.

**Arguments**
 * **string1** {String}: is the string that will be trimmed
 * (Optional) _[**trimCharacter** {String = " "}]_: is an optional character that will be trimmed from the string. By default, this character will be the space character " "

**Returns**
 * {String}: Returns the origional string with characters trimmed from the left and right

**Examples**
 *  =TrimChar("   Hello World   ") -> "Hello World"
 *  =TrimChar("---Hello World---", "-") -> "Hello World"


========================================

### TrimLeft

This function takes a string trims characters to the left of the string, similar to Excel's Built-in TRIM() function, except that an optional different character can be used for the trim.

**Arguments**
 * **string1** {String}: is the string that will be trimmed
 * (Optional) _[**trimCharacter** {String = " "}]_: is an optional character that will be trimmed from the string. By default, this character will be the space character " "

**Returns**
 * {String}: Returns the origional string with characters trimmed from the left only

**Examples**
 *  =TrimLeft("   Hello World   ") -> "Hello World   "
 *  =TrimLeft("---Hello World---", "-") -> "Hello World---"


========================================

### TrimRight

This function takes a string trims characters to the right of the string, similar to Excel's Built-in TRIM() function, except that an optional different character can be used for the trim.

**Arguments**
 * **string1** {String}: is the string that will be trimmed
 * (Optional) _[**trimCharacter** {String = " "}]_: is an optional character that will be trimmed from the string. By default, this character will be the space character " "

**Returns**
 * {String}: Returns the origional string with characters trimmed from the right only

**Examples**
 *  =TrimRight("   Hello World   ") -> "   Hello World"
 *  =TrimRight("---Hello World---", "-") -> "---Hello World"


========================================

### Zfill

This function pads zeros to the left of a string until the string is at least the length of the fill length. Optional parameters can be used to pad with a different character than 0, and to pad from right to left instead of from the default left to right.

**Arguments**
 * **string1** {String}: is the string that will be filled
 * **fillLength** {Byte}: is the length that string1 will be padded to. In cases where string1 is of greater length than this argument, no padding will occur.
 * (Optional) _[**fillCharacter** {String = "0"}]_: is an optional string that will change the character that will be padded with
 * (Optional) _[**rightToLeftFlag** {Boolean}]_: is a Boolean parameter that if set to TRUE will result in padding from right to leftt instead of left to right

**Returns**
 * {String}: Returns a new padded string of the length of specified by fillLength at minimum

**Examples**
 *  =Zfill(123, 5) -> "00123"
 *  =Zfill(5678, 5) -> "05678"
 *  =Zfill(12345678, 5) -> "12345678"
 *  =Zfill(123, 5, "X") -> "XX123"
 *  =Zfill(123, 5, "X", TRUE) -> "123XX"


---

## StringMetrics Module

This module contains a set of functions for performing fuzzy string matches. It can be useful when you have 2 columns containing text that is close but not 100% the same. However, since the functions in this module only perform fuzzy matches, there is no guarantee that there will be 100% accuracy in the matches. However, for small groups of string where each string is very different than the other (such as a small group of fairly dissimilar names), these functions can be highly accurate. Finally, some of the functions in this Module will take a long time to calculate for large numbers of cells, as the number of calculations for some functions will grow exponentially, but for small sets of data (such as 100 strings to compare), these functions perform fairly quickly.

========================================

### Damerau

This function takes two strings of any length and calculates the Damerau-Levenshtein Distance between them. Damerau-Levenshtein Distance differs from Levenshtein Distance in that it includes an additional operation, called Transpositions, which occurs when two adjacent characters are swapped. Thus, Damerau-Levenshtein Distance calculates the number of Insertions, Deletions, Substitutions, and Transpositons needed to convert string1 into string2. As a result, this function is good when it is likely that spelling errors have occured between two string where the error is simply a transposition of 2 adjacent characters.

**Arguments**
 * **string1** {String}: is the first string
 * **string2** {String}: is the second string that will be compared to the first string

**Returns**
 * {Integer}: Returns an integer of the Damerau-Levenshtein Distance between two string

**Examples**
 *  =Damerau("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
 *  =Damerau("Cat", "Ca") -> 1; Since only one Insertion needs to occur (adding a "t" at the end)
 *  =Damerau("Cat", "Cta") -> 1; Since the "t" and "a" can be transposed as they are adjacent to each other. Notice how LEVENSHTEIN("Cat","Cta")=2 but DAMERAU("Cat","Cta")=1


========================================

### Hamming

This function takes two strings of the same length and calculates the Hamming Distance between them. Hamming Distance measures how close two strings are by checking how many Substitutions are needed to turn one string into the other. Lower numbers mean the strings are closer than high numbers.

**Arguments**
 * **string1** {String}: is the first string
 * **string2** {String}: is the second string that will be compared to the first string

**Returns**
 * {Integer}: Returns an integer of the Hamming Distance between two string

**Examples**
 *  =Hamming("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
 *  =Hamming("Cat", "Bag") -> 2; 2 changes are needed, changing the "B" to "C" and the "g" to "t"
 *  =Hamming("Cat", "Dog") -> 3; Every single character needs to be substituted in this case


========================================

### Levenshtein

This function takes two strings of any length and calculates the Levenshtein Distance between them. Levenshtein Distance measures how close two strings are by checking how many Insertions, Deletions, or Substitutions are needed to turn one string into the other. Lower numbers mean the strings are closer than high numbers. Unlike Hamming Distance, Levenshtein Distance works for strings of any length and includes 2 more operations. However, calculation time will be slower than Hamming Distance for same length strings, so if you know the two strings are the same length, its preferred to use Hamming Distance.

**Arguments**
 * **string1** {String}: is the first string
 * **string2** {String}: is the second string that will be compared to the first string

**Returns**
 * {Long}: Returns an integer of the Levenshtein Distance between two string

**Examples**
 *  =Levenshtein("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
 *  =Levenshtein("Cat", "Ca") -> 1; Since only one Insertion needs to occur (adding a "t" at the end)
 *  =Levenshtein("Cat", "Cta") -> 2; Since the "t" in "Cta" needs to be substituted into an "a", and the final character "a" needs to be substituted into a "t"


---

## Utilities Module

This module contains a set of basic miscellaneous utility functions

========================================

### BigDec2Hex

This function is an implementation of Dec2Hex that allows big integers up to 14-byte to be used

**Arguments**
 * **number** {Variant}: is the integer that will be converted to a hex string

**Returns**
 * {String}: Returns the number rounded down to the nearest integer

**Examples**
 *  =Dec2Hex(255, 8) -> "000000FF"
 *  =Dec2Hex(3000000000, 16) -> Error; As Dec2Hex does not support integers this large
 *  =BigDec2Hex(3000000000, 16) -> "00000000B2D05E00"


========================================

### BigHex

This function is an implementation of the Hex() function that allows for 14-byte integers to be used

**Arguments**
 * **number** {Variant}: is the number that will be converted to hex

**Returns**
 * {String}: Returns a string of the number converted to hex

**Examples**
 *  =BigHex(255) -> "FF"
 *  =Hex(3000000000) -> Error; As hex does not support big integers
 *  =BigHex(3000000000) -> "B2D05E00"


========================================

### Dec2Hex

This function takes an integer and converts it to a hex string, with the option to specify the number of leading zeros for the hex string returned

**Arguments**
 * **number** {Long}: is the integer that will be converted to a hex string

**Returns**
 * {String}: Returns the number rounded down to the nearest integer

**Examples**
 *  =Dec2Hex(5) -> "5"
 *  =Dec2Hex(5, 2) -> "05"
 *  =Dec2Hex(255, 2) -> "FF"
 *  =Dec2Hex(255, 8) -> "000000FF"


========================================

### Hex2Dec

This function takes a hex number as a string and converts it to a decimal long

**Arguments**
 * **hexNumber** {String}: is the hex number that will be converted to a long

**Returns**
 * {Long}: Returns a decimal base number converted from the hex number

**Examples**
 *  =Hex2Dec("FF") -> 255
 *  =Hex2Dec("FFFF") -> 65535


========================================

### HideText

This function takes the value in a cell and visibly hides it if the HideText flag set to TRUE. If TRUE, the value will appear as "********", with the option to set the HideText characters to a different set of text.

**Arguments**
 * **string1** {String}: is the string that will be HideText
 * **hiddenFlag** {Boolean}: if set to TRUE will hide string1
 * (Optional) _[**hideString** {String}]_: is an optional string that if set will be used instead of "********"

**Returns**
 * {String}: Returns a string to hide string1 if hideFlag is TRUE

**Examples**
 *  =HideText("Hello World", FALSE) -> "Hello World"
 *  =HideText("Hello World", TRUE) -> "********"
 *  =HideText("Hello World", TRUE, "[Hidden Text]") -> "[Hidden Text]"
 *  =HideText("Hello World", UserName()="Anthony") -> "********"


========================================

### HtmlEscape

This function takes a string and escapes the HTML characters in it. For example, the character ">" will be escaped into "%gt;"

**Arguments**
 * **string1** {String}: is the string that will have its characters HTML escaped

**Returns**
 * {String}: Returns an HTML escaped string

**Examples**
 *  =HtmlEscape("<p>Hello World</p>") -> "&lt;p&gt;Hello World&lt;/p&gt;"


========================================

### HtmlUnescape

This function takes a string and unescapes the HTML characters in it. For example, the character "%gt;" will be escaped into ">"

**Arguments**
 * **string1** {String}: is the string that will have its characters HTML unescaped

**Returns**
 * {String}: Returns an HTML unescaped string

**Examples**
 *  =HtmlUnescape("&lt;p&gt;Hello World&lt;/p&gt;") -> "<p>Hello World</p>"


========================================

### JavaScript

This function executes JavaScript code using Microsoft's JScript scripting language. It takes 3 arguments, the JavaScript code that will be executed, the name of the JavaScript function that will be executed, and up to 16 optional arguments to be used in the JavaScript function that is called. One thing to note is that ES5 syntax should be used in the JavaScript code, as ES6 features are unlikely to be supported.

**Arguments**
 * **jsFuncCode** {String}: is a string of the JavaScript source code that will be executed
 * **jsFuncName** {String}: is the name of the JavaScript function that will be called
 * (Optional) _[**argument1** {Variant}]_: - argument16 are optional arguments used in the JScript function call

**Returns**
 * {Variant}: Returns the result of the JavaScript function that is called

**Examples**
 *  =JavaScript("function helloFunc(){return 'Hello World!'}", "helloFunc") -> "Hello World!"
 *  =JavaScript("function addTwo(a, b){return a + b}","addTwo",12,24) -> 36


========================================

### Jsonify

This function takes an array of strings and numbers and returns the array as a JSON string. This function takes into account formatting for numbers, and supports specifying the indentation level.

**Arguments**
 * **indentLevel** {Byte}: is an optional number that specifying the indentation level. Leaving this argument out will result in no indentation
 * **stringArray** {Variant}: is an array of strings and number in the following format: {"Hello", "World"}

**Returns**
 * {Variant}: Returns a JSON valid string of all elements in the array

**Examples**
 *  =Jsonify(0, "Hello", "World", "1", "2", 3, 4.5) -> "["Hello","World",1,2,3,4.5]"
 *  =Jsonify(0, {"Hello", "World", "1", "2", 3, 4.5}, 10) -> "["Hello","World",1,2,3,4.5]"


========================================

### Len2

This function is an extension on the Len() function by returning the length of strings, arrays, numbers, and many other objects in Excel, Word, PowerPoint, and Access, including Objects such as Dictionaries. Internally, any Object that implements a .Count property will have a length returned by this function. Also, any number used within this function will be converted to a string and then its length returned.

**Arguments**
 * **val** {Variant}: is the value you want the length from

**Returns**
 * {Integer}: Returns an integer of the length of the value specified

**Examples**
 *  =Len2("Hello") -> 5; As the string is 5 characters long
 *  =Len2(arr) -> 3; Where arr is an array with {1, 2, 3} in it, and the array has 3 values in it
 *  =Len2("100") -> 3; As the string is 3 characters long
 *  =Len2(100) -> 3; As the integer is 3 characters long when converted to a string
 *  =Len2(Range("A1:A3")) -> 3; As the Excel Range has 3
 *  =Len2(col) -> 5; Where col is a Collection with 5 items in it
 *  =Len2(dict) -> 2; Where dict is a Dictionary with 2 key/value pairs in it
 *  =Len2(Application.Documents) -> 3; Where we currently have 3 documents open
 *  =Len2(Application.ActivePresentation.Slides) -> 10; Where the active PowerPoint Presentation has 10 slides


========================================

### SpeakText

This function takes the range of the cell that this function resides, and then an array of text, and when this function is recalculated manually by the user (for example when pressing the F2 key while on the cell) this function will use Microsoft's text-to-speech to speak out the text through the speakers or microphone.

**Arguments**
 * **textArray** {Variant}: is an array of ranges, strings, or number that will be displayed

**Returns**
 * {Variant}: Returns all the strings in the text array combined as well as displays all the text in the text array

**Examples**
 *  =SpeakText("Hello", "World") -> "Hello World" and the text will be spoken through the speaker


========================================

### UuidFour

This function generates a unique ID based on the UUID V4 specification. This function is useful for generating unique IDs of a fixed character length.

**Arguments**
 * None

**Returns**
 * {String}: Returns a string unique ID based on UUID V4. The format of the string will always be in the form of "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx" where each x is a hex digit, and y is either 8, 9, A, or B.

**Examples**
 *  =UuidFour() -> "3B4BDC26-E76E-4D6C-9E05-7AE7D2D68304"
 *  =UuidFour() -> "D5761256-8385-4FDA-AD56-6AEF0AD6B9A5"
 *  =UuidFour() -> "CDCAE2F5-B52F-4C90-A38A-42BD58BCED4B"


---

## Validators Module

This module contains a set of functions for validating some commonly used string, such as validators for email addresses and phone numbers.

========================================

### CreditCardName

This function checks if a string is a valid credit card from one of the major card issuing companies, and then returns the name of the credit card name. This function assumes no spaces or hyphens (if you have card numbers with spaces or hyphens you can remove these using =SUBSTITUTE("-", "") function.

**Arguments**
 * **string1** {String}: is the credit card string

**Returns**
 * {String}: Returns the name of the credit card. Currently supports these cards: Visa, MasterCard, Discover, Amex, Diners, JCB

**Examples**
 *  =CreditCardName("5111567856785678") -> "MasterCard"; This is a valid Mastercard number
 *  =CreditCardName("not\_a\_card\_number") -> #VALUE!


========================================

### FormatCreditCard

This function checks if a string is a valid credit card, and if it is formats it in a more readable way. The format used is XXXX-XXXX-XXXX-XXXX.

**Arguments**
 * **string1** {String}: is credit card number

**Returns**
 * {String}: Returns a string formatted as a more readable credit card number

**Examples**
 *  =FormatCreditCard("5111567856785678") -> "5111-5678-5678-5678"


========================================

### IsCreditCard

This function checks if a string is a valid credit card from one of the major card issuing companies.

**Arguments**
 * **string1** {String}: is the string we are checking if its a valid credit card number

**Returns**
 * {Boolean}: Returns TRUE if the string is a valid credit card number, and FALSE if its invalid. Currently supports these cards: Visa, MasterCard, Discover, Amex, Diners, JCB

**Examples**
 *  =IsCreditCard("5111567856785678") -> TRUE; This is a valid Mastercard number
 *  =IsCreditCard("511156785678567") -> FALSE; Not enough digits
 *  =IsCreditCard("9999999999999999") -> FALSE; Enough digits, but not a valid card number
 *  =IsCreditCard("Hello World") -> FALSE


========================================

### IsEmail

This function checks if a string is a valid email address.

**Arguments**
 * **string1** {String}: is the string we are checking if its a valid email

**Returns**
 * {Boolean}: Returns TRUE if the string is a valid email, and FALSE if its invalid

**Examples**
 *  =IsEmail("JohnDoe@testmail.com") -> TRUE
 *  =IsEmail("JohnDoe@test/mail.com") -> FALSE
 *  =IsEmail("not\_an\_email\_address") -> FALSE


========================================

### IsIPFour

This function checks if a string is a valid IPv4 address.

**Arguments**
 * **string1** {String}: is the string we are checking if its a valid IPv4 address

**Returns**
 * {Boolean}: Returns TRUE if the string is a valid IPv4, and FALSE if its invalid

**Examples**
 *  =IsIPFour("0.0.0.0") -> TRUE
 *  =IsIPFour("100.100.100.100") -> TRUE
 *  =IsIPFour("255.255.255.255") -> TRUE
 *  =IsIPFour("255.255.255.256") -> FALSE; as the final 256 makes the address outside of the bounds of IPv4
 *  =IsIPFour("0.0.0") -> FALSE; as the fourth octet is missing


========================================

### IsMacAddress

This function checks if a string is a valid 48-bit Mac Address.

**Arguments**
 * **string1** {String}: is the string we are checking if its a valid 48-bit Mac Address

**Returns**
 * {Boolean}: Returns TRUE if the string is a valid 48-bit Mac Address, and FALSE if its invalid

**Examples**
 *  =IsMacAddress("00:25:96:12:34:56") -> TRUE
 *  =IsMacAddress("FF:FF:FF:FF:FF:FF") -> TRUE
 *  =IsMacAddress("00-25-96-12-34-56") -> TRUE
 *  =IsMacAddress("123.789.abc.DEF") -> TRUE
 *  =IsMacAddress("Not A Mac Address") -> FALSE
 *  =IsMacAddress("FF:FF:FF:FF:FF:FH") -> FALSE; the H at the end is not a valid Hex number


========================================

### IsPhone

This function checks if a string is a phone number is valid.

**Arguments**
 * **string1** {String}: is the string we are checking if its a valid phone number

**Returns**
 * {Boolean}: Returns TRUE if the string is a valid phone number, and FALSE if its invalid

**Examples**
 *  =IsPhone("123 456 7890") -> TRUE
 *  =IsPhone("1234567890") -> TRUE
 *  =IsPhone("1-234-567-890") -> FALSE; Not enough digits
 *  =IsPhone("1-234-567-8905") -> TRUE
 *  =IsPhone("+1-234-567-890") -> FALSE; Not enough digits
 *  =IsPhone("+1-234-567-8905") -> TRUE
 *  =IsPhone("+1-(234)-567-8905") -> TRUE
 *  =IsPhone("+1 (234) 567 8905") -> TRUE
 *  =IsPhone("1(234)5678905") -> TRUE
 *  =IsPhone("123-456-789") -> FALSE; Not enough digits
 *  =IsPhone("Hello World") -> FALSE; Not a phone number


========================================

### IsUrl

This function checks if a string is a valid URL address.

**Arguments**
 * **string1** {String}: is the string we are checking if its a valid URL

**Returns**
 * {Boolean}: Returns TRUE if the string is a valid URL, and FALSE if its invalid

**Examples**
 *  =IsUrl("https://www.wikipedia.org/") -> TRUE
 *  =IsUrl("http://www.wikipedia.org/") -> TRUE
 *  =IsUrl("hello\_world") -> FALSE


---

	
