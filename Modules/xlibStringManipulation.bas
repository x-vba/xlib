Attribute VB_Name = "xlibStringManipulation"
'@Module: This module contains a set of basic functions for manipulating strings.

Option Explicit


Public Function Capitalize( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and returns the same string with the first character capitalized and all other characters lowercased
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that the capitalization will be performed on
    '@Returns: Returns a new string with the first character capitalized and all others lowercased
    '@Example: =Capitalize("hello World") -> "Hello world"

    Capitalize = UCase(Left(string1, 1)) & LCase(Mid(string1, 2))
    
End Function


Public Function LeftFind( _
    ByVal string1 As String, _
    ByVal searchString As String) _
As String

    '@Description: This function takes a string and a search string, and returns a string with all characters to the left of the first search string found within string1. Similar to Excel's built-in =SEARCH() function, this function is case-sensitive. For a case-insensitive version of this function, see =LeftSearch().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be searched
    '@Param: searchString is the string that will be used to search within string1
    '@Returns: Returns a new string with all characters to the left of the first search string within string1
    '@Example: =LeftFind("Hello World", "r") -> "Hello Wo"
    '@Example: =LeftFind("Hello World", "R") -> "#VALUE!"; Since string1 does not contain "R" in it.

    LeftFind = Left(string1, InStr(1, string1, searchString) - 1)

End Function


Public Function RightFind( _
    ByVal string1 As String, _
    ByVal searchString As String) _
As String

    '@Description: This function takes a string and a search string, and returns a string with all characters to the right of the last search string found within string 1. Similar to Excel's built-in =SEARCH() function, this function is case-sensitive. For a case-insensitive version of this function, see =RightSearch().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be searched
    '@Param: searchString is the string that will be used to search within string1
    '@Returns: Returns a new string with all characters to the right of the last search string within string1
    '@Example: =RightFind("Hello World", "o") -> "rld"
    '@Example: =RightFind("Hello World", "O") -> "#VALUE!"; Since string1 does not contain "O" in it.

    RightFind = Right(string1, Len(string1) - InStrRev(string1, searchString))

End Function


Public Function LeftSearch( _
    ByVal string1 As String, _
    ByVal searchString As String) _
As String

    '@Description: This function takes a string and a search string, and returns a string with all characters to the left of the first search string found within string1. Similar to Excel's built-in =FIND() function, this function is NOT case-sensitive (it's case-insensitive). For a case-sensitive version of this function, see =LeftFind().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be searched
    '@Param: searchString is the string that will be used to search within string1
    '@Returns: Returns a new string with all characters to the left of the first search string within string1
    '@Example: =LeftSearch("Hello World", "r") -> "Hello Wo"
    '@Example: =LeftSearch("Hello World", "R") -> "Hello Wo"

    LeftSearch = Left(string1, InStr(1, string1, searchString, vbTextCompare) - 1)

End Function


Public Function RightSearch( _
    ByVal string1 As String, _
    ByVal searchString As String) _
As String

    '@Description: This function takes a string and a search string, and returns a string with all characters to the right of the last search string found within string 1. Similar to Excel's built-in =FIND() function, this function is NOT case-sensitive (it's case-insensitive). For a case-sensitive version of this function, see =RightFind().
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be searched
    '@Param: searchString is the string that will be used to search within string1
    '@Returns: Returns a new string with all characters to the right of the last search string within string1
    '@Example: =RightSearch("Hello World", "o") -> "rld"
    '@Example: =RightSearch("Hello World", "O") -> "rld"

    RightSearch = Right(string1, Len(string1) - InStrRev(string1, searchString, Compare:=vbTextCompare))

End Function


Public Function Substr( _
    ByVal string1 As String, _
    ByVal startCharacterNumber As Integer, _
    ByVal endCharacterNumber As Integer) _
As String

    '@Description: This function takes a string and a starting character number and ending character number, and returns the substring between these two numbers. The total number of characters returned will be endCharacterNumber - startCharacterNumber.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that we will get a substring from
    '@Param: startCharacterNumber is the character number of the start of the substring, with 1 being the first character in the string
    '@Param: endCharacterNumber is the character number of the end of the substring
    '@Returns: Returns a substring between the two numbers.
    '@Example: =Substr("Hello World", 2, 6) -> "ello"

    Substr = Mid(string1, startCharacterNumber, endCharacterNumber - startCharacterNumber)

End Function


Public Function SubstrFind( _
    ByVal string1 As String, _
    ByVal RightFindString As String, _
    ByVal rightSearchString As String, _
    Optional ByVal noninclusiveFlag As Boolean) _
As String

    '@Description: This function takes a string and a left string and right string, and returns a substring between those two strings. The left string will find the first matching string starting from the left, and the right string will find the first matching string starting from the right. Finally, and optional final parameter can be set to TRUE to make the substring noninclusive of the two searched strings. SubstrFind is case-sensitive. For case-insensitive version, see SubstrSearch
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that we will get a substring from
    '@Param: RightFindString is the string that will be searched from the left
    '@Param: rightSearchString is the string that will be searched from the right
    '@Param: noninclusiveFlag is an optional parameter that if set to TRUE will result in the substring not including the left and right searched characters
    '@Returns: Returns a substring between the two strings.
    '@Example: =SubstrFind("Hello World", "e", "o") -> "ello Wo"
    '@Example: =SubstrFind("Hello World", "e", "o", TRUE) -> "llo W"
    '@Example: =SubstrFind("One Two Three", "ne ", " Thr") -> "ne Two Thr"
    '@Example: =SubstrFind("One Two Three", "NE ", " THR") -> "#VALUE!"; Since SubstrFind() is case-sensitive
    '@Example: =SubstrFind("One Two Three", "ne ", " Thr", TRUE) -> "Two"
    '@Example: =SubstrFind("Country Code: +51; Area Code: 315; Phone Number: 762-5929;", "Area Code: ", "; Phone", TRUE) -> 315
    '@Example: =SubstrFind("Country Code: +313; Area Code: 423; Phone Number: 284-2468;", "Area Code: ", "; Phone", TRUE) -> 423
    '@Example: =SubstrFind("Country Code: +171; Area Code: 629; Phone Number: 731-5456;", "Area Code: ", "; Phone", TRUE) -> 629

    Dim leftCharacterNumber As Integer
    Dim rightCharacterNumber As Integer
    
    leftCharacterNumber = InStr(1, string1, RightFindString)
    rightCharacterNumber = InStrRev(string1, rightSearchString)
    
    If noninclusiveFlag = True Then
        leftCharacterNumber = leftCharacterNumber + Len(RightFindString)
        rightCharacterNumber = rightCharacterNumber - Len(rightSearchString)
    End If
    
    SubstrFind = Mid(string1, leftCharacterNumber, rightCharacterNumber - leftCharacterNumber + Len(rightSearchString))

End Function


Public Function SubstrSearch( _
    ByVal string1 As String, _
    ByVal RightFindString As String, _
    ByVal rightSearchString As String, _
    Optional ByVal noninclusiveFlag As Boolean) _
As String

    '@Description: This function takes a string and a left string and right string, and returns a substring between those two strings. The left string will find the first matching string starting from the left, and the right string will find the first matching string starting from the right. Finally, and optional final parameter can be set to TRUE to make the substring noninclusive of the two searched strings. SubstrSearch is case-insensitive. For case-sensitive version, see SubstrFind
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Param: string1 is the string that we will get a substring from
    '@Param: RightFindString is the string that will be searched from the left
    '@Param: rightSearchString is the string that will be searched from the right
    '@Param: noninclusiveFlag is an optional parameter that if set to TRUE will result in the substring not including the left and right searched characters
    '@Returns: Returns a substring between the two strings.
    '@Example: =SubstrSearch("Hello World", "e", "o") -> "ello Wo"
    '@Example: =SubstrSearch("Hello World", "e", "o", TRUE) -> "llo W"
    '@Example: =SubstrSearch("One Two Three", "ne ", " Thr") -> "ne Two Thr"
    '@Example: =SubstrSearch("One Two Three", "NE ", " THR") -> "ne Two Thr"; No error, since SubstrSearch is case-insensitive
    '@Example: =SubstrSearch("One Two Three", "ne ", " Thr", TRUE) -> "Two"
    '@Example: =SubstrSearch("Country Code: +51; Area Code: 315; Phone Number: 762-5929;", "Area Code: ", "; Phone", TRUE) -> 315
    '@Example: =SubstrSearch("Country Code: +313; Area Code: 423; Phone Number: 284-2468;", "Area Code: ", "; Phone", TRUE) -> 423
    '@Example: =SubstrSearch("Country Code: +171; Area Code: 629; Phone Number: 731-5456;", "Area Code: ", "; Phone", TRUE) -> 629

    Dim leftCharacterNumber As Integer
    Dim rightCharacterNumber As Integer
    
    leftCharacterNumber = InStr(1, string1, RightFindString, vbTextCompare)
    rightCharacterNumber = InStrRev(string1, rightSearchString, Compare:=vbTextCompare)
    
    If noninclusiveFlag = True Then
        leftCharacterNumber = leftCharacterNumber + Len(RightFindString)
        rightCharacterNumber = rightCharacterNumber - Len(rightSearchString)
    End If
    
    SubstrSearch = Mid(string1, leftCharacterNumber, rightCharacterNumber - leftCharacterNumber + Len(rightSearchString))

End Function

    
Public Function Repeat( _
    ByVal string1 As String, _
    ByVal numberOfRepeats As Integer) _
As String

    '@Description: This function repeats string1 based on the number of repeats specified in the second argument
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be repeated
    '@Param: numberOfRepeats is the number of times string1 will be repeated
    '@Returns: Returns a string repeated multiple times based on the numberOfRepeats
    '@Example: =Repeat("Hello", 2) -> HelloHello"
    '@Example: =Repeat("=", 10) -> "=========="

    Dim i As Integer
    Dim combinedString As String

    For i = 1 To numberOfRepeats
        combinedString = combinedString & string1
    Next

    Repeat = combinedString

End Function


Public Function Formatter( _
    ByVal formatString As String, _
    ParamArray textArray() As Variant) _
As String

    '@Description: This function takes a Formatter string and then an array of ranges or strings, and replaces the format placeholders with the values in the range or strings. The format syntax is "{1} - {2}" where the "{1}" and "{2}" will be replaced with the values given in the text array.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: formatString is the string that will be used as the format and which will be replaced with the individual strings
    '@Param: textArray are the ranges or strings that will be placed within the slots of the format string
    '@Returns: Returns a new string with the individual strings in the placeholder slots of the format string
    '@Example: =Formatter("Hello {1}", "World") -> "Hello World"
    '@Example: =Formatter("{1} {2}", "Hello", "World") -> "Hello World"
    '@Example: =Formatter("{1}.{2}@{3}", "FirstName", "LastName", "email.com") -> "FirstName.LastName@email.com"
    '@Example: =Formatter("{1}.{2}@{3}", A1:A3) -> "FirstName.LastName@email.com"; where A1="FirstName", A2="LastName", and A3="email.com"
    '@Example: =Formatter("{1}.{2}@{3}", A1, A2, A3) -> "FirstName.LastName@email.com"; where A1="FirstName", A2="LastName", and A3="email.com"

    Dim i As Byte
    Dim individualTextItem As Variant
    Dim individualValue As Variant
    
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


Public Function Zfill( _
    ByVal string1 As String, _
    ByVal fillLength As Byte, _
    Optional ByVal fillCharacter As String = "0", _
    Optional ByVal rightToLeftFlag As Boolean) _
As String

    '@Description: This function pads zeros to the left of a string until the string is at least the length of the fill length. Optional parameters can be used to pad with a different character than 0, and to pad from right to left instead of from the default left to right.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be filled
    '@Param: fillLength is the length that string1 will be padded to. In cases where string1 is of greater length than this argument, no padding will occur.
    '@Param: fillCharacter is an optional string that will change the character that will be padded with
    '@Param: rightToLeftFlag is a Boolean parameter that if set to TRUE will result in padding from right to leftt instead of left to right
    '@Returns: Returns a new padded string of the length of specified by fillLength at minimum
    '@Example: =Zfill(123, 5) -> "00123"
    '@Example: =Zfill(5678, 5) -> "05678"
    '@Example: =Zfill(12345678, 5) -> "12345678"
    '@Example: =Zfill(123, 5, "X") -> "XX123"
    '@Example: =Zfill(123, 5, "X", TRUE) -> "123XX"
    
    While Len(string1) < fillLength
        If rightToLeftFlag = False Then
            string1 = fillCharacter + string1
        Else
            string1 = string1 + fillCharacter
        End If
    Wend
    
    Zfill = string1

End Function


Public Function SplitText( _
    ByVal string1 As String, _
    ByVal substringNumber As Integer, _
    Optional ByVal delimiterString As String = " ") _
As String
    
    '@Description: This function takes a string and a number, splits the string by the space characters, and returns the substring in the position of the number specified in the second argument.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be split and a substring returned
    '@Param: substringNumber is the number of the substring that will be chosen
    '@Param: delimiterString is an optional parameter that can be used to specify a different delimiter
    '@Returns: Returns a substring of the split text in the location specified
    '@Example: =SplitText("Hello World", 1) -> "Hello"
    '@Example: =SplitText("Hello World", 2) -> "World"
    '@Example: =SplitText("One Two Three", 2) -> "Two"
    '@Example: =SplitText("One-Two-Three", 2, "-") -> "Two"
    
    SplitText = Split(string1, delimiterString)(substringNumber - 1)

End Function


Public Function CountWords( _
    ByVal string1 As String, _
    Optional ByVal delimiterString As String = " ") _
As Integer

    '@Description: This function takes a string and returns the number of words in the string
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Note: If the number given is higher than the number of words, its possible that the string contains excess whitespace. Try using the =TRIM() function first to remove the excess whitespace
    '@Param: string1 is the string whose number of words will be counted
    '@Param: delimiterString is an optional parameter that can be used to specify a different delimiter
    '@Returns: Returns the number of words in the string
    '@Example: =CountWords("Hello World") -> 2
    '@Example: =CountWords("One Two Three") -> 3
    '@Example: =CountWords("One-Two-Three", "-") -> 3

    Dim stringArray() As String

    stringArray = Split(string1, delimiterString)
    
    CountWords = UBound(stringArray) - LBound(stringArray) + 1

End Function


Public Function CamelCase( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and returns the same string in camel case, removing all the spaces.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be camel cased
    '@Returns: Returns a new string in camel case, where the first character of the first word is lowercase, and uppercased for all other words
    '@Example: =CamelCase("Hello World") -> "helloWorld"
    '@Example: =CamelCase("One Two Three") -> "oneTwoThree"

    Dim i As Integer
    Dim stringArray() As String
    
    stringArray = Split(string1, " ")
    stringArray(0) = LCase(stringArray(0))
    
    For i = 1 To (UBound(stringArray) - LBound(stringArray))
        stringArray(i) = UCase(Left(stringArray(i), 1)) & LCase(Mid(stringArray(i), 2))
    Next
    
    CamelCase = Join(stringArray, "")

End Function


Public Function KebabCase( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and returns the same string in kebab case.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be kebab cased
    '@Returns: Returns a new string in kebab case, where all letters are lowercase and seperated by a "-" character
    '@Example: =KebabCase("Hello World") -> "hello-world"
    '@Example: =KebabCase("One Two Three") -> "one-two-three"

    KebabCase = LCase(Join(Split(string1, " "), "-"))

End Function


Public Function RemoveCharacters( _
    ByVal string1 As String, _
    ParamArray removedCharacters() As Variant) _
As String

    '@Description: This function takes a string and either another string or multiple strings and removes all characters from the first string that are in the second string.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Consider adding a Boolean flag that will make non-case sensitive replacements
    '@Note: This function is case sensitive. If you want to remove the "H" from "Hello World" you would need to use "H" as a removed character, not "h".
    '@Param: string1 is the string that will have characters removed
    '@Param: removedCharacters is an array of strings that will be removed from string1
    '@Returns: Returns the origional string with characters removed
    '@Example: =RemoveCharacters("Hello World", "l") -> "Heo Word"
    '@Example: =RemoveCharacters("Hello World", "lo") -> "He Wrd"
    '@Example: =RemoveCharacters("Hello World", "l", "o") -> "He Wrd"
    '@Example: =RemoveCharacters("Hello World", "lod") -> "He Wr"
    '@Example: =RemoveCharacters("Two Three Four", "f", "t") -> "Two Three Four"; Nothing is replaced since this function is case sensitive
    '@Example: =RemoveCharacters("Two Three Four", "F", "T") -> "wo hree our"

    Dim i As Integer
    Dim individualCharacter As Variant
    
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


Private Function NumberOfUppercaseLetters( _
    ByVal string1 As String) _
As Integer

    '@Description: This function returns the number of uppercase letter found within a string based on the ASCII character code range for uppercase letters
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose uppercase letters will be counted
    '@Returns: Returns the number of uppercase letters

    Dim i As Integer
    Dim numberOfUppercase As Integer
    
    For i = 1 To Len(string1)
        If Asc(Mid(string1, i, 1)) >= 65 Then
            If Asc(Mid(string1, i, 1)) <= 90 Then
                numberOfUppercase = numberOfUppercase + 1
            End If
        End If
    Next
    
    NumberOfUppercaseLetters = numberOfUppercase

End Function


Public Function CompanyCase( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and uses an algorithm to return the string in Company Case. The standard =PROPER() function in Excel will not capitalize company names properly, as it only capitalizes based on space characters, so a name like "j.p. morgan" will be incorrectly formatted as "J.p. Morgan" instead of the correct "J.P. Morgan". Additionally =PROPER() may incorrectly lowercase company abbreviations, such as the last "H" in "GmbH", as =PROPER() returns "Gmbh" instead of the correct "GmbH". This function attempts to adjust for these issues when a string is a company name.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Warning: There is no perfect algorithm for correctly formatting company names, and while this function can give better performance for correct formatting when compared to =PROPER(), if the performance of this function isn't as accurate as one needs, another solution would be to try Partial Lookup functions in the String Metrics Module and compare that to a known list of well formatted company strings.
    '@Param: string1 is the string that will be formatted
    '@Returns: Returns the origional string in a Company Case format
    '@Example: =CompanyCase("hello world") -> "Hello World"
    '@Example: =CompanyCase("x.y.z company & co.") -> "X.Y.Z Company & Co."
    '@Example: =CompanyCase("x.y.z plc") -> "X.Y.Z PLC"
    '@Example: =CompanyCase("one company gmbh") -> "One Company GmbH"
    '@Example: =CompanyCase("three company s. en n.c.") -> "Three Company S. en N.C."
    '@Example: =CompanyCase("FOUR COMPANY SPOL S.R.O.") -> "Four Company spol s.r.o."
    '@Example: =CompanyCase("five company bvba") -> "Five Company BVBA"

    Dim i As Integer
    Dim k As Integer
    Dim origionalString As String
    Dim stringArray() As String
    Dim splitCharacters As String
    
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
    
    
    ' Checking the final words in the string to see if they are one of the
    ' company abbreviation strings, and if it is, replace the ending with
    ' the correct cases of the company abbreviation
    Dim companyAbbreviationArray() As String
    companyAbbreviationArray = Split("AB|AG|GmbH|LLC|LLP|NV|PLC|SA|A. en P.|ACE|AD|AE|AL|AmbA|ANS|ApS|AS|ASA|AVV|BVBA|CA|CVA|d.d.|d.n.o.|d.o.o.|DA|e.V.|EE|EEG|EIRL|ELP|EOOD|EPE|EURL|GbR|GCV|GesmbH|GIE|HB|hf|IBC|j.t.d.|k.d.|k.d.d.|k.s.|KA/S|KB|KD|KDA|KG|KGaA|KK|Kol. SrK|Kom. SrK|LDC|Ltée.|NT|OE|OHG|Oy|OYJ|OÜ|PC Ltd|PMA|PMDN|PrC|PT|RAS|S. de R.L.|S. en N.C.|SA de CV|SAFI|SAS|SC|SCA|SCP|SCS|SENC|SGPS|SK|SNC|SOPARFI|sp|Sp. z.o.o.|SpA|spol s.r.o.|SPRL|TD|TLS|v.o.s.|VEB|VOF|BYSHR", "|")

    Dim stringArrayLength As Integer

    stringArray = Split(string1, " ")
    stringArrayLength = UBound(stringArray) - LBound(stringArray)

    Dim companyAbbreviationString As Variant
    
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


Public Function ReverseText( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and reverses all the characters in it so that the returned string is backwards
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be reversed
    '@Returns: Returns the origional string in reverse
    '@Example: =ReverseText("Hello World") -> "dlroW olleH"

    Dim i As Integer
    Dim reversedString As String
    
    For i = 1 To Len(string1)
        reversedString = reversedString & Mid(string1, Len(string1) - i + 1, 1)
    Next
    
    ReverseText = reversedString

End Function


Public Function ReverseWords( _
    ByVal string1 As String, _
    Optional ByVal delimiterCharacter As String = " ") _
As String

    '@Description: This function takes a string and reverses all the words in it so that the returned string's words are backwards. By default, this function uses the space character as a delimiter, but you can optionally specify a different delimiter.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose words will be reversed
    '@Param: delimiterCharacter is the delimiter that will be used, with the default being " "
    '@Returns: Returns the origional string with it's words reversed
    '@Example: =ReverseWords("Hello World") -> "World Hello"
    '@Example: =ReverseWords("One Two Three") -> "Three Two One"
    '@Example: =ReverseWords("One-Two-Three", "-") -> "Three-Two-One"

    Dim i As Integer
    Dim stringArray() As String
    Dim stringArrayLength As Integer
    Dim reversedStringArray() As String
    
    stringArray = Split(string1, delimiterCharacter)
    stringArrayLength = (UBound(stringArray) - LBound(stringArray))
    
    ReDim reversedStringArray(stringArrayLength)
    
    For i = 0 To stringArrayLength
        reversedStringArray(i) = stringArray(stringArrayLength - i)
    Next
    
    ReverseWords = Join(reversedStringArray, delimiterCharacter)

End Function


Public Function IndentText( _
    ByVal string1 As String, _
    Optional ByVal indentAmount As Byte = 4) _
As String

    '@Description: This function takes a string and indents all of its lines by a specified number of space characters (or 4 space characters if left blank)
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be indented
    '@Param: indentAmount is the amount of " " characters that will be indented to the left of string1
    '@Returns: Returns the origional string indented by a specified number of space characters
    '@Example: =IndentText("Hello") -> "    Hello"
    '@Example: =IndentText("Hello", 4) -> "    Hello"
    '@Example: =IndentText("Hello", 3) -> "   Hello"
    '@Example: =IndentText("Hello", 2) -> "  Hello"
    '@Example: =IndentText("Hello", 1) -> " Hello"

    Dim i As Integer
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


Public Function DedentText( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and dedents all of its lines so that there are no space characters to the left or right of each line
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be dedented
    '@Returns: Returns the origional string dedented on each line
    '@Note: Unlike the Excel built-in TRIM() function, this function will dedent every single line, so for strings that span multiple lines in a cell, this will dedent all lines.
    '@Example: =DedentText("    Hello") -> "Hello"

    Dim i As Integer
    Dim stringArray() As String

    stringArray = Split(string1, Chr(10))
    
    For i = 0 To (UBound(stringArray) - LBound(stringArray))
        stringArray(i) = Trim(stringArray(i))
    Next

    DedentText = Join(stringArray, Chr(10))

End Function


Public Function ShortenText( _
    ByVal string1 As String, _
    Optional ByVal shortenWidth As Integer = 80, _
    Optional ByVal placeholderText As String = "[...]", _
    Optional ByVal delimiterCharacter As String = " ") _
As String

    '@Description: This function takes a string and shortens it with placeholder text so that it is no longer in length than the specified width.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be shortened
    '@Param: shortenWidth is the max width of the string. By default this is set to 80
    '@Param: placeholderText is the text that will be placed at the end of the string if it is longer than the shortenWidth. By default this placeholder string is "[...]
    '@Param: delimiterCharacter is the character that will be used as the word delimiter. By default this is the space character " "
    '@Returns: Returns a shortened string with placeholder text if it is longer than the shorten width
    '@Example: =ShortenText("Hello World One Two Three", 20) -> "Hello World [...]"; Only the first two words and the placeholder will result in a string that is less than or equal to 20 in length
    '@Example: =ShortenText("Hello World One Two Three", 15) -> "Hello [...]"; Only the first word and the placeholder will result in a string that is less than or equal to 15 in length
    '@Example: =ShortenText("Hello World One Two Three") -> "Hello World One Two Three"; Since this string is shorter than the default 80 shorten width value, no placeholder will be used and the string wont be shortened
    '@Example: =ShortenText("Hello World One Two Three", 15, "-->") -> "Hello World -->"; A new placeholder is used
    '@Example: =ShortenText("Hello_World_One_Two_Three", 15, "-->", "_") -> "Hello_World_-->"; A new placeholder andd delimiter is used

    Dim shortenedString As String
    Dim individualString As Variant
    Dim stringArray() As String
    
    ' In cases where the origional string is less than the threshold needed to
    ' shorten the string, simply return the origional string
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


Public Function InSplit( _
    ByVal string1 As String, _
    ByVal splitString As String, _
    Optional ByVal delimiterCharacter As String = " ") _
As Boolean

    '@Description: This function takes a search string and checks if it exists within a larger string that is split by a delimiter character.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be checked if it exists within the splitString after the split
    '@Param: splitString is the string that will be split and of which string1 will be searched in
    '@Param: delimiterCharacter is the character that will be used as the delimiter for the split. By default this is the space character " "
    '@Returns: Returns TRUE if string1 is found in splitString after the split occurs
    '@Example: =InSplit("Hello", "Hello World One Two Three") -> TRUE; Since "Hello" is found within the searchString after being split
    '@Example: =InSplit("NotInString", "Hello World One Two Three") -> FALSE; Since "NotInString" is not found within the searchString after being split
    '@Example: =InSplit("Hello", "Hello-World-One-Two-Three", "-") -> TRUE; Since "Hello" is found and since the delimiter is set to "-"

    Dim individualString As Variant
    
    For Each individualString In Split(splitString, delimiterCharacter)
        If string1 = individualString Then
            InSplit = True
            Exit Function
        End If
    Next
    
    InSplit = False

End Function


Public Function EliteCase( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and returns the string with characters replaced by similar in appearance numbers
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will have characters replaced
    '@Returns: Returns the string with characters replaced with similar in appearance numbers
    '@Example: =EliteCase("Hello World") -> "H3110 W0r1d"

    string1 = Replace(string1, "o", "0", Compare:=vbTextCompare)
    string1 = Replace(string1, "l", "1", Compare:=vbTextCompare)
    string1 = Replace(string1, "z", "2", Compare:=vbTextCompare)
    string1 = Replace(string1, "e", "3", Compare:=vbTextCompare)
    string1 = Replace(string1, "a", "4", Compare:=vbTextCompare)
    string1 = Replace(string1, "s", "5", Compare:=vbTextCompare)
    string1 = Replace(string1, "t", "7", Compare:=vbTextCompare)

    EliteCase = string1

End Function


Public Function ScrambleCase( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string scrambles the case on each character in the string
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose character's cases will be scrambled
    '@Returns: Returns the origional string with cases scrambled
    '@Example: =ScrambleCase("Hello World") -> "helLo WORlD"
    '@Example: =ScrambleCase("Hello World") -> "HElLo WorLD"
    '@Example: =ScrambleCase("Hello World") -> "hELlo WOrLd"

    Dim i As Integer

    For i = 1 To Len(string1)
        If RandBetween(0, 1) = 1 Then
            Mid(string1, i, 1) = UCase(Mid(string1, i, 1))
        Else
            Mid(string1, i, 1) = LCase(Mid(string1, i, 1))
        End If
    Next
    
    ScrambleCase = string1

End Function


Public Function LeftSplit( _
    ByVal string1 As String, _
    ByVal numberOfSplit As Integer, _
    Optional ByVal delimiterCharacter As String = " ") _
As String

    '@Description: This function takes a string, splits it based on a delimiter, and returns all characters to the left of the specified position of the split.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be split to get a substring
    '@Param: numberOfSplit is the number of the location within the split that we will get all characters to the left of
    '@Param: delimiterCharacter is the delimiter that will be used for the split. By default, the delimiter will be the space character " "
    '@Returns: Returns all characters to the left of the number of the split
    '@Example: =LeftSplit("Hello World One Two Three", 1) -> "Hello"
    '@Example: =LeftSplit("Hello World One Two Three", 2) -> "Hello World"
    '@Example: =LeftSplit("Hello World One Two Three", 3) -> "Hello World One"
    '@Example: =LeftSplit("Hello World One Two Three", 10) -> "Hello World One Two Three"
    '@Example: =LeftSplit("Hello-World-One-Two-Three", 2, "-") -> "Hello-World"

    Dim i As Integer
    Dim newString As String
    Dim stringArray() As String
    Dim stringArrayLength As Integer
    
    numberOfSplit = numberOfSplit - 1
    stringArray = Split(string1, delimiterCharacter)
    stringArrayLength = (UBound(stringArray) - LBound(stringArray) + 1)
    
    ' Checking if the number of split is greater than the length of the split
    ' array, and if so returns the origional string
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


Public Function RightSplit( _
    ByVal string1 As String, _
    ByVal numberOfSplit As Integer, _
    Optional ByVal delimiterCharacter As String = " ") _
As String

    '@Description: This function takes a string, splits it based on a delimiter, and returns all characters to the right of the specified position of the split.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be split to get a substring
    '@Param: numberOfSplit is the number of the location within the split that we will get all characters to the right of
    '@Param: delimiterCharacter is the delimiter that will be used for the split. By default, the delimiter will be the space character " "
    '@Returns: Returns all characters to the right of the number of the split
    '@Example: =RightSplit("Hello World One Two Three", 1) -> "Three"
    '@Example: =RightSplit("Hello World One Two Three", 2) -> "Two Three"
    '@Example: =RightSplit("Hello World One Two Three", 3) -> "One Two Three"
    '@Example: =RightSplit("Hello World One Two Three", 10) -> "Hello World One Two Three"
    '@Example: =RightSplit("Hello-World-One-Two-Three", 2, "-") -> "Two-Three"

    Dim i As Integer
    Dim newString As String
    Dim stringArray() As String
    Dim stringArrayLength As Integer
    
    numberOfSplit = numberOfSplit - 1
    stringArray = Split(string1, delimiterCharacter)
    stringArrayLength = (UBound(stringArray) - LBound(stringArray) + 1)
    
    ' Checking if the number of split is greater than the length of the split
    ' array, and if so returns the origional string
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


Public Function TrimChar( _
    ByVal string1 As String, _
    Optional ByVal trimCharacter As String = " ") _
As String

    '@Description: This function takes a string trims characters to the left and right of the string, similar to Excel's Built-in TRIM() function, except that an optional different character can be used for the trim.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Allow more than 1 character to be used for trimming
    '@Param: string1 is the string that will be trimmed
    '@Param: trimCharacter is an optional character that will be trimmed from the string. By default, this character will be the space character " "
    '@Returns: Returns the origional string with characters trimmed from the left and right
    '@Note: This function currently supports only single characters for trimming
    '@Example: =TrimChar("   Hello World   ") -> "Hello World"
    '@Example: =TrimChar("---Hello World---", "-") -> "Hello World"

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


Public Function TrimLeft( _
    ByVal string1 As String, _
    Optional ByVal trimCharacter As String = " ") _
As String

    '@Description: This function takes a string trims characters to the left of the string, similar to Excel's Built-in TRIM() function, except that an optional different character can be used for the trim.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Allow more than 1 character to be used for trimming
    '@Param: string1 is the string that will be trimmed
    '@Param: trimCharacter is an optional character that will be trimmed from the string. By default, this character will be the space character " "
    '@Returns: Returns the origional string with characters trimmed from the left only
    '@Note: This function currently supports only single characters for trimming
    '@Example: =TrimLeft("   Hello World   ") -> "Hello World   "
    '@Example: =TrimLeft("---Hello World---", "-") -> "Hello World---"

    While Left(string1, 1) = trimCharacter
        Mid(string1, 1) = Chr(1)
        string1 = Replace(string1, Chr(1), "")
    Wend
    
    TrimLeft = string1

End Function


Public Function TrimRight( _
    ByVal string1 As String, _
    Optional ByVal trimCharacter As String = " ") _
As String

    '@Description: This function takes a string trims characters to the right of the string, similar to Excel's Built-in TRIM() function, except that an optional different character can be used for the trim.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Allow more than 1 character to be used for trimming
    '@Param: string1 is the string that will be trimmed
    '@Param: trimCharacter is an optional character that will be trimmed from the string. By default, this character will be the space character " "
    '@Returns: Returns the origional string with characters trimmed from the right only
    '@Note: This function currently supports only single characters for trimming
    '@Example: =TrimRight("   Hello World   ") -> "   Hello World"
    '@Example: =TrimRight("---Hello World---", "-") -> "---Hello World"
    
    While Right(string1, 1) = trimCharacter
        Mid(string1, Len(string1)) = Chr(1)
        string1 = Replace(string1, Chr(1), "")
    Wend
    
    TrimRight = string1

End Function


Public Function CountUppercaseCharacters( _
    ByVal string1 As String) _
As Integer

    '@Description: This function takes a string and counts the number of uppercase characters in it
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose characters will be counted
    '@Returns: Returns the number of uppercase characters in the string
    '@Example: =CountUppercaseCharacters("Hello World") -> 2; As the "H" and the "E" are the only 2 uppercase characters in the string

    Dim i As Integer
    Dim characterAsciiCode As Byte
    Dim uppercaseCounter As Integer
    
    For i = 1 To Len(string1)
        characterAsciiCode = Asc(Mid(string1, i, 1))
        If characterAsciiCode >= 65 And characterAsciiCode <= 90 Then
            uppercaseCounter = uppercaseCounter + 1
        End If
    Next
    
    CountUppercaseCharacters = uppercaseCounter

End Function


Public Function CountLowercaseCharacters( _
    ByVal string1 As String) _
As Integer

    '@Description: This function takes a string and counts the number of lowercase characters in it
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string whose characters will be counted
    '@Returns: Returns the number of lowercase characters in the string
    '@Example: =CountLowercaseCharacters("Hello World") -> 8; As the "ello" and the "orld" are lowercase

    Dim i As Integer
    Dim characterAsciiCode As Byte
    Dim lowercaseCounter As Integer
    
    For i = 1 To Len(string1)
        characterAsciiCode = Asc(Mid(string1, i, 1))
        If characterAsciiCode >= 97 And characterAsciiCode <= 122 Then
            lowercaseCounter = lowercaseCounter + 1
        End If
    Next
    
    CountLowercaseCharacters = lowercaseCounter

End Function


Public Function TextJoin( _
    ByVal stringArray As Variant, _
    Optional ByVal delimiterCharacter As String, _
    Optional ByVal ignoreEmptyCellsFlag As Boolean) _
As String

    '@Description: This function takes a range of cells and combines all the text together, optionally allowing a character delimiter between all the combined strings, and optionally allowing blank cells to be ignored when combining the text. Finally note that this function is very similar to the TEXTJOIN function available in Excel 2019, and thus is a polyfill for that function for earlier versions of Excel.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: stringArray is the range with all the strings we want to combine
    '@Param: delimiterCharacter is an optional character that will be used as the delimiter between the combined text. By default, no delimiter character will be used.
    '@Param: ignoreEmptyCellsFlag if set to TRUE will skip combining empty cells into the combined string, and is useful when specifying a delimiter so that the delimiter does not repeat for empty cells.
    '@Returns: Returns a new combined string containing the strings in the range delimited by the delimiter character.
    '@Example: =TextJoin(A1:A3) -> "123"; Where A1:A3 contains ["1", "2", "3"]
    '@Example: =TextJoin(A1:A3, "--") -> "1--2--3"; Where A1:A3 contains ["1", "2", "3"]
    '@Example: =TextJoin(A1:A3, "--") -> "1----3"; Where A1:A3 contains ["1", "", "3"]
    '@Example: =TextJoin(A1:A3, "-") -> "1--3"; Where A1:A3 contains ["1", "", "3"]
    '@Example: =TextJoin(A1:A3, "-", TRUE) -> "1-3"; Where A1:A3 contains ["1", "", "3"]

    Dim individualString As Variant
    Dim combinedString As String
    
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
