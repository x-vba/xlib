Attribute VB_Name = "xlibUtilities"
'@Module: This module contains a set of basic miscellaneous utility functions

Option Explicit


Public Function Jsonify( _
    ByVal indentLevel As Byte, _
    ParamArray stringArray() As Variant) _
As String

    '@Description: This function takes an array of strings and numbers and returns the array as a JSON string. This function takes into account formatting for numbers, and supports specifying the indentation level.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: indentLevel is an optional number that specifying the indentation level. Leaving this argument out will result in no indentation
    '@Param: stringArray() is an array of strings and number in the following format: {"Hello", "World"}
    '@Returns: Returns a JSON valid string of all elements in the array
    '@Example: =Jsonify(0, "Hello", "World", "1", "2", 3, 4.5) -> "["Hello","World",1,2,3,4.5]"
    '@Example: =Jsonify(0, {"Hello", "World", "1", "2", 3, 4.5}, 10) -> "["Hello","World",1,2,3,4.5]"

    Dim i As Byte
    Dim jsonString As String
    Dim individualTextItem As Variant
    Dim individualValue As Variant
    Dim indentString As String
    
    ' Setting up some base JSON features and the indenting
    jsonString = "["
    
    For i = 1 To indentLevel
        indentString = indentString & " "
    Next
    
    If indentLevel > 0 Then
        jsonString = jsonString & Chr(10)
    End If
    
    
    ' Creating the contents of the JSON string
    For Each individualTextItem In stringArray
    
        ' In cases of ranges
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
            
        ' In cases of text
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

    '@Description: This function generates a unique ID based on the UUID V4 specification. This function is useful for generating unique IDs of a fixed character length.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string unique ID based on UUID V4. The format of the string will always be in the form of "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx" where each x is a hex digit, and y is either 8, 9, A, or B.
    '@Example: =UuidFour() -> "3B4BDC26-E76E-4D6C-9E05-7AE7D2D68304"
    '@Example: =UuidFour() -> "D5761256-8385-4FDA-AD56-6AEF0AD6B9A5"
    '@Example: =UuidFour() -> "CDCAE2F5-B52F-4C90-A38A-42BD58BCED4B"

    Dim firstGroup As String
    Dim secondGroup As String
    Dim thirdGroup As String
    Dim fourthGroup As String
    Dim fifthGroup As String
    Dim sixthGroup As String

    firstGroup = BigDec2Hex(BigRandBetween(0, 4294967295#), 8) & "-"
    secondGroup = Dec2Hex(RandBetween(0, 65535), 4) & "-"
    thirdGroup = Dec2Hex(RandBetween(16384, 20479), 4) & "-"
    fourthGroup = Dec2Hex(RandBetween(32768, 49151), 4) & "-"
    fifthGroup = Dec2Hex(RandBetween(0, 65535), 4)
    sixthGroup = BigDec2Hex(BigRandBetween(0, 4294967295#), 8)

    UuidFour = firstGroup & secondGroup & thirdGroup & fourthGroup & fifthGroup & sixthGroup

End Function


Public Function HideText( _
    ByVal string1 As String, _
    ByVal hiddenFlag As Boolean, _
    Optional ByVal hideString As String) _
As String

    '@Description: This function takes the value in a cell and visibly hides it if the HideText flag set to TRUE. If TRUE, the value will appear as "********", with the option to set the HideText characters to a different set of text.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will be HideText
    '@Param: hiddenFlag if set to TRUE will hide string1
    '@Param: hideString is an optional string that if set will be used instead of "********"
    '@Returns: Returns a string to hide string1 if hideFlag is TRUE
    '@Example: =HideText("Hello World", FALSE) -> "Hello World"
    '@Example: =HideText("Hello World", TRUE) -> "********"
    '@Example: =HideText("Hello World", TRUE, "[Hidden Text]") -> "[Hidden Text]"
    '@Example: =HideText("Hello World", UserName()="Anthony") -> "********"

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


Public Function JavaScript( _
    ByVal jsFuncCode As String, _
    ByVal jsFuncName As String, _
    Optional ByVal argument1 As Variant, _
    Optional ByVal argument2 As Variant, _
    Optional ByVal argument3 As Variant, _
    Optional ByVal argument4 As Variant, _
    Optional ByVal argument5 As Variant, _
    Optional ByVal argument6 As Variant, _
    Optional ByVal argument7 As Variant, _
    Optional ByVal argument8 As Variant, _
    Optional ByVal argument9 As Variant, _
    Optional ByVal argument10 As Variant, _
    Optional ByVal argument11 As Variant, _
    Optional ByVal argument12 As Variant, _
    Optional ByVal argument13 As Variant, _
    Optional ByVal argument14 As Variant, _
    Optional ByVal argument15 As Variant, _
    Optional ByVal argument16 As Variant) _
As Variant

    '@Description: This function executes JavaScript code using Microsoft's JScript scripting language. It takes 3 arguments, the JavaScript code that will be executed, the name of the JavaScript function that will be executed, and up to 16 optional arguments to be used in the JavaScript function that is called. One thing to note is that ES5 syntax should be used in the JavaScript code, as ES6 features are unlikely to be supported.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: jsFuncCode is a string of the JavaScript source code that will be executed
    '@Param: jsFuncName is the name of the JavaScript function that will be called
    '@Param: argument1 - argument16 are optional arguments used in the JScript function call
    '@Returns: Returns the result of the JavaScript function that is called
    '@Example: =JavaScript("function helloFunc(){return 'Hello World!'}", "helloFunc") -> "Hello World!"
    '@Example: =JavaScript("function addTwo(a, b){return a + b}","addTwo",12,24) -> 36

    Dim ScriptContoller As Object
    Set ScriptContoller = CreateObject("ScriptControl")
    
    ScriptContoller.Language = "JScript"
    ScriptContoller.addCode jsFuncCode

    JavaScript = ScriptContoller.Run(jsFuncName, _
        argument1, argument2, argument3, argument4, _
        argument5, argument6, argument7, argument8, _
        argument9, argument10, argument11, argument12, _
        argument13, argument14, argument15, argument16)

End Function

Public Function HtmlEscape( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and escapes the HTML characters in it. For example, the character ">" will be escaped into "%gt;"
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will have its characters HTML escaped
    '@Returns: Returns an HTML escaped string
    '@Example: =HtmlEscape("<p>Hello World</p>") -> "&lt;p&gt;Hello World&lt;/p&gt;"

    string1 = Replace(string1, "&", "&amp;")
    string1 = Replace(string1, Chr(34), "&quot;")
    string1 = Replace(string1, "'", "&apos;")
    string1 = Replace(string1, "<", "&lt;")
    string1 = Replace(string1, ">", "&gt;")
    
    HtmlEscape = string1

End Function


Public Function HtmlUnescape( _
    ByVal string1 As String) _
As String

    '@Description: This function takes a string and unescapes the HTML characters in it. For example, the character "%gt;" will be escaped into ">"
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that will have its characters HTML unescaped
    '@Returns: Returns an HTML unescaped string
    '@Example: =HtmlUnescape("&lt;p&gt;Hello World&lt;/p&gt;") -> "<p>Hello World</p>"

    string1 = Replace(string1, "&amp;", "&")
    string1 = Replace(string1, "&quot;", Chr(34))
    string1 = Replace(string1, "&apos;", "'")
    string1 = Replace(string1, "&lt;", "<")
    string1 = Replace(string1, "&gt;", ">")

    HtmlUnescape = string1

End Function


Private Sub CallTextToSpeech(combinedString)

    '@Description: This subroutine simply calls the text-to-speech API
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: combinedString is the string that will be spoken

    Application.Speech.Speak combinedString, True

End Sub


Public Function SpeakText( _
    ParamArray textArray() As Variant) _
As String

    '@Description: This function takes the range of the cell that this function resides, and then an array of text, and when this function is recalculated manually by the user (for example when pressing the F2 key while on the cell) this function will use Microsoft's text-to-speech to speak out the text through the speakers or microphone.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: textArray() is an array of ranges, strings, or number that will be displayed
    '@Note: Note that text-to-speech is only available on Microsoft Excel. This function will still return the combined string from the text array, but will only result in speech through the speakers in Microsoft Excel
    '@Returns: Returns all the strings in the text array combined as well as displays all the text in the text array
    '@Example: =SpeakText("Hello", "World") -> "Hello World" and the text will be spoken through the speaker

    Dim combinedString As String
    Dim individualTextItem As Variant
    
    For Each individualTextItem In textArray
        combinedString = combinedString & individualTextItem & " "
    Next
    
    If Application.Name = "Microsoft Excel" Then
        CallTextToSpeech combinedString
    End If

    SpeakText = Trim(combinedString)

End Function


Public Function Dec2Hex( _
    ByVal number As Long, _
    Optional ByVal zeroFillAmount As Integer) _
As String

    '@Description: This function takes an integer and converts it to a hex string, with the option to specify the number of leading zeros for the hex string returned
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: number is the integer that will be converted to a hex string
    '@Returns: Returns the number rounded down to the nearest integer
    '@Example: =Dec2Hex(5) -> "5"
    '@Example: =Dec2Hex(5, 2) -> "05"
    '@Example: =Dec2Hex(255, 2) -> "FF"
    '@Example: =Dec2Hex(255, 8) -> "000000FF"

    Dim i As Integer
    Dim hexString As String
    
    hexString = Hex(number)
    
    If zeroFillAmount > 0 Then
        While Len(hexString) < zeroFillAmount
            hexString = "0" & hexString
        Wend
    End If
    
    Dec2Hex = hexString

End Function


Public Function BigDec2Hex( _
    ByVal number As Variant, _
    Optional ByVal zeroFillAmount As Integer) _
As String

    '@Description: This function is an implementation of Dec2Hex that allows big integers up to 14-byte to be used
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: number is the integer that will be converted to a hex string
    '@Returns: Returns the number rounded down to the nearest integer
    '@Example: =Dec2Hex(255, 8) -> "000000FF"
    '@Example: =Dec2Hex(3000000000, 16) -> Error; As Dec2Hex does not support integers this large
    '@Example: =BigDec2Hex(3000000000, 16) -> "00000000B2D05E00"

    Dim i As Integer
    Dim hexString As String
    
    hexString = BigHex(number)
    
    If zeroFillAmount > 0 Then
        While Len(hexString) < zeroFillAmount
            hexString = "0" & hexString
        Wend
    End If
    
    BigDec2Hex = hexString

End Function


Public Function BigHex( _
    ByVal number As Variant) _
As String

    '@Description: This function is an implementation of the Hex() function that allows for 14-byte integers to be used
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: number is the number that will be converted to hex
    '@Returns: Returns a string of the number converted to hex
    '@Example: =BigHex(255) -> "FF"
    '@Example: =Hex(3000000000) -> Error; As hex does not support big integers
    '@Example: =BigHex(3000000000) -> "B2D05E00"

    Dim integerString As String
    Dim decimalString As String
    Dim hexString As String

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

Public Function Hex2Dec( _
    ByVal hexNumber As String) _
As Long

    '@Description: This function takes a hex number as a string and converts it to a decimal long
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: hexNumber is the hex number that will be converted to a long
    '@Returns: Returns a decimal base number converted from the hex number
    '@Example: =Hex2Dec("FF") -> 255
    '@Example: =Hex2Dec("FFFF") -> 65535

    Hex2Dec = CLng("&H" & hexNumber)

End Function


Public Function Len2( _
    ByVal val As Variant) _
As Integer

    '@Description: This function is an extension on the Len() function by returning the length of strings, arrays, numbers, and many other objects in Excel, Word, PowerPoint, and Access, including Objects such as Dictionaries. Internally, any Object that implements a .Count property will have a length returned by this function. Also, any number used within this function will be converted to a string and then its length returned.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: val is the value you want the length from
    '@Returns: Returns an integer of the length of the value specified
    '@Example: =Len2("Hello") -> 5; As the string is 5 characters long
    '@Example: =Len2(arr) -> 3; Where arr is an array with {1, 2, 3} in it, and the array has 3 values in it
    '@Example: =Len2("100") -> 3; As the string is 3 characters long
    '@Example: =Len2(100) -> 3; As the integer is 3 characters long when converted to a string
    '@Example: =Len2(Range("A1:A3")) -> 3; As the Excel Range has 3
    '@Example: =Len2(col) -> 5; Where col is a Collection with 5 items in it
    '@Example: =Len2(dict) -> 2; Where dict is a Dictionary with 2 key/value pairs in it
    '@Example: =Len2(Application.Documents) -> 3; Where we currently have 3 documents open
    '@Example: =Len2(Application.ActivePresentation.Slides) -> 10; Where the active PowerPoint Presentation has 10 slides

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

