Attribute VB_Name = "xlibRegex"
'@Module: This module contains a set of functions for performing Regular Expressions, which are a type of string pattern matching. For more info on Regular Expression Pattern matching, please check "https://docs.microsoft.com/en-us/dotnet/standard/base-types/regular-expression-language-quick-reference"

Option Explicit


Public Function RegexSearch( _
    ByVal string1 As String, _
    ByVal stringPattern As String, _
    Optional ByVal globalFlag As Boolean, _
    Optional ByVal ignoreCaseFlag As Boolean, _
    Optional ByVal multilineFlag As Boolean) _
As String

    '@Description: This function takes a string that we will perform the Regular Expression on and a Regular Expression string pattern, and returns the first value of the matched string. This function also contains optional arguments for various Regular Expression flags.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that the regex will be performed on
    '@Param: stringPattern is the regex pattern
    '@Param: globalFlag is a boolean value that if set TRUE will perform a global search
    '@Param: ignoreCaseFlag is a boolean value that if set TRUE will perform a case insensitive search
    '@Param: multilineFlag is a boolean value that if set TRUE will perform a mulitline search
    '@Returns: Returns a string of the regex value that is found
    '@Example: =RegexSearch("Hello World","[a-z]{2}\s[W]") -> "lo W";

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


Public Function RegexTest( _
    ByVal string1 As String, _
    ByVal stringPattern As String, _
    Optional ByVal globalFlag As Boolean, _
    Optional ByVal ignoreCaseFlag As Boolean, _
    Optional ByVal multilineFlag As Boolean) _
As Boolean

    '@Description: This function takes a string that we will perform the Regular Expression on and a Regular Expression string pattern, and returns TRUE if the pattern is found in the string. This function also contains optional arguments for various Regular Expression flags.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that the regex will be performed on
    '@Param: stringPattern is the regex pattern
    '@Param: globalFlag is a boolean value that if set TRUE will perform a global search
    '@Param: ignoreCaseFlag is a boolean value that if set TRUE will perform a case insensitive search
    '@Param: multilineFlag is a boolean value that if set TRUE will perform a mulitline search
    '@Returns: Returns TRUE if the regex value that is found, or FALSE if it isn't
    '@Example: =RegexTest("Hello World","[a-z]{2}\s[W]") -> TRUE;

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


Public Function RegexReplace( _
    ByVal string1 As String, _
    ByVal stringPattern As String, _
    ByVal replacementString As String, _
    Optional ByVal globalFlag As Boolean, _
    Optional ByVal ignoreCaseFlag As Boolean, _
    Optional ByVal multilineFlag As Boolean) _
As String

    '@Description: This function takes a string that we will perform the Regular Expression on, a Regular Expression string pattern, and a string that we will replace if the pattern is found, and returns a new string with the replacement string in place of the pattern. This function also contains optional arguments for various Regular Expression flags.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string that the regex will be performed on
    '@Param: stringPattern is the regex pattern
    '@Param: replacementString is a string that will be replaced if the pattern is found
    '@Param: globalFlag is a boolean value that if set TRUE will perform a global search
    '@Param: ignoreCaseFlag is a boolean value that if set TRUE will perform a case insensitive search
    '@Param: multilineFlag is a boolean value that if set TRUE will perform a mulitline search
    '@Returns: Returns a new string with the replaced string values
    '@Example: =RegexReplace("Hello World","[W][a-z]{4}", "VBA") -> "Hello VBA"

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

