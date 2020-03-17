Attribute VB_Name = "xlibValidators"
'@Module: This module contains a set of functions for validating some commonly used string, such as validators for email addresses and phone numbers.

Option Explicit


Public Function IsEmail( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid email address.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Improve regex robustness
    '@Param: string1 is the string we are checking if its a valid email
    '@Returns: Returns TRUE if the string is a valid email, and FALSE if its invalid
    '@Example: =IsEmail("JohnDoe@testmail.com") -> TRUE
    '@Example: =IsEmail("JohnDoe@test/mail.com") -> FALSE
    '@Example: =IsEmail("not_an_email_address") -> FALSE

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


Public Function IsPhone( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a phone number is valid.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Improve regex robustness
    '@Todo: Add a second argument that lets the user add a country code and uses a different regex for phone number formats for that country. Also make the regx more robust so it can include more common formats.
    '@Param: string1 is the string we are checking if its a valid phone number
    '@Returns: Returns TRUE if the string is a valid phone number, and FALSE if its invalid
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


Public Function IsCreditCard( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid credit card from one of the major card issuing companies.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string we are checking if its a valid credit card number
    '@Returns: Returns TRUE if the string is a valid credit card number, and FALSE if its invalid. Currently supports these cards: Visa, MasterCard, Discover, Amex, Diners, JCB
    '@Example: =IsCreditCard("5111567856785678") -> TRUE; This is a valid Mastercard number
    '@Example: =IsCreditCard("511156785678567") -> FALSE; Not enough digits
    '@Example: =IsCreditCard("9999999999999999") -> FALSE; Enough digits, but not a valid card number
    '@Example: =IsCreditCard("Hello World") -> FALSE

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
        
    Dim regexPattern As String
    
    ' Regex for Amex
    regexPattern = regexPattern & "(3[47][0-9]{13})|"
    
    ' Regex for Diners
    regexPattern = regexPattern & "(3(0[0-5]|[68][0-9])?[0-9]{11})|"
    
    ' Regex for Discover
    regexPattern = regexPattern & "(6(011|5[0-9]{2})[0-9]{12})|"
    
    ' Regex for JCB
    regexPattern = regexPattern & "((2131|1800|35[0-9]{3})[0-9]{11})|"
    
    ' Regex for MasterCard
    regexPattern = regexPattern & "(5[1-5][0-9]{14})|"
    
    ' Regex for Visa
    regexPattern = regexPattern & "(4[0-9]{12}([0-9]{3})?)"
    
    With Regex
        .Global = True
        .IgnoreCase = True
        .MultiLine = True
        .Pattern = regexPattern
    End With

    IsCreditCard = Regex.Test(string1)

End Function


Public Function IsUrl( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid URL address.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Improve regex robustness
    '@Param: string1 is the string we are checking if its a valid URL
    '@Returns: Returns TRUE if the string is a valid URL, and FALSE if its invalid
    '@Example: =IsUrl("https://www.wikipedia.org/") -> TRUE
    '@Example: =IsUrl("http://www.wikipedia.org/") -> TRUE
    '@Example: =IsUrl("hello_world") -> FALSE

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


Public Function IsIPFour( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid IPv4 address.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Improve regex robustness
    '@Param: string1 is the string we are checking if its a valid IPv4 address
    '@Returns: Returns TRUE if the string is a valid IPv4, and FALSE if its invalid
    '@Example: =IsIPFour("0.0.0.0") -> TRUE
    '@Example: =IsIPFour("100.100.100.100") -> TRUE
    '@Example: =IsIPFour("255.255.255.255") -> TRUE
    '@Example: =IsIPFour("255.255.255.256") -> FALSE; as the final 256 makes the address outside of the bounds of IPv4
    '@Example: =IsIPFour("0.0.0") -> FALSE; as the fourth octet is missing

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


Public Function IsMacAddress( _
    ByVal string1 As String) _
As Boolean

    '@Description: This function checks if a string is a valid 48-bit Mac Address.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the string we are checking if its a valid 48-bit Mac Address
    '@Returns: Returns TRUE if the string is a valid 48-bit Mac Address, and FALSE if its invalid
    '@Example: =IsMacAddress("00:25:96:12:34:56") -> TRUE
    '@Example: =IsMacAddress("FF:FF:FF:FF:FF:FF") -> TRUE
    '@Example: =IsMacAddress("00-25-96-12-34-56") -> TRUE
    '@Example: =IsMacAddress("123.789.abc.DEF") -> TRUE
    '@Example: =IsMacAddress("Not A Mac Address") -> FALSE
    '@Example: =IsMacAddress("FF:FF:FF:FF:FF:FH") -> FALSE; the H at the end is not a valid Hex number

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


Public Function CreditCardName( _
    ByVal string1 As String) _
As String

    '@Description: This function checks if a string is a valid credit card from one of the major card issuing companies, and then returns the name of the credit card name. This function assumes no spaces or hyphens (if you have card numbers with spaces or hyphens you can remove these using =SUBSTITUTE("-", "") function.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the credit card string
    '@Returns: Returns the name of the credit card. Currently supports these cards: Visa, MasterCard, Discover, Amex, Diners, JCB
    '@Example: =CreditCardName("5111567856785678") -> "MasterCard"; This is a valid Mastercard number
    '@Example: =CreditCardName("not_a_card_number") -> #VALUE!

    Dim Regex As Object
    Set Regex = CreateObject("VBScript.RegExp")
    
    Regex.Global = True
    Regex.IgnoreCase = True
    Regex.MultiLine = True

    ' Regex for Amex
    Regex.Pattern = "(3[47][0-9]{13})"
    If Regex.Test(string1) Then
        CreditCardName = "Amex"
        Exit Function
    End If
    
    ' Regex for Diners
    Regex.Pattern = "(3(0[0-5]|[68][0-9])?[0-9]{11})"
    If Regex.Test(string1) Then
        CreditCardName = "Diners"
        Exit Function
    End If
    
    ' Regex for Discover
    Regex.Pattern = "(6(011|5[0-9]{2})[0-9]{12})"
    If Regex.Test(string1) Then
        CreditCardName = "Discover"
        Exit Function
    End If
    
    ' Regex for JCB
    Regex.Pattern = "((2131|1800|35[0-9]{3})[0-9]{11})"
    If Regex.Test(string1) Then
        CreditCardName = "JCB"
        Exit Function
    End If
    
    ' Regex for MasterCard
    Regex.Pattern = "(5[1-5][0-9]{14})"
    If Regex.Test(string1) Then
        CreditCardName = "MasterCard"
        Exit Function
    End If
    
    ' Regex for Visa
    Regex.Pattern = "(4[0-9]{12}([0-9]{3})?)"
    If Regex.Test(string1) Then
        CreditCardName = "Visa"
        Exit Function
    End If
    
    CreditCardName = "#NotAValidCreditCardNumber!"

End Function


Public Function FormatCreditCard( _
    ByVal string1 As String) _
As String

    '@Description: This function checks if a string is a valid credit card, and if it is formats it in a more readable way. The format used is XXXX-XXXX-XXXX-XXXX.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is credit card number
    '@Returns: Returns a string formatted as a more readable credit card number
    '@Example: =FormatCreditCard("5111567856785678") -> "5111-5678-5678-5678"

    If IsCreditCard(string1) Then
        FormatCreditCard = Left(string1, 4) & "-" & Mid(string1, 5, 4) & "-" & Mid(string1, 9, 4) & "-" & Mid(string1, 13)
    Else
        FormatCreditCard = "#NotAValidCreditCardNumber!"
    End If

End Function
