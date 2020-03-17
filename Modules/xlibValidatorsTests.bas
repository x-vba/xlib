Attribute VB_Name = "xlibValidatorsTests"
Option Explicit

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


