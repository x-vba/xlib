Attribute VB_Name = "xlibNetworkTests"
Option Explicit

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


