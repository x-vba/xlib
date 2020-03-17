Attribute VB_Name = "xlibNetwork"
'@Module: This module contains a set of functions for performing networking tasks such as performing HTTP requests and parsing HTML.

Option Explicit


Public Function Http( _
    ByVal url As String, _
    Optional ByVal httpMethod As String = "GET", _
    Optional ByVal headers As Variant, _
    Optional ByVal postData As Variant = "", _
    Optional ByVal asyncFlag As Boolean, _
    Optional ByVal statusErrorHandlerFlag As Boolean, _
    Optional ByVal parseArguments As Variant) _
As String

    '@Description: This function performs an HTTP request to the web and returns the response as a string. It provides many options to change the http method, provide data for a POST request, change the headers, handle errors for non-successful requests, and parse out text from a request using a light parsing language.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: url is a string of the URL of the website you want to fetch data from
    '@Param: httpMethod is a string with the http method, with the default being a GET request. For POST requests, use "POST", for PUT use "PUT", and for DELETE use "DELETE"
    '@Param: headers is either an array or a Scripting Dictionary of headers that will be used in the request. For an array, the 1st, 3rd, 5th... will be used as the key and the 2nd, 4th, 6th... will be used as the values. For a Scripting Dictionary, the dictionary keys will be used as header keys, and the values as values. Finally, in the case when no headers are set, the User-Agent will be set to "XPlus" as a courtesy to the web server.
    '@Param: postData is a string that will contain data for a POST request
    '@Param: asyncFlag is a Boolean value that if set to TRUE will make the request asynchronous. By default requests will be synchronous, which will lock Excel while fetching but will also prevent errors when performing calculations based on fetched data.
    '@Param: statusErrorHandlerFlag is a Boolean value that if set to TRUE will result in a User-Defined Error String being returned for all non 200 requests that tells the user the status code that occured. This flag is useful in cases where requests need to be successful and if not errors should be thrown.
    '@Param: parseArguments is an array of arguments that perform string parsing on the response. It uses a light scripting language that includes commands similar to the Excel Built-in LEFT(), RIGHT(), and MID() that allow you to parse the request before it gets returned. See the Note on the scripting language, and the Warning on why this argument should be used.
    '@Returns: Returns the parsed HTTP response as a string
    '@Note: The parseArguments parameter uses a light scripting language to perform string manipulations on the HTTP response text that allows you to parse out the relevant information to you. The language contains 5 commands that can be used for parsing. Please check out the examples as well below for a better understanding of how to use the parsing language:<br><br> {"ID", "idOfAnElement"} -> HTML inside of the element with the specified ID <br> {"TAG", "div", 2} -> HTML inside of the second div tag found <br> {"LEFT", 100} -> The 100 leftmost characters <br> {"LEFT", "Hello World"} -> All characters left of the first "Hello World" found in the HTML <br> {"RIGHT", 100} -> The 100 rightmost characters <br> {"RIGHT", "Hello World"} -> All characters right of the last "Hello World" found in the HTML <br> {"MID", 100} -> All character to the right of the 100th character in the string <br> {"MID", "Hello World"} -> All characters right of the first "Hello World" found in the HTML
    '@Warning: Excel has a limit on the number of characters that can be placed within a cell. This limit is a max of 32767 characters. If the request returns any more than this, a #VALUE! error will be returned. Most webpages surpass this number of characters, which makes the Excel Built-in function WEBSERVICE() not very useful. However, internally VBA can handle around 2,000,000,000 characters, which more characters that found on virtually every single webpage. As a result, parsing arguments should be used with this function so that you can parse out the relevant information for a request without this function failing. See the Note on the syntax of the light parsing language.
    '@Example: =Http("https://httpbin.org/uuid") -> "{"uuid: "41416bcf-ef11-4256-9490-63853d14e4e8"}"
    '@Example: =Http("https://httpbin.org/user-agent", "GET", {"User-Agent","MicrosoftExcel"}) -> "{"user-agent": "MicrosoftExcel"}"
    '@Example: =Http("https://httpbin.org/status/404",,,,,TRUE) -> "#RequestFailedStatusCode404!"; Since the status error handler flag is set and since this URL returns a 404 status code. Also note that this formula is easier to construct using the Excel Formula Builder
    '@Example: =Http("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000}) -> Returning a string with the leftmost 3000 characters found within the element with the ID "mw-content-text" (we are trying to get the release date of VBA from the VBA wikipedia page, but we need to do more parsing first)
    '@Example: =Http("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000,"MID","appeared"}) -> Returns the prior string, but now with all characters right of the first occurance of the word "appeared" in the HTML (getting closer to parsing the VBA creation date)
    '@Example: =Http("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>"}) -> From the prior result, now returning everything after the first occurance of the "<TD>" in the prior string
    '@Example: =Http("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications",,{"User-Agent","MicrosoftExcel"},,,,{"ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>","LEFT","<span"}) -> "1993"; Finally this is all the parsing needed to be able to return the date 1993 that we were looking for

    Dim WinHttpRequest As Object
    Set WinHttpRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    WinHttpRequest.Open httpMethod, url, asyncFlag
    
    ' Setting the request headers
    ' Case where headers come in the form of an Array
    If IsArray(headers) Then
        Dim i As Integer
        
        For i = 0 To UBound(headers) - LBound(headers) Step 2
            WinHttpRequest.SetRequestHeader headers(i), headers(i + 1)
        Next
        
    ' Case where headers come in the form of a Dictionary
    ElseIf TypeName(headers) = "Dictionary" Then
        Dim dictKey As Variant
        
        For Each dictKey In headers.Keys()
            WinHttpRequest.SetRequestHeader dictKey, headers(dictKey)
        Next
        
    ' In cases where no headers are given by the user, set a base User-Agent to
    ' "XPlus" as a courtesy to the webserver
    Else
        WinHttpRequest.SetRequestHeader "User-Agent", "XLib"
    End If
    
    ' Sending the HTTP request
    If postData = "" Then
        WinHttpRequest.Send
    Else
        WinHttpRequest.Send postData
    End If
    
    ' If the status error handler flag is set to True, then enable error returns
    ' in cases where the status code is not a 200
    If statusErrorHandlerFlag Then
        If WinHttpRequest.Status = 200 Then
            Http = WinHttpRequest.ResponseText
        Else
            Http = "#RequestFailedStatusCode" & WinHttpRequest.Status & "!"
        End If
    
    ' Case when the status code error handler is not used
    Else
        Http = WinHttpRequest.ResponseText
    End If
    
    ' Parsing Html Response
    If IsArray(parseArguments) Then
        Dim reorderedParseArguments() As Variant
        i = UBound(parseArguments) - LBound(parseArguments)
        ReDim reorderedParseArguments(i)
        
        ' Reordering here, as possibly had some name collision with the name parseArguments somewhere

        For i = 0 To UBound(parseArguments) - LBound(parseArguments)
            reorderedParseArguments(i) = parseArguments(i)
        Next
        
        Http = ParseHtmlString(Http, reorderedParseArguments)
    
    End If

End Function


Public Function SimpleHttp( _
    ByVal url As String, _
    ParamArray parseArguments() As Variant) _
As String

    '@Description: This function performs an HTTP request to the web and returns the response as a string, similar to the HTTP() function, except that only requires one parameter, the URL, and then takes an infinite number of strings after it as the parsing arguments instead of requiring an Array to use. Essentially, this function is a little cleaner to set up when performing very basic GET requests.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: url is a string of the URL of the website you want to fetch data from
    '@Param: parseArguments is an array of arguments that perform string parsing on the response. It uses a light scripting language that includes commands similar to the Excel Built-in LEFT(), RIGHT(), and MID() that allow you to parse the request before it gets returned. See the Note on the HTTP() function, and the Warning on the HTTP() function on why this argument should be used.
    '@Returns: Returns the parsed HTTP response as a string
    '@Example: =SimpleHttp("https://en.wikipedia.org/wiki/Visual_Basic_for_Applications","ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>","LEFT","<span") -> "1993"; See the examples in the HTTP() function, as this example has the same result as the example in the HTTP() function. You can see that this function is cleaner and easier to set up than the corresponding HTTP() function.

    ' Case where parse arguments are provided
    If UBound(parseArguments) > 0 Then
        ' Need to reorder the arguments of the Array since when the caller is a
        ' Range, the Array is 1-based, where as when the caller is another VBA function,
        ' the Array is 0-based
        Dim i As Integer
        Dim reorderedParseArguments() As Variant
        i = UBound(parseArguments) - LBound(parseArguments)
        ReDim reorderedParseArguments(i)
        
        ' Reordering for Range
        For i = 0 To UBound(parseArguments) - LBound(parseArguments)
            reorderedParseArguments(i) = parseArguments(i)
        Next
        
        SimpleHttp = ParseHtmlString(Http(url), reorderedParseArguments)
    
    ' In case of no parse arguments, simply perform an HTTP request
    Else
        SimpleHttp = Http(url)
    End If

End Function


Public Function ParseHtmlString( _
    ByVal htmlString As String, _
    ByVal parseArguments As Variant) _
As Variant

    '@Description: This function parses an HTML string using the same parsing language that the HTTP() function uses. See the HTTP() function for more information on how to use this function.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: htmlString is a string of the HTML
    '@Param: parseArguments is an array of arguments that perform string parsing on the response. It uses a light scripting language that includes commands similar to the Excel Built-in LEFT(), RIGHT(), and MID() that allow you to parse the request before it gets returned. See the Note on the HTTP() function, and the Warning on the HTTP() function on why this argument should be used.
    '@Returns: Returns the parsed HTTP response as a string
    '@Example: =ParseHtmlString("HTML String from the webpage: https://en.wikipedia.org/wiki/Visual_Basic_for_Applications","ID","mw-content-text","LEFT",3000,"MID","appeared","MID","<TD>","LEFT","<span") -> "1993"

    Dim partialHtml As String
    Dim html As Object
    Set html = CreateObject("HtmlFile")
    
    ' Setting the HTML Document
    html.body.innerHTML = htmlString
    
    ' Parsing out info from the HTML Document
    Dim i As Integer
    
    For i = LBound(parseArguments) To UBound(parseArguments)
        ' Note that id and tag will truncate poorly formatted HTML
        ' Works with late bindings
        If LCase(parseArguments(i)) = "id" Then
            If partialHtml <> "" Then
                html.body.innerHTML = partialHtml
            End If
            partialHtml = html.getElementById(parseArguments(i + 1)).innerHTML
            html.body.innerHTML = partialHtml
            i = i + 1
            
        ' Requires early bindings. Don't include in final code, but potentially consider for future updates
        'ElseIf LCase(parseArguments(i)) = "class" Then
        '    partialHtml = html.getElementsByClassName(parseArguments(i + 1))(i + 2).innerHTML
        '    i = i + 2
        
        ' Works with late bindings
        ElseIf LCase(parseArguments(i)) = "tag" Then
            If partialHtml <> "" Then
                html.body.innerHTML = partialHtml
            End If
            partialHtml = html.getElementsByTagName(parseArguments(i + 1))(i + 2).innerHTML
            html.body.innerHTML = partialHtml
            i = i + 2
            
        ' Left string manipulation
        ElseIf LCase(parseArguments(i)) = "left" Then
            If IsNumeric(parseArguments(i + 1)) And TypeName(parseArguments(i + 1)) <> "String" Then
                partialHtml = Left(partialHtml, parseArguments(i + 1))
            Else
                partialHtml = Left(partialHtml, InStr(1, partialHtml, CStr(parseArguments(i + 1)), vbTextCompare) - 1)
            End If
            i = i + 1
            
        ' Right string manipulation
        ElseIf LCase(parseArguments(i)) = "right" Then
            If IsNumeric(parseArguments(i + 1)) And TypeName(parseArguments(i + 1)) <> "String" Then
                partialHtml = Right(partialHtml, parseArguments(i + 1))
            Else
                partialHtml = Right(partialHtml, Len(partialHtml) - Len(parseArguments(i + 1)) + 1 - InStrRev(partialHtml, CStr(parseArguments(i + 1)), Compare:=vbTextCompare))
            End If
            i = i + 1
            
        ' Mid string manipulation. Possibly update this to allow Mid length argument
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


