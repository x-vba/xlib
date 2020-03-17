Attribute VB_Name = "xlibColorTests"
Option Explicit

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

    '@Example: =Rgb2Hex(255, 255, 255) -> "FFFFFF"
    
    Rgb2HexTest = True

    Rgb2HexTest = Rgb2HexTest And Rgb2Hex(255, 255, 255) = "FFFFFF"
    
End Function


Private Function Hex2RgbTest() As Boolean

    '@Example: =Hex2Rgb("FFFFFF") -> "(255, 255, 255)"
    '@Example: =Hex2Rgb("FF0109", 0) -> 255; The red color
    '@Example: =Hex2Rgb("FF0109", "Red") -> 255; The red color
    '@Example: =Hex2Rgb("FF0109", 1) -> 1; The green color
    '@Example: =Hex2Rgb("FF0109", "Green") -> 1; The green color
    '@Example: =Hex2Rgb("FF0109", 2) -> 9; The blue color
    '@Example: =Hex2Rgb("FF0109", "Blue") -> 9; The blue color

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

    '@Example: =Rgb2Hsl(8, 64, 128) -> "(212.0°, 88.2%, 26.7%)"
    '@Example: =Rgb2Hsl(8, 64, 128, 0) -> 212
    '@Example: =Rgb2Hsl(8, 64, 128, "Hue") -> 212
    '@Example: =Rgb2Hsl(8, 64, 128, 1) -> .882
    '@Example: =Rgb2Hsl(8, 64, 128, "Saturation") -> .882
    '@Example: =Rgb2Hsl(8, 64, 128, 2) -> .267
    '@Example: =Rgb2Hsl(8, 64, 128, "Lightness") -> .267

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

    '@Example: =Hex2Hsl("084080") -> "(212.0, 88.2%, 26.7%)"
    '@Example: =Hex2Hsl("#084080") -> "(212.0, 88.2%, 26.7%)"

    Hex2HslTest = True

    Hex2HslTest = Hex2HslTest And Hex2Hsl("084080") = "(212.0, 88.2%, 26.7%)"
    Hex2HslTest = Hex2HslTest And Hex2Hsl("#084080") = "(212.0, 88.2%, 26.7%)"

End Function


Private Function Hsl2RgbTest() As Boolean

    '@Example: =Hsl2Rgb(212, .882, .267) -> "(8, 64, 128)"
    '@Example: =Hsl2Rgb(212, .882, .267, 0) -> 8
    '@Example: =Hsl2Rgb(212, .882, .267, "Red") -> 8
    '@Example: =Hsl2Rgb(212, .882, .267, 1) -> 64
    '@Example: =Hsl2Rgb(212, .882, .267, "Green") -> 64
    '@Example: =Hsl2Rgb(212, .882, .267, 2) -> 128
    '@Example: =Hsl2Rgb(212, .882, .267, "Blue") -> 128

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

    '@Example: =Hsl2Hex(212, .882, .267) -> "084080"

    Hsl2HexTest = True

    Hsl2HexTest = Hsl2HexTest And Hsl2Hex(212, 0.882, 0.267) = "084080"

End Function


Private Function Rgb2HsvTest() As Boolean

    '@Example: =Rgb2Hsv(8, 64, 128) -> "(212.0, 93.8%, 50.2%)"
    '@Example: =Rgb2Hsv(8, 64, 128, 0) -> 212
    '@Example: =Rgb2Hsv(8, 64, 128, "Hue") -> 212
    '@Example: =Rgb2Hsv(8, 64, 128, 1) -> .938
    '@Example: =Rgb2Hsv(8, 64, 128, "Saturation") -> .938
    '@Example: =Rgb2Hsv(8, 64, 128, 2) -> .502
    '@Example: =Rgb2Hsv(8, 64, 128, "Value") -> .502

    Rgb2HsvTest = True

    Rgb2HsvTest = Rgb2HsvTest And Rgb2Hsv(8, 64, 128) = "(212.0, 93.8%, 50.2%)"
    Rgb2HsvTest = Rgb2HsvTest And Rgb2Hsv(8, 64, 128, 0) = 212
    Rgb2HsvTest = Rgb2HsvTest And Rgb2Hsv(8, 64, 128, "Hue") = 212
    Rgb2HsvTest = Rgb2HsvTest And Round(Rgb2Hsv(8, 64, 128, 1), 3) = 0.938
    Rgb2HsvTest = Rgb2HsvTest And Round(Rgb2Hsv(8, 64, 128, "Saturation"), 3) = 0.938
    Rgb2HsvTest = Rgb2HsvTest And Round(Rgb2Hsv(8, 64, 128, "Value"), 3) = 0.502
    
End Function
