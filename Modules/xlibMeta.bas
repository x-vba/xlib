Attribute VB_Name = "xlibMeta"
'@Module: This module contains a set of functions that return information on the Xlib library, such as the version number, credits, and a link to the documentation.

Option Explicit


Public Function XlibVersion() As String

    '@Description: This function returns the version number of XPlus
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns the XPlus version number
    '@Example: =XlibVersion() -> "1.0.0"; Where the version of XPlus you are using is 1.0.0

    XlibVersion = "1.0.0"

End Function


Public Function XlibCredits() As String

    '@Description: This function returns credits for the XPlus library
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns the XPlus credits
    '@Example: =XlibCredits() -> "Copyright (c) 2020 Anthony Mancini. XLib is Licensed under an MIT License."

    XlibCredits = "Copyright (c) 2020 Anthony Mancini. XLib is Licensed under an MIT License."

End Function


Public Function XlibDocumentation() As String

    '@Description: This function returns a link to the Documentation for XPlus
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns the XPlus Documentation link
    '@Example: =XlibDocumentation() -> "https://x-vba.com/xlib"

    XlibDocumentation = "https://x-vba.com/xlib"

End Function

