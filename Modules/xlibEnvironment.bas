Attribute VB_Name = "xlibEnvironment"
'@Module: This module contains a set of functions for gathering information on the environment that Excel is being run on, such as the UserName of the computer, the OS Excel is being run on, and other Environment Variable values.

Option Private Module
Option Explicit


Public Function OS() As String

    '@Description: This function returns the Operating System name. Currently it will return either "Windows" or "Mac" depending on the OS used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns the name of the Operating System
    '@Example: =OS() -> "Windows"; When running this function on Windows
    '@Example: =OS() -> "Mac"; When running this function on MacOS

    #If Mac Then
        OS = "Mac"
    #Else
        OS = "Windows"
    #End If

End Function


Public Function UserName() As String

    '@Description: This function takes no arguments and returns a string of the USERNAME of the computer
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Returns: Returns a string of the username
    '@Example: =UserName() -> "Anthony"
    
    #If Mac Then
        UserName = Environ("USER")
    #Else
        UserName = Environ("USERNAME")
    #End If

End Function


Public Function UserDomain() As String

    '@Description: This function takes no arguments and returns a string of the USERDOMAIN of the computer
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Returns: Returns a string of the user domain of the computer
    '@Example: =UserDomain() -> "DESKTOP-XYZ1234"
    
    #If Mac Then
        UserDomain = Environ("HOST")
    #Else
        UserDomain = Environ("USERDOMAIN")
    #End If

End Function


Public Function ComputerName() As String

    '@Description: This function takes no arguments and returns a string of the COMPUTERNAME of the computer
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Returns: Returns a string of the computer name of the computer
    '@Example: =ComputerName() -> "DESKTOP-XYZ1234"

    ComputerName = Environ("COMPUTERNAME")

End Function
