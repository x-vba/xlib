Attribute VB_Name = "xlibEnvironmentTests"
Option Explicit

Public Function AllXlibEnvironmentTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not OSTest() Then
        Debug.Print "Failed: OSTest"
        TestStatus = False
    Else
        Debug.Print "Passed: OSTest"
    End If
    
    If Not UserNameTest() Then
        Debug.Print "Failed: UserNameTest"
        TestStatus = False
    Else
        Debug.Print "Passed: UserNameTest"
    End If
    
    If Not UserDomainTest() Then
        Debug.Print "Failed: UserDomainTest"
        TestStatus = False
    Else
        Debug.Print "Passed: UserDomainTest"
    End If
    
    If Not ComputerNameTest() Then
        Debug.Print "Failed: ComputerNameTest"
        TestStatus = False
    Else
        Debug.Print "Passed: ComputerNameTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibEnvironmentTests = TestStatus
    
End Function



Private Function OSTest() As Boolean

    '@Example: =OS() -> "Windows"; When running this function on Windows
    '@Example: =OS() -> "Mac"; When running this function on MacOS

    OSTest = True

    #If Mac Then
        OSTest = OSTest And OS() = "Mac"
    #Else
        OSTest = OSTest And OS() = "Windows"
    #End If

End Function


Private Function UserNameTest() As Boolean

    '@Example: =UserName() -> "Anthony"
    
    UserNameTest = True

    #If Mac Then
        UserNameTest = UserNameTest And UserName() = Environ("USER")
    #Else
        UserNameTest = UserNameTest And UserName() = Environ("USERNAME")
    #End If

End Function


Private Function UserDomainTest() As Boolean

    '@Example: =UserDomain() -> "DESKTOP-XYZ1234"
    
    UserDomainTest = True
    
    #If Mac Then
        UserDomainTest = UserDomainTest And UserDomain() = Environ("HOST")
    #Else
        UserDomainTest = UserDomainTest And UserDomain() = Environ("USERDOMAIN")
    #End If

End Function


Private Function ComputerNameTest() As Boolean

    '@Example: =ComputerName() -> "DESKTOP-XYZ1234"

    ComputerNameTest = True
    
    ComputerNameTest = ComputerNameTest And ComputerName() = Environ("COMPUTERNAME")

End Function

