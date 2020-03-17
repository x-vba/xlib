Attribute VB_Name = "xlibFile"
'@Module: This module contains a set of functions for gathering info on files. It includes functions for gathering file info on the current workbook presentation, document, or database, as well as functions for reading and writing to files, and functions for manipulating file path strings.

Option Private Module
Option Explicit


Public Function GetActivePathAndName() As String

    '@Description: This function returns the path of the file of the office program that is calling this function. It currently supports Excel, Word, PowerPoint, Access, and Publisher.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current path
    '@Example: =GetActivePathAndName() -> "C:\Users\UserName\Documents\XLib.xlsm"

    If Application.Name = "Microsoft Excel" Then
        GetActivePathAndName = GetActivePathAndNameExcel()
        
    ElseIf Application.Name = "Microsoft Word" Then
        GetActivePathAndName = GetActivePathAndNameWord()
        
    ElseIf Application.Name = "Microsoft PowerPoint" Then
        GetActivePathAndName = GetActivePathAndNamePowerPoint()
        
    ElseIf Application.Name = "Microsoft Access" Then
        GetActivePathAndName = GetActivePathAndNameAccess()
        
    ElseIf Application.Name = "Microsoft Publisher" Then
        GetActivePathAndName = GetActivePathAndNamePublisher()
        
    End If

End Function


Private Function GetActivePathAndNameExcel() As String

    '@Description: This function returns the path of the workbook calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current workbook path
    
    #If Mac Then
        GetActivePathAndNameExcel = ThisWorkbook.Path & "/" & ThisWorkbook.Name
    #Else
        GetActivePathAndNameExcel = ThisWorkbook.Path & "\" & ThisWorkbook.Name
    #End If

End Function


Private Function GetActivePathAndNameWord() As String

    '@Description: This function returns the path of the document calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current document path
    
    #If Mac Then
        GetActivePathAndNameWord = ThisDocument.Path & "/" & ThisDocument.Name
    #Else
        GetActivePathAndNameWord = ThisDocument.Path & "\" & ThisDocument.Name
    #End If

End Function


Private Function GetActivePathAndNamePowerPoint() As String

    '@Description: This function returns the path of the presentation calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current presentation path
    
    #If Mac Then
        GetActivePathAndNamePowerPoint = ActivePresentation.Path & "/" & ActivePresentation.Name
    #Else
        GetActivePathAndNamePowerPoint = ActivePresentation.Path & "\" & ActivePresentation.Name
    #End If

End Function


Private Function GetActivePathAndNameAccess() As String

    '@Description: This function returns the path of the database calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current database path
    
    #If Mac Then
        GetActivePathAndNameAccess = CurrentProject.Path & "/" & CurrentProject.Name
    #Else
        GetActivePathAndNameAccess = CurrentProject.Path & "\" & CurrentProject.Name
    #End If

End Function


Private Function GetActivePathAndNamePublisher() As String

    '@Description: This function returns the path of the publisher file calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current publisher file path
    
    #If Mac Then
        GetActivePathAndNamePublisher = ThisDocument.Path & "/" & ThisDocument.Name
    #Else
        GetActivePathAndNamePublisher = ThisDocument.Path & "\" & ThisDocument.Name
    #End If

End Function


Public Function GetActivePath() As String

    '@Description: This function returns the path of the folder of the office program that is calling this function. It currently supports Excel, Word, PowerPoint, Access, and Publisher.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current folder path
    '@Example: =GetActivePath() -> "C:\Users\UserName\Documents\"; Where the file resides in the Documents folder

    If Application.Name = "Microsoft Excel" Then
        GetActivePath = GetActivePathExcel()
        
    ElseIf Application.Name = "Microsoft Word" Then
        GetActivePath = GetActivePathWord()
        
    ElseIf Application.Name = "Microsoft PowerPoint" Then
        GetActivePath = GetActivePathPowerPoint()
        
    ElseIf Application.Name = "Microsoft Access" Then
        GetActivePath = GetActivePathAccess()
        
    ElseIf Application.Name = "Microsoft Publisher" Then
        GetActivePath = GetActivePathPublisher()
        
    End If

End Function


Private Function GetActivePathExcel() As String

    '@Description: This function returns the folder path of the workbook calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current workbook folder path
    
    #If Mac Then
        GetActivePathExcel = ThisWorkbook.Path & "/"
    #Else
        GetActivePathExcel = ThisWorkbook.Path & "\"
    #End If

End Function


Private Function GetActivePathWord() As String

    '@Description: This function returns the folder path of the document calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current document folder path
    
    #If Mac Then
        GetActivePathWord = ThisDocument.Path & "/"
    #Else
        GetActivePathWord = ThisDocument.Path & "\"
    #End If

End Function


Private Function GetActivePathPowerPoint() As String

    '@Description: This function returns the folder path of the presentation calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current presentation folder path
    
    #If Mac Then
        GetActivePathPowerPoint = ActivePresentation.Path & "/"
    #Else
        GetActivePathPowerPoint = ActivePresentation.Path & "\"
    #End If

End Function


Private Function GetActivePathAccess() As String

    '@Description: This function returns the folder path of the database calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current database folder path
    
    #If Mac Then
        GetActivePathAccess = CurrentProject.Path & "/"
    #Else
        GetActivePathAccess = CurrentProject.Path & "\"
    #End If

End Function


Private Function GetActivePathPublisher() As String

    '@Description: This function returns the folder path of the publisher file calling this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Returns: Returns a string of the current publisher folder path
    
    #If Mac Then
        GetActivePathPublisher = ThisDocument.Path & "/"
    #Else
        GetActivePathPublisher = ThisDocument.Path & "\"
    #End If

End Function


Public Function FileCreationTime( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the file creation time of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the file creation time of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileCreationTime() -> "1/1/2020 1:23:45 PM"
    '@Example: =FileCreationTime("C:\hello\world.txt") -> "1/1/2020 5:55:55 PM"
    '@Example: =FileCreationTime("vba.txt") -> "12/25/2000 1:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        FileCreationTime = FSO.GetFile(GetActivePathAndName()).DateCreated
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileCreationTime = FSO.GetFile(GetActivePath() & filePath).DateCreated
        Else
            FileCreationTime = FSO.GetFile(filePath).DateCreated
        End If
    End If

End Function


Public Function FileLastModifiedTime( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the file last modified time of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the file last modified time of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileLastModifiedTime() -> "1/1/2020 2:23:45 PM"
    '@Example: =FileLastModifiedTime("C:\hello\world.txt") -> "1/1/2020 7:55:55 PM"
    '@Example: =FileLastModifiedTime("vba.txt") -> "12/25/2000 3:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        FileLastModifiedTime = FSO.GetFile(GetActivePathAndName()).DateLastModified
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileLastModifiedTime = FSO.GetFile(GetActivePath() & filePath).DateLastModified
        Else
            FileLastModifiedTime = FSO.GetFile(filePath).DateLastModified
        End If
    End If

End Function


Public Function FileDrive( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the drive of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the file drive of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileDrive() -> "A:"; Where the current workbook resides on the A: drive
    '@Example: =FileDrive("C:\hello\world.txt") -> "C:"
    '@Example: =FileDrive("vba.txt") -> "B:"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in, and where the workbook resides in the B: drive

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        FileDrive = FSO.GetFile(GetActivePathAndName()).Drive
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileDrive = FSO.GetFile(GetActivePath() & filePath).Drive
        Else
            FileDrive = FSO.GetFile(filePath).Drive
        End If
    End If

End Function


Public Function FileName( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the name of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the name of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileName() -> "MyWorkbook.xlsm"
    '@Example: =FileName("C:\hello\world.txt") -> "world.txt"
    '@Example: =FileName("vba.txt") -> "vba.txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        FileName = FSO.GetFile(GetActivePathAndName()).Name
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileName = FSO.GetFile(GetActivePath() & filePath).Name
        Else
            FileName = FSO.GetFile(filePath).Name
        End If
    End If

End Function


Public Function FileFolder( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the path of the folder of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the path of the folder where the file resides in as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileFolder() -> "C:\my_excel_files"
    '@Example: =FileFolder("C:\hello\world.txt") -> "C:\hello"
    '@Example: =FileFolder("vba.txt") -> "C:\my_excel_files"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        FileFolder = FSO.GetFile(GetActivePathAndName()).ParentFolder
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileFolder = FSO.GetFile(GetActivePath() & filePath).ParentFolder
        Else
            FileFolder = FSO.GetFile(filePath).ParentFolder
        End If
    End If

End Function


Public Function CurrentFilePath( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the path of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the path of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =CurrentFilePath() -> "C:\my_excel_files\MyWorkbook.xlsx"
    '@Example: =CurrentFilePath("C:\hello\world.txt") -> "C:\hello\world.txt"
    '@Example: =CurrentFilePath("vba.txt") -> "C:\hello\world.txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        CurrentFilePath = FSO.GetFile(GetActivePathAndName()).Path
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            CurrentFilePath = FSO.GetFile(GetActivePath() & filePath).Path
        Else
            CurrentFilePath = FSO.GetFile(filePath).Path
        End If
    End If

End Function


Public Function FileSize( _
    Optional ByVal filePath As String, _
    Optional ByVal byteSize As String) _
As Double

    '@Description: This function returns the file size of the file specified in the file path argument, with the option to set if the file size is returned in Bytes, Kilobytes, Megabytes, or Gigabytes. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Param: byteSize is a string of value "KB", "MB", or "GB"
    '@Returns: Returns the size of the file as a Double
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileSize() -> 1024
    '@Example: =FileSize(,"KB") -> 1
    '@Example: =FileSize("vba.txt", "KB") -> 0.25; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim totalBytes As Double
    
    If filePath = "" Then
        totalBytes = FSO.GetFile(GetActivePathAndName()).Size
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            totalBytes = FSO.GetFile(GetActivePath() & filePath).Size
        Else
            totalBytes = FSO.GetFile(filePath).Size
        End If
    End If
    
    Select Case LCase(byteSize)
        Case "kb"
            totalBytes = totalBytes / (2 ^ 10)
        Case "mb"
            totalBytes = totalBytes / (2 ^ 20)
        Case "gb"
            totalBytes = totalBytes / (2 ^ 30)
    End Select

    FileSize = totalBytes

End Function


Public Function FileType( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the file type of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the file type of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: FileType() -> "Microsoft Excel Macro-Enabled Worksheet"
    '@Example: FileType("C:\hello\world.txt") -> "Text Document"
    '@Example: FileType("vba.txt") -> "Text Document"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")

    If filePath = "" Then
        FileType = FSO.GetFile(GetActivePathAndName()).Type
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileType = FSO.GetFile(GetActivePath() & filePath).Type
        Else
            FileType = FSO.GetFile(filePath).Type
        End If
    End If

End Function


Public Function FileExtension( _
    Optional ByVal filePath As String) _
As String

    '@Description: This function returns the extension of the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the extension of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Example: =FileExtension() = "xlsx"
    '@Example: =FileExtension("C:\hello\world.txt") -> "txt"
    '@Example: =FileExtension("vba.txt") -> "txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim FileName As String
    If filePath = "" Then
        FileName = FSO.GetFile(GetActivePathAndName()).Name
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            FileName = FSO.GetFile(GetActivePath() & filePath).Name
        Else
            FileName = FSO.GetFile(filePath).Name
        End If
    End If
    
    FileExtension = Right(FileName, Len(FileName) - InStrRev(FileName, "."))

End Function


Public Function ReadFile( _
    ByVal filePath As String, _
    Optional ByVal lineNumber As Integer) _
As String

    '@Description: This function reads the file specified in the file path argument and returns it's contents. Optionally, a line number can be specified so that only a single line is read. If a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Param: lineNumber is the number of the line that will be read, and if left blank all the file contents will be read. Note that the first line starts at line number 1.
    '@Returns: Returns the contents of the file as a string
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function may run very slowly when running it on large files. Also, for files that are not in text format (such as compressed zip files) this file contents returned will not be in a usable format.
    '@Example: =ReadFile("C:\hello\world.txt") -> "Hello" World
    '@Example: =ReadFile("vba.txt") -> "This is my VBA text file"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in
    '@Example: =ReadFile("multline.txt", 1) -> "This is line 1";
    '@Example: =ReadFile("multline.txt", 2) -> "This is line 2";

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim FileName As String
    Dim fileStream As Object
    
    ' Checking if the file exists in the current directory, and then if it
    ' exists in the path specified, and if it doesn't exist in either, returns
    ' a "#FileDoesntExist!"
    If FSO.FileExists(GetActivePath() & filePath) Then
        filePath = GetActivePath() & filePath
    ElseIf FSO.FileExists(filePath) Then
        filePath = filePath
    Else
        ReadFile = "#FileDoesntExist!"
    End If
    
    Set fileStream = FSO.GetFile(filePath)
    Set fileStream = fileStream.OpenAsTextStream(1, -2)
    
    
    ' If lineNumber is positive, read a line, else read the whole contents
    If lineNumber > 0 Then
        Dim fileLinesArray() As String
        
        fileLinesArray = Split(fileStream.ReadAll(), vbCrLf)
        ReadFile = fileLinesArray(lineNumber)
    Else
        ReadFile = fileStream.ReadAll()
    End If

End Function


Public Function WriteFile( _
    ByVal filePath As String, _
    ByVal fileText As String, _
    Optional ByVal appendModeFlag As Boolean) _
As Boolean

    '@Description: This function creates and writes to the file specified in the file path argument. If no file path is specified, the current workbook, document, presentation or database is used. Also, if a full path isn't used, a path relative to the folder the workbook, document, presentation or database resides in will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Todo: Refactor a bit shorten the code a bit, such as the area where the file is written to.
    '@Param: filePath is a string path of the file on the system, such as "C:\hello\world.txt"
    '@Param: fileText is the text that will be written to the file
    '@Param: appendModeFlag is a Boolean value that if set to TRUE will append to the existing file instead of creating a new file and writing over the contents.
    '@Returns: Returns a message stating the file written to successfully
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: Be careful when writing files, as misuse of this function can results in files being overwritten accidently as well as creating large numbers of files accidently.
    '@Example: =WriteFile("C:\MyWorkbookFolder\hello.txt", "Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"
    '@Example: =WriteFile("hello.txt", "Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"; Where the Workbook resides in "C:\MyWorkbookFolder\"

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    ' Checking if the folder exists if the path is an absolute path
    If InStr(filePath, "\") = 0 Then
        If InStr(filePath, "/") = 0 Then
            filePath = GetActivePath() & filePath
        End If
    ElseIf Right(filePath, 1) = "\" Or Right(filePath, 1) = "/" Then
        If Not FSO.FolderExists(Left(filePath, InStrRev(filePath, "\"))) Then
            WriteFile = False
            Exit Function
        End If
    ElseIf Not FSO.FolderExists(filePath) Then
        WriteFile = False
        Exit Function
    End If
    
    
    ' Writing to the file
    Dim fileStream As Object
    
    If appendModeFlag = False Then
        Set fileStream = FSO.CreateTextFile(filePath, True)
        fileStream.Write fileText
        
    Else
        Dim fileObject As Object
        
        Set fileObject = FSO.GetFile(filePath)
        Set fileStream = fileObject.OpenAsTextStream(8)
        fileStream.Write fileText
    End If
    
    WriteFile = True

End Function


Public Function PathSeparator() As String

    '@Description: This function returns the path separator character of the OS running this function
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Example: =PathSeparator() -> "\"; When running this code on Windows
    '@Example: =PathSeparator() -> "/"; When running this code on Mac
    
    #If Mac Then
        PathSeparator = "/"
    #Else
        PathSeparator = "\"
    #End If

End Function


Public Function PathJoin( _
    ParamArray pathArray() As Variant) _
As String

    '@Description: This function combines multiple strings into a file path by placing the path separator character between the arguments
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: pathArray is an array of strings that will be combined into a path
    '@Returns: Returns a string with the combined file path
    '@Example: =PathJoin("C:", "hello", "world.txt") -> "C:\hello\world.txt"; On Windows
    '@Example: =PathJoin("hello", "world.txt") -> "/hello/world.txt"; On Mac

    Dim individualPath As Variant
    Dim combinedPath As String
    Dim individualValue As Variant

    For Each individualPath In pathArray
        If IsArray(individualPath) Then
            For Each individualValue In individualPath
                combinedPath = combinedPath & individualValue & PathSeparator()
            Next
        Else
            combinedPath = combinedPath & CStr(individualPath) & PathSeparator()
        End If
    Next
    
    combinedPath = Left(combinedPath, Len(combinedPath) - 1)
    
    PathJoin = combinedPath
    
End Function


Public Function CountFiles( _
    Optional ByVal filePath As String) _
As Integer

    '@Description: This function returns the number of files at the specified folder path. If no path is given, the current workbook path will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the number of files in the folder
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function includes the count for hidden files as well. For example, when a workbook is open, a hidden file for the workbook is created, so if you run this function in the same folder as the workbook and notice the file count is one higher than expected, it is likely due to the hidden file.
    '@Example: =CountFiles() -> 6
    '@Example: =CountFiles("C:\hello") -> 10

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        CountFiles = FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).Files.Count
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            CountFiles = FSO.GetFolder(GetActivePath() & filePath).Files.Count
        Else
            CountFiles = FSO.GetFolder(filePath).Files.Count
        End If
    End If

End Function


Public Function CountFolders( _
    Optional ByVal filePath As String) _
As Integer

    '@Description: This function returns the number of folders at the specified folder path. If no path is given, the current workbook path will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the number of folders in the folder
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function includes the count for hidden folders as well. Hidden folders are often prefixed with a . character at the beginning
    '@Example: =CountFolders() -> 2
    '@Example: =CountFolders("C:\hello") -> 20

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        CountFolders = FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).SubFolders.Count
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            CountFolders = FSO.GetFolder(GetActivePath() & filePath).SubFolders.Count
        Else
            CountFolders = FSO.GetFolder(filePath).SubFolders.Count
        End If
    End If

End Function


Public Function CountFilesAndFolders( _
    Optional ByVal filePath As String) _
As Integer

    '@Description: This function returns the number of files and folders at the specified folder path. If no path is given, the current workbook path will be used.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Returns: Returns the number of files and folders in the folder
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function includes the count for hidden files and folders as well
    '@Example: =CountFilesAndFolders() -> 8
    '@Example: =CountFilesAndFolders("C:\hello") -> 30

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If filePath = "" Then
        CountFilesAndFolders = FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).Files.Count + FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).SubFolders.Count
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            CountFilesAndFolders = FSO.GetFolder(GetActivePath() & filePath).Files.Count + FSO.GetFolder(GetActivePath() & filePath).SubFolders.Count
        Else
            CountFilesAndFolders = FSO.GetFolder(filePath).Files.Count + FSO.GetFolder(filePath).SubFolders.Count
        End If
    End If

End Function


Public Function GetFileNameByNumber( _
    Optional ByVal filePath As String, _
    Optional ByVal fileNumber As Integer = -1) _
As String

    '@Description: This function returns the name of a file in a folder given the number of the file in the list of all files
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: filePath is a string path to the file on the system, such as "C:\hello\world.txt"
    '@Param: fileNumber is the number of the file in the folder. For example, if there are 3 files in a folder, this should be a number between 1 and 3
    '@Returns: Returns the name of the specified file
    '@Note: You can find the path of a file via Shift+RightClick -> Copy as Path; on a file in the Windows Explorer
    '@Warning: This function includes hidden files as well. For example, when a workbook is open, a hidden file for the workbook is created, so if you run this function in the same folder as the workbook and notice the file count is one higher than expected, it is likely due to the hidden file.
    '@Example: =GetFileName(,1) -> "hello.txt"
    '@Example: =GetFileName(,1) -> "world.txt"
    '@Example: =GetFileName("C:\hello", 1) -> "one.txt"
    '@Example: =GetFileName("C:\hello", 1) -> "two.txt"
    '@Example: =GetFileName("C:\hello", 1) -> "three.txt"

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Dim fileCounter As Integer
    Dim individualFile As Object
    Dim fileCollection As Object
    
    If filePath = "" Then
        Set fileCollection = FSO.GetFolder(FSO.GetParentFolderName(GetActivePathAndName())).Files
    Else
        If FSO.FileExists(GetActivePath() & filePath) Then
            Set fileCollection = FSO.GetFolder(GetActivePath() & filePath).Files
        Else
            Set fileCollection = FSO.GetFolder(filePath).Files
        End If
    End If
    
    For Each individualFile In fileCollection
        fileCounter = fileCounter + 1
        If fileNumber = -1 Then
            GetFileNameByNumber = individualFile.Name
            Exit Function
        ElseIf fileCounter = fileNumber Then
            GetFileNameByNumber = individualFile.Name
            Exit Function
        End If
    Next

End Function
