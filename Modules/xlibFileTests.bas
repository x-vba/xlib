Attribute VB_Name = "xlibFileTests"
Option Explicit

Public Function AllXlibFileTests()

    Dim TestStatus As Boolean
    TestStatus = True
    
    Debug.Print "========================================"
    
    ' Begin Tests
    If Not GetActivePathAndNameTest() Then
        Debug.Print "Failed: GetActivePathAndNameTest"
        TestStatus = False
    Else
        Debug.Print "Passed: GetActivePathAndNameTest"
    End If
    
    If Not GetActivePathTest() Then
        Debug.Print "Failed: GetActivePathTest"
        TestStatus = False
    Else
        Debug.Print "Passed: GetActivePathTest"
    End If
    
    If Not FileCreationTimeTest() Then
        Debug.Print "Failed: FileCreationTimeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileCreationTimeTest"
    End If
    
    If Not FileLastModifiedTimeTest() Then
        Debug.Print "Failed: FileLastModifiedTimeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileLastModifiedTimeTest"
    End If
    
    If Not FileDriveTest() Then
        Debug.Print "Failed: FileDriveTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileDriveTest"
    End If
    
    If Not FileNameTest() Then
        Debug.Print "Failed: FileNameTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileNameTest"
    End If
    
    If Not FileFolderTest() Then
        Debug.Print "Failed: FileFolderTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileFolderTest"
    End If
    
    If Not CurrentFilePathTest() Then
        Debug.Print "Failed: CurrentFilePathTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CurrentFilePathTest"
    End If
    
    If Not FileSizeTest() Then
        Debug.Print "Failed: FileSizeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileSizeTest"
    End If
    
    If Not FileTypeTest() Then
        Debug.Print "Failed: FileTypeTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileTypeTest"
    End If
    
    If Not FileExtensionTest() Then
        Debug.Print "Failed: FileExtensionTest"
        TestStatus = False
    Else
        Debug.Print "Passed: FileExtensionTest"
    End If
    
    If Not WriteFileTest() Then
        Debug.Print "Failed: WriteFileTest"
        TestStatus = False
    Else
        Debug.Print "Passed: WriteFileTest"
    End If
    
    If Not ReadFileTest() Then
        Debug.Print "Failed: ReadFileTest"
        TestStatus = False
    Else
        Debug.Print "Passed: ReadFileTest"
    End If
    
    If Not PathSeparatorTest() Then
        Debug.Print "Failed: PathSeparatorTest"
        TestStatus = False
    Else
        Debug.Print "Passed: PathSeparatorTest"
    End If
    
    If Not PathJoinTest() Then
        Debug.Print "Failed: PathJoinTest"
        TestStatus = False
    Else
        Debug.Print "Passed: PathJoinTest"
    End If
    
    If Not CountFilesTest() Then
        Debug.Print "Failed: CountFilesTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CountFilesTest"
    End If
    
    If Not CountFilesAndFoldersTest() Then
        Debug.Print "Failed: CountFilesAndFoldersTest"
        TestStatus = False
    Else
        Debug.Print "Passed: CountFilesAndFoldersTest"
    End If
    
    If Not GetFileNameByNumberTest() Then
        Debug.Print "Failed: GetFileNameByNumberTest"
        TestStatus = False
    Else
        Debug.Print "Passed: GetFileNameByNumberTest"
    End If
    ' End Tests
    
    Debug.Print "----------------------------------------"
    
    If TestStatus Then
        Debug.Print "Passed All Tests"
    Else
        Debug.Print "!!! FAILED TESTS !!!"
    End If
    
    Debug.Print "========================================"
    
    AllXlibFileTests = TestStatus
    
End Function



Private Function GetActivePathAndNameTest() As Boolean

    '@Example: =GetActivePathAndName() -> "C:\Users\UserName\Documents\XLib.xlsm"
    
    GetActivePathAndNameTest = True
    
    If InStr(1, GetActivePathAndName(), ".") > 0 Then
        #If Mac Then
            If InStr(1, GetActivePathAndName(), "/") > 0 Then
                GetActivePathAndNameTest = True
            Else
                GetActivePathAndNameTest = False
            End If
        #Else
            If InStr(1, GetActivePathAndName(), ":\") > 0 Then
                GetActivePathAndNameTest = True
            Else
                GetActivePathAndNameTest = False
            End If
        #End If
    End If

End Function


Private Function GetActivePathTest() As Boolean

    '@Example: =GetActivePath() -> "C:\Users\UserName\Documents\"

    GetActivePathTest = True
        
    #If Mac Then
        If InStr(1, GetActivePath(), "/") > 0 Then
            GetActivePathTest = True
        Else
            GetActivePathTest = False
        End If
    #Else
        If InStr(1, GetActivePath(), ":\") > 0 Then
            GetActivePathTest = True
        Else
            GetActivePathTest = False
        End If
    #End If

End Function


Private Function FileCreationTimeTest() As Boolean

    '@Example: =FileCreationTime() -> "1/1/2020 1:23:45 PM"
    '@Example: =FileCreationTime("C:\hello\world.txt") -> "1/1/2020 5:55:55 PM"
    '@Example: =FileCreationTime("vba.txt") -> "12/25/2000 1:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    FileCreationTimeTest = False
    
    If InStr(1, FileCreationTime(), " ") > 0 Then
        If InStr(1, FileCreationTime(), ":") > 0 Then
            FileCreationTimeTest = True
        End If
    End If

End Function


Private Function FileLastModifiedTimeTest() As Boolean

    '@Example: =FileLastModifiedTime() -> "1/1/2020 2:23:45 PM"
    '@Example: =FileLastModifiedTime("C:\hello\world.txt") -> "1/1/2020 7:55:55 PM"
    '@Example: =FileLastModifiedTime("vba.txt") -> "12/25/2000 3:00:00 PM"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    FileLastModifiedTimeTest = False
    
    If InStr(1, FileLastModifiedTime(), " ") > 0 Then
        If InStr(1, FileLastModifiedTime(), ":") > 0 Then
            FileLastModifiedTimeTest = True
        End If
    End If

End Function


Private Function FileDriveTest() As Boolean

    '@Example: =FileDrive() -> "A:"; Where the current workbook resides on the A: drive
    '@Example: =FileDrive("C:\hello\world.txt") -> "C:"
    '@Example: =FileDrive("vba.txt") -> "B:"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in, and where the workbook resides in the B: drive

    If InStr(1, FileDrive(), ":") > 0 Then
        FileDriveTest = True
    End If

End Function


Private Function FileNameTest() As Boolean

    '@Example: =FileName() -> "MyWorkbook.xlsm"
    '@Example: =FileName("C:\hello\world.txt") -> "world.txt"
    '@Example: =FileName("vba.txt") -> "vba.txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    If Len(FileName()) > 0 Then
        FileNameTest = True
    End If

End Function


Private Function FileFolderTest() As Boolean

    '@Example: =FileFolder() -> "C:\my_excel_files"
    '@Example: =FileFolder("C:\hello\world.txt") -> "C:\hello"
    '@Example: =FileFolder("vba.txt") -> "C:\my_excel_files"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in
    
    #If Mac Then
        If InStr(1, FileFolder(), "/") > 0 Then
            FileFolderTest = True
        End If
    #Else
        If InStr(1, FileFolder(), ":\") > 0 Then
            FileFolderTest = True
        End If
    #End If

End Function


Private Function CurrentFilePathTest() As Boolean

    '@Example: =CurrentFilePath() -> "C:\my_excel_files\MyWorkbook.xlsx"
    '@Example: =CurrentFilePath("C:\hello\world.txt") -> "C:\hello\world.txt"
    '@Example: =CurrentFilePath("vba.txt") -> "C:\hello\world.txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    #If Mac Then
        If InStr(1, CurrentFilePath(), "/") > 0 Then
            CurrentFilePathTest = True
        End If
    #Else
        If InStr(1, CurrentFilePath(), ":\") > 0 Then
            CurrentFilePathTest = True
        End If
    #End If

End Function


Private Function FileSizeTest() As Boolean

    '@Example: =FileSize() -> 1024
    '@Example: =FileSize(,"KB") -> 1
    '@Example: =FileSize("vba.txt", "KB") -> 0.25; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in
    
    If FileSize() > 0 Then
        FileSizeTest = True
    End If

End Function


Private Function FileTypeTest() As Boolean

    '@Example: FileType() -> "Microsoft Excel Macro-Enabled Worksheet"
    '@Example: FileType("C:\hello\world.txt") -> "Text Document"
    '@Example: FileType("vba.txt") -> "Text Document"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in
    
    If Len(FileType()) > 0 Then
        FileTypeTest = True
    End If

End Function


Private Function FileExtensionTest() As Boolean

    '@Example: =FileExtension() = "xlsx"
    '@Example: =FileExtension("C:\hello\world.txt") -> "txt"
    '@Example: =FileExtension("vba.txt") -> "txt"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in

    If Len(FileExtension()) > 0 Then
        FileExtensionTest = True
    End If

End Function


Private Function WriteFileTest() As Boolean

    '@Example: =WriteFile("C:\MyWorkbookFolder\hello.txt", "Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"
    '@Example: =WriteFile("hello.txt", "Hello World") -> "Successfully wrote to: C:\MyWorkbookFolder\hello.txt"; Where the Workbook resides in "C:\MyWorkbookFolder\"
    
    If WriteFile("TempTestFile.txt", "Hello World") Then
        WriteFileTest = True
    End If

End Function


Private Function ReadFileTest() As Boolean

    '@Example: =ReadFile("C:\hello\world.txt") -> "Hello" World
    '@Example: =ReadFile("vba.txt") -> "This is my VBA text file"; Where "vba.txt" resides in the same folder as the workbook, document, presentation, or database this function resides in
    '@Example: =ReadFile("multline.txt", 1) -> "This is line 1";
    '@Example: =ReadFile("multline.txt", 2) -> "This is line 2";

    ReadFileTest = IIf(ReadFile("TempTestFile.txt") = "Hello World", True, False)
    Kill (GetActivePath() & "TempTestFile.txt")

End Function


Private Function PathSeparatorTest() As Boolean

    '@Example: =PathSeparator() -> "\"; When running this code on Windows
    '@Example: =PathSeparator() -> "/"; When running this code on Mac
    
    If PathSeparator() = "\" Or PathSeparator() = "/" Then
        PathSeparatorTest = True
    End If

End Function


Private Function PathJoinTest() As Boolean

    '@Example: =PathJoin("C:", "hello", "world.txt") -> "C:\hello\world.txt"; On Windows
    '@Example: =PathJoin("hello", "world.txt") -> "/hello/world.txt"; On Mac

    If PathJoin("hello", "world") = "hello/world" Or PathJoin("hello", "world") = "hello\world" Then
        PathJoinTest = True
    End If
    
End Function


Private Function CountFilesTest() As Boolean

    '@Example: =CountFiles() -> 6
    '@Example: =CountFiles("C:\hello") -> 10

    If CountFiles() > 0 Then
        CountFilesTest = True
    End If

End Function


Private Function CountFilesAndFoldersTest() As Boolean

    '@Example: =CountFilesAndFolders() -> 8
    '@Example: =CountFilesAndFolders("C:\hello") -> 30
    
    If CountFilesAndFolders() > 0 Then
        CountFilesAndFoldersTest = True
    End If

End Function


Private Function GetFileNameByNumberTest() As Boolean

    '@Example: =GetFileName(,1) -> "hello.txt"
    '@Example: =GetFileName(,1) -> "world.txt"
    '@Example: =GetFileName("C:\hello", 1) -> "one.txt"
    '@Example: =GetFileName("C:\hello", 1) -> "two.txt"
    '@Example: =GetFileName("C:\hello", 1) -> "three.txt"

    If Len(GetFileNameByNumber()) > 0 Then
        GetFileNameByNumberTest = True
    End If

End Function


