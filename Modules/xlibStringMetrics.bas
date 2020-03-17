Attribute VB_Name = "xlibStringMetrics"
'@Module: This module contains a set of functions for performing fuzzy string matches. It can be useful when you have 2 columns containing text that is close but not 100% the same. However, since the functions in this module only perform fuzzy matches, there is no guarantee that there will be 100% accuracy in the matches. However, for small groups of string where each string is very different than the other (such as a small group of fairly dissimilar names), these functions can be highly accurate. Finally, some of the functions in this Module will take a long time to calculate for large numbers of cells, as the number of calculations for some functions will grow exponentially, but for small sets of data (such as 100 strings to compare), these functions perform fairly quickly.

Option Explicit


'========================================
'  Hamming Distance
'========================================

Public Function Hamming( _
    string1 As String, _
    string2 As String) _
As Integer

    '@Description: This function takes two strings of the same length and calculates the Hamming Distance between them. Hamming Distance measures how close two strings are by checking how many Substitutions are needed to turn one string into the other. Lower numbers mean the strings are closer than high numbers.
    '@Author: Anthony Mancini
    '@Version: 1.0.0
    '@License: MIT
    '@Param: string1 is the first string
    '@Param: string2 is the second string that will be compared to the first string
    '@Returns: Returns an integer of the Hamming Distance between two string
    '@Example: =Hamming("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
    '@Example: =Hamming("Cat", "Bag") -> 2; 2 changes are needed, changing the "B" to "C" and the "g" to "t"
    '@Example: =Hamming("Cat", "Dog") -> 3; Every single character needs to be substituted in this case

    If Len(string1) <> Len(string2) Then
        Hamming = CVErr(2015)
    End If
    
    Dim totalDistance As Integer
    totalDistance = 0
    
    Dim i As Integer
    
    For i = 1 To Len(string1)
        If Mid(string1, i, 1) <> Mid(string2, i, 1) Then
            totalDistance = totalDistance + 1
        End If
    Next
    
    Hamming = totalDistance
    
End Function



'========================================
'  Levenshtein Distance
'========================================

Public Function Levenshtein( _
    string1 As String, _
    string2 As String) _
As Long

    '@Description: This function takes two strings of any length and calculates the Levenshtein Distance between them. Levenshtein Distance measures how close two strings are by checking how many Insertions, Deletions, or Substitutions are needed to turn one string into the other. Lower numbers mean the strings are closer than high numbers. Unlike Hamming Distance, Levenshtein Distance works for strings of any length and includes 2 more operations. However, calculation time will be slower than Hamming Distance for same length strings, so if you know the two strings are the same length, its preferred to use Hamming Distance.
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Param: string1 is the first string
    '@Param: string2 is the second string that will be compared to the first string
    '@Returns: Returns an integer of the Levenshtein Distance between two string
    '@Example: =Levenshtein("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
    '@Example: =Levenshtein("Cat", "Ca") -> 1; Since only one Insertion needs to occur (adding a "t" at the end)
    '@Example: =Levenshtein("Cat", "Cta") -> 2; Since the "t" in "Cta" needs to be substituted into an "a", and the final character "a" needs to be substituted into a "t"

    ' **Error Checking**
    ' Quick returns for common errors
    If string1 = string2 Then
        Levenshtein = 0
        Exit Function
    ElseIf string1 = Empty Then
        Levenshtein = Len(string2)
        Exit Function
    ElseIf string2 = Empty Then
        Levenshtein = Len(string1)
        Exit Function
    End If
    

    ' **Algorithm Code**
    ' Creating the distance metrix and filling it with values
    Dim numberOfRows As Integer
    Dim numberOfColumns As Integer
    
    numberOfRows = Len(string1)
    numberOfColumns = Len(string2)
    
    Dim distanceArray() As Integer
    ReDim distanceArray(numberOfRows, numberOfColumns)
    
    Dim r As Integer
    Dim c As Integer
    
    For r = 0 To numberOfRows
        For c = 0 To numberOfColumns
            distanceArray(r, c) = 0
        Next
    Next
    
    For r = 1 To numberOfRows
        distanceArray(r, 0) = r
    Next
    
    For c = 1 To numberOfColumns
        distanceArray(0, c) = c
    Next
    
    ' Non-recursive Levenshtein Distance matrix walk
    Dim operationCost As Integer
    
    For c = 1 To numberOfColumns
        For r = 1 To numberOfRows
            If Mid(string1, r, 1) = Mid(string2, c, 1) Then
                operationCost = 0
            Else
                operationCost = 1
            End If
                                                           
            distanceArray(r, c) = Min(distanceArray(r - 1, c) + 1, distanceArray(r, c - 1) + 1, distanceArray(r - 1, c - 1) + operationCost)
        Next
    Next
    
    Levenshtein = distanceArray(numberOfRows, numberOfColumns)

End Function



'========================================
'  Damerau-Levenshtein Distance
'========================================

Public Function Damerau( _
    string1 As String, _
    string2 As String) _
As Integer

    '@Description: This function takes two strings of any length and calculates the Damerau-Levenshtein Distance between them. Damerau-Levenshtein Distance differs from Levenshtein Distance in that it includes an additional operation, called Transpositions, which occurs when two adjacent characters are swapped. Thus, Damerau-Levenshtein Distance calculates the number of Insertions, Deletions, Substitutions, and Transpositons needed to convert string1 into string2. As a result, this function is good when it is likely that spelling errors have occured between two string where the error is simply a transposition of 2 adjacent characters.
    '@Author: Anthony Mancini
    '@Version: 1.1.0
    '@License: MIT
    '@Param: string1 is the first string
    '@Param: string2 is the second string that will be compared to the first string
    '@Returns: Returns an integer of the Damerau-Levenshtein Distance between two string
    '@Example: =Damerau("Cat", "Bat") -> 1; Since all that is needed is 1 change (changing the "B" in Bat to "C")
    '@Example: =Damerau("Cat", "Ca") -> 1; Since only one Insertion needs to occur (adding a "t" at the end)
    '@Example: =Damerau("Cat", "Cta") -> 1; Since the "t" and "a" can be transposed as they are adjacent to each other. Notice how LEVENSHTEIN("Cat","Cta")=2 but DAMERAU("Cat","Cta")=1

    ' **Error Checking**
    ' Quick returns for common errors
    If string1 = string2 Then
        Damerau = 0
    ElseIf string1 = Empty Then
        Damerau = Len(string2)
    ElseIf string2 = Empty Then
        Damerau = Len(string1)
    End If
    
    Dim inf As Long
    Dim da As Object
    inf = Len(string1) + Len(string2)
    Set da = CreateObject("Scripting.Dictionary")
    
    ' 35 - 38 = filling the dictionary
    Dim i As Integer
    For i = 1 To Len(string1)
        If da.exists(Mid(string1, i, 1)) = False Then
            da.Add Mid(string1, i, 1), "0"
        End If
    Next
    
    For i = 1 To Len(string2)
        If da.exists(Mid(string2, i, 1)) = False Then
            da.Add Mid(string2, i, 1), "0"
        End If
    Next
    
    ' 39 = creating h matrix
    Dim H() As Long
    ReDim H(Len(string1) + 1, Len(string2) + 1)
    
    Dim k As Integer
    For i = 0 To (Len(string1) + 1)
        For k = 0 To (Len(string2) + 1)
            H(i, k) = 0
        Next
    Next
    
    ' 40 - 45 = updating the matrix
    For i = 0 To Len(string1)
        H(i + 1, 0) = inf
        H(i + 1, 1) = i
    Next
    For k = 0 To Len(string2)
        H(0, k + 1) = inf
        H(1, k + 1) = k
    Next
    

    ' 46 - 60 = running the array
    Dim db As Long
    Dim i1 As Long
    Dim k1 As Long
    Dim cost As Long
    
    For i = 1 To Len(string1)
        db = 0
        For k = 1 To Len(string2)
            i1 = CInt(da(Mid(string2, k, 1)))
            k1 = db
            cost = 1
            
            If Mid(string1, i, 1) = Mid(string2, k, 1) Then
                cost = 0
                db = k
            End If
            
            H(i + 1, k + 1) = Min(H(i, k) + cost, _
                                  H(i + 1, k) + 1, _
                                  H(i, k + 1) + 1, _
                                  H(i1, k1) + (i - i1 - 1) + 1 + (k - k1 - 1))
                           
            
        Next
        
        If da.exists(Mid(string1, i, 1)) Then
            da.Remove Mid(string1, i, 1)
            da.Add Mid(string1, i, 1), CStr(i)
        Else
            da.Add Mid(string1, i, 1), CStr(i)
        End If
        
    Next

    Damerau = H(Len(string1) + 1, Len(string2) + 1)

End Function

