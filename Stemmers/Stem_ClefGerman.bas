Attribute VB_Name = "Stem_ClefGerman"
Option Explicit
Function ClefGerman(strWord As String) As String
    strWord = LCase(strWord)
    If InStr(strWord, "ä") > 0 Then strWord = Replace(strWord, "ä", "a")
    If InStr(strWord, "ö") > 0 Then strWord = Replace(strWord, "ö", "o")
    If InStr(strWord, "ü") > 0 Then strWord = Replace(strWord, "ü", "u")
    
    Dim intLen As Integer
    intLen = Len(strWord) - 1
    
    If intLen > 3 Then
        If intLen > 5 Then
            If right(strWord, 3) = "nen" Then
                ClefGerman = left(strWord, Len(strWord) - 3)
                Exit Function
            End If
        End If
        
        If intLen > 4 Then
            If right(strWord, 2) = "en" Or right(strWord, 2) = "se" Or right(strWord, 2) = "es" Or right(strWord, 2) = "er" Then
                ClefGerman = left(strWord, Len(strWord) - 2)
                Exit Function
            End If
        End If
        
        If right(strWord, 1) Like "[enrs]" Then
            ClefGerman = left(strWord, Len(strWord) - 1)
            Exit Function
        End If
    End If
    
    ClefGerman = strWord
End Function
