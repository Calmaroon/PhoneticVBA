Attribute VB_Name = "_PhoneticFunctions"
Option Explicit
Function DeleteConsecutiveRepeats(ByVal strWord As String) As String
    If Len(strWord) = 0 Then
        DeleteConsecutiveRepeats = vbNullString
        Exit Function
    End If
    
    Dim i As Long
    Dim strResult As String
    
    strResult = Mid$(strWord, 1, 1)
    
    Dim strChar As String
    Dim strPrev As String
    strPrev = Left$(strWord, 1)
    For i = 2 To Len(strWord)
        strChar = Mid$(strWord, i, 1)
        If strChar <> strPrev Then
            strResult = strResult & strChar
        End If
        strPrev = strChar
    Next i

    DeleteConsecutiveRepeats = strResult
End Function
Function GetAlphaOnly(strInput As String) As String
    Dim strResult As String, i As Long
    For i = 1 To Len(strInput)
        If Mid$(strInput, i, 1) Like "[A-Za-z]" Then
            strResult = strResult & Mid$(strInput, i, 1)
        End If
    Next
    GetAlphaOnly = strResult
End Function
