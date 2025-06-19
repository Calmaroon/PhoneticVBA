Attribute VB_Name = "PhoneticFunctions"
Option Explicit
Function DeleteConsecutiveRepeats(ByVal strWord As String) As String
    Dim i As Long
    Dim strResult As String

    If Len(strWord) = 0 Then
        DeleteConsecutiveRepeats = ""
        Exit Function
    End If

    strResult = Mid$(strWord, 1, 1)
    
    For i = 2 To Len(strWord)
        If Mid$(strWord, i, 1) <> Mid$(strWord, i - 1, 1) Then
            strResult = strResult & Mid$(strWord, i, 1)
        End If
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
