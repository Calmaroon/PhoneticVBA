Attribute VB_Name = "MRA_Encode"
Option Explicit
Function MRA(strWord As String) As String
    Dim strEncoding As String
    Dim i As Long
    strWord = UCase(strWord)
    
    strEncoding = Left(strWord, 1)
    For i = 2 To Len(strWord)
        If Not Mid(strWord, i, 1) Like "[AEIOU]" Then
            strEncoding = strEncoding & Mid(strWord, i, 1)
        End If
    Next
    
    strEncoding = PhoneticFunctions.DeleteConsecutiveRepeats(strEncoding)
    
    If Len(strEncoding) > 6 Then
        strEncoding = Left(strEncoding, 3) & Right(strEncoding, 3)
    End If
    MRA = strEncoding
End Function
