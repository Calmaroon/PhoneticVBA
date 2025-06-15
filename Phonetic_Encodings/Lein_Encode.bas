Attribute VB_Name = "Lein_Encode"
Option Explicit
Const strTranscodeIn = "BCDFGJKLMNPQRSTVXZ"
Const strTranscodeOut = "451455532245351455"
Function Lein(strWord As String, Optional intMaxLength As Integer = 4, Optional boolZeroPad As Boolean = True) As String
    strWord = UCase$(UnicodeFunctions.UnicodeStrip(strWord))
    Dim strEncoding As String
    Dim i As Long
    
    strEncoding = Left(strWord, 1)
    strWord = Mid(strWord, 2, Len(strWord))
    
    Dim strWordExcluded As String
    For i = 1 To Len(strWord)
        If InStr(strTranscodeIn, Mid$(strWord, i, 1)) > 0 Then
            strWordExcluded = strWordExcluded & Mid$(strWord, i, 1)
        End If
    Next
    strWordExcluded = PhoneticFunctions.DeleteConsecutiveRepeats(strWordExcluded)
    For i = 1 To Len(strWordExcluded)
        Mid$(strWordExcluded, i, 1) = Mid$(strTranscodeOut, InStr(strTranscodeIn, Mid$(strWordExcluded, i, 1)), 1)
    Next
    
    strEncoding = strEncoding & strWordExcluded
    
    If boolZeroPad Then
        strEncoding = Left(strEncoding & String(intMaxLength, "0"), intMaxLength)
    End If
    Lein = strEncoding
End Function
