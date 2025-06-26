Attribute VB_Name = "Encode_Lein"
Option Explicit
Const strTranscodeIn = "BCDFGJKLMNPQRSTVXZ"
Const strTranscodeOut = "451455532245351455"
Function Lein(strWord As String, Optional intMaxLength As Integer = 4, Optional boolZeroPad As Boolean = True) As String
    strWord = UCase$(UnicodeStrip(strWord))
    Dim StrEncoding As String
    Dim i As Long
    
    StrEncoding = Left(strWord, 1)
    strWord = Mid(strWord, 2, Len(strWord))
    
    Dim strWordExcluded As String
    For i = 1 To Len(strWord)
        If InStr(strTranscodeIn, Mid$(strWord, i, 1)) > 0 Then
            strWordExcluded = strWordExcluded & Mid$(strWord, i, 1)
        End If
    Next
    strWordExcluded = DeleteConsecutiveRepeats(strWordExcluded)
    For i = 1 To Len(strWordExcluded)
        Mid$(strWordExcluded, i, 1) = Mid$(strTranscodeOut, InStr(strTranscodeIn, Mid$(strWordExcluded, i, 1)), 1)
    Next
    
    StrEncoding = StrEncoding & strWordExcluded
    
    If boolZeroPad Then
        StrEncoding = Left(StrEncoding & String(intMaxLength, "0"), intMaxLength)
    End If
    Lein = StrEncoding
End Function
