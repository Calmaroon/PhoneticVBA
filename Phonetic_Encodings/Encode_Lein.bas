Attribute VB_Name = "Encode_Lein"
Option Explicit
Const strTranscodeIn = "BCDFGJKLMNPQRSTVXZ"
Const strTranscodeOut = "451455532245351455"
Function Lein(strWord As String, Optional intMaxLength As Integer = 4, Optional boolZeroPad As Boolean = True) As String
    strWord = UCase$(UnicodeStrip(strWord))
    
    Dim strEncoding As String
    strEncoding = left$(strWord, 1)
    strWord = Mid$(strWord, 2)
    
    Dim strWordExcluded As String: Dim i As Long
    For i = 1 To Len(strWord)
        If InStr(strTranscodeIn, Mid$(strWord, i, 1)) > 0 Then strWordExcluded = strWordExcluded & Mid$(strWord, i, 1)
    Next
    
    strWordExcluded = DeleteConsecutiveRepeats(strWordExcluded)
    For i = 1 To Len(strWordExcluded)
        Mid$(strWordExcluded, i, 1) = Mid$(strTranscodeOut, InStr(strTranscodeIn, Mid$(strWordExcluded, i, 1)), 1)
    Next
    
    strEncoding = strEncoding & strWordExcluded
    
    If boolZeroPad Then strEncoding = left$(strEncoding & String(intMaxLength, "0"), intMaxLength)
    
    Lein = strEncoding
End Function
