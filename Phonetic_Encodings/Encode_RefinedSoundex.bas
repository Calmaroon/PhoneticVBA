Attribute VB_Name = "Encode_RefinedSoundex"
Option Explicit
Const strTranscodeIn = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const strTranscodeOut = "01360240043788015936020505"
Function RefinedSoundex(strWord As String, Optional intMaxLength As Integer = -1, Optional boolZeroPad As Boolean = False, Optional boolRetainVowels As Boolean = False) As String
    strWord = UnicodeStrip(strWord)
    strWord = GetAlphaOnly(UCase$(strWord))
    
    Dim strSoundex As String
    strSoundex = Left(strWord, 1)
    
    Dim i As Long
    For i = 2 To Len(strWord)
        Mid(strWord, i, 1) = Mid(strTranscodeOut, InStr(strTranscodeIn, Mid(strWord, i, 1)), 1)
    Next
    strSoundex = strSoundex & Mid(strWord, 2)
    strSoundex = DeleteConsecutiveRepeats(strSoundex)
    
    If Not boolRetainVowels Then
        strSoundex = Replace(strSoundex, "0", "")
    End If
    
    If intMaxLength > 0 Then
        If boolZeroPad Then
            strSoundex = strSoundex & String(intMaxLength, "0")
        End If
        strSoundex = Left(strSoundex, intMaxLength)
    End If
    RefinedSoundex = strSoundex
End Function
