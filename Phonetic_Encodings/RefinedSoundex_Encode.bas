Attribute VB_Name = "RefinedSoundex_Encode"
Option Explicit
Const strTranscodeIn = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const strTranscodeOut = "01360240043788015936020505"
Function RefinedSoundex(StrWord As String, Optional intMaxLength As Integer = -1, Optional boolZeroPad As Boolean = False, Optional boolRetainVowels As Boolean = False) As String
    StrWord = UnicodeFunctions.UnicodeStrip(StrWord)
    StrWord = PhoneticFunctions.GetAlphaOnly(UCase$(StrWord))
    
    Dim strSoundex As String
    strSoundex = Left(StrWord, 1)
    
    Dim i As Long
    For i = 2 To Len(StrWord)
        Mid(StrWord, i, 1) = Mid(strTranscodeOut, InStr(strTranscodeIn, Mid(StrWord, i, 1)), 1)
    Next
    strSoundex = strSoundex & Mid(StrWord, 2)
    strSoundex = PhoneticFunctions.DeleteConsecutiveRepeats(strSoundex)
    
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
