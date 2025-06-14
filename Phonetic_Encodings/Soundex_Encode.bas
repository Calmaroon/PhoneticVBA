Attribute VB_Name = "Soundex_Encode"
Const strTranscodeIn = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const strTranscodeOut = "01230129022455012623019202"
                         
Const strAlphaIn = "0123456789"
Const strAlphaOut = "APKTLNRH"

Public Enum SoundexVariant
    vAmerican = 1
    vSpecial = 2
    vCensus = 3
End Enum
Function Soundex(strWord As String, Optional vSoundexVariant As SoundexVariant = vAmerican, Optional intMaxLength As Integer = 4, Optional boolReverse As Boolean = False, Optional boolZeroPad As Boolean = True) As String
    Dim strEncoding As String
    If intMaxLength < 0 Then
        intMaxLength = 4
    Else
        If intMaxLength > 64 Then intMaxLength = 64
    End If
    
    'Remove markings from unicode characters
    strWord = UnicodeFunctions.UnicodeStrip(strWord)
    
    strWord = UCase(strWord)
    
    If vSoundexVariant = vCensus Then
        If Left$(strWord, 3) = "VAN" Or Left$(strWord, 3) = "CON" And Len(strWord) > 4 Then
            strWord = Mid$(strWord, 4, Len(strWord))
        ElseIf Left$(strWord, 2) = "DE" Or Left$(strWord, 2) = "DI" Or Left$(strWord, 2) = "LA" Or Left$(strWord, 2) = "LE" Then
            strWord = Mid$(strWord, 3, Len(strWord))
        End If
    End If
    
    Dim i As Long
    For i = 1 To Len(strWord)
        If InStr(strTranscodeIn, Mid$(strWord, i, 1)) > -1 Then
            strEncoding = strEncoding & Mid$(strWord, i, 1)
        End If
    Next
    
    If Len(strEncoding) = 0 Then
        Soundex = "0"
        Exit Function
    End If
    
    If boolReverse Then
        strEncoding = StrReverse(strEncoding)
    End If
    
    
    For i = 1 To Len(strEncoding)
        Mid$(strEncoding, i, 1) = Mid$(strTranscodeOut, InStr(strTranscodeIn, Mid$(strEncoding, i, 1)), 1)
    Next
    
    If vSoundexVariant = vSpecial Then
        strEncoding = Replace(strEncoding, "9", "0")
    Else
        strEncoding = Replace(strEncoding, "9", "")
    End If
    
    strEncoding = PhoneticFunctions.DeleteConsecutiveRepeats(strEncoding)

    If Left$(strWord, 1) = "H" Or Left$(strWord, 1) = "W" Then
        strEncoding = Left$(strWord, 1) & strEncoding
    Else
        strEncoding = Left$(strWord, 1) & Mid$(strEncoding, 2, Len(strEncoding))
    End If
    
    strEncoding = Replace(strEncoding, "0", "")
    
    If boolZeroPad Then
        strEncoding = strEncoding & String(intMaxLength, "0")
    End If
    
    strEncoding = Left$(strEncoding, intMaxLength)
    Soundex = strEncoding
End Function
