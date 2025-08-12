Attribute VB_Name = "Encode_Soundex"
Option Explicit
Const SoundexCodes = "01230129022455012623019202"
Public Enum SoundexVariant
    vAmerican = 1
    vSpecial = 2
    vCensus = 3
End Enum
Function Soundex(strWord As String, Optional vSoundexVariant As SoundexVariant = vAmerican, Optional intMaxLength As Integer = 4, Optional boolReverse As Boolean = False, Optional boolZeroPad As Boolean = True) As String
    Dim strEncoding As String, strChar As String, strPrevCode As String, strCurrCode As String
    Dim i As Long, intWordLength As Long, intOutPos As Long
    
    If intMaxLength < 0 Then intMaxLength = 4
    If intMaxLength > 64 Then intMaxLength = 64

    intWordLength = Len(strWord)
    If intWordLength = 0 Then Soundex = "": Exit Function
    
    strWord = UCase$(strWord)
    
    strWord = GetAlphaOnly(strWord)
    If strWord = "" Then
        Soundex = ""
        Exit Function
    End If
    
    If vSoundexVariant = vCensus Then
        If Left$(strWord, 3) = "VAN" Or Left$(strWord, 3) = "CON" And Len(strWord) > 4 Then
            strWord = Mid$(strWord, 4, Len(strWord))
        ElseIf Left$(strWord, 2) = "DE" Or Left$(strWord, 2) = "DI" Or Left$(strWord, 2) = "LA" Or Left$(strWord, 2) = "LE" Then
            strWord = Mid$(strWord, 3, Len(strWord))
        End If
    End If

    If boolReverse Then strWord = StrReverse(strWord)
    
    For i = 1 To intWordLength
        strChar = Mid$(strWord, i, 1)
        If (Asc(strChar) >= 65 And Asc(strChar) <= 90) Then
            strEncoding = strChar
            strPrevCode = Mid$(SoundexCodes, Asc(strEncoding) - 64, 1)
            Exit For
        ElseIf AscW(strChar) >= 192 And AscW(strChar) <= 383 Then 'Check if the first char has an accent
            strEncoding = Left$(UnicodeStrip(strChar), 1)
            strPrevCode = Mid$(SoundexCodes, AscW(strEncoding) - 64, 1)
            Exit For
        End If
    Next
    
    intOutPos = 2
    Dim intAsc As Long
    For i = 2 To Len(strWord)
        strChar = Mid$(strWord, i, 1)
        
        strCurrCode = Mid$(SoundexCodes, Asc(strChar) - 64, 1)
        If (Asc(strChar) >= 65 And Asc(strChar) <= 90) Or (AscW(strChar) >= 192 And AscW(strChar) <= 383) Then
            If AscW(strChar) >= 192 And AscW(strChar) <= 383 Then strChar = Left$(UnicodeStrip(strChar), 1)
            strCurrCode = Mid$(SoundexCodes, Asc(strChar) - 64, 1)
            
            If strCurrCode <> "9" And strCurrCode <> "0" And strCurrCode <> strPrevCode Then
                strEncoding = strEncoding & strCurrCode
                intOutPos = intOutPos + 1
                If intOutPos > intMaxLength Then Exit For
            End If
 
            If strCurrCode <> "9" Then strPrevCode = strCurrCode
        End If
    Next
    
    If vSoundexVariant = vSpecial Then
        strEncoding = Replace$(strEncoding, "9", "0")
    Else
        strEncoding = Replace$(strEncoding, "9", "")
    End If
    
    If boolZeroPad And Len(strEncoding) < intMaxLength Then strEncoding = strEncoding & String$(intMaxLength - Len(strEncoding), "0")
    
    Soundex = strEncoding
End Function

