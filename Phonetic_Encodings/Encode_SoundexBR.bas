Attribute VB_Name = "Encode_SoundexBR"
Option Explicit
Const strTransIn As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const strTransOut As String = "01230120022455012623010202"
Function SoundexBR(strInput As String, Optional intMaxLength As Integer = 4, Optional boolZeroPad As Boolean = True) As String
    strInput = UCase$(strInput)
    Dim i As Long

    For i = 1 To Len(strInput)
        If AscW(Mid$(strInput, i, 1)) >= 192 And AscW(Mid$(strInput, i, 1)) <= 383 Then 'Check if the first char has an accent
            strInput = UnicodeStrip(strInput)
            Exit For
        End If
    Next

    Dim strSoundex As String
    If left$(strInput, 2) = "WA" Then
        strSoundex = "V"
    ElseIf left$(strInput, 1) = "K" And Mid$(strInput, 2, 1) Like "[AOU]" Then
        strSoundex = "C"
    ElseIf left$(strInput, 1) = "C" And Mid$(strInput, 2, 1) Like "[EI]" Then
        strSoundex = "S"
    ElseIf left$(strInput, 1) = "G" And Mid$(strInput, 2, 1) Like "[EI]" Then
        strSoundex = "H"
    ElseIf left$(strInput, 1) = "Y" Then
        strSoundex = "I"
    ElseIf left$(strInput, 1) = "H" Then
        strSoundex = Mid$(strInput, 2, 1)
        strInput = Mid$(strInput, 1)
    Else
        strSoundex = left$(strInput, 1)
    End If
    
    For i = 2 To Len(strInput)
        strSoundex = strSoundex & Mid$(strTransOut, InStr(strTransIn, Mid$(strInput, i, 1)), 1)
    Next
    
    strSoundex = DeleteConsecutiveRepeats(strSoundex)
    strSoundex = Replace$(strSoundex, "0", "")
    
    If boolZeroPad Then strSoundex = strSoundex & String(intMaxLength, "0")
    
    SoundexBR = left$(strSoundex, intMaxLength)
End Function
