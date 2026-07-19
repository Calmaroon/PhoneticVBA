Attribute VB_Name = "Encode_SoundexBR"
Option Explicit
Const strTransIn As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const strTransOut As String = "01230120022455012623010202"
Function SoundexBR(strInput As String, Optional intMaxLength As Integer = 4, Optional boolZeroPad As Boolean = True) As String
    strInput = UCase$(strInput)
    strInput = GetAlphaOnly(strInput)
    Dim strSoundex As String
    If Left$(strInput, 2) = "WA" Then
        strSoundex = "V"
    ElseIf Left$(strInput, 1) = "K" And Mid$(strInput, 2, 1) Like "[AOU]" Then
        strSoundex = "C"
    ElseIf Left$(strInput, 1) = "C" And Mid$(strInput, 2, 1) Like "[EI]" Then
        strSoundex = "S"
    ElseIf Left$(strInput, 1) = "G" And Mid$(strInput, 2, 1) Like "[EI]" Then
        strSoundex = "H"
    ElseIf Left$(strInput, 1) = "Y" Then
        strSoundex = "I"
    ElseIf Left$(strInput, 1) = "H" Then
        strSoundex = Mid$(strInput, 2, 1)
        strInput = Mid$(strInput, 2)
    Else
        strSoundex = Left$(strInput, 1)
    End If
    
    Dim i As Long
    For i = 2 To Len(strInput)
        strSoundex = strSoundex & Mid$(strTransOut, InStr(strTransIn, Mid$(strInput, i, 1)), 1)
    Next
    
    strSoundex = DeleteConsecutiveRepeats(strSoundex)
    strSoundex = Replace$(strSoundex, "0", "")
    
    If boolZeroPad Then
        strSoundex = strSoundex & String(intMaxLength, "0")
    End If
    
    SoundexBR = Left$(strSoundex, intMaxLength)
End Function
