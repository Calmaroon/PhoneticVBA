Attribute VB_Name = "Encode_PhoneticSpanish"
Option Explicit
Const strTranscodeIn = "BCDFGHJKLMNPQRSTVXYZ"
Const strTranscodeOut = "14328287566079431454"

Function PhoneticSpanish(strWord As String, Optional intMaxLength As Integer = -1) As String
    strWord = UnicodeStrip(UCase(strWord))
    Dim strWordUCSet As String
    Dim i As Long
    For i = 1 To Len(strWord)
        If Mid$(strWord, i, 1) Like "[BCDFGHJKLMNPQRSTVXYZ]" Then
            strWordUCSet = strWordUCSet & Mid$(strWord, i, 1)
        End If
    Next
    
    strWord = strWordUCSet
    strWord = Replace(strWord, "LL", "L")
    strWord = Replace(strWord, "RR", "R") 'The Abydos origin replaces "R" with "R", Based on the comment preceding that line, this is an oversight
    For i = 1 To Len(strWord)
        Mid(strWord, i, 1) = Mid(strTranscodeOut, InStr(strTranscodeIn, Mid(strWord, i, 1)), 1)
    Next
    
    If intMaxLength > 0 Then
        strWord = Left(strWord & String(intMaxLength, "0"), intMaxLength)
    End If
    
    PhoneticSpanish = strWord
End Function
