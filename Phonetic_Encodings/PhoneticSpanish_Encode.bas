Attribute VB_Name = "PhoneticSpanish_Encode"
Option Explicit
Const strTranscodeIn = "BCDFGHJKLMNPQRSTVXYZ"
Const strTranscodeOut = "14328287566079431454"

Function PhoneticSpanish(StrWord As String, Optional intMaxLength As Integer = -1) As String
    StrWord = UnicodeFunctions.UnicodeStrip(UCase(StrWord))
    Dim strWordUCSet As String
    Dim i As Long
    For i = 1 To Len(StrWord)
        If Mid(StrWord, i, 1) Like "[BCDFGHJKLMNPQRSTVXYZ]" Then
            strWordUCSet = strWordUCSet & Mid(StrWord, i, 1)
        End If
    Next
    
    StrWord = strWordUCSet
    StrWord = Replace(StrWord, "LL", "L")
    StrWord = Replace(StrWord, "RR", "R") 'The Abydos origin replaces "R" with "R", Based on the comment preceding that line, this is an oversight
    For i = 1 To Len(StrWord)
        Mid(StrWord, i, 1) = Mid(strTranscodeOut, InStr(strTranscodeIn, Mid(StrWord, i, 1)), 1)
    Next
    
    If intMaxLength > 0 Then
        StrWord = Left(StrWord & String(intMaxLength, "0"), intMaxLength)
    End If
    
    PhoneticSpanish = StrWord
End Function
