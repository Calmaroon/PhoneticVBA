Attribute VB_Name = "Encode_MRA"
Option Explicit
Function MRA(strWord As String) As String
    'does not strip out numerics/accents/etc.
    Dim StrEncoding As String
    Dim i As Long
    strWord = UCase(strWord)
    
    StrEncoding = Left(strWord, 1)
    For i = 2 To Len(strWord)
        If Not Mid(strWord, i, 1) Like "[AEIOU]" Then
            StrEncoding = StrEncoding & Mid(strWord, i, 1)
        End If
    Next
    
    StrEncoding = DeleteConsecutiveRepeats(StrEncoding)
    
    If Len(StrEncoding) > 6 Then
        StrEncoding = Left(StrEncoding, 3) & Right(StrEncoding, 3)
    End If
    MRA = StrEncoding
End Function
