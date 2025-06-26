Attribute VB_Name = "Encode_Davidson"
Option Explicit
Function Davidson(strLName As String, Optional strFName As String = ".", Optional boolOmitFName As Boolean = False)
    Dim StrEncoding As String
    Dim i As Long
    strLName = UCase$(strLName)
    StrEncoding = Left$(strLName, 1)
    For i = 2 To Len(strLName)
        If Not Mid$(strLName, i, 1) Like "[AEIOUWHY]" Then
            StrEncoding = StrEncoding & Mid$(strLName, i, 1)
        End If
    Next

    StrEncoding = DeleteConsecutiveRepeats(StrEncoding)
    
    If Len(StrEncoding) < 4 Then
        StrEncoding = StrEncoding & Space(4 - Len(StrEncoding))
    Else
        StrEncoding = Left$(StrEncoding, 4)
    End If
    
    If Not boolOmitFName Then
        StrEncoding = StrEncoding & Left$(strFName, 1)
    End If
    
    Davidson = StrEncoding
End Function
