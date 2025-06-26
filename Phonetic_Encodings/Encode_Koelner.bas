Attribute VB_Name = "Encode_Koelner"
Option Explicit
Const strVSet = "AEIOUJY"
Function Koelner(strWord As String) As String
    Dim StrEncoding As String
    

    strWord = UCase$(strWord)
    strWord = Replace(strWord, "Ä", "AE")
    strWord = Replace(strWord, "Ö", "OE")
    strWord = Replace(strWord, "Ü", "UE")
    
    strWord = UnicodeStrip(strWord)
    
    If strWord = "" Then
        strWord = ""
        Exit Function
    End If
    
    Dim i As Long
    Dim strChar As String
    For i = 1 To Len(strWord)
        strChar = Mid(strWord, i, 1)
        If InStr(strVSet, strChar) > 0 Then
            StrEncoding = StrEncoding & "0"
        ElseIf strChar = "B" Then
            StrEncoding = StrEncoding & "1"
        ElseIf strChar = "P" Then
            If i + 1 < Len(strWord) And Mid(strWord, i + 1, 1) = "H" Then
                StrEncoding = StrEncoding & "3"
            Else
                StrEncoding = StrEncoding & "3"
            End If
        ElseIf strChar Like "[DT]" Then
            If i < Len(strWord) And Mid(strWord, i + 1, 1) Like "[CSZ]" Then
                StrEncoding = StrEncoding & "8"
            Else
                StrEncoding = StrEncoding & "2"
            End If
        ElseIf strChar Like "[FVW]" Then
              StrEncoding = StrEncoding & "3"
        ElseIf strChar Like "[GKQ]" Then
            StrEncoding = StrEncoding & "4"
        ElseIf strChar = "C" Then
            If i > 1 Then
                If Mid(strWord, i - 1, 1) Like "[SZ]" Then
                    StrEncoding = StrEncoding & "8"
                ElseIf i < Len(strWord) And Mid(strWord, i + 1, 1) Like "[AHKOQUX]" Then
                    StrEncoding = StrEncoding & "4"
                Else
                    StrEncoding = StrEncoding & "8"
                End If
            ElseIf i = 1 Then
                If i < Len(strWord) And Mid(strWord, i + 1, 1) Like "[AHKLOQRUX]" Then
                    StrEncoding = StrEncoding & "4"
                Else
                    StrEncoding = StrEncoding & "8"
                End If
            End If
        ElseIf strChar = "X" Then
            If i > 1 Then
                If Mid(strWord, i - 1, 1) Like "[CKQ]" Then
                    StrEncoding = StrEncoding & "8"
                Else
                    StrEncoding = StrEncoding & "48"
                End If
            Else
                StrEncoding = StrEncoding & "48"
            End If
        ElseIf strChar = "L" Then
            StrEncoding = StrEncoding & "5"
        ElseIf strChar Like "[MN]" Then
            StrEncoding = StrEncoding & "6"
        ElseIf strChar = "R" Then
            StrEncoding = StrEncoding & "7"
        ElseIf strChar Like "[SZ]" Then
            StrEncoding = StrEncoding & "8"
        End If
    Next
    
    StrEncoding = DeleteConsecutiveRepeats(StrEncoding)
    StrEncoding = Left(StrEncoding, 1) & Replace(Mid(StrEncoding, 2, Len(StrEncoding)), "0", "")
    
    Koelner = StrEncoding
End Function
