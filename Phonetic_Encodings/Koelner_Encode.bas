Attribute VB_Name = "Koelner_Encode"
Option Explicit
Const strVSet = "AEIOUJY"
Function Koelner(strWord As String) As String
    Dim strEncoding As String
    

    strWord = UCase$(strWord)
    strWord = Replace(strWord, "Ä", "AE")
    strWord = Replace(strWord, "Ö", "OE")
    strWord = Replace(strWord, "Ü", "UE")
    
    strWord = UnicodeFunctions.UnicodeStrip(strWord)
    
    If strWord = "" Then
        strWord = ""
        Exit Function
    End If
    
    Dim i As Long
    Dim strChar As String
    For i = 1 To Len(strWord)
        strChar = Mid(strWord, i, 1)
        If InStr(strVSet, strChar) > 0 Then
            strEncoding = strEncoding & "0"
        ElseIf strChar = "B" Then
            strEncoding = strEncoding & "1"
        ElseIf strChar = "P" Then
            If i + 1 < Len(strWord) And Mid(strWord, i + 1, 1) = "H" Then
                strEncoding = strEncoding & "3"
            Else
                strEncoding = strEncoding & "3"
            End If
        ElseIf strChar Like "[DT]" Then
            If i < Len(strWord) And Mid(strWord, i + 1, 1) Like "[CSZ]" Then
                strEncoding = strEncoding & "8"
            Else
                strEncoding = strEncoding & "2"
            End If
        ElseIf strChar Like "[FVW]" Then
              strEncoding = strEncoding & "3"
        ElseIf strChar Like "[GKQ]" Then
            strEncoding = strEncoding & "4"
        ElseIf strChar = "C" Then
            If i > 1 Then
                If Mid(strWord, i - 1, 1) Like "[SZ]" Then
                    strEncoding = strEncoding & "8"
                ElseIf i < Len(strWord) And Mid(strWord, i + 1, 1) Like "[AHKOQUX]" Then
                    strEncoding = strEncoding & "4"
                Else
                    strEncoding = strEncoding & "8"
                End If
            ElseIf i = 1 Then
                If i < Len(strWord) And Mid(strWord, i + 1, 1) Like "[AHKLOQRUX]" Then
                    strEncoding = strEncoding & "4"
                Else
                    strEncoding = strEncoding & "8"
                End If
            End If
        ElseIf strChar = "X" Then
            If i > 1 Then
                If Mid(strWord, i - 1, 1) Like "[CKQ]" Then
                    strEncoding = strEncoding & "8"
                Else
                    strEncoding = strEncoding & "48"
                End If
            Else
                strEncoding = strEncoding & "48"
            End If
        ElseIf strChar = "L" Then
            strEncoding = strEncoding & "5"
        ElseIf strChar Like "[MN]" Then
            strEncoding = strEncoding & "6"
        ElseIf strChar = "R" Then
            strEncoding = strEncoding & "7"
        ElseIf strChar Like "[SZ]" Then
            strEncoding = strEncoding & "8"
        End If
    Next
    
    strEncoding = PhoneticFunctions.DeleteConsecutiveRepeats(strEncoding)
    strEncoding = Left(strEncoding, 1) & Replace(Mid(strEncoding, 2, Len(strEncoding)), "0", "")
    
    Koelner = strEncoding
End Function
