Attribute VB_Name = "Phonex_Encode"
Option Explicit
Function Phonex(strWord As String, Optional intMaxLength As Integer = 4, Optional boolZeroPad As Boolean = True)
    If intMaxLength < 0 Then intMaxLength = 64
    If intMaxLength = 0 Then intMaxLength = 4
    
    strWord = UnicodeFunctions.UnicodeStrip(UCase$(strWord))
    Dim strNameCode As String, strLast As String
    
    While Right(strWord, 1) = "S"
        strWord = Left(strWord, Len(strWord) - 1)
    Wend
    
    If Left(strWord, 2) = "KN" Then
        strWord = "N" & Mid(strWord, 3)
    ElseIf Left(strWord, 2) = "PH" Then
        strWord = "F" & Mid(strWord, 3)
    ElseIf Left(strWord, 2) = "WR" Then
        strWord = "R" & Mid(strWord, 3)
    End If
    
    If Left(strWord, 1) = "H" Then
        strWord = Mid(strWord, 2)
    End If
    
    If Left(strWord, 1) Like "[AEIOUY]" Then
        strWord = "A" & Mid(strWord, 2)
    ElseIf Left(strWord, 1) Like "[BP]" Then
        strWord = "B" & Mid(strWord, 2)
    ElseIf Left(strWord, 1) Like "[VF]" Then
        strWord = "F" & Mid(strWord, 2)
    ElseIf Left(strWord, 1) Like "[CKQ]" Then
        strWord = "C" & Mid(strWord, 2)
    ElseIf Left(strWord, 1) Like "[GJ]" Then
        strWord = "G" & Mid(strWord, 2)
    ElseIf Left(strWord, 1) Like "[SZ]" Then
        strWord = "S" & Mid(strWord, 2)
    End If
    
    strNameCode = Left(strWord, 1)
    strLast = strNameCode
    
    Dim i As Long
    Dim strCode As String
    For i = 2 To Len(strWord)
        strCode = "0"
        If Mid(strWord, i, 1) Like "[BFPV]" Then
            strCode = "1"
        ElseIf Mid(strWord, i, 1) Like "[CGJKQSXZ]" Then
            strCode = "2"
        ElseIf Mid(strWord, i, 1) Like "[DT]" Then
            If Mid(strWord, i + 1, 1) <> "C" Then
                strCode = "3"
            End If
        ElseIf Mid(strWord, i, 1) = "L" Then
            If Mid(strWord, i + 1, 1) Like "[AEIOUY]" Or i = Len(strWord) Then
                strCode = "4"
            End If
         ElseIf Mid(strWord, i, 1) Like "[MN]" Then
            If Mid(strWord, i + 1, 1) Like "[DG]" Then
                Mid(strWord, i + 1) = Mid(strWord, i, 1)
            End If
            
            strCode = "5"
        ElseIf Mid(strWord, i, 1) = "R" Then
            If Mid(strWord, i + 1, 1) Like "[AEIOUY]" Or i = Len(strWord) Then
                strCode = "6"
            End If
        End If
        
        If strLast <> strCode And strCode <> "0" Then
            strNameCode = strNameCode & strCode
        End If
        
        strLast = Right(strNameCode, 1)
    Next
    
    If boolZeroPad Then
        strNameCode = strNameCode & String(intMaxLength, "0")
    End If
    
    If strNameCode = "" Then strNameCode = "0"
    
    Phonex = Left(strNameCode, intMaxLength)
End Function

