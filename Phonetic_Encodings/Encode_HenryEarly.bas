Attribute VB_Name = "Encode_HenryEarly"
Option Explicit
Function HenryEarly(strWord As String, Optional intMaxLength As Integer = 3) As String
    strWord = UCase(strWord)
    
    Dim strCharOnly As String
    Dim i As Integer
    For i = 1 To Len(strWord)
        If Mid(strWord, i, 1) Like "[ABCDEFGHIJKLMNOPQRSTUVWXYZ]" Then
            strCharOnly = strCharOnly & Mid(strWord, i, 1)
        End If
    Next
    strWord = strCharOnly
    
    If strWord = vbNullString Then
        HenryEarly = vbNullString
        Exit Function
    End If
    
    'Rule 1B
    If Left(strWord, 1) Like "[AEIOUY]" Then
        If (Mid(strWord, 2, 1) Like "[BCDFGHJKLPQRSTVWXZ]" And Mid(strWord, 3, 1) Like "[BCDFGHJKLMNPQRSTVWXZ]") Or (Mid(strWord, 2, 1) Like "[BCDFGHJKLMNPQRSTVWXZ]" And Not Mid(strWord, 3, 1) Like "[BCDFGHJKLMNPQRSTVWXZ]") Then
            If Left(strWord, 1) = "Y" Then
                strWord = "I" & Mid(strWord, 2)
            End If
        ElseIf Mid(strWord, 2, 1) Like "[MN]" And Mid(strWord, 3, 1) Like "[BCDFGHJKLMNPQRSTVWXZ]" Then
            If Left(strWord, 1) = "E" Then
                strWord = "A" & Mid(strWord, 2)
            ElseIf Left(strWord, 1) Like "[IUY]" Then
                strWord = "E" & Mid(strWord, 2)
            End If
        ElseIf Left(strWord, 2) = "AI" Or Left(strWord, 2) = "AY" Or Left(strWord, 2) = "EI" Or Left(strWord, 2) = "AU" Or Left(strWord, 2) = "OI" Or Left(strWord, 2) = "OU" Or Left(strWord, 2) = "EU" Then
            If Left(strWord, 2) = "AI" Or Left(strWord, 2) = "AY" Or Left(strWord, 2) = "EI" Then
                strWord = "E" & Mid(strWord, 3)
            ElseIf Left(strWord, 2) = "AU" Or Left(strWord, 2) = "OI" Or Left(strWord, 2) = "OU" Then
                strWord = "O" & Mid(strWord, 3)
            Else
                strWord = "U" & Mid(strWord, 3)
            End If
        ElseIf Mid(strWord, 2, 1) Like "[AEIOUY]" And Left(strWord, 1) = "Y" Then
            strWord = "I" & Mid(strWord, 2)
        End If
    End If
    
    Dim strCode As String
    Dim intSkip As Integer
    
    Dim strChar As String
    Dim strNextChar As String
    Dim strPrevChar As String
    For i = 1 To Len(strWord)
        strPrevChar = strChar
        strChar = Mid(strWord, i, 1)
        strNextChar = Mid(strWord, i + 1, 1)
        
        If intSkip > 0 Then
            intSkip = intSkip - 1
        ElseIf strChar Like "[AEIOUY]" Then
            strCode = strCode & strChar
        ElseIf strChar = strNextChar Then
            intSkip = 1
            strCode = strCode & strChar
        ElseIf Mid(strWord, i, 2) = "CQ" Or Mid(strWord, i, 2) = "DT" Or Mid(strWord, i, 2) = "SC" Then
            'Continue
        ElseIf strChar Like "[WXZ]" Then
            If strChar = "W" Then
                strCode = strCode & "V"
            Else
                strCode = strCode & "S"
            End If
        ElseIf strChar Like "[CGPQS]" Then
            If strChar = "C" Then
                If strNextChar Like "[AOULR]" Then
                    strCode = strCode & "K"
                ElseIf strNextChar Like "[EIY]" Then
                    strCode = strCode & "S"
                ElseIf strNextChar = "H" Then
                    If Mid(strWord, i + 2, 1) Like "[AEIOUY]" Then
                        strCode = strCode & "C"
                    Else
                        strCode = strCode & "K"
                    End If
                Else
                    strCode = strCode & "C"
                End If
            ElseIf strChar = "G" Then
                If strNextChar Like "[AOULR]" Then
                    strCode = strCode & "G"
                ElseIf strNextChar Like "[EIY]" Then
                    strCode = strCode & "J"
                ElseIf strNextChar = "N" Then
                    strCode = strCode & "N"
                End If
            ElseIf strChar = "P" Then
                If strNextChar <> "H" Then
                    strCode = strCode & "P"
                Else
                    strCode = strCode & "F"
                End If
            ElseIf strChar = "Q" Then
                If Mid(strWord, i + 1, 2) Like "U[EIY]" Then
                    strCode = strCode & "G"
                Else
                    strCode = strCode & "K"
                End If
            Else
                If Mid(strWord, i, 6) = "SAINTE" Then
                    strCode = strCode & "X"
                    intSkip = 5
                ElseIf Mid(strWord, i, 5) = "SAINT" Then
                    strCode = strCode & "X"
                    intSkip = 4
                ElseIf Mid(strWord, i, 3) = "STE" Then
                    strCode = strCode & "X"
                    intSkip = 2
                ElseIf Mid(strWord, i, 2) = "ST" Then
                    strCode = strCode & "X"
                    intSkip = 1
                ElseIf strNextChar Like "[BCDFGHJKLMNPQRSTVWXZ]" Then
                    'Continue
                Else
                    strCode = strCode & "S"
                End If
            End If
        ElseIf strChar = "H" And strPrevChar Like "[BCDFGHJKLMNPQRSTVWXZ]" Then
            'Continue
        ElseIf strChar Like "[BCDFGHJKMNPQSTVWXZ]" And strNextChar Like "[BCDFGHJKMNPQSTVWXZ]" Then
            'Continue
        ElseIf strChar = "L" And strNextChar Like "[MN]" Then
            'Continue
        ElseIf strChar Like "[MN]" And strPrevChar Like "[AEIOUY]" And strNextChar Like "[BCDFGHJKLMNPQRSTVWXZ]" Then
            'continue
        Else
            strCode = strCode & strChar
        End If
    Next
    
    If Right(strCode, 4) Like "[AEO]ULT" Then
        strCode = Left(strCode, Len(strCode) - 2)
    ElseIf Mid(strCode, Len(strCode) - 1, 1) = "R" And Right(strCode, 1) Like "[BCDFGHJKLMNPQRSTVWXZ]" Then
        strCode = Left(strCode, Len(strCode) - 1)
    ElseIf Mid(strCode, Len(strCode) - 1, 1) Like "[AEIOUY]" And Right(strCode, 1) Like "[DMNST]" Then
        strCode = Left(strCode, Len(strCode) - 1)
    ElseIf Right(strCode, 2) = "ER" Then
        strCode = Left(strCode, Len(strCode) - 1)
    End If
    
    strCode = Left(strCode, 1) & Replace(Replace(Replace(Replace(Replace(Replace(Mid(strCode, 2), "A", ""), "E", ""), "I", ""), "O", ""), "U", ""), "Y", "")
    
    If intMaxLength > 0 Then
        strCode = Left(strCode, intMaxLength)
    End If
    
    HenryEarly = strCode

End Function
