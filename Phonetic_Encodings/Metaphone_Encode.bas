Attribute VB_Name = "Metaphone_Encode"
Option Explicit
Function Metaphone(strWord As String, Optional intMaxLength As Integer = -1) As String
    If intMaxLength <> -1 Then
        intMaxLength = IIf(intMaxLength < 4, 4, intMaxLength)
    Else
        intMaxLength = 64
    End If
    
    Dim strEncodeName As String
    Dim i As Long
    strWord = UCase(strWord)
    For i = 1 To Len(strWord)
        If Mid(strWord, i, 1) Like "[A-Z]" Then
            strEncodeName = strEncodeName & Mid(strWord, i, 1)
        End If
    Next
    
    If Len(strEncodeName) = 0 Then
        Metaphone = ""
        Exit Function
    End If
    
    Select Case Left(strEncodeName, 2)
        Case "PN", "AE", "KN", "GN", "WR":
            strEncodeName = Mid(strEncodeName, 2)
    End Select
    If Left(strEncodeName, 1) = "X" Then
        strEncodeName = "S" & Mid(strEncodeName, 2)
    ElseIf Left(strEncodeName, 2) = "WH" Then
        strEncodeName = "W" & Mid(strEncodeName, 3)
    End If
    
    Dim strMetaphone As String
    Dim strChar As String, strPrev As String
    For i = 1 To Len(strEncodeName)
        If Len(strMetaphone) > intMaxLength Then Exit For
        strChar = Mid(strEncodeName, i, 1)
        If i > 1 Then strPrev = Mid(strEncodeName, i - 1, 1)
        If strChar Like "[!GT]" And i > 1 And strPrev = strChar Then
            'continue
        Else
            If i = 1 And strChar Like "[AEIOU]" Then
                strMetaphone = Left(strEncodeName, 1)
            ElseIf strChar = "B" Then
                If i <> Len(strEncodeName) Or strPrev <> "M" Then
                    strMetaphone = strMetaphone & "B"
                End If
            ElseIf strChar = "T" Then
                If i > 1 And i + 2 <= Len(strEncodeName) And Mid(strEncodeName, i + 1, 1) = "I" And Mid(strEncodeName, i + 2, 1) Like "[AO]" Then
                    strMetaphone = strMetaphone & "X"
                ElseIf Mid(strEncodeName, i + 1, 1) = "H" Then
                    strMetaphone = strMetaphone & "0"
                ElseIf Mid(strEncodeName, i + 1, 2) <> "CH" Then
                    If strPrev <> "T" Then
                        strMetaphone = strMetaphone & "T"
                    End If
                End If
             ElseIf strChar = "Q" Then
                strMetaphone = strMetaphone & "K"
            ElseIf strChar = "V" Then
                strMetaphone = strMetaphone & "F"
            ElseIf strChar = "K" Then
                If i > 1 And strPrev = "C" Then
                    'continue
                Else
                    strMetaphone = strMetaphone & "K"
                End If
            ElseIf strChar = "W" Or strChar = "Y" Then
                If Mid(strEncodeName, i + 1, 1) Like "[AEIOU]" Then
                    strMetaphone = strMetaphone & strChar
                End If
            ElseIf strChar = "D" Then
                If Mid(strEncodeName, i + 1, 1) = "G" And Mid(strEncodeName, i + 2, 1) Like "[EIY]" Then
                    strMetaphone = strMetaphone & "J"
                Else
                    strMetaphone = strMetaphone & "T"
                End If
            ElseIf strChar = "P" Then
                If Mid(strEncodeName, i + 1, 1) = "H" Then strMetaphone = strMetaphone & "F" Else strMetaphone = strMetaphone & "P"
            ElseIf strChar Like "[FJLMNR]" Then
                strMetaphone = strMetaphone & strChar
            ElseIf strChar = "X" Then
                strMetaphone = strMetaphone & "KS"
            ElseIf strChar = "Z" Then
                strMetaphone = strMetaphone & "S"
            ElseIf strChar = "S" Then
                If i > 1 And i + 2 <= Len(strEncodeName) And Mid(strEncodeName, i + 1, 1) = "I" And Mid(strEncodeName, i + 2, 1) Like "[OA]" Then
                    strMetaphone = strMetaphone & "X"
                ElseIf Mid(strEncodeName, i + 1, 1) = "H" Then
                    strMetaphone = strMetaphone & "X"
                Else
                    strMetaphone = strMetaphone & "S"
                End If
            ElseIf strChar = "H" Then
                If i > 1 And strPrev Like "[AEIOU]" And Mid(strEncodeName, i + 1, 1) Like "[!AEIOU]" Then
                    'continue
                ElseIf i > 1 And strPrev Like "[CGPST]" Then
                    'continue
                Else
                    strMetaphone = strMetaphone & "H"
                End If
            ElseIf strChar = "G" Then
                If Mid(strEncodeName, i + 1, 1) = "H" And Not (i + 1 = Len(strEncodeName) Or Mid(strEncodeName, i + 2, 1) Like "[!AEIOU]") Then
                    'Continue
                ElseIf i > 1 And (i + 1 = Len(strEncodeName) And Mid(strEncodeName, i + 1, 1) = "N") Or (i + 3 = Len(strEncodeName) And Mid(strEncodeName, i + 1, 3) = "NED") Then
                    'continue
                ElseIf i - 1 > 1 And i + 1 <= Len(strEncodeName) And strPrev = "D" And Mid(strEncodeName, i + 1, 1) Like "[EIY]" Then
                    'continue
                ElseIf Mid(strEncodeName, i + 1, 1) = "G" Then
                    'continue
                ElseIf Mid(strEncodeName, i + 1, 1) Like "[EIY]" Then
                    If i = 1 Or strPrev <> "G" Then
                        strMetaphone = strMetaphone & "J"
                    Else
                        strMetaphone = strMetaphone & "K"
                    End If
                Else
                    strMetaphone = strMetaphone & "K"
                End If
                
            ElseIf strChar = "C" Then
                If Not (i > 0 And strPrev = "S" And Mid(strEncodeName, i + 1, 1) Like "[EIY]") Then
                    If Mid(strEncodeName, i + 1, 2) = "IA" Then
                        strMetaphone = strMetaphone & "X"
                    ElseIf Mid(strEncodeName, i + 1, 1) Like "[EIY]" Then
                        strMetaphone = strMetaphone & "S"
                    ElseIf i > 0 And strPrev = "C" And Mid(strEncodeName, i, 2) = "CH" Then
                        strMetaphone = strMetaphone & "K"
                    ElseIf Mid(strEncodeName, i + 1, 1) = "H" Then
                        If (i = 1 And i + 1 < Len(strEncodeName) And Mid(strEncodeName, i + 2, 1) Like "[!AEIOU]") Then
                            strMetaphone = strMetaphone & "K"
                        Else
                            strMetaphone = strMetaphone & "X"
                        End If
                    Else
                        strMetaphone = strMetaphone & "K"
                    End If
                End If
            
            End If
        End If
    Next
    
    Metaphone = strMetaphone
End Function

