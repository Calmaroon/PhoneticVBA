Attribute VB_Name = "Encode_NYSIIS"
Option Explicit
Function NYSIIS(strWord As String, Optional intMaxLength As Integer = 6, Optional boolModified As Boolean = False) As String
    If intMaxLength < 6 Then intMaxLength = 6
    Dim strWordAlpha As String

    Dim i As Long
    
    strWord = UCase$(strWord)
    For i = 1 To Len(strWord)
        If Mid(strWord, i, 1) Like "[A-Z]" Then
            strWordAlpha = strWordAlpha & Mid(strWord, i, 1)
        End If
    Next
    
    If Len(strWordAlpha) = 0 Then
        NYSIIS = ""
        Exit Function
    End If
    
    Dim strOriginalFirstChar As String
    strOriginalFirstChar = Left(strWord, 1)
    
    If Left(strWordAlpha, 3) = "MAC" Then
        strWordAlpha = "MCC" & Mid(strWordAlpha, 4, Len(strWordAlpha))
    ElseIf Left(strWordAlpha, 2) = "KN" Then
        strWordAlpha = "NN" & Mid(strWordAlpha, 3, Len(strWordAlpha))
    ElseIf Left(strWordAlpha, 1) = "K" Then
        strWordAlpha = "C" & Mid(strWordAlpha, 2, Len(strWordAlpha))
    ElseIf Left(strWordAlpha, 2) = "PH" Or Left(strWordAlpha, 2) = "PF" Then
        strWordAlpha = "FF" & Mid(strWordAlpha, 3, Len(strWordAlpha))
    ElseIf Left(strWordAlpha, 3) = "SCH" Then
        strWordAlpha = "SSS" & Mid(strWordAlpha, 4, Len(strWordAlpha))
    ElseIf boolModified Then
        If Left(strWordAlpha, 2) = "WR" Then
            strWordAlpha = "RR" & Mid(strWordAlpha, 3, Len(strWordAlpha))
        ElseIf Left(strWordAlpha, 2) = "RH" Then
            strWordAlpha = "R" & Mid(strWordAlpha, 3, Len(strWordAlpha))
        ElseIf Left(strWordAlpha, 2) = "DG" Then
            strWordAlpha = "GG" & Mid(strWordAlpha, 3, Len(strWordAlpha))
        ElseIf Left(strWordAlpha, 1) Like "[AEIOU]" Then
            strWordAlpha = "A" & Mid(strWordAlpha, 2, Len(strWordAlpha))
        End If
    End If
    
    
    If boolModified And Right(strWordAlpha, 1) Like "[SZ]" Then
        strWordAlpha = Left(strWordAlpha, Len(strWordAlpha) - 1)
    End If
    
    If Right(strWordAlpha, 2) = "EE" Or Right(strWordAlpha, 2) = "IE" Or (boolModified And Right(strWordAlpha, 2) = "YE") Then
        strWordAlpha = Left(strWordAlpha, Len(strWordAlpha) - 2) & "Y"
    ElseIf Right(strWordAlpha, 2) = "DT" Or Right(strWordAlpha, 2) = "RT" Or Right(strWordAlpha, 2) = "RD" Then
        strWordAlpha = Left(strWordAlpha, Len(strWordAlpha) - 2) & "D"
    ElseIf Right(strWordAlpha, 2) = "NT" Or Right(strWordAlpha, 2) = "ND" Then
        If boolModified Then
            strWordAlpha = Left(strWordAlpha, Len(strWordAlpha) - 2) & "N"
        Else
            strWordAlpha = Left(strWordAlpha, Len(strWordAlpha) - 2) & "D"
        End If
    ElseIf boolModified Then
        If Right(strWordAlpha, 2) = "IX" Then
            strWordAlpha = Left(strWordAlpha, Len(strWordAlpha) - 2) & "ICK"
        ElseIf Right(strWordAlpha, 2) = "EX" Then
            strWordAlpha = Left(strWordAlpha, Len(strWordAlpha) - 2) & "ECK"
        ElseIf Right(strWordAlpha, 2) = "JR" Or Right(strWordAlpha, 2) = "SR" Then
            'Throw error
        End If
        
    End If
    
    Dim strKey As String
    strKey = Left(strWordAlpha, 1)
    
    
    
    Dim intSkip As Integer
    intSkip = 0
    
    i = 2
    'Debug.Print strWordAlpha
    Do While i <= Len(strWordAlpha)
        If intSkip > 0 Then
            intSkip = intSkip - 1
            i = i + 1
            GoTo NextChar
        End If
         
        If Mid(strWordAlpha, i, 2) = "EV" Then
            Mid(strWordAlpha, i, 2) = "AF"
            intSkip = 0
        ElseIf Mid(strWordAlpha, i, 1) Like "[AEIOU]" Then
            Mid(strWordAlpha, i, 1) = "A"
        ElseIf boolModified And i <> Len(strWordAlpha) And Mid(strWordAlpha, i, 1) = "Y" Then
            Mid(strWordAlpha, i, 1) = "A"
        ElseIf Mid(strWordAlpha, i, 1) = "Q" Then
            Mid(strWordAlpha, i, 1) = "G"
        ElseIf Mid(strWordAlpha, i, 1) = "Z" Then
            Mid(strWordAlpha, i, 1) = "S"
        ElseIf Mid(strWordAlpha, i, 1) = "M" Then
            Mid(strWordAlpha, i, 1) = "N"
        ElseIf i < Len(strWordAlpha) And Mid(strWordAlpha, i, 2) = "KN" Then
            Mid(strWordAlpha, i, 2) = "N"
            intSkip = 1
        ElseIf Mid(strWordAlpha, i, 1) = "K" Then
            Mid(strWordAlpha, i, 1) = "C"

        ElseIf boolModified And i = Len(strWordAlpha) - 3 And Mid(strWordAlpha, i, 3) = "SCH" Then
            Mid(strWordAlpha, i, 3) = "SSA"
            intSkip = 1
        ElseIf Mid(strWordAlpha, i, 3) = "SCH" Then
            Mid(strWordAlpha, i, 3) = "SSS"
            intSkip = 1
            
        ElseIf boolModified And i = Len(strWordAlpha) - 2 And Mid(strWordAlpha, i, 2) = "SH" Then
            Mid(strWordAlpha, i, 2) = "SA"
            intSkip = 0
        ElseIf Mid(strWordAlpha, i, 2) = "SH" Then
            Mid(strWordAlpha, i, 2) = "SS"
            intSkip = 1
        ElseIf Mid(strWordAlpha, i, 2) = "PH" Then
            Mid(strWordAlpha, i, 2) = "FF"
            intSkip = 1
        ElseIf boolModified And Mid(strWordAlpha, i, 3) = "GHT" Then
            Mid(strWordAlpha, i, 3) = "TTT"
            intSkip = 2
        ElseIf boolModified And Mid(strWordAlpha, i, 2) = "DG" Then
            Mid(strWordAlpha, i, 2) = "GG"
            intSkip = 1
        ElseIf boolModified And Mid(strWordAlpha, i, 2) = "WR" Then
            Mid(strWordAlpha, i, 2) = "RR"
            intSkip = 1
        ElseIf Mid(strWordAlpha, i, 1) = "H" And (Not Mid(strWordAlpha, i - 1, 1) Like "[AEIOU]" Or Not Mid(strWordAlpha, i + 1, 1) Like "[AEIOU]") Then
            strWordAlpha = Left(strWordAlpha, i - 1) & Mid(strWordAlpha, i - 1, 1) & Mid(strWordAlpha, i + 1)
        ElseIf Mid(strWordAlpha, i, 1) = "W" And Mid(strWordAlpha, i - 1, 1) Like "[AEIOU]" Then
            strWordAlpha = Left(strWordAlpha, i - 1) & Mid(strWordAlpha, i - 1, 1) & Mid(strWordAlpha, i + 1)
            intSkip = 0
        End If
        
        If Right(strKey, 1) <> Right(Mid(strWordAlpha, i, intSkip + 1), 1) Then
            strKey = strKey & Mid(strWordAlpha, i, intSkip + 1)
        End If
        i = i + 1
NextChar:
    Loop

    strKey = DeleteConsecutiveRepeats(strKey)
    
    If Right(strKey, 1) = "S" Then
        strKey = Left(strKey, Len(strKey) - 1)
    End If
    
    If Right(strKey, 2) = "AY" Then
        strKey = Left(strKey, Len(strKey) - 2) & "Y"
    End If
    
    If Right(strKey, 1) = "A" Then
        strKey = Left(strKey, Len(strKey) - 1)
    End If
    
    If boolModified And Left(strKey, 1) = "A" Then
        strKey = strOriginalFirstChar & Mid(strKey, 2)
    End If
    
    strKey = Left(strKey, intMaxLength)
    

    
    NYSIIS = strKey

End Function
