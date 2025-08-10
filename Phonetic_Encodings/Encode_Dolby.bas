Attribute VB_Name = "Encode_Dolby"
Option Explicit
Function Dolby(strWord As String, Optional intMaxLength As Integer = -1, Optional boolKeepVowels As Boolean = False, Optional strVowelChar As String = "*") As String
    strWord = UCase$(strWord)
    Dim strCharOnly As String
    Dim i As Integer
    For i = 1 To Len(strWord)
        If Mid$(strWord, i, 1) Like "[ABCDEFGHIJKLMNOPQRSTUVWXYZ]" Then
            strCharOnly = strCharOnly & Mid$(strWord, i, 1)
        End If
    Next
    strWord = strCharOnly
    
    If Left$(strWord, 3) = "MCG" Or Left$(strWord, 3) Like "MA[GC]" Then
        strWord = "MK" & Mid$(strWord, 4)
    ElseIf Left$(strWord, 2) = "MC" Then
        strWord = "MK" & Mid$(strWord, 3)
    End If
    
    Dim intPos As Integer
    intPos = Len(strWord) - 1
    
    Do While intPos >= 1
        Select Case Mid$(strWord, intPos, 2)
            Case "DT", "LD", "ND", "NT", "RC", "RD", "RT", "SC", "SK", "ST"
                strWord = Left$(strWord, intPos) & Mid$(strWord, intPos + 2)
            Case Else
                intPos = intPos - 1
        End Select
    Loop
    
    strWord = Replace(strWord, "X", "KS")
    strWord = Replace(strWord, "CE", "SE")
    strWord = Replace(strWord, "CI", "SI")
    strWord = Replace(strWord, "CY", "SI")
    strWord = Replace(strWord, "TCH", "CH")
    
    intPos = InStr(2, strWord, "CH")
    Do While intPos > 0
        If Not Mid$(strWord, intPos - 1, 1) Like "[AEIOUY]" Then
            Mid$(strWord, intPos, 1) = "S"
        End If
        
        intPos = InStr(intPos + 1, strWord, "CH")
    Loop
    
    strWord = Replace(strWord, "C", "K")
    strWord = Replace(strWord, "Z", "S")
    
    strWord = Replace(strWord, "WR", "R")
    strWord = Replace(strWord, "DG", "G")
    strWord = Replace(strWord, "QU", "K")
    strWord = Replace(strWord, "T", "D")
    strWord = Replace(strWord, "PH", "F")
    
    intPos = InStr(1, strWord, "K")
    Do While intPos > 0
        If intPos > 2 Then
            If Not (Mid$(strWord, intPos - 1, 1) Like "[AEIOUYLNR]") Then
                strWord = Left$(strWord, intPos - 2) & Mid$(strWord, intPos)
                intPos = intPos - 1
            End If
        End If

        intPos = InStr(intPos + 1, strWord, "K")
    Loop
    
    If intMaxLength > 0 And Right$(strWord, 1) = "E" Then strWord = Left$(strWord, Len(strWord) - 1)

    strWord = DeleteConsecutiveRepeats(strWord)
    If Left$(strWord, 2) = "PF" Then
        strWord = Mid$(strWord, 2)
    End If

    If Right$(strWord, 2) = "PF" Then
        strWord = Left$(strWord, Len(strWord) - 1)
    ElseIf Right$(strWord, 2) = "GH" Then
        If Mid$(strWord, Len(strWord) - 2, 1) Like "[AEIOUY]" Then
            strWord = Left$(strWord, Len(strWord) - 2) & "F"
        Else
            strWord = Left$(strWord, Len(strWord) - 2) & "G"
        End If
    End If
    strWord = Replace$(strWord, "GH", "")
    
    If intMaxLength > 0 Then strWord = Replace(strWord, "V", "F")
    
    Dim intFirst As Integer
    intFirst = 1 + IIf(intMaxLength > 0, 1, 0)
    Dim strCode As String
    For i = 1 To Len(strWord)
        If Mid$(strWord, i, 1) Like "[AEIOUY]" Then
            If intFirst > 0 Or boolKeepVowels Then
                strCode = strCode & strVowelChar
                intFirst = intFirst - 1
            End If
        ElseIf i > 1 And Mid$(strWord, i, 1) Like "[WH]" Then
            'Continue
        Else
            strCode = strCode & Mid$(strWord, i, 1)
        End If
    Next
    
    If intMaxLength > 0 Then
        If Len(strCode) > intMaxLength And Right$(strWord, 1) = "S" Then
            strCode = Left$(strCode, Len(strCode) - 1)
        End If
        If boolKeepVowels Then
            strCode = Left$(strCode, intMaxLength)
        Else
            Dim intVowels As Integer
            Dim intExcess As Integer
            Dim strTmp As String
            strCode = Left$(strCode, intMaxLength + 2)
            
            Do While Len(strCode) > intMaxLength
                intVowels = Len(strCode) - intMaxLength
                intExcess = intVowels - 1
                
                strTmp = strCode
                strCode = ""
                For i = 1 To Len(strTmp)
                    If Mid$(strTmp, i, 1) = strVowelChar Then
                        If intVowels > 0 Then
                            strCode = strCode & Mid$(strTmp, i, 1)
                            intVowels = intVowels - 1
                        End If
                    Else
                        strCode = strCode & Mid$(strTmp, i, 1)
                    End If
                Next i
                
                strCode = Left$(strCode, intMaxLength + intExcess)
            Loop
        End If
    End If
    Dolby = strCode
    
End Function
