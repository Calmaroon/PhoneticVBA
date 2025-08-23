Attribute VB_Name = "Encode_Norphone"
Option Explicit
Function Norphone(strWord As String) As String
    strWord = UCase$(strWord)
    
    Dim strCode As String
    Dim intSkip As Integer
    
    If Left(strWord, 2) = "AA" Then
        strCode = "Å"
        intSkip = 2
    ElseIf Left(strWord, 2) = "GI" Then
        strCode = "J"
        intSkip = 2
    ElseIf Left(strWord, 3) = "SKY" Then
        strCode = "X"
        intSkip = 3
    ElseIf Left(strWord, 2) = "EI" Then
        strCode = "Æ"
        intSkip = 2
    ElseIf Left(strWord, 2) = "KY" Then
        strCode = "X"
        intSkip = 2
    ElseIf Left(strWord, 1) = "C" Then
        strCode = "K"
        intSkip = 1
    ElseIf Left(strWord, 1) = "Ä" Then
        strCode = "Æ"
        intSkip = 1
    ElseIf Left(strWord, 1) = "Ö" Then
        strCode = "Ø"
        intSkip = 1
    End If
    
    If Right(strWord, 2) = "DT" Then
        strWord = Left(strWord, Len(strWord) - 2) & "T"
    ElseIf Mid(strWord, Len(strWord) - 1, 1) Like "[AEIOUYÅÆØÄÖ]" And Right(strWord, 1) = "D" Then
        strWord = Left(strWord, Len(strWord) - 2)
    End If
    
    Dim i As Long
    Dim boolMatched As Boolean
    For i = intSkip + 1 To Len(strWord)
            boolMatched = False
            Select Case Mid(strWord, i, 4)
                Case "SKEI":
                    strCode = strCode & "X"
                    boolMatched = True
                    i = i + 3
            End Select
            
            If Not boolMatched Then
                Select Case Mid(strWord, i, 3)
                    Case "SKJ", "KEI":
                        strCode = strCode & "X"
                        boolMatched = True
                        i = i + 2
                End Select
            End If
            
            If Not boolMatched Then
                Select Case Mid(strWord, i, 2)
                    Case "CH", "CK", "GH", "HG":
                        strCode = strCode & "K"
                        boolMatched = True
                        i = i + 1
                    Case "GJ", "HJ":
                        strCode = strCode & "J"
                        boolMatched = True
                        i = i + 1
                    Case "HL", "LD":
                        strCode = strCode & "L"
                        boolMatched = True
                        i = i + 1
                    Case "HR":
                        strCode = strCode & "R"
                        boolMatched = True
                        i = i + 1
                    Case "KJ", "KI", "SJ":
                        strCode = strCode & "X"
                        boolMatched = True
                        i = i + 1
                    Case "ND":
                        strCode = strCode & "N"
                        boolMatched = True
                        i = i + 1
                    Case "PH":
                        strCode = strCode & "F"
                        boolMatched = True
                        i = i + 1
                    Case "TH":
                        strCode = strCode & "T"
                        boolMatched = True
                        i = i + 1
                End Select
            End If
            
            If Not boolMatched Then
                Select Case Mid(strWord, i, 1)
                    Case "W": strCode = strCode & "V": boolMatched = True
                    Case "X": strCode = strCode & "KS": boolMatched = True
                    Case "Z": strCode = strCode & "S": boolMatched = True
                    Case "D": strCode = strCode & "T": boolMatched = True
                    Case "G": strCode = strCode & "K": boolMatched = True
                End Select
            End If

            If (i = 1 Or Mid(strWord, i, 1) Like "[!AEIOUYÅÆØÄÖ]") And Not boolMatched Then
                strCode = strCode & Mid(strWord, i, 1)
            End If
    Next
    
    strCode = DeleteConsecutiveRepeats(strCode)
    
    Norphone = strCode
End Function
