Attribute VB_Name = "Encode_DoubleMetaphone"
Option Explicit
Function DoubleMetaphone(strWord As String, Optional intMaxLength As Integer = -1) As String
    Dim strPrimary As String
    Dim strSecondary As String
    Dim intCurrent As Integer: intCurrent = 1
    Dim intLength As Integer: intLength = Len(strWord)
    
    If intLength < 1 Then
        DoubleMetaphone = ","
        Exit Function
    End If
    
    Dim intLast As Integer
    intLast = intLength - 1
    strWord = UCase$(strWord)
    
    If left(strWord, 2) = "GN" Or left(strWord, 2) = "KN" Or left(strWord, 2) = "PN" Or left(strWord, 2) = "WR" Or left(strWord, 2) = "PS" Then intCurrent = intCurrent + 1
    
    If left(strWord, 1) = "X" Then
        strPrimary = strPrimary & "S": strSecondary = strSecondary & "S"
        intCurrent = intCurrent + 1
    End If
    
    Dim i As Integer
    
    Dim currChar As String
    For i = intCurrent To Len(strWord)
        currChar = Mid(strWord, i, 1)
        
        Select Case currChar
            Case "A", "E", "I", "O", "U", "Y"
                If i = 1 Then
                    strPrimary = strPrimary & "A": strSecondary = strSecondary & "A"
                End If
            Case "B":
                strPrimary = strPrimary & "P": strSecondary = strSecondary & "P"
                If Mid(strWord, i + 1, 1) = "B" Then i = i + 1
            Case "Ç":
                strPrimary = strPrimary & "S": strSecondary = strSecondary & "S"
            Case "C":
                If i > 2 And Not IsVowel(strWord, i - 2) And StringAt(strWord, i - 1, 3, "ACH") And (Mid(strWord, i + 2, 1) <> "I" And (Mid(strWord, i + 2, 1) <> "E" Or StringAt(strWord, i - 2, 6, "BACHER,MACHER"))) Then
                    strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                    i = i + 1
                ElseIf i = 1 And StringAt(strWord, i, 6, "CAESAR") Then
                    strPrimary = strPrimary & "S": strSecondary = strSecondary & "S"
                    i = i + 1
                ElseIf StringAt(strWord, i, 4, "CHIA") Then
                    strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                    i = i + 1
                ElseIf StringAt(strWord, i, 2, "CH") Then
                    If i > 1 And StringAt(strWord, i, 4, "CHAE") Then
                        strPrimary = strPrimary & "K": strSecondary = strSecondary & "X"
                        i = i + 1
                    ElseIf (i = 1 And (StringAt(strWord, i + 1, 5, "HARAC,HARIS") Or StringAt(strWord, i + 1, 3, "HOR,HYM,HIA,HEM"))) And Not StringAt(strWord, 1, 5, "CHORE") Then
                        strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                        i = i + 1
                    ElseIf (StringAt(strWord, 1, 4, "VAN ,VON ") Or StringAt(strWord, 1, 3, "SCH")) Or StringAt(strWord, i - 2, 6, "ORCHES,ARCHIT,ORCHID") Or StringAt(strWord, i + 2, 1, "T,S") Or (StringAt(strWord, i - 1, 1, "A,O,U,E") Or i = 1) And StringAt(strWord, i + 2, 1, "L,R,N,M,B,H,F,V,W, ") Then 'long one
                        strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                        i = i + 1
                    Else
                        If i > 1 Then
                            If StringAt(strWord, 1, 2, "MC") Then
                                strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                                 i = i + 1
                            Else
                                strPrimary = strPrimary & "X": strSecondary = strSecondary & "K"
                                 i = i + 1
                            End If
                        Else
                            strPrimary = strPrimary & "X": strSecondary = strSecondary & "X"
                             i = i + 1
                        End If
                    End If
                ElseIf StringAt(strWord, i, 2, "CZ") And Not StringAt(strWord, i - 2, 4, "WICZ") Then
                    strPrimary = strPrimary & "S": strSecondary = strSecondary & "X"
                    i = i + 1
                ElseIf StringAt(strWord, i + 1, 3, "CIA") Then
                    strPrimary = strPrimary & "X": strSecondary = strSecondary & "X"
                    i = i + 2
                ElseIf StringAt(strWord, i, 2, "CC") And Not (i = 2 And left(strWord, 1) = "M") Then
                    If StringAt(strWord, i + 2, 1, "I,E,H") And Not StringAt(strWord, i + 2, 2, "HU") Then
                        If (i = 2 And left(strWord, 1) = "A") Or StringAt(strWord, i - 1, 5, "UCCEE,UCCES") Then
                            strPrimary = strPrimary & "KS": strSecondary = strSecondary & "KS"
                        Else
                            strPrimary = strPrimary & "X": strSecondary = strSecondary & "X"
                        End If
                        i = i + 2
                    Else
                        strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                        i = i + 1
                    End If
                ElseIf StringAt(strWord, i, 2, "CK,CG,CQ") Then
                    strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                    i = i + 1
                ElseIf StringAt(strWord, i, 2, "CI,CE,CY") Then
                    If StringAt(strWord, i, 3, "CIO,CIE,CIA") Then
                        strPrimary = strPrimary & "S": strSecondary = strSecondary & "X"
                    Else
                        strPrimary = strPrimary & "S": strSecondary = strSecondary & "S"
                    End If
                    i = i + 1
                Else
                    strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                    If StringAt(strWord, i + 1, 2, " C, Q, G") Then
                        i = i + 2
                    ElseIf StringAt(strWord, i + 1, 1, "C,K,Q") And Not StringAt(strWord, i + 1, 2, "CE,CI") Then
                        i = i + 1
                    End If
                End If
            Case "D":
                If Mid(strWord, i, 2) = "DG" Then
                    If Mid(strWord, i + 2, 1) Like "[IEY]" Then
                        strPrimary = strPrimary & "J": strSecondary = strSecondary & "J"
                        i = i + 2
                    Else
                        strPrimary = strPrimary & "TK": strSecondary = strSecondary & "TK"
                        i = i + 1
                    End If
                ElseIf Mid(strWord, i, 2) Like "D[TD]" Then
                    strPrimary = strPrimary & "T": strSecondary = strSecondary & "T"
                    i = i + 1
                Else
                    strPrimary = strPrimary & "T": strSecondary = strSecondary & "T"
                End If
            Case "F":
                If Mid(strWord, i + 1, 1) = "F" Then i = i + 1
                strPrimary = strPrimary & "F": strSecondary = strSecondary & "F"
            Case "G":
                If Mid(strWord, i + 1, 1) = "H" Then
                    If i > 1 And Not IsVowel(strWord, i - 1) Then
                        strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                        i = i + 1
                    ElseIf i = 1 Then
                        If Mid(strWord, i + 2, 1) = "I" Then
                            strPrimary = strPrimary & "J": strSecondary = strSecondary & "J"
                        Else
                            strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                        End If
                        i = i + 1
                    ElseIf (i > 2 And StringAt(strWord, i - 2, 1, "B,H,D")) Or (i > 3 And StringAt(strWord, i - 3, 1, "B,H,D")) Or (i > 4 And StringAt(strWord, i - 4, 1, "B,H")) Then '
                        i = i + 1
                    Else
                        If i > 3 And StringAt(strWord, i - 1, 1, "U") And StringAt(strWord, i - 3, 1, "C,G,L,R,T") Then
                            strPrimary = strPrimary & "F": strSecondary = strSecondary & "F"
                        ElseIf i > 1 And Not StringAt(strWord, i - 1, 1, "I") Then
                            strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                        End If
                        i = i + 1
                    End If
                ElseIf Mid(strWord, i + 1, 1) = "N" Then
                    If i = 2 And IsVowel(strWord, 1) And Not SlavoGermanic(strWord) Then
                        strPrimary = strPrimary & "KN": strSecondary = strSecondary & "N"
                    ElseIf (Not StringAt(strWord, i + 2, 2, "EY") And Mid(strWord, i + 1, 1) <> "Y") And Not SlavoGermanic(strWord) Then
                        strPrimary = strPrimary & "N": strSecondary = strSecondary & "KN"
                    Else
                        strPrimary = strPrimary & "KN": strSecondary = strSecondary & "KN"
                    End If
                    i = i + 1
                ElseIf StringAt(strWord, i + 1, 2, "LI") And Not SlavoGermanic(strWord) Then
                    strPrimary = strPrimary & "KL": strSecondary = strSecondary & "L"
                    i = i + 1
                ElseIf i = 1 And (Mid(strWord, i + 1, 1) = "Y" Or StringAt(strWord, i + 1, 2, "ES,EP,EB,EL,EY,IB,IL,IN,IE,EI,ER")) Then
                    strPrimary = strPrimary & "K": strSecondary = strSecondary & "J"
                    i = i + 1
                ElseIf (StringAt(strWord, i + 1, 2, "ER") Or Mid(strWord, i + 1, 1) = "Y") And Not StringAt(strWord, 1, 6, "DANGER,RANGER,MANGER") And Not StringAt(strWord, i - 1, 1, "E,I") And Not StringAt(strWord, i - 1, 3, "RGY,OGY") Then
                    strPrimary = strPrimary & "K": strSecondary = strSecondary & "J"
                    i = i + 1
                ElseIf StringAt(strWord, i + 1, 1, "E,I,Y") Or StringAt(strWord, i - 1, 4, "AGGI,OGGI") Then
                    If StringAt(strWord, 1, 4, "VAN ,VON ") Or StringAt(strWord, 1, 3, "SCH") Or StringAt(strWord, i + 1, 2, "ET") Then
                        strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                    ElseIf StringAt(strWord, i + 1, 4, "IER ") Then
                        strPrimary = strPrimary & "J": strSecondary = strSecondary & "J"
                    Else
                        strPrimary = strPrimary & "J": strSecondary = strSecondary & "K"
                    End If
                    i = i + 1
                Else
                    If Mid(strWord, i + 1, 1) = "G" Then
                        i = i + 1
                    End If
                    strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
                End If
            Case "H":
                If (i = 1 Or IsVowel(strWord, i - 1)) And IsVowel(strWord, i + 1) Then
                    strPrimary = strPrimary & "H": strSecondary = strSecondary & "H"
                    i = i + 1
                End If
             Case "J":
                If StringAt(strWord, i, 4, "JOSE") Or StringAt(strWord, 1, 4, "SAN ") Then
                    If i = 1 And Mid(strWord & "   ", i + 4, 1) = " " Or left(strWord, 4) = "SAN " Then
                        strPrimary = strPrimary & "H": strSecondary = strSecondary & "H"
                    Else
                        strPrimary = strPrimary & "J": strSecondary = strSecondary & "H"
                    End If
                ElseIf i = 1 And Not StringAt(strWord, i, 4, "JOSE") Then
                    strPrimary = strPrimary & "J": strSecondary = strSecondary & "A"
                ElseIf IsVowel(strWord, i - 1) And Not SlavoGermanic(strWord) And Mid(strWord, i + 1, 1) Like "[AO]" Then
                    strPrimary = strPrimary & "J": strSecondary = strSecondary & "H"
                ElseIf i = Len(strWord) Then
                    strPrimary = strPrimary & "J": strSecondary = strSecondary & " "
                ElseIf Not StringAt(strWord, i + 1, 1, "L,T,K,S,N,M,B,Z") And Not StringAt(strWord, i - 1, 1, "S,K,L") Then
                    strPrimary = strPrimary & "J": strSecondary = strSecondary & "J"
                End If
                
                If Mid(strWord, i + 1, 1) = "J" Then i = i + 1
            Case "K":
                If Mid(strWord, i + 1, 1) = "K" Then i = i + 1
                strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
            Case "L":
                If Mid(strWord, i + 1, 1) = "L" Then
                    If (i = Len(strWord) - 2 And StringAt(strWord, i - 1, 4, "ILLO,ILLA,ALLE")) Or ((right(strWord, 2) Like "[AO]S" Or right(strWord, 1) Like "[AO]") And StringAt(strWord, i - 1, 4, "ALLE")) Then
                        strPrimary = strPrimary & "L": strSecondary = strSecondary & ""
                        i = i + 1
                     Else
                        strPrimary = strPrimary & "L": strSecondary = strSecondary & "L"
                        i = i + 1
                     End If
                Else
                    strPrimary = strPrimary & "L": strSecondary = strSecondary & "L"
                End If
            Case "M":
                If (StringAt(strWord, i - 1, 3, "UMB") And (i + 1 = Len(strWord) Or StringAt(strWord, i + 2, 2, "ER"))) Or Mid(strWord, i + 1, 1) = "M" Then i = i + 1
                strPrimary = strPrimary & "M": strSecondary = strSecondary & "M"
            Case "N":
                If Mid(strWord, i + 1, 1) = "N" Then i = i + 1
                strPrimary = strPrimary & "N": strSecondary = strSecondary & "N"
            Case "Ñ":
                strPrimary = strPrimary & "N": strSecondary = strSecondary & "N"
            Case "P":
                If Mid(strWord, i + 1, 1) = "H" Then
                     strPrimary = strPrimary & "F": strSecondary = strSecondary & "F"
                     i = i + 1
                ElseIf Mid(strWord, i + 1, 1) Like "[PB]" Then
                     strPrimary = strPrimary & "P": strSecondary = strSecondary & "P"
                     i = i + 1
                Else
                     strPrimary = strPrimary & "P": strSecondary = strSecondary & "P"
                End If
            Case "Q":
                If Mid(strWord, i + 1, 1) = "Q" Then i = i + 1
                strPrimary = strPrimary & "K": strSecondary = strSecondary & "K"
            Case "R":
                If i = Len(strWord) And Not SlavoGermanic(strWord) And StringAt(strWord, i - 2, 2, "IE") And Not StringAt(strWord, i - 4, 2, "ME,MA") Then
                    strPrimary = strPrimary & "": strSecondary = strSecondary & "R"
                Else
                    strPrimary = strPrimary & "R": strSecondary = strSecondary & "R"
                End If
                
                If Mid(strWord, i + 1, 1) = "R" Then i = i + 1
            Case "S":
                If StringAt(strWord, i - 1, 3, "ISL,YSL") Then
                    'Do Nothing
                ElseIf i = 1 And StringAt(strWord, i, 5, "SUGAR") Then
                    strPrimary = strPrimary & "X": strSecondary = strSecondary & "S"
                ElseIf StringAt(strWord, i, 2, "SH") Then
                    If StringAt(strWord, i + 1, 4, "HEIM,HOEK,HOLM,HOLZ") Then
                        strPrimary = strPrimary & "S": strSecondary = strSecondary & "S"
                    Else
                        strPrimary = strPrimary & "X": strSecondary = strSecondary & "X"
                    End If
                    i = i + 1
                ElseIf StringAt(strWord, i, 3, "SIO,SIA") Or StringAt(strWord, i, 4, "SIAN") Then
                    If Not SlavoGermanic(strWord) Then
                        strPrimary = strPrimary & "S": strSecondary = strSecondary & "X"
                    Else
                        strPrimary = strPrimary & "S": strSecondary = strSecondary & "S"
                    End If
                ElseIf (i = 1) And StringAt(strWord, i + 1, 1, "M,N,L,W") Or StringAt(strWord, i + 1, 1, "Z") Then
                    strPrimary = strPrimary & "S": strSecondary = strSecondary & "X"
                    If StringAt(strWord, i + 1, 1, "Z") Then i = i + 1
                ElseIf StringAt(strWord, i, 2, "SC") Then
                    If Mid(strWord, i + 2, 1) = "H" Then
                        If StringAt(strWord, i + 3, 2, "OO,ER,EN,UY,ED,EM") Then
                            If StringAt(strWord, i + 3, 2, "ER,EN") Then
                                strPrimary = strPrimary & "X": strSecondary = strSecondary & "SK"
                            Else
                                strPrimary = strPrimary & "SK": strSecondary = strSecondary & "SK"
                            End If
                            i = i + 2
                        Else
                            If i = 1 And Not IsVowel(strWord, 4) And Mid(strWord, 4, 1) <> "W" Then
                                strPrimary = strPrimary & "X": strSecondary = strSecondary & "S"
                            Else
                                strPrimary = strPrimary & "X": strSecondary = strSecondary & "X"
                            End If
                            i = i + 1
                        End If
                    ElseIf StringAt(strWord, i + 2, 1, "I,E,Y") Then
                        strPrimary = strPrimary & "S": strSecondary = strSecondary & "S"
                        i = i + 2
                    Else
                        strPrimary = strPrimary & "SK": strSecondary = strSecondary & "SK"
                        i = i + 2
                    End If
                Else
                    If i = Len(strWord) And StringAt(strWord, i - 2, 2, ("AI,OI")) Then
                        strPrimary = strPrimary & "": strSecondary = strSecondary & "S"
                    Else
                        strPrimary = strPrimary & "S": strSecondary = strSecondary & "S"
                    End If
                    
                    If StringAt(strWord, i + 1, 1, "S,Z") Then
                        i = i + 1
                    End If
                End If
            Case "T":
                If StringAt(strWord, i, 4, "TION") Then
                    strPrimary = strPrimary & "X": strSecondary = strSecondary & "X"
                    i = i + 2
                ElseIf StringAt(strWord, i, 3, "TIA,TCH") Then
                    strPrimary = strPrimary & "X": strSecondary = strSecondary & "X"
                    i = i + 2
                ElseIf StringAt(strWord, i, 2, "TH") Or StringAt(strWord, i, 3, "TTH") Then
                    If StringAt(strWord, i + 2, 2, "OM,AM") Or StringAt(strWord, 1, 4, "VAN ,VON ") Or StringAt(strWord, 1, 3, "SCH") Then
                        strPrimary = strPrimary & "T": strSecondary = strSecondary & "T"
                    Else
                        strPrimary = strPrimary & "0": strSecondary = strSecondary & "T"
                    End If
                    i = i + 1
                ElseIf StringAt(strWord, i + 1, 1, "T,D") Then
                    i = i + 1
                    strPrimary = strPrimary & "T": strSecondary = strSecondary & "T"
                Else
                    strPrimary = strPrimary & "T": strSecondary = strSecondary & "T"
                End If
            Case "V":
                If Mid(strWord, i + 1, 1) = "V" Then i = i + 1
                strPrimary = strPrimary & "F": strSecondary = strSecondary & "F"
            Case "W":
                If StringAt(strWord, i, 2, "WR") Then
                    strPrimary = strPrimary & "R": strSecondary = strSecondary & "R"
                    i = i + 1
                ElseIf i = 1 And (IsVowel(strWord, i + 1) Or StringAt(strWord, i, 2, "WH")) Then
                    If IsVowel(strWord, i + 1) Then
                        strPrimary = strPrimary & "A": strSecondary = strSecondary & "F"
                    Else
                        strPrimary = strPrimary & "A": strSecondary = strSecondary & "A"
                    End If
                End If
                
                If (i = Len(strWord) And IsVowel(strWord, i - 1)) Or StringAt(strWord, i - 1, 5, "EWSKI,EWSKY,OWSKI,OWSKY") Or StringAt(strWord, 1, 3, "SCH") Then
                    strPrimary = strPrimary & "": strSecondary = strSecondary & "F"
                ElseIf StringAt(strWord, i, 4, "WICZ,WITZ") Then
                    strPrimary = strPrimary & "TS": strSecondary = strSecondary & "FX"
                    i = i + 3
                End If
            Case "X":
                If Not ((i = Len(strWord) And (StringAt(strWord, i - 3, 3, "IAU,EAU") Or StringAt(strWord, i - 2, 2, "AU,OU")))) Then
                  strPrimary = strPrimary & "KS": strSecondary = strSecondary & "KS"
                End If
                If Mid(strWord, i + 1, 1) Like "[CX]" Then i = i + 1
            Case "Z":
                If Mid(strWord, i + 1, 1) = "H" Then
                    strPrimary = strPrimary & "J": strSecondary = strSecondary & "J"
                    i = i + 1
                ElseIf StringAt(strWord, i + 1, 2, "ZO,ZI,ZA") Or (SlavoGermanic(strWord) And i > 1 And Not StringAt(strWord, i - 1, 1, "T")) Then
                    strPrimary = strPrimary & "S": strSecondary = strSecondary & "TS"
                Else
                    strPrimary = strPrimary & "S": strSecondary = strSecondary & "S"
                End If
                
                If Mid(strWord, i + 1, 1) = "Z" Then i = i + 1
        End Select

    Next
    
    If intMaxLength > 0 Then
        strPrimary = left(strPrimary, intMaxLength)
        strSecondary = left(strSecondary, intMaxLength)
    End If
    
    If strPrimary = strSecondary Then strSecondary = vbNullString
    
    DoubleMetaphone = IIf(strSecondary = vbNullString, strPrimary, strPrimary & "," & strSecondary)
End Function

Function SlavoGermanic(strWord As String) As Boolean
    SlavoGermanic = InStr(strWord, "W") > 0 Or InStr(strWord, "K") > 0 Or InStr(strWord, "CZ") > 0
End Function
Function IsVowel(strWord As String, intPos As Integer) As Boolean
    If intPos >= 1 Then
        IsVowel = Mid(strWord, intPos, 1) Like "[AEIOUY]"
    Else
        IsVowel = False
    End If
End Function
Function StringAt(ByVal strWord As String, intPos As Integer, intLen As Integer, strSubstrings As String) As Boolean
    strWord = strWord & "       "
    If intPos < 1 Then
        StringAt = False
        Exit Function
    End If
    
    Dim substr As Variant
    For Each substr In Split(strSubstrings, ",")
        If Mid(strWord, intPos, intLen) = substr Then
            StringAt = True
            Exit Function
        End If
    Next
    
    StringAt = False
End Function
