Attribute VB_Name = "Encode_RethSchek"
Option Explicit
Function RethSchek(strWord As String) As String
    strWord = UCase$(strWord)
    
    If InStr(strWord, "Ä") > 0 Then strWord = Replace$(strWord, "Ä", "AE")
    If InStr(strWord, "Ö") > 0 Then strWord = Replace$(strWord, "Ö", "OE")
    If InStr(strWord, "Ü") > 0 Then strWord = Replace$(strWord, "Ü", "UE")
    
    Dim i As Long
    i = 1
    
    Dim matched As Boolean
    Do While i < Len(strWord)
        matched = False
        
        If i + 2 <= Len(strWord) Then
            Select Case Mid$(strWord, i, 3)
                Case "AEH": strWord = left$(strWord, i - 1) & "E" & Mid$(strWord, i + 3): i = i + 1: matched = True
                Case "IEH": strWord = left$(strWord, i - 1) & "I" & Mid$(strWord, i + 3): i = i + 1: matched = True
                Case "OEH": strWord = left$(strWord, i - 1) & "OE" & Mid$(strWord, i + 3): i = i + 1: matched = True
                Case "UEH": strWord = left$(strWord, i - 1) & "UE" & Mid$(strWord, i + 3): i = i + 1: matched = True
                Case "SCH": strWord = left$(strWord, i - 1) & "CH" & Mid$(strWord, i + 3): i = i + 1: matched = True
                Case "ZIO", "TIU", "ZIU": strWord = left$(strWord, i - 1) & "TIO" & Mid$(strWord, i + 3): i = i + 1: matched = True
                Case "CHS", "CKS": strWord = left$(strWord, i - 1) & "X" & Mid$(strWord, i + 3): i = i + 1: matched = True
                Case "AEU": strWord = left$(strWord, i - 1) & "OI" & Mid$(strWord, i + 3): i = i + 1: matched = True
            End Select
        End If
        
        If i + 1 <= Len(strWord) And Not matched Then
            Select Case Mid$(strWord, i, 2)
                Case "LL": strWord = left$(strWord, i - 1) & "L" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "AA", "AH": strWord = left$(strWord, i - 1) & "A" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "BB", "PP", "BP", "PB": strWord = left$(strWord, i - 1) & "B" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "DD", "DT", "TT", "TH": strWord = left$(strWord, i - 1) & "D" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "EE", "EH", "AE": strWord = left$(strWord, i - 1) & "E" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "FF", "PH": strWord = left$(strWord, i - 1) & "F" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "KK": strWord = left$(strWord, i - 1) & "K" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "GG", "GK", "KG", "CK": strWord = left$(strWord, i - 1) & "G" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "CC": strWord = left$(strWord, i - 1) & "C" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "IE", "IH": strWord = left$(strWord, i - 1) & "I" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "MM": strWord = left$(strWord, i - 1) & "M" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "NN": strWord = left$(strWord, i - 1) & "N" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "OO", "OH": strWord = left$(strWord, i - 1) & "O" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "SZ": strWord = left$(strWord, i - 1) & "S" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "UH": strWord = left$(strWord, i - 1) & "U" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "GS", "KS": strWord = left$(strWord, i - 1) & "X" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "TZ": strWord = left$(strWord, i - 1) & "Z" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "AY", "EI", "EY": strWord = left$(strWord, i - 1) & "AI" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "EU": strWord = left$(strWord, i - 1) & "OI" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "RR": strWord = left$(strWord, i - 1) & "R" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "SS": strWord = left$(strWord, i - 1) & "S" & Mid$(strWord, i + 2): i = i + 1: matched = True
                Case "KQ": strWord = left$(strWord, i - 1) & "QU" & Mid$(strWord, i + 2): i = i + 1: matched = True
            End Select
        End If
        
        If i <= Len(strWord) And Not matched Then
            Select Case Mid$(strWord, i, 1)
                Case "P": strWord = left$(strWord, i - 1) & "B" & Mid$(strWord, i + 1): i = i + 1: matched = True
                Case "T": strWord = left$(strWord, i - 1) & "D" & Mid$(strWord, i + 1): i = i + 1: matched = True
                Case "V", "W": strWord = left$(strWord, i - 1) & "F" & Mid$(strWord, i + 1): i = i + 1: matched = True
                Case "C", "K": strWord = left$(strWord, i - 1) & "G" & Mid$(strWord, i + 1): i = i + 1: matched = True
                Case "Y": strWord = left$(strWord, i - 1) & "I" & Mid$(strWord, i + 1): i = i + 1: matched = True
            End Select
        End If

        If Not matched Then i = i + 1
    Loop
        
    If InStr(strWord, "CH") > 0 Then strWord = Replace$(strWord, "CH", "SCH")
    
    If right$(strWord, 2) = "ER" Then
        strWord = left$(strWord, Len(strWord) - 2) & "R"
    ElseIf right(strWord, 2) = "EL" Then
        strWord = left$(strWord, Len(strWord) - 2) & "L"
    ElseIf right$(strWord, 1) = "H" Then
        strWord = left$(strWord, Len(strWord) - 1)
    End If
    RethSchek = strWord
End Function
