Attribute VB_Name = "Encode_FuzzySoundex"
Option Explicit
Const strTranscodeIn = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const strTranscodeOut = "0193017-07745501769301-7-9"
Function FuzzySoundex(strWord As String, Optional intMaxLength As Integer = 5, Optional boolZeroPad As Boolean = True) As String
    'Abydos documentation issue, states max_length defaults to 4 but it defaults to 5
    If intMaxLength <> -1 Then
        If intMaxLength > 64 Then intMaxLength = 64
        If intMaxLength < 4 Then intMaxLength = 4
    Else
        intMaxLength = 64
    End If
    
    
    strWord = UnicodeStrip(strWord)
    strWord = UCase$(strWord)

    Dim strWordOriginal As String
    strWordOriginal = strWord

    If Len(strWord) = 0 Then
        If boolZeroPad Then
            FuzzySoundex = String(intMaxLength, "0")
        Else
            FuzzySoundex = "0"
        End If
        Exit Function
    End If

    Select Case Left(strWord, 2)
        Case "CS", "CZ", "TS", "TZ":
            strWord = "SS" & Mid(strWord, 3)
        Case "GN":
            strWord = "NN" & Mid(strWord, 3)
        Case "HR", "WR":
            strWord = "RR" & Mid(strWord, 3)
        Case "HW":
            strWord = "WW" & Mid(strWord, 3)
        Case "KN", "NG":
            strWord = "NN" & Mid(strWord, 3)
    End Select

    If Right(strWord, 2) = "CH" Then
        strWord = Left(strWord, Len(strWord) - 2) & "KK"
    ElseIf Right(strWord, 2) = "NT" Then
        strWord = Left(strWord, Len(strWord) - 2) & "TT"
    ElseIf Right(strWord, 2) = "RT" Then
        strWord = Left(strWord, Len(strWord) - 2) & "RR"
    ElseIf Right(strWord, 3) = "RDT" Then
        strWord = Left(strWord, Len(strWord) - 3) & "RR"
    End If
    
    strWord = Replace(strWord, "CA", "KA")
    strWord = Replace(strWord, "CC", "KK")
    strWord = Replace(strWord, "CK", "KK")
    strWord = Replace(strWord, "CE", "SE")
    strWord = Replace(strWord, "CHL", "KL")
    strWord = Replace(strWord, "CL", "KL")
    strWord = Replace(strWord, "CHR", "KR")
    strWord = Replace(strWord, "CR", "KR")
    strWord = Replace(strWord, "CI", "SI")
    strWord = Replace(strWord, "CO", "KO")
    strWord = Replace(strWord, "CU", "KU")
    strWord = Replace(strWord, "CY", "SY")
    strWord = Replace(strWord, "DG", "GG")
    strWord = Replace(strWord, "GH", "HH")
    strWord = Replace(strWord, "MAC", "MK")
    strWord = Replace(strWord, "MC", "MK")
    strWord = Replace(strWord, "NST", "NSS")
    strWord = Replace(strWord, "PF", "FF")
    strWord = Replace(strWord, "PH", "FF")
    strWord = Replace(strWord, "SCH", "SSS")
    strWord = Replace(strWord, "TIO", "SIO")
    strWord = Replace(strWord, "TIA", "SIO")
    strWord = Replace(strWord, "TCH", "CHH")
    
    Dim i As Long

    For i = 1 To Len(strWord)
        If InStr(strTranscodeIn, Mid(strWord, i, 1)) > 0 Then
            Mid(strWord, i, 1) = Mid(strTranscodeOut, InStr(strTranscodeIn, Mid(strWord, i, 1)), 1)
        End If
    Next
    
    strWord = Replace(strWord, "-", "")
    strWord = DeleteConsecutiveRepeats(strWord)
    
    If Left(strWordOriginal, 1) Like "[HWY]" Then
        strWord = Left(strWordOriginal, 1) & strWord
    Else
       strWord = Left(strWordOriginal, 1) & Mid(strWord, 2)
    End If
    
    strWord = Replace(strWord, "0", "")
    
    If boolZeroPad Then
       strWord = strWord & String(intMaxLength, "0")
    End If
       
    
    FuzzySoundex = Left(strWord, intMaxLength)
End Function
