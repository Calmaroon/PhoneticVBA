Attribute VB_Name = "Encode_PSHPLast"
Option Explicit
Const StrTransIn As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const StrTransOut As String = "01230120022455012523010202"
Function PSHPLast(strInput As String, Optional intMaxLength As Integer = 4, Optional boolGerman As Boolean = False) As String
    strInput = UCase(strInput)
    Dim strInputAlphaOnly As String
    Dim i As Long
    For i = 1 To Len(strInput)
        If Mid(strInput, i, 1) Like "[A-Z]" Then strInputAlphaOnly = strInputAlphaOnly & Mid(strInput, i, 1)
    Next
    strInput = strInputAlphaOnly
    
    If Left(strInput, 3) Like "V[AO]N" Then strInput = Mid(strInput, 4)

    If Not boolGerman Then
        If Left(strInput, 3) = "MAC" Then
            strInput = "M" & Mid(strInput, 4)
        ElseIf Left(strInput, 2) = "MC" Then
            strInput = "M" & Mid(strInput, 3)
        End If
    End If
    
    If Left(strInput, 1) Like "[EIOU]" Then
        strInput = "A" & Mid(strInput, 2)
    ElseIf Left(strInput, 2) Like "G[EIY]" Then
        strInput = "J" & Mid(strInput, 2)
    ElseIf Left(strInput, 2) Like "C[EIY]" Then
        strInput = "S" & Mid(strInput, 2)
    ElseIf Left(strInput, 3) = "CHR" Then
        strInput = "S" & Mid(strInput, 2)
    ElseIf Left(strInput, 1) = "C" And Left(strInput, 2) <> "CH" Then
        strInput = "S" & Mid(strInput, 2)
    End If
    
    If Left(strInput, 2) = "KN" Then
        strInput = "N" & Mid(strInput, 2)
    ElseIf Left(strInput, 2) = "PH" Then
        strInput = "F" & Mid(strInput, 2)
    ElseIf Left(strInput, 3) = "WIE" Or Left(strInput, 3) = "WEI" Then
        strInput = "V" & Mid(strInput, 2)
    End If
    
    If boolGerman And Left(strInput, 1) Like "[WMYZ]" Then
        If Left(strInput, 1) = "W" Then strInput = "V" & Mid(strInput, 2)
        If Left(strInput, 1) = "M" Then strInput = "N" & Mid(strInput, 2)
        If Left(strInput, 1) = "Y" Then strInput = "J" & Mid(strInput, 2)
        If Left(strInput, 1) = "Z" Then strInput = "S" & Mid(strInput, 2)
    End If
    
    Dim strCode As String
    strCode = Left(strInput, 1)
    
    If boolGerman Then
        If Right(strInput, 3) = "TES" Then
            strInput = Left(strInput, Len(strInput) - 3)
        ElseIf Right(strInput, 2) = "TS" Then
            strInput = Left(strInput, Len(strInput) - 2)
        End If
        
        If Right(strInput, 3) = "TZE" Then
            strInput = Left(strInput, Len(strInput) - 3)
        ElseIf Right(strInput, 2) = "ZE" Then
            strInput = Left(strInput, Len(strInput) - 2)
        End If
        
        If Right(strInput, 1) = "Z" Then
            strInput = Left(strInput, Len(strInput) - 1)
        ElseIf Right(strInput, 2) = "TE" Then
            strInput = Left(strInput, Len(strInput) - 2)
        End If
    End If
    
    If Right(strInput, 1) = "R" Then
        strInput = Left(strInput, Len(strInput) - 1) & "N"
    ElseIf Right(strInput, 2) Like "[SC]E" Then
        strInput = Left(strInput, Len(strInput) - 2)
    End If
    
    If Right(strInput, 2) = "SS" Then
        strInput = Left(strInput, Len(strInput) - 2)
    ElseIf Right(strInput, 1) = "S" Then
        strInput = Left(strInput, Len(strInput) - 1)
    End If
    
    If Not boolGerman Then
        If Right(strInput, 5) = "STOWN" Then strInput = Left(strInput, Len(strInput) - 5) & "SAWON"
        If Right(strInput, 5) = "MPSON" Then strInput = Left(strInput, Len(strInput) - 5) & "MASON"
        If Right(strInput, 4) = "NSEN" Then strInput = Left(strInput, Len(strInput) - 4) & "ASEN"
        If Right(strInput, 4) = "MSON" Then strInput = Left(strInput, Len(strInput) - 4) & "ASON"
        If Right(strInput, 4) = "STEN" Then strInput = Left(strInput, Len(strInput) - 4) & "SAEN"
        If Right(strInput, 4) = "STON" Then strInput = Left(strInput, Len(strInput) - 4) & "SAON"
    End If
    
    If Left(strInput, 1) Like "N[GD]" Then strInput = Left(strInput, Len(strInput) - 1)
    
    If Not boolGerman And Right(strInput, 3) Like "G[EA]N" Then strInput = Left(strInput, Len(strInput) - 3) & "A" & Right(strInput, 2)
    
    strInput = Replace(strInput, "CK", "C")
    strInput = Replace(strInput, "SCH", "S")
    strInput = Replace(strInput, "DT", "T")
    strInput = Replace(strInput, "ND", "N")
    strInput = Replace(strInput, "NG", "N")
    strInput = Replace(strInput, "LM", "M")
    strInput = Replace(strInput, "MN", "M")
    strInput = Replace(strInput, "WIE", "VIE")
    strInput = Replace(strInput, "WEI", "VEI")
    
    For i = 1 To Len(strInput)
        Mid(strInput, i, 1) = Mid(StrTransOut, InStr(StrTransIn, Mid(strInput, i, 1)), 1)
    Next
    
    strInput = DeleteConsecutiveRepeats(strInput)
    
    strCode = strCode & Mid(strInput, 2)
    strCode = Replace(strCode, "0", "")
    
    If intMaxLength > 0 Then
        If Len(strCode) < intMaxLength Then
            strCode = strCode & String(intMaxLength - Len(strCode), "0")
        Else
            strCode = Left(strCode, intMaxLength)
        End If
    End If
    PSHPLast = strCode
End Function
