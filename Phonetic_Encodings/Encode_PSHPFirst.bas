Attribute VB_Name = "Encode_PSHPFirst"
Option Explicit
Const StrTransIn As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const StrTransOut As String = "01230120022455012523010202"

Function PSHPSoundexFirst(strInput As String, Optional intMaxLength As Integer = 4, Optional boolGerman As Boolean = False) As String
    strInput = UCase$(strInput)
    Dim strInputAlphaOnly As String
    Dim i As Long
    For i = 1 To Len(strInput)
        If Mid(strInput, i, 1) Like "[A-Z]" Then strInputAlphaOnly = strInputAlphaOnly & Mid(strInput, i, 1)
    Next
    strInput = strInputAlphaOnly
    
    Dim strCode As String
    
    If strInput = "JAMES" Then
        strCode = "J7"
    ElseIf strInput = "PAT" Then
        strCode = "P7"
    Else
        If Left(strInput, 2) Like "G[EIY]" Then
            strInput = "J" & Mid(strInput, 2)
        ElseIf Left(strInput, 2) Like "C[EIY]" Then
            strInput = "S" & Mid(strInput, 2)
        ElseIf Left(strInput, 3) = "CHR" Then
            strInput = "K" & Mid(strInput, 2)
        ElseIf Left(strInput, 1) = "C" And Left(strInput, 2) <> "CH" Then
            strInput = "K" & Mid(strInput, 2)
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
        
        strCode = Left(strInput, 1)
        For i = 1 To Len(strInput)
            Mid(strInput, i, 1) = Mid(StrTransOut, InStr(StrTransIn, Mid(strInput, i, 1)), 1)
        Next
    
        strInput = DeleteConsecutiveRepeats(strInput)
        
        strCode = strCode & Mid(strInput, 2)

        Dim syl_ptr As Long: Dim syl2_ptr As Long
        syl_ptr = InStr(strCode, "0")
        If syl_ptr <> 0 Then
            syl2_ptr = InStr(syl_ptr + 1, strCode, "0")
            If syl_ptr <> 0 And syl2_ptr <> 0 And syl2_ptr - syl_ptr > -1 Then strCode = Left(strCode, syl_ptr + 1)
        End If
        
        strCode = Replace(strCode, "0", "")
        If intMaxLength > 0 Then
            If Len(strCode) < intMaxLength Then
               strCode = strCode & String(intMaxLength - Len(strCode), "0")
            Else
                strCode = Left(strCode, intMaxLength)
            End If
        End If
    End If
    
    PSHPSoundexFirst = strCode
End Function
