Attribute VB_Name = "Encode_SoundD"
Option Explicit
Const strTranscodeIn = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const strTranscodeOut = "01230120022455012623010202"

Function SoundD(strWord As String, Optional intMaxLength As Integer = 4) As String
    strWord = UCase$(UnicodeStrip(strWord))
    
    Dim strWordAlpha As String
    Dim i As Long
    
    For i = 1 To Len(strWord)
        If Mid(strWord, i, 1) Like "[A-Z]" Then
            strWordAlpha = strWordAlpha & Mid(strWord, i, 1)
        End If
    Next
    
    Select Case Left(strWordAlpha, 2)
        Case "KN", "GN", "PN", "AC", "WR"
            strWordAlpha = Mid$(strWordAlpha, 2)
    End Select
    
    If Left(strWordAlpha, 1) = "X" Then
        strWordAlpha = "S" & Mid$(strWordAlpha, 2)
    ElseIf Left(strWordAlpha, 2) = "WH" Then
        strWordAlpha = "W" & Mid$(strWordAlpha, 3)
    End If

    strWordAlpha = Replace(Replace(Replace(strWordAlpha, "DGE", "20"), "DGI", "20"), "GH", "0")
    
    For i = 1 To Len(strWordAlpha)
        If InStr(strTranscodeIn, Mid$(strWordAlpha, i, 1)) > 0 Then
            Mid(strWordAlpha, i, 1) = Mid$(strTranscodeOut, InStr(strTranscodeIn, Mid$(strWordAlpha, i, 1)), 1)
        End If
    Next
    
    strWordAlpha = DeleteConsecutiveRepeats(strWordAlpha)
    strWordAlpha = Replace(strWordAlpha, "0", "")
    
    If intMaxLength <> -1 Then
        If Len(strWordAlpha) < intMaxLength Then
            strWordAlpha = strWordAlpha & String(intMaxLength - Len(strWordAlpha), "0")
        Else
            strWordAlpha = Left(strWordAlpha, intMaxLength)
        End If
    End If
    SoundD = strWordAlpha

End Function
