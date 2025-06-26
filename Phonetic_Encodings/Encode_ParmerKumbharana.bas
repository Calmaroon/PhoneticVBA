Attribute VB_Name = "Encode_ParmerKumbharana"
Option Explicit
Function ParmarKumbarana(strWord As String) As String
    strWord = UCase(strWord)
    strWord = DeleteConsecutiveRepeats(strWord)
    Dim i As Long
    i = 1
    
    Dim matched As Boolean
    
    Do While i <= Len(strWord)
         matched = False
    
        If i + 3 <= Len(strWord) And Mid(strWord, i, 4) = "OUGH" Then
            strWord = Left(strWord, i - 1) & "F" & Mid(strWord, i + 4)
            i = i + 1
            matched = True
        ElseIf i + 2 <= Len(strWord) Then
            Select Case Mid(strWord, i, 3)
                Case "DGE": strWord = Left(strWord, i - 1) & "J" & Mid(strWord, i + 3): i = i + 1: matched = True
                Case "OUL": strWord = Left(strWord, i - 1) & "U" & Mid(strWord, i + 3): i = i + 1: matched = True
                Case "GHT": strWord = Left(strWord, i - 1) & "T" & Mid(strWord, i + 3): i = i + 1: matched = True
            End Select
        End If
    
        If Not matched And i + 1 <= Len(strWord) Then
            Select Case Mid(strWord, i, 2)
                Case "CE", "CI", "CY": strWord = Left(strWord, i - 1) & "S" & Mid(strWord, i + 2): i = i + 1: matched = True
                Case "GE", "GI", "GY": strWord = Left(strWord, i - 1) & "J" & Mid(strWord, i + 2): i = i + 1: matched = True
                Case "WR": strWord = Left(strWord, i - 1) & "R" & Mid(strWord, i + 2): i = i + 1: matched = True
                Case "GN", "KN", "PN": strWord = Left(strWord, i - 1) & "N" & Mid(strWord, i + 2): i = i + 1: matched = True
                Case "CK": strWord = Left(strWord, i - 1) & "K" & Mid(strWord, i + 2): i = i + 1: matched = True
                Case "SH": strWord = Left(strWord, i - 1) & "S" & Mid(strWord, i + 2): i = i + 1: matched = True
            End Select
        End If
    
        If Not matched Then i = i + 1
    Loop
    
    Dim strVowels As Variant
    Dim strTail As String
    Dim v As Variant 'vowels
    strVowels = Array("A", "E", "I", "O", "U", "Y")
    strTail = Mid(strWord, 2)
    
    For Each v In strVowels
        strTail = Replace(strTail, v, "")
    Next
    ParmarKumbarana = Left(strWord, 1) & strTail
End Function
