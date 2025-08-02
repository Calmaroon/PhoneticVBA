Attribute VB_Name = "Stem_Caumanns"
Option Explicit
Function Caumanns(strWord As String) As String
    If Len(strWord) = 0 Then
        Exit Function
    End If
    
    Dim boolUpperInitial As Boolean
    boolUpperInitial = StrComp(left(strWord, 1), UCase(left(strWord, 1)), vbBinaryCompare) = 0
    
    '1. Substitution
    strWord = LCase$(strWord)
    strWord = Replace(strWord, "ä", "a")
    strWord = Replace(strWord, "ö", "o")
    strWord = Replace(strWord, "ü", "u")
    strWord = Replace(strWord, "ß", "ss")
    
    '2. Change second of doubled characters to *
    Dim strNewWord As String
    strNewWord = left(strWord, 1)
    
    Dim i As Integer
    For i = 2 To Len(strWord)
        If Mid(strNewWord, i - 1, 1) = Mid(strWord, i, 1) Then
            strNewWord = strNewWord & "*"
        Else
            strNewWord = strNewWord & Mid(strWord, i, 1)
        End If
    Next
    strWord = strNewWord
    
    '3. Replace sch, ch, ei, ie with $, §, %, &
    strWord = Replace(strWord, "sch", "$")
    strWord = Replace(strWord, "ch", "§")
    strWord = Replace(strWord, "ei", "%")
    strWord = Replace(strWord, "ie", "&")
    strWord = Replace(strWord, "ig", "#")
    strWord = Replace(strWord, "st", "!")
    
    Do While Len(strWord) > 3
        If (Len(strWord) > 4 And right(strWord, 2) Like "e[mr]") Or Len(strWord) > 5 And right(strWord, 2) = "nd" Then
            strWord = left(strWord, Len(strWord) - 2)
        ElseIf right(strWord, 1) Like "[ens]" Or (Not boolUpperInitial And right(strWord, 1) Like "[t!]") Then
            strWord = left(strWord, Len(strWord) - 1)
        Else
            Exit Do
        End If
    Loop
    
    If Len(strWord) > 5 And right(strWord, 5) = "erin*" Then strWord = left(strWord, Len(strWord) - 1)
    If right(strWord, 1) = "Z" Then strWord = left(strWord, Len(strWord) - 1) & "x"
    
    strWord = Replace(strWord, "$", "sch")
    strWord = Replace(strWord, "§", "ch")
    strWord = Replace(strWord, "%", "ei")
    strWord = Replace(strWord, "&", "ie")
    strWord = Replace(strWord, "#", "ig")
    strWord = Replace(strWord, "!", "st")
    
    strNewWord = left(strWord, 1)
    For i = 2 To Len(strWord)
        If Mid(strWord, i, 1) = "*" Then
            strNewWord = strNewWord & Mid(strWord, i - 1, 1)
        Else
            strNewWord = strNewWord & Mid(strWord, i, 1)
        End If
    Next
    
    strWord = strNewWord
    
    If Len(strWord) > 4 Then
        strWord = Replace(strWord, "gege", "ge", , 1)
    End If
    
    Caumanns = strWord
End Function
