Attribute VB_Name = "Stem_Porter"
Option Explicit
Function Porter(strWord As String, Optional boolEarlyEnglish As Boolean = False) As String
    strWord = LCase(strWord)
    
    If Len(strWord) < 3 Then
        Porter = strWord
        Exit Function
    End If

    If left(strWord, 1) = "y" Then strWord = "Y" & Mid(strWord, 2)
    
    Dim i As Integer
    For i = 2 To Len(strWord)
        If Mid(strWord, i, 1) = "y" And Mid(strWord, i - 1, 1) Like "[aeiouy]" Then strWord = left(strWord, i - 1) & "Y" & Mid(strWord, i + 1)
    Next
    
    'step 1a
    If right(strWord, 1) = "s" Then
        If right(strWord, 4) = "sses" Then
            strWord = left(strWord, Len(strWord) - 2)
        ElseIf right(strWord, 3) = "ies" Then
            strWord = left(strWord, Len(strWord) - 2)
        ElseIf right(strWord, 2) = "ss" Then
            'Pass
        Else
            strWord = left(strWord, Len(strWord) - 1)
        End If
    End If

    Dim boolStep1bFlag As Boolean
    If right(strWord, 3) = "eed" Then
        If m_degree(left(strWord, Len(strWord) - 3)) > 0 Then
            strWord = left(strWord, Len(strWord) - 1)
        End If
    ElseIf right(strWord, 2) = "ed" Then
        If has_vowel(left(strWord, Len(strWord) - 2)) Then
            strWord = left(strWord, Len(strWord) - 2)
            boolStep1bFlag = True
        End If
    ElseIf right(strWord, 3) = "ing" Then
        If has_vowel(left(strWord, Len(strWord) - 3)) Then
            strWord = left(strWord, Len(strWord) - 3)
            boolStep1bFlag = True
        End If
    ElseIf boolEarlyEnglish Then
        If right(strWord, 3) = "est" Then
            If has_vowel(left(strWord, Len(strWord) - 3)) Then
                strWord = left(strWord, Len(strWord) - 3)
                boolStep1bFlag = True
            End If
        ElseIf right(strWord, 3) = "eth" Then
            If has_vowel(left(strWord, Len(strWord) - 3)) Then
                strWord = left(strWord, Len(strWord) - 3)
                boolStep1bFlag = True
            End If
        End If
    End If
    
    If boolStep1bFlag Then
        If right(strWord, 2) = "at" Or right(strWord, 2) = "bl" Or right(strWord, 2) = "iz" Then
            strWord = strWord & "e"
        ElseIf ends_in_doubled_cons(strWord) And Not right(strWord, 1) Like "[lsz]" Then
            strWord = left(strWord, Len(strWord) - 1)
        ElseIf m_degree(strWord) = 1 And ends_in_cvc(strWord) Then
            strWord = strWord & "e"
        End If
    End If

    'Step 1c
    If right(strWord, 1) Like "[yY]" And has_vowel(left(strWord, Len(strWord) - 1)) Then
        strWord = left(strWord, Len(strWord) - 1) & "i"
    End If

    'Step 2
    If Len(strWord) > 1 Then
        Select Case Mid(strWord, Len(strWord) - 1, 1)
            Case "a":
                If right(strWord, 7) = "ational" Then
                    If m_degree(left(strWord, Len(strWord) - 7)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 5) & "e"
                    End If
                ElseIf right(strWord, 6) = "tional" Then
                    If m_degree(left(strWord, Len(strWord) - 6)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 2)
                    End If
                End If
            Case "c":
                If right(strWord, 4) Like "[ae]nci" Then
                    If m_degree(left(strWord, Len(strWord) - 4)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 1) & "e"
                    End If
                End If
            Case "e":
                If right(strWord, 4) = "izer" Then
                    If m_degree(left(strWord, Len(strWord) - 4)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 1)
                    End If
                End If
            Case "g":
                If right(strWord, 4) = "logi" Then
                    If m_degree(left(strWord, Len(strWord) - 4)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 1)
                    End If
                End If
            Case "l":
                If right(strWord, 3) = "bli" Then
                    If m_degree(left(strWord, Len(strWord) - 3)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 1) & "e"
                    End If
                ElseIf right(strWord, 4) = "alli" Then
                    If m_degree(left(strWord, Len(strWord) - 4)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 2)
                    End If
                ElseIf right(strWord, 5) = "entli" Then
                    If m_degree(left(strWord, Len(strWord) - 5)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 2)
                    End If
                ElseIf right(strWord, 3) = "eli" Then
                    If m_degree(left(strWord, Len(strWord) - 3)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 2)
                    End If
                ElseIf right(strWord, 5) = "ousli" Then
                    If m_degree(left(strWord, Len(strWord) - 5)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 2)
                    End If
                End If
            Case "o":
                If right(strWord, 7) = "ization" Then
                    If m_degree(left(strWord, Len(strWord) - 7)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 5) & "e"
                    End If
                ElseIf right(strWord, 5) = "ation" Then
                    If m_degree(left(strWord, Len(strWord) - 5)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 3) & "e"
                    End If
                ElseIf right(strWord, 4) = "ator" Then
                    If m_degree(left(strWord, Len(strWord) - 4)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 2) & "e"
                    End If
                End If
            Case "s":
                If right(strWord, 5) = "alism" Then
                    If m_degree(left(strWord, Len(strWord) - 5)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 3)
                    End If
                ElseIf right(strWord, 7) = "iveness" Or right(strWord, 7) = "fulness" Or right(strWord, 7) = "ousness" Then
                    If m_degree(left(strWord, Len(strWord) - 7)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 4)
                    End If
                End If
            Case "t":
                If right(strWord, 5) = "aliti" Then
                    If m_degree(left(strWord, Len(strWord) - 5)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 3)
                    End If
                ElseIf right(strWord, 5) = "iviti" Then
                    If m_degree(left(strWord, Len(strWord) - 5)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 3) & "e"
                    End If
                ElseIf right(strWord, 6) = "biliti" Then
                    If m_degree(left(strWord, Len(strWord) - 6)) > 0 Then
                        strWord = left(strWord, Len(strWord) - 5) & "le"
                    End If
                End If
                
        End Select
    End If
 
    'Step 3
    If right(strWord, 5) = "icate" Then
        If m_degree(left(strWord, Len(strWord) - 5)) > 0 Then strWord = left(strWord, Len(strWord) - 3)
    ElseIf right(strWord, 5) = "ative" Then
        If m_degree(left(strWord, Len(strWord) - 5)) > 0 Then strWord = left(strWord, Len(strWord) - 5)
    ElseIf right(strWord, 5) = "alize" Or right(strWord, 5) = "iciti" Then
        If m_degree(left(strWord, Len(strWord) - 5)) > 0 Then strWord = left(strWord, Len(strWord) - 3)
    ElseIf right(strWord, 4) = "ical" Then
        If m_degree(left(strWord, Len(strWord) - 4)) > 0 Then strWord = left(strWord, Len(strWord) - 2)
    ElseIf right(strWord, 3) = "ful" Then
        If m_degree(left(strWord, Len(strWord) - 3)) > 0 Then strWord = left(strWord, Len(strWord) - 3)
    ElseIf right(strWord, 4) = "ness" Then
        If m_degree(left(strWord, Len(strWord) - 4)) > 0 Then strWord = left(strWord, Len(strWord) - 4)
    End If

    'Step 4
    If right(strWord, 2) = "al" Then
        If m_degree(left(strWord, Len(strWord) - 2)) > 1 Then strWord = left(strWord, Len(strWord) - 2)
    ElseIf right(strWord, 4) Like "[ae]nce" Then
        If m_degree(left(strWord, Len(strWord) - 4)) > 1 Then strWord = left(strWord, Len(strWord) - 4)
    ElseIf right(strWord, 2) = "er" Or right(strWord, 2) = "ic" Then
        If m_degree(left(strWord, Len(strWord) - 2)) > 1 Then strWord = left(strWord, Len(strWord) - 2)
    ElseIf right(strWord, 4) Like "[ai]ble" Then
        If m_degree(left(strWord, Len(strWord) - 4)) > 1 Then strWord = left(strWord, Len(strWord) - 4)
    ElseIf right(strWord, 3) = "ant" Then
        If m_degree(left(strWord, Len(strWord) - 3)) > 1 Then strWord = left(strWord, Len(strWord) - 3)
    ElseIf right(strWord, 5) = "ement" Then
        If m_degree(left(strWord, Len(strWord) - 5)) > 1 Then strWord = left(strWord, Len(strWord) - 5)
    ElseIf right(strWord, 4) = "ment" Then
        If m_degree(left(strWord, Len(strWord) - 4)) > 1 Then strWord = left(strWord, Len(strWord) - 4)
    ElseIf right(strWord, 3) = "ent" Then
        If m_degree(left(strWord, Len(strWord) - 3)) > 1 Then strWord = left(strWord, Len(strWord) - 3)
    ElseIf right(strWord, 4) Like "[st]ion" Then
        If m_degree(left(strWord, Len(strWord) - 3)) > 1 Then strWord = left(strWord, Len(strWord) - 3)
    ElseIf right(strWord, 2) = "ou" Then
        If m_degree(left(strWord, Len(strWord) - 2)) > 1 Then strWord = left(strWord, Len(strWord) - 2)
    ElseIf right(strWord, 3) = "ism" Or right(strWord, 3) = "ate" Or right(strWord, 3) = "iti" Or right(strWord, 3) = "ous" Or right(strWord, 3) = "ive" Or right(strWord, 3) = "ize" Then
        If m_degree(left(strWord, Len(strWord) - 3)) > 1 Then strWord = left(strWord, Len(strWord) - 3)
    End If

    'Step 5a
    If right(strWord, 1) = "e" Then
        If m_degree(left(strWord, Len(strWord) - 1)) > 1 Then
            strWord = left(strWord, Len(strWord) - 1)
        ElseIf m_degree(left(strWord, Len(strWord) - 1)) = 1 And Not ends_in_cvc(left(strWord, Len(strWord) - 1)) Then
            strWord = left(strWord, Len(strWord) - 1)
        End If
    End If

    'Step 5b
    If right(strWord, 2) = "ll" And m_degree(strWord) > 1 Then
        strWord = left(strWord, Len(strWord) - 1)
    End If

    strWord = Replace(strWord, "Y", "y")
    Porter = strWord
End Function
Function m_degree(strTerm As String) As Integer
    Dim boolLastWasVowel As Boolean
    Dim i As Integer
    For i = 1 To Len(strTerm)
        If Mid(strTerm, i, 1) Like "[aeiouy]" Then
            boolLastWasVowel = True
        Else
            If boolLastWasVowel Then m_degree = m_degree + 1
            boolLastWasVowel = False
        End If
    Next
End Function
Function has_vowel(strTerm As String) As Boolean
    Dim i As Integer
    For i = 1 To Len(strTerm)
        If Mid(strTerm, i, 1) Like "[aeiouy]" Then
            has_vowel = True
            Exit Function
        End If
    Next
End Function
Function ends_in_doubled_cons(strTerm As String) As Boolean
    If Len(strTerm) > 1 Then
         ends_in_doubled_cons = Not right(strTerm, 1) Like "[aeiouy]" And right(strTerm, 1) = Mid(strTerm, Len(strTerm) - 1, 1)
    Else
        ends_in_doubled_cons = False
    End If
End Function
Function ends_in_cvc(strTerm As String) As Boolean
    If Len(strTerm) > 2 Then
        ends_in_cvc = Not right(strTerm, 1) Like "[aeiouy]" And Mid(strTerm, Len(strTerm) - 1, 1) Like "[aeiouy]" And Not Mid(strTerm, Len(strTerm) - 2, 1) Like "[aeiouy]" And Not right(strTerm, 1) Like "[wxY]"
        
    Else
        ends_in_cvc = False
    End If
End Function
