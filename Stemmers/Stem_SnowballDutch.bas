Attribute VB_Name = "Stem_SnowballDutch"
Option Explicit
Function SnowballDutch(strWord As String) As String
    strWord = LCase(strWord)
    Dim i As Integer
    For i = 1 To Len(strWord)
        If InStr("äëïöüáéíóú", Mid(strWord, i, 1)) > 0 Then
            Mid(strWord, i, 1) = Mid("aeiouaeiou", InStr("äëïöüáéíóú", Mid(strWord, i, 1)), 1)
        End If
    Next
    
    If Left(strWord, 1) = "y" Then strWord = "Y" & Mid(strWord, 2)
    
    For i = 2 To Len(strWord)
        If Mid(strWord, i, 1) = "y" And Mid(strWord, i - 1, 1) Like "[aeiouyè]" Then
            strWord = Left(strWord, i - 1) & "Y" & Mid(strWord, i + 1)
        ElseIf Mid(strWord, i, 1) = "i" And Mid(strWord, i - 1, 1) Like "[aeiouyè]" And i + 1 <= Len(strWord) And Mid(strWord, i + 1, 1) Like "[aeiouyè]" Then
            strWord = Left(strWord, i - 1) & "I" & Mid(strWord, i + 1)
        End If
    Next
    
    Dim r1Prefix() As String: ReDim r1Prefix(0)
    Dim r1Start As Integer: r1Start = sbR1(strWord, r1Prefix())
    If r1Start < 3 Then r1Start = 3
    Dim r2Start As Integer: r2Start = sbR2(strWord, r1Prefix())
    
    'Step 1
    If Right(strWord, 5) = "heden" Then
        If Len(strWord) - r1Start >= 5 Then
            strWord = Left(strWord, Len(strWord) - 3) & "id"
        End If
    ElseIf Right(strWord, 3) = "ene" Then
        If Len(strWord) >= 6 Then
            If Len(strWord) - r1Start >= 3 And (Not Mid(strWord, Len(strWord) - 3, 1) Like "[aeiouyè]" And Mid(strWord, Len(strWord) - 5, 3) <> "gem") Then
                strWord = undouble(Left(strWord, Len(strWord) - 3))
            End If
        ElseIf Len(strWord) > 3 Then
            If Len(strWord) - r1Start >= 3 And Not Mid(strWord, Len(strWord) - 3, 1) Like "[aeiouyè]" Then
                strWord = undouble(Left(strWord, Len(strWord) - 3))
            End If
        End If
    ElseIf Right(strWord, 2) = "en" Then
        If Len(strWord) > 5 Then
            If Len(strWord) - r1Start >= 2 And (Not Mid(strWord, Len(strWord) - 2, 1) Like "[aeiouyè]" And Mid(strWord, Len(strWord) - 4, 3) <> "gem") Then
                strWord = undouble(Left(strWord, Len(strWord) - 2))
            End If
        ElseIf Len(strWord) > 3 Then
            If Len(strWord) - r1Start >= 2 And (Not Mid(strWord, Len(strWord) - 2, 1) Like "[aeiouyè]") Then
                strWord = undouble(Left(strWord, Len(strWord) - 2))
            End If
        End If
    ElseIf Right(strWord, 2) = "se" Then
        If Len(strWord) > 2 Then
            If Len(strWord) - r1Start >= 2 And Not Mid(strWord, Len(strWord) - 2, 1) Like "[aeijouyè]" Then
                 strWord = Left(strWord, Len(strWord) - 2)
            End If
        End If
    ElseIf Right(strWord, 1) = "s" Then
        If Len(strWord) > 1 Then
            If Len(strWord) - r1Start >= 1 And Not Mid(strWord, Len(strWord) - 1, 1) Like "[aeijouyè]" Then
                strWord = Left(strWord, Len(strWord) - 1)
            End If
        End If
    End If
    
    'Step 2
    Dim boolERemoved As Boolean
    If Right(strWord, 1) = "e" And Len(strWord) > 1 Then
        If Len(strWord) - r1Start >= 1 And Not Mid(strWord, Len(strWord) - 1, 1) Like "[aeiouyè]" Then
            strWord = undouble(Left(strWord, Len(strWord) - 1))
            boolERemoved = True
        End If
    End If
    
    'Step 3a
    If Right(strWord, 4) = "heid" And Len(strWord) > 4 Then
        If Len(strWord) - r2Start >= 4 And Mid(strWord, Len(strWord) - 4, 1) <> "c" Then
            strWord = Left(strWord, Len(strWord) - 4)
            If Right(strWord, 2) = "en" And Len(strWord) > 5 Then
                If Len(strWord) - r1Start >= 2 And (Not Mid(strWord, Len(strWord) - 2, 1) Like "[aeijouyè]" And Mid(strWord, Len(strWord) - 4, 3) <> "gem") Then
                    strWord = undouble(Left(strWord, Len(strWord) - 2))
                End If
            End If
        End If
    End If
    
    'step 3b
    If Right(strWord, 4) = "lijk" Then
        If Len(strWord) - r2Start >= 4 Then
            strWord = Left(strWord, Len(strWord) - 4)
            If Right(strWord, 1) = "e" And Len(strWord) > 2 Then
                If Len(strWord) - r1Start >= 1 And Not Mid(strWord, Len(strWord) - 1, 1) Like "[aeiouyè]" Then
                    strWord = undouble(Left(strWord, Len(strWord) - 1))
                End If
            End If
        End If
    ElseIf Right(strWord, 4) = "baar" Then
        If Len(strWord) - r2Start >= 4 Then
            strWord = Left(strWord, Len(strWord) - 4)
        End If
    ElseIf Right(strWord, 3) = "end" Or Right(strWord, 3) = "ing" Then
        If Len(strWord) - r2Start >= 3 Then
            strWord = Left(strWord, Len(strWord) - 3)
            If Len(strWord) > 2 Then
                If Right(strWord, 2) = "ig" And Len(strWord) - r2Start >= 2 And Mid(strWord, Len(strWord) - 2, 1) <> "e" Then
                    strWord = Left(strWord, Len(strWord) - 2)
                Else
                    strWord = undouble(strWord)
                End If
            End If
        End If
    ElseIf Right(strWord, 3) = "bar" Then
        If Len(strWord) - r2Start >= 3 And boolERemoved Then
            strWord = Left(strWord, Len(strWord) - 3)
        End If
    ElseIf Right(strWord, 2) = "ig" Then
        If Len(strWord) - r2Start >= 2 And Mid(strWord, Len(strWord) - 2, 1) <> "e" Then
            strWord = Left(strWord, Len(strWord) - 2)
        End If
    End If
    
    'step 4
    If Len(strWord) >= 4 Then
        If Mid(strWord, Len(strWord) - 2, 1) = Mid(strWord, Len(strWord) - 1, 1) And Mid(strWord, Len(strWord) - 1, 1) Like "[aeou]" And Not Mid(strWord, Len(strWord) - 3, 1) Like "[aeiouyè]" And Not Right(strWord, 1) Like "[aeiouyè]" And Right(strWord, 1) <> "I" Then
            strWord = Left(strWord, Len(strWord) - 2) & Right(strWord, 1)
        End If
    End If
    
    strWord = Replace(strWord, "Y", "y")
    strWord = Replace(strWord, "I", "i")
    
    SnowballDutch = strWord
End Function
Function undouble(strWord As String) As String
    If Len(strWord) > 1 Then
        If Right(strWord, 1) = Mid(strWord, Len(strWord) - 1, 1) And Right(strWord, 1) Like "[dkt]" Then
            strWord = Left(strWord, Len(strWord) - 1)
        End If
    End If
    
    undouble = strWord
End Function
Function sbR1(strTerm As String, r1Prefix() As String) As Integer
    Dim boolVowelFound As Boolean
    Dim strPrefix As String
    Dim i As Integer
    If UBound(r1Prefix) > 0 Then
        For i = LBound(r1Prefix) To UBound(r1Prefix)
            If Left$(strTerm, Len(r1Prefix(i))) = r1Prefix(i) Then
                sbR1 = Len(r1Prefix(i))
                Exit Function
            End If
        Next
    End If
    
    For i = 1 To Len(strTerm)
        If Not boolVowelFound And Mid$(strTerm, i, 1) Like "[aeiouyè]" Then
            boolVowelFound = True
        ElseIf boolVowelFound And Not Mid$(strTerm, i, 1) Like "[aeiouyè]" Then
            sbR1 = i
            Exit Function
        End If
    Next
         
    sbR1 = Len(strTerm)
End Function
Function sbR2(strTerm As String, ByRef r1Prefix() As String) As Integer
    Dim r1Start As Integer
    r1Start = sbR1(strTerm, r1Prefix())
    Dim r2Prefix() As String
    ReDim r2Prefix(0)
    sbR2 = r1Start + sbR1(Mid$(strTerm, r1Start + 1), r2Prefix)
End Function
