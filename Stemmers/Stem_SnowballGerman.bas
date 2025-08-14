Attribute VB_Name = "Stem_SnowballGerman"
Option Explicit
Function SnowballGerman(strWord As String, Optional BoolAlternateVowels As Boolean = False) As String
    strWord = LCase$(strWord)
    strWord = Replace(strWord, "ß", "ss")
    
    
    Dim i As Integer
    If Len(strWord) > 2 Then
        For i = 3 To Len(strWord)
            If Mid(strWord, i, 1) Like "[aeiouyäöü]" And Mid(strWord, i - 2, 1) Like "[aeiouyäöü]" Then
                If Mid(strWord, i - 1, 1) = "u" Then
                    strWord = Left(strWord, i - 2) & "U" & Mid(strWord, i)
                ElseIf Mid(strWord, i - 1, 1) = "y" Then
                    strWord = Left(strWord, i - 2) & "Y" & Mid(strWord, i)
                End If
            End If
        Next
    End If
    
    If BoolAlternateVowels Then
        strWord = Replace(strWord, "ae", "ä")
        strWord = Replace(strWord, "oe", "ö")
        strWord = Replace(strWord, "que", "Q")
        strWord = Replace(strWord, "ue", "ü")
        strWord = Replace(strWord, "Q", "que")
    End If
    
    Dim r1Prefix() As String: ReDim r1Prefix(0)
    Dim r1Start As Integer: r1Start = sbR1(strWord, r1Prefix())
    If r1Start < 3 Then r1Start = 3
    
    
    Dim r2Start As Integer: r2Start = sbR2(strWord, r1Prefix())
    
    'Step 1
    Dim boolNissFlag As Boolean
    If Right(strWord, 3) = "ern" Then
        If Len(strWord) - r1Start >= 3 Then strWord = Left(strWord, Len(strWord) - 3)
    ElseIf Right(strWord, 2) = "em" Or Right(strWord, 2) = "er" Then
        If Len(strWord) - r1Start >= 2 Then strWord = Left(strWord, Len(strWord) - 2)
    ElseIf Right(strWord, 2) = "en" Or Right(strWord, 2) = "es" Then
        If Len(strWord) - r1Start >= 2 Then
            strWord = Left(strWord, Len(strWord) - 2)
            boolNissFlag = True
        End If
    ElseIf Right(strWord, 1) = "e" Then
        If Len(strWord) - r1Start >= 1 Then
            strWord = Left(strWord, Len(strWord) - 1)
            boolNissFlag = True
        End If
    ElseIf Right(strWord, 1) = "s" Then
        If Len(strWord) >= 2 And Len(strWord) - r1Start >= 1 Then
            If Mid(strWord, Len(strWord) - 1, 1) Like "[bdfghklmnrt]" Then
                strWord = Left(strWord, Len(strWord) - 1)
            End If
        End If
    End If
    
    If boolNissFlag And Right(strWord, 4) = "niss" Then
        strWord = Left(strWord, Len(strWord) - 1)
    End If
    
    'step2
    If Right(strWord, 3) = "est" Then
        If Len(strWord) - r1Start >= 3 Then strWord = Left(strWord, Len(strWord) - 3)
    ElseIf Right(strWord, 2) Like "e[nr]" Then
        If Len(strWord) - r1Start >= 2 Then strWord = Left(strWord, Len(strWord) - 2)
    ElseIf Right(strWord, 2) = "st" Then
        If Len(strWord) >= 6 Then
            If Len(strWord) - r1Start >= 2 And Mid(strWord, Len(strWord) - 2, 1) Like "[bdfghklmnt]" Then
                strWord = Left(strWord, Len(strWord) - 2)
            End If
        End If
    End If
    
    'step 3
    If Right(strWord, 4) = "isch" Then
        If Len(strWord) > 5 Then
            If Len(strWord) - r2Start >= 4 And Mid(strWord, Len(strWord) - 4, 1) <> "e" Then
                strWord = Left(strWord, Len(strWord) - 4)
            End If
        End If
    ElseIf Right(strWord, 4) = "lich" Or Right(strWord, 4) = "heit" Then
        If Len(strWord) - r2Start >= 4 Then
            strWord = Left(strWord, Len(strWord) - 4)
            If Right(strWord, 2) Like "e[rn]" And Len(strWord) - r1Start >= 2 Then
                strWord = Left(strWord, Len(strWord) - 2)
            End If
        End If
    ElseIf Right(strWord, 4) = "keit" Then
        If Len(strWord) - r2Start >= 4 Then
            strWord = Left(strWord, Len(strWord) - 4)
            If Right(strWord, 4) = "lich" And Len(strWord) - r2Start >= 4 Then
                strWord = Left(strWord, Len(strWord) - 4)
            ElseIf Right(strWord, 2) = "ig" And Len(strWord) - r2Start >= 2 Then
                strWord = Left(strWord, Len(strWord) - 2)
            End If
        End If
    ElseIf Right(strWord, 3) = "end" Or Right(strWord, 3) = "ung" Then
        If Len(strWord) - r2Start >= 3 Then
            strWord = Left(strWord, Len(strWord) - 3)
            
            If Len(strWord) > 3 Then
                If Right(strWord, 2) = "ig" And Len(strWord) - r2Start >= 2 And Mid(strWord, Len(strWord) - 2, 1) <> "e" Then
                    strWord = Left(strWord, Len(strWord) - 2)
                End If
            End If
        End If
    ElseIf Right(strWord, 2) Like "i[gk]" Then
        If Len(strWord) > 2 Then
            If Len(strWord) - r2Start >= 2 And Mid(strWord, Len(strWord) - 2, 1) <> "e" Then
                strWord = Left(strWord, Len(strWord) - 2)
            End If
        End If
    End If
    
    strWord = Replace(strWord, "Y", "y")
    strWord = Replace(strWord, "U", "u")
    
    
    strWord = Replace(strWord, "ä", "a")
    strWord = Replace(strWord, "ö", "o")
    strWord = Replace(strWord, "ü", "u")
    
    
    SnowballGerman = strWord
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
        If Not boolVowelFound And Mid$(strTerm, i, 1) Like "[aeiouyäöü]" Then
            boolVowelFound = True
        ElseIf boolVowelFound And Not Mid$(strTerm, i, 1) Like "[aeiouyäöü]" Then
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
