Attribute VB_Name = "Stem_Porter2"
Option Explicit
Function Porter2(strWord As String, Optional boolEarlyEnglish As Boolean) As String
    strWord = LCase(strWord)
    strWord = Replace(strWord, "’", "'") 'they appear to be the same thing...will look into specific ASC codes
    
    Dim Doubles As New Dictionary
    Doubles.Add "bb", "": Doubles.Add "dd", "": Doubles.Add "ff", "": Doubles.Add "gg", "": Doubles.Add "mm", "": Doubles.Add "nn", "": Doubles.Add "pp", "": Doubles.Add "rr", "": Doubles.Add "tt", ""
    
    Dim exception1Dict As New Dictionary
    exception1Dict.Add "skis", "ski"
    exception1Dict.Add "skies", "sky"
    exception1Dict.Add "dying", "die"
    exception1Dict.Add "lying", "lie"
    exception1Dict.Add "tying", "tie"
    exception1Dict.Add "idly", "idl"
    exception1Dict.Add "gently", "gentl"
    exception1Dict.Add "ugly", "ugli"
    exception1Dict.Add "early", "earli"
    exception1Dict.Add "only", "onli"
    exception1Dict.Add "singly", "singl"
    
    Dim exception1Set As New Dictionary
    exception1Set.Add "sky", ""
    exception1Set.Add "news", ""
    exception1Set.Add "howe", ""
    exception1Set.Add "atlas", ""
    exception1Set.Add "cosmos", ""
    exception1Set.Add "bias", ""
    exception1Set.Add "andes", ""
    
    Dim exception2Set As New Dictionary
    exception2Set.Add "inning", ""
    exception2Set.Add "outing", ""
    exception2Set.Add "canning", ""
    exception2Set.Add "herring", ""
    exception2Set.Add "earring", ""
    exception2Set.Add "proceed", ""
    exception2Set.Add "exceed", ""
    exception2Set.Add "succeed", ""
    
    If exception1Dict.Exists(strWord) Then
        Porter2 = exception1Dict.item(strWord)
        Exit Function
    ElseIf exception1Set.Exists(strWord) Then
        Porter2 = strWord
        Exit Function
    End If

    If Len(strWord) < 3 Then
        Porter2 = strWord
        Exit Function
    End If

    Do While Len(strWord) > 0 And Left$(strWord, 1) = "'"
        strWord = Mid$(strWord, 2)
        If Len(strWord) < 2 Then
            Porter2 = strWord
            Exit Function
        End If
    Loop

    If Left$(strWord, 1) = "y" Then
        strWord = "Y" & Mid$(strWord, 2)
    End If
    
    Dim i As Integer
    For i = 2 To Len(strWord)
        If Mid$(strWord, i, 1) = "y" And Mid$(strWord, i - 1, 1) Like "[aeiouy]" Then
            strWord = Left$(strWord, i - 1) & "Y" & Mid$(strWord, i + 1)
        End If
    Next
    Dim strPrefixes() As String: strPrefixes = Split("commun,gener,arsen", ",")
    Dim r1Start As Integer: r1Start = sbR1(strWord, strPrefixes)
    Dim r2Start As Integer: r2Start = sbR2(strWord, strPrefixes)
    
    'Step0
    If Len(strWord) > 2 Then
        If Mid$(strWord, Len(strWord) - 2) = "'s'" Then
            strWord = Left$(strWord, Len(strWord) - 3)
        ElseIf Mid$(strWord, Len(strWord) - 1) = "'s" Then
            strWord = Left$(strWord, Len(strWord) - 2)
        ElseIf Right(strWord, 1) = "'" Then
            strWord = Left$(strWord, Len(strWord) - 1)
        End If
    End If
    
    'Step1A
    If Right(strWord, 4) = "sses" Then
        strWord = Left$(strWord, Len(strWord) - 2)
    ElseIf Right(strWord, 3) Like "ie[sd]" Then
        If Len(strWord) > 4 Then
            strWord = Left$(strWord, Len(strWord) - 2)
        Else
            strWord = Left$(strWord, Len(strWord) - 1)
        End If
    ElseIf Right(strWord, 2) = "us" Or Right(strWord, 2) = "ss" Then
        'pass
    ElseIf Right(strWord, 1) = "s" Then
        If sbHasVowel(Left$(strWord, Len(strWord) - 2)) Then
            strWord = Left$(strWord, Len(strWord) - 1)
        End If
    End If
    
    If exception2Set.Exists(strWord) Then
        Porter2 = strWord
        Exit Function
    End If
    
    'Step1B
    Dim boolStep1bFlag As Boolean
    If Right(strWord, 5) = "eedly" Then
        If Len(strWord) - r1Start >= 5 Then strWord = Left$(strWord, Len(strWord) - 3)
    ElseIf Right(strWord, 5) = "ingly" Then
        If sbHasVowel(Left$(strWord, Len(strWord) - 5)) Then
            strWord = Left$(strWord, Len(strWord) - 5)
            boolStep1bFlag = True
        End If
    ElseIf Right(strWord, 4) = "edly" Then
        If sbHasVowel(Left$(strWord, Len(strWord) - 4)) Then
            strWord = Left$(strWord, Len(strWord) - 4)
            boolStep1bFlag = True
        End If
    ElseIf Right(strWord, 3) = "eed" Then
        If Len(strWord) - r1Start >= 3 Then strWord = Left$(strWord, Len(strWord) - 1)
    ElseIf Right(strWord, 3) = "ing" Then
        If sbHasVowel(Left$(strWord, Len(strWord) - 3)) Then
            strWord = Left$(strWord, Len(strWord) - 3)
            boolStep1bFlag = True
        End If
    ElseIf Right(strWord, 2) = "ed" Then
        If sbHasVowel(Left$(strWord, Len(strWord) - 2)) Then
            strWord = Left$(strWord, Len(strWord) - 2)
            boolStep1bFlag = True
        End If
    ElseIf boolEarlyEnglish Then
        If Right(strWord, 3) = "est" Or Right(strWord, 3) = "eth" Then
            If sbHasVowel(Left$(strWord, Len(strWord) - 3)) Then
                strWord = Left$(strWord, Len(strWord) - 3)
                boolStep1bFlag = True
            End If
        End If
    End If
    
    If boolStep1bFlag Then
        If Right(strWord, 2) = "at" Or Right(strWord, 2) = "bl" Or Right(strWord, 2) = "iz" Then
            strWord = strWord & "e"
        ElseIf Doubles.Exists(Right(strWord, 2)) Then
            strWord = Left$(strWord, Len(strWord) - 1)
        ElseIf sbShortWord(strWord, strPrefixes) Then
            strWord = strWord & "e"
        End If
    End If
    
    'step1C
    If Len(strWord) > 2 And Right(strWord, 1) Like "[yY]" And Not Mid$(strWord, Len(strWord) - 1, 1) Like "[aeiouy]" Then
        strWord = Left$(strWord, Len(strWord) - 1) & "i"
    End If
    'step 2
    If Mid$(strWord, Len(strWord) - 1, 1) = "a" Then
        If Right(strWord, 7) = "ational" Then
            If Len(strWord) - r1Start >= 7 Then strWord = Left$(strWord, Len(strWord) - 5) & "e"
        ElseIf Right(strWord, 6) = "tional" Then
            If Len(strWord) - r1Start >= 6 Then strWord = Left$(strWord, Len(strWord) - 2)
        End If
    ElseIf Mid$(strWord, Len(strWord) - 1, 1) = "c" Then
        If Right(strWord, 4) Like "[ae]nci" Then If Len(strWord) - r1Start >= 4 Then strWord = Left$(strWord, Len(strWord) - 1) & "e"
    ElseIf Mid$(strWord, Len(strWord) - 1, 1) = "e" Then
        If Right(strWord, 4) = "izer" Then If Len(strWord) - r1Start >= 4 Then strWord = Left$(strWord, Len(strWord) - 1)
    ElseIf Mid$(strWord, Len(strWord) - 1, 1) = "g" Then
        If Right(strWord, 3) = "ogi" And Len(strWord) > 4 Then
            If r1Start >= 1 And Len(strWord) - r1Start >= 3 And Mid$(strWord, Len(strWord) - 3, 1) = "l" Then strWord = Left$(strWord, Len(strWord) - 1)
        End If
    ElseIf Mid$(strWord, Len(strWord) - 1, 1) = "l" Then
        If Right(strWord, 6) = "lessli" Then
            If Len(strWord) - r1Start >= 6 Then strWord = Left$(strWord, Len(strWord) - 2)
        ElseIf Right(strWord, 5) = "entli" Or Right(strWord, 5) = "fulli" Or Right(strWord, 5) = "ousli" Then
            If Len(strWord) - r1Start >= 5 Then strWord = Left$(strWord, Len(strWord) - 2)
        ElseIf Right(strWord, 4) = "abli" Then
            If Len(strWord) - r1Start >= 4 Then strWord = Left$(strWord, Len(strWord) - 1) & "e"
        ElseIf Right(strWord, 4) = "alli" Then
            If Len(strWord) - r1Start >= 4 Then strWord = Left$(strWord, Len(strWord) - 2)
        ElseIf Right(strWord, 3) = "bli" Then
            If Len(strWord) - r1Start >= 3 Then strWord = Left$(strWord, Len(strWord) - 1) & "e"
        ElseIf Right(strWord, 2) = "li" Then
            If r1Start >= 1 And Len(strWord) - r1Start >= 2 And Mid$(strWord, Len(strWord) - 2, 1) Like "[cdeghkmnrt]" Then strWord = Left$(strWord, Len(strWord) - 2)
        End If
    ElseIf Mid$(strWord, Len(strWord) - 1, 1) = "o" Then
        If Right(strWord, 7) = "ization" Then
            If Len(strWord) - r1Start >= 7 Then strWord = Left$(strWord, Len(strWord) - 5) & "e"
        ElseIf Right(strWord, 5) = "ation" Then
            If Len(strWord) - r1Start >= 5 Then strWord = Left$(strWord, Len(strWord) - 3) & "e"
        ElseIf Right(strWord, 4) = "ator" Then
            If Len(strWord) - r1Start >= 4 Then strWord = Left$(strWord, Len(strWord) - 2) & "e"
        End If
    ElseIf Mid$(strWord, Len(strWord) - 1, 1) = "s" Then
        If Right(strWord, 7) = "fulness" Or Right(strWord, 7) = "ousness" Or Right(strWord, 7) = "iveness" Then
            If Len(strWord) - r1Start >= 7 Then strWord = Left$(strWord, Len(strWord) - 4)
        ElseIf Right(strWord, 5) = "alism" Then
            If Len(strWord) - r1Start >= 5 Then strWord = Left$(strWord, Len(strWord) - 3)
        End If
    ElseIf Mid$(strWord, Len(strWord) - 1, 1) = "t" Then
        If Right(strWord, 6) = "biliti" Then
            If Len(strWord) - r1Start >= 6 Then strWord = Left$(strWord, Len(strWord) - 5) & "le"
        ElseIf Right(strWord, 5) = "aliti" Then
            If Len(strWord) - r1Start >= 5 Then strWord = Left$(strWord, Len(strWord) - 3)
        ElseIf Right(strWord, 5) = "iviti" Then
            If Len(strWord) - r1Start >= 5 Then strWord = Left$(strWord, Len(strWord) - 3) & "e"
        End If
    End If
    
    'Step3
    If Right(strWord, 7) = "ational" Then
        If Len(strWord) - r1Start >= 7 Then strWord = Left$(strWord, Len(strWord) - 5) & "e"
    ElseIf Right(strWord, 6) = "tional" Then
        If Len(strWord) - r1Start >= 6 Then strWord = Left$(strWord, Len(strWord) - 2)
    ElseIf Right(strWord, 5) = "alize" Or Right(strWord, 5) = "icate" Or Right(strWord, 5) = "iciti" Then
        If Len(strWord) - r1Start >= 5 Then strWord = Left$(strWord, Len(strWord) - 3)
    ElseIf Right(strWord, 5) = "ative" Then
        If Len(strWord) - r2Start >= 5 Then strWord = Left$(strWord, Len(strWord) - 5)
    ElseIf Right(strWord, 4) = "ical" Then
        If Len(strWord) - r1Start >= 4 Then strWord = Left$(strWord, Len(strWord) - 2)
    ElseIf Right(strWord, 4) = "ness" Then
        If Len(strWord) - r1Start >= 4 Then strWord = Left$(strWord, Len(strWord) - 4)
    ElseIf Right(strWord, 3) = "ful" Then
        If Len(strWord) - r1Start >= 3 Then strWord = Left$(strWord, Len(strWord) - 3)
    End If
    
    'step 4
    Dim strSuffix As Variant
    Dim boolSuffixRemoved As Boolean
    For Each strSuffix In Split("ement,ance,ence,able,ible,ment,ant,ent,ism,ate,iti,ous,ive,ize,al,er,ic", ",")
        If Right(strWord, Len(strSuffix)) = strSuffix Then
            If Len(strWord) - r2Start >= Len(strSuffix) Then
                strWord = Left$(strWord, Len(strWord) - Len(strSuffix))
                boolSuffixRemoved = True
            End If
            Exit For
        End If
    Next
    If Right(strWord, 3) = "ion" And Not boolSuffixRemoved Then
        If Len(strWord) - r2Start >= 3 And Len(strWord) >= 4 And Mid$(strWord, Len(strWord) - 3, 1) Like "[st]" Then
            strWord = Left$(strWord, Len(strWord) - 3)
        End If
    End If
    
    'Step 5
    If Right(strWord, 1) = "e" Then
        If Len(strWord) - r2Start >= 1 Or (Len(strWord) - r1Start >= 1 And Not endsInShortSyllable(Left$(strWord, Len(strWord) - 1))) Then strWord = Left$(strWord, Len(strWord) - 1)
    ElseIf Right(strWord, 1) = "l" Then
        If Len(strWord) - r2Start >= 1 And Mid$(strWord, Len(strWord) - 1, 1) = "l" Then strWord = Left$(strWord, Len(strWord) - 1)
    End If
    
    strWord = Replace(strWord, "Y", "y")
    
    Porter2 = strWord
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
        If Not boolVowelFound And Mid$(strTerm, i, 1) Like "[aeiouy]" Then
            boolVowelFound = True
        ElseIf boolVowelFound And Not Mid$(strTerm, i, 1) Like "[aeiouy]" Then
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
Function endsInShortSyllable(strTerm As String) As Boolean
    If strTerm = "" Then
        endsInShortSyllable = False
        Exit Function
    End If
    
    If Len(strTerm) = 2 Then
        If Left$(strTerm, 1) Like "[aeiouy]" And Not Right(strTerm, 1) Like "[aeiouy]" Then
            endsInShortSyllable = True
            Exit Function
        End If
    ElseIf Len(strTerm) >= 3 Then
        If Not Mid$(strTerm, Len(strTerm) - 2, 1) Like "[aeiouy]" And Mid$(strTerm, Len(strTerm) - 1, 1) Like "[aeiouy]" And Right(strTerm, 1) Like "[bcdfghjklmnpqrstvz]" Then
            endsInShortSyllable = True
            Exit Function
        End If
    End If
    
    endsInShortSyllable = False
End Function
Function sbShortWord(strTerm As String, r1Prefix() As String) As Boolean
    If sbR1(strTerm, r1Prefix()) = Len(strTerm) And endsInShortSyllable(strTerm) Then
        sbShortWord = True
        Exit Function
    End If
    sbShortWord = False
End Function
Function sbHasVowel(strTerm As String) As Boolean
    sbHasVowel = strTerm Like "*[aeiouy]*"
End Function
