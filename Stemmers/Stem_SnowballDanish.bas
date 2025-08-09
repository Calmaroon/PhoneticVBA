Attribute VB_Name = "Stem_SnowballDanish"
Option Explicit
Function SnowballDanish(strWord As String) As String
    strWord = LCase(strWord)
    Dim r1Start As Integer
    Dim r1Prefix() As String
    ReDim r1Prefix(0)
    r1Start = sbR1(strWord, r1Prefix)
    If r1Start < 3 Then r1Start = 3
    If r1Start > Len(strWord) Then r1Start = Len(strWord)
    
    'Step 1
    Dim r1 As String
    r1 = Mid(strWord, r1Start + 1)
    
    If Right(r1, 7) = "erendes" Then
        strWord = Left(strWord, Len(strWord) - 7)
    ElseIf Right(r1, 6) = "erende" Or Right(r1, 6) = "hedens" Then
        strWord = Left(strWord, Len(strWord) - 6)
    ElseIf Right(r1, 5) = "ethed" Or Right(r1, 5) = "erede" Or Right(r1, 5) = "heden" Or Right(r1, 5) = "heder" Or Right(r1, 5) = "endes" Or Right(r1, 5) = "ernes" Or Right(r1, 5) = "erens" Or Right(r1, 5) = "erets" Then
        strWord = Left(strWord, Len(strWord) - 5)
    ElseIf Right(r1, 4) = "ered" Or Right(r1, 4) = "ende" Or Right(r1, 4) = "erne" Or Right(r1, 4) = "eren" Or Right(r1, 4) = "erer" Or Right(r1, 4) = "heds" Or Right(r1, 4) = "enes" Or Right(r1, 4) = "eres" Or Right(r1, 4) = "eret" Then
        strWord = Left(strWord, Len(strWord) - 4)
    ElseIf Right(r1, 3) = "hed" Or Right(r1, 3) = "ene" Or Right(r1, 3) = "ere" Or Right(r1, 3) = "ens" Or Right(r1, 3) = "ers" Or Right(r1, 3) = "ets" Then
        strWord = Left(strWord, Len(strWord) - 3)
    ElseIf Right(r1, 2) Like "e[nrst]" Then
        strWord = Left(strWord, Len(strWord) - 2)
    ElseIf Right(r1, 1) = "e" Then
        strWord = Left(strWord, Len(strWord) - 1)
    ElseIf Right(r1, 1) = "s" Then
        If Len(strWord) > 1 Then
            If Mid(strWord, Len(strWord) - 1, 1) Like "[abcdfghjklmnoprtvyzå]" Then strWord = Left(strWord, Len(strWord) - 1)
        End If
    End If
    
    'step 2
    If Right(Mid(strWord, r1Start + 1), 2) = "gd" Or Right(Mid(strWord, r1Start + 1), 2) = "dt" Or Right(Mid(strWord, r1Start + 1), 2) = "gt" Or Right(Mid(strWord, r1Start + 1), 2) = "kt" Then strWord = Left(strWord, Len(strWord) - 1)

    
    'step 3
    If Right(strWord, 4) = "igst" Then strWord = Left(strWord, Len(strWord) - 2)
    
    r1 = Mid(strWord, r1Start + 1)
    Dim boolRepeatStep2 As Boolean
    If Right(r1, 4) = "elig" Then
        strWord = Left(strWord, Len(strWord) - 4)
        boolRepeatStep2 = True
    ElseIf Right(r1, 4) = "løst" Then
        strWord = Left(strWord, Len(strWord) - 1)
    ElseIf Right(r1, 3) = "lig" Or Right(r1, 3) = "els" Then
        strWord = Left(strWord, Len(strWord) - 3)
        boolRepeatStep2 = True
    ElseIf Right(r1, 2) = "ig" Then
        strWord = Left(strWord, Len(strWord) - 2)
        boolRepeatStep2 = True
    End If
    
    If boolRepeatStep2 Then
        If Right(Mid(strWord, r1Start + 1), 2) = "gd" Or Right(Mid(strWord, r1Start + 1), 2) = "dt" Or Right(Mid(strWord, r1Start + 1), 2) = "gt" Or Right(Mid(strWord, r1Start + 1), 2) = "kt" Then
            strWord = Left(strWord, Len(strWord) - 1)
        End If
    End If
    
    'Step 4
    If Len(strWord) - r1Start >= 1 And Len(strWord) >= 2 Then
        If Right(strWord, 1) = Mid(strWord, Len(strWord) - 1, 1) And Not Right(strWord, 1) Like "[aeiouyåæø]" Then
            strWord = Left(strWord, Len(strWord) - 1)
        End If
    
    End If
    
    SnowballDanish = strWord
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
        If Not boolVowelFound And Mid$(strTerm, i, 1) Like "[aeiouyåæø]" Then
            boolVowelFound = True
        ElseIf boolVowelFound And Not Mid$(strTerm, i, 1) Like "[aeiouyåæø]" Then
            sbR1 = i
            Exit Function
        End If
    Next
         
    sbR1 = Len(strTerm)
End Function
