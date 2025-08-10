Attribute VB_Name = "Stem_SnowballNorwegian"
Option Explicit
Function SnowballNorwegian(strWord As String) As String
    strWord = LCase$(strWord)
    Dim r1Prefix() As String: ReDim r1Prefix(0)
    Dim r1Start As Integer: r1Start = sbR1(strWord, r1Prefix())
    If r1Start < 3 Then r1Start = 3
    If r1Start > Len(strWord) Then r1Start = Len(strWord)
    
    Dim r1 As String: r1 = Mid(strWord, r1Start + 1)
    
    'Step1
    If Right(r1, 7) = "hetenes" Then
        strWord = Left(strWord, Len(strWord) - 7)
    ElseIf Right(r1, 6) = "hetene" Or Right(r1, 6) = "hetens" Then
        strWord = Left(strWord, Len(strWord) - 6)
    ElseIf Right(r1, 5) = "heten" Or Right(r1, 5) = "heter" Or Right(r1, 5) = "endes" Then
        strWord = Left(strWord, Len(strWord) - 5)
    ElseIf Right(r1, 4) = "ande" Or Right(r1, 4) = "ende" Or Right(r1, 4) = "edes" Or Right(r1, 4) = "enes" Or Right(r1, 4) = "erte" Then
        If Right(strWord, 4) = "erte" Then
            strWord = Left(strWord, Len(strWord) - 2)
        Else
            strWord = Left(strWord, Len(strWord) - 4)
        End If
    ElseIf Right(r1, 3) = "ede" Or Right(r1, 3) = "ane" Or Right(r1, 3) = "ene" Or Right(r1, 3) = "ens" Or Right(r1, 3) = "ers" Or Right(r1, 3) = "ets" Or Right(r1, 3) = "het" Or Right(r1, 3) = "ast" Or Right(r1, 3) = "ert" Then
        If Right(strWord, 3) = "ert" Then
            strWord = Left(strWord, Len(strWord) - 1)
        Else
            strWord = Left(strWord, Len(strWord) - 3)
        End If
    ElseIf Right(r1, 2) = "en" Or Right(r1, 2) = "ar" Or Right(r1, 2) = "er" Or Right(r1, 2) = "as" Or Right(r1, 2) = "es" Or Right(r1, 2) = "et" Then
        strWord = Left(strWord, Len(strWord) - 2)
    ElseIf Right(r1, 1) Like "[ae]" Then
        strWord = Left(strWord, Len(strWord) - 1)
    ElseIf Right(r1, 1) = "s" Then
        If (Len(strWord) > 1 And Mid(strWord, Len(strWord) - 1, 1) Like "[bcdfghjlmnoprtvyz]") Or (Len(strWord) > 2 And Mid(strWord, Len(strWord) - 1, 1) = "k" And Not Mid(strWord, Len(strWord) - 2, 1) Like "[aeiouyåæø]") Then
            strWord = Left(strWord, Len(strWord) - 1)
        End If
    End If
    
    'step2
    If Right(Mid(strWord, r1Start + 1), 2) Like "[dv]t" Then
        strWord = Left(strWord, Len(strWord) - 1)
    End If
    
    'Step 3
    r1 = Mid(strWord, r1Start + 1)
    If Right(r1, 7) = "hetslov" Then
        strWord = Left(strWord, Len(strWord) - 7)
    ElseIf Right(r1, 4) = "eleg" Or Right(r1, 4) = "elig" Or Right(r1, 4) = "elov" Or Right(r1, 4) = "slov" Then
        strWord = Left(strWord, Len(strWord) - 4)
    ElseIf Right(r1, 3) = "leg" Or Right(r1, 3) = "eig" Or Right(r1, 3) = "lig" Or Right(r1, 3) = "els" Or Right(r1, 3) = "lov" Then
        strWord = Left(strWord, Len(strWord) - 3)
    ElseIf Right(r1, 2) = "ig" Then
        strWord = Left(strWord, Len(strWord) - 2)
    End If
    
    SnowballNorwegian = strWord
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
