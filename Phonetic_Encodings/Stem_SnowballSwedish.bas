Attribute VB_Name = "Stem_SnowballSwedish"
Option Explicit
Function SnowballSwedish(strWord As String) As String

    strWord = LCase(strWord)
    Dim r1Prefix() As String: ReDim r1Prefix(0)
    
    Dim r1Start As Integer: r1Start = sbR1(strWord, r1Prefix())
    If r1Start < 3 Then r1Start = 3
    If Len(strWord) < 3 Then r1Start = Len(strWord)
    
    Dim r1 As String
    r1 = Mid(strWord, r1Start + 1)
    
    'step 1
    If Right(r1, 7) = "heterna" Then
        strWord = Left(strWord, Len(strWord) - 7)
    ElseIf Right(r1, 6) = "hetens" Then
        strWord = Left(strWord, Len(strWord) - 6)
    ElseIf Right(r1, 5) = "anden" Or Right(r1, 5) = "heten" Or Right(r1, 5) = "heter" Or Right(r1, 5) = "arnas" Or Right(r1, 5) = "ernas" Or Right(r1, 5) = "ornas" Or Right(r1, 5) = "andes" Or Right(r1, 5) = "arens" Or Right(r1, 5) = "andet" Then
        strWord = Left(strWord, Len(strWord) - 5)
    ElseIf Right(r1, 4) Like "[aeo]rna" Or Right(r1, 4) = "ande" Or Right(r1, 4) = "arne" Or Right(r1, 4) = "aste" Or Right(r1, 4) = "aren" Or Right(r1, 4) = "ades" Or Right(r1, 4) = "erns" Then
        strWord = Left(strWord, Len(strWord) - 4)
    ElseIf Right(r1, 3) = "ade" Or Right(r1, 3) = "are" Or Right(r1, 3) = "ern" Or Right(r1, 3) = "ens" Or Right(r1, 3) = "het" Or Right(r1, 3) = "ast" Then
        strWord = Left(strWord, Len(strWord) - 3)
    ElseIf Right(r1, 2) = "ad" Or Right(r1, 2) = "en" Or Right(r1, 2) = "ar" Or Right(r1, 2) = "er" Or Right(r1, 2) = "or" Or Right(r1, 2) = "as" Or Right(r1, 2) = "es" Or Right(r1, 2) = "at" Then
        strWord = Left(strWord, Len(strWord) - 2)
    ElseIf Right(r1, 1) Like "[ae]" Then
         strWord = Left(strWord, Len(strWord) - 1)
    ElseIf Right(r1, 1) = "s" Then
        If Len(strWord) > 1 Then
            If Mid(strWord, Len(strWord) - 1, 1) Like "[bcdfghjklmnoprtvy]" Then
                strWord = Left(strWord, Len(strWord) - 1)
            End If
        End If
    End If
    
    'step 2
    Dim strStep2
    strStep2 = Mid(strWord, r1Start + 1)
    If Right(strStep2, 2) = "dd" Or Right(strStep2, 2) = "gd" Or Right(strStep2, 2) = "nn" Or Right(strStep2, 2) = "dt" Or Right(strStep2, 2) = "gt" Or Right(strStep2, 2) = "kt" Or Right(strStep2, 2) = "tt" Then
        strWord = Left(strWord, Len(strWord) - 1)
    End If
    
    'step 3
    r1 = Mid(strWord, r1Start + 1)
    If Right(r1, 5) = "fullt" Then
        strWord = Left(strWord, Len(strWord) - 1)
    ElseIf Right(r1, 4) = "löst" Then
        strWord = Left(strWord, Len(strWord) - 1)
    ElseIf Right(r1, 3) = "lig" Or Right(r1, 3) = "els" Then
        strWord = Left(strWord, Len(strWord) - 3)
    ElseIf Right(r1, 2) = "ig" Then
        strWord = Left(strWord, Len(strWord) - 2)
    End If
    
    SnowballSwedish = strWord

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
        If Not boolVowelFound And Mid$(strTerm, i, 1) Like "[aeiouyäåö]" Then
            boolVowelFound = True
        ElseIf boolVowelFound And Not Mid$(strTerm, i, 1) Like "[aeiouäåö]" Then
            sbR1 = i
            Exit Function
        End If
    Next
         
    sbR1 = Len(strTerm)
End Function

