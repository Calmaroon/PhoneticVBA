Attribute VB_Name = "Stem_ClefSwedish"
Option Explicit
Function CLEFSwedish(strWord As String) As String
    If Len(strWord) > 4 And Right$(strWord, 1) = "s" Then
        strWord = Left$(strWord, Len(strWord) - 1)
    End If

    If Right$(strWord, 5) = "elser" Or Right$(strWord, 5) = "heten" Then
        CLEFSwedish = Left$(strWord, Len(strWord) - 5)
        Exit Function
    End If
    If Right$(strWord, 4) = "arne" Or Right$(strWord, 4) = "erna" Or Right$(strWord, 4) = "ande" Or Right$(strWord, 4) = "else" Or Right$(strWord, 4) = "aste" Or Right$(strWord, 4) = "orna" Or Right$(strWord, 4) = "aren" Then
        CLEFSwedish = Left$(strWord, Len(strWord) - 4)
        Exit Function
    End If
    If Right$(strWord, 3) = "are" Or Right$(strWord, 3) = "ast" Or Right$(strWord, 3) = "het" Then
        CLEFSwedish = Left$(strWord, Len(strWord) - 3)
        Exit Function
    End If
    If Right$(strWord, 2) = "ar" Or Right$(strWord, 2) = "er" Or Right$(strWord, 2) = "en" Or Right$(strWord, 2) = "at" Or Right$(strWord, 2) = "te" Or Right$(strWord, 2) = "et" Then
        CLEFSwedish = Left$(strWord, Len(strWord) - 2)
        Exit Function
    End If
    If Right$(strWord, 1) Like "[aent]" Then
        CLEFSwedish = Left$(strWord, Len(strWord) - 1)
        Exit Function
    End If
    CLEFSwedish = strWord
End Function
