Attribute VB_Name = "Stem_ClefGermanPlus"
Option Explicit
Const strAccentsIn As String = "äàáâöòóôïìíîüùúû"
Const strAccentsOut As String = "aaaaooooiiiiuuuu"

Function CLEFGermanPlus(strWord As String) As String
    strWord = LCase$(strWord)
    Dim i As Integer
    For i = 1 To Len(strAccentsIn)
        If InStr(strWord, Mid(strAccentsIn, i, 1)) > 0 Then strWord = Replace(strWord, Mid(strAccentsIn, i, 1), Mid(strAccentsOut, i, 1))
    Next
    
    Dim intLen As Integer: intLen = Len(strWord) - 1
    If intLen > 4 And Right(strWord, 3) = "ern" Then
        strWord = Left(strWord, Len(strWord) - 3)
    ElseIf intLen > 3 And Right(strWord, 2) Like "e[mnrs]" Then
        strWord = Left(strWord, Len(strWord) - 2)
    ElseIf intLen > 2 Then
        If (Right(strWord, 1) = "e" Or (Right(strWord, 1) = "s" And Mid(strWord, Len(strWord) - 2, 1) Like "[bdfghklmnt]")) Then
            strWord = Left(strWord, Len(strWord) - 1)
        End If
    End If
    
    
    intLen = Len(strWord) - 1
    If intLen > 4 And Right(strWord, 3) = "est" Then
        strWord = Left(strWord, Len(strWord) - 3)
    ElseIf intLen > 3 Then
        If (Right(strWord, 2) Like "e[rn]" And Mid(strWord, Len(strWord) - 3, 1) Like "[bdfghklmnt]") Then
            strWord = Left(strWord, Len(strWord) - 2)
        End If
    End If
    
    CLEFGermanPlus = strWord
End Function
