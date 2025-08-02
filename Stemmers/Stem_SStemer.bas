Attribute VB_Name = "Stem_SStemer"
Option Explicit
Function SStemmer(strWord As String) As String
    Dim strlowered As String
    strlowered = LCase$(strWord)
    If Len(strWord) >= 3 And right$(strlowered, 3) = "ies" Then
        If Mid$(strlowered, Len(strWord) - 3 + 1, 1) <> "e" And Mid$(strlowered, Len(strWord) - 3 + 1, 1) <> "a" Then
            If StrComp(Mid$(strWord, Len(strWord), 1), UCase$(right$(strWord, 1)), vbBinaryCompare) = 0 Then
                SStemmer = left$(strWord, Len(strWord) - 3) & "Y"
            Else
                SStemmer = left$(strWord, Len(strWord) - 3) & "y"
            End If
            Exit Function
        End If
    End If
    
    If Len(strWord) >= 2 And right$(strlowered, 2) = "es" Then
        If Mid$(strlowered, Len(strWord) - 2 + 1, 1) <> "a" And Mid$(strlowered, Len(strWord) - 2 + 1, 1) <> "e" And Mid$(strlowered, Len(strWord) - 2 + 1, 1) <> "o" Then
            SStemmer = left$(strWord, Len(strWord) - 1)
            Exit Function
        End If
    End If
    
    If Len(strWord) >= 1 And right$(strlowered, 1) = "s" Then
        If Mid$(strlowered, Len(strWord) - 1, 1) <> "u" And Mid$(strlowered, Len(strWord) - 1, 1) <> "s" Then
            SStemmer = left$(strWord, Len(strWord) - 1)
            Exit Function
        End If
    End If
    
    SStemmer = strWord
End Function
