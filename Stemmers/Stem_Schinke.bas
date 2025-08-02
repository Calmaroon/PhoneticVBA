Attribute VB_Name = "Stem_Schinke"
Option Explicit
Const strKeepQue As String = "'at','quo','ne','ita','abs','aps','abus','adae','adus','deni','de','sus','obli','perae','plenis','quando','quis','quae','cuius','cui','quem','quam','qua','qui','quorum','quarum','quibus','quos','quas','quotusquis','quous','ubi','undi','us','uter','uti','utro','utribi','tor','co','conco','contor','detor','deco','exco','extor','obtor','optor','retor','reco','attor','inco','intor','praetor'"
Function Schinke(strWord As String) As String
    strWord = LCase(strWord)
    Dim strAlphaOnly As String
    Dim i As Integer
    For i = 1 To Len(strWord)
        If Mid(strWord, i, 1) Like "[a-z]" Then
            strAlphaOnly = strAlphaOnly & Mid(strWord, i, 1)
        End If
    Next
    strWord = strAlphaOnly
    
    strWord = Replace(strWord, "j", "i")
    strWord = Replace(strWord, "v", "u")
    
    If right(strWord, 3) = "que" Then
        If strWord = "que" Or InStr(strKeepQue, "'" & left(strWord, Len(strWord) - 3) & "'") > 0 Then
            Schinke = strWord & "," & strWord
            Exit Function
        Else
            strWord = left(strWord, Len(strWord) - 3)
        End If
    End If
        
    Dim strNoun As String: strNoun = strWord
    Dim strVerb As String: strVerb = strWord
    
    Dim nEndings As New Dictionary
    nEndings.Add 4, Array("ibus")
    nEndings.Add 3, Array("ius")
    nEndings.Add 2, Array("is", "nt", "ae", "os", "am", "ud", "as", "um", "em", "us", "es", "ia")
    nEndings.Add 1, Array("a", "e", "i", "o", "u")
    Dim boolExitFor As Boolean
    Dim ending As Variant
    For i = 4 To 1 Step -1
       For Each ending In nEndings(i)
            If right(strWord, i) = ending Then
                If Len(strWord) - 2 >= i Then
                    strNoun = left(strWord, Len(strWord) - i)
                Else
                    strNoun = strWord
                End If
                boolExitFor = True
                Exit For
            End If
       Next
       If boolExitFor Then Exit For
    Next
    
    Dim vEndingsStrip As New Dictionary
    vEndingsStrip.Add 6, Array()
    vEndingsStrip.Add 5, Array()
    vEndingsStrip.Add 4, Array("mini", "ntur", "stis")
    vEndingsStrip.Add 3, Array("mur", "mus", "ris", "sti", "tis", "tur")
    vEndingsStrip.Add 2, Array("ns", "nt", "ri")
    vEndingsStrip.Add 1, Array("m", "r", "s", "t")
    
    Dim vEndingsAlter As New Dictionary
    vEndingsAlter.Add 6, Array("iuntur")
    vEndingsAlter.Add 5, Array("beris", "erunt", "untur")
    vEndingsAlter.Add 4, Array("iunt")
    vEndingsAlter.Add 3, Array("bor", "ero", "unt")
    vEndingsAlter.Add 2, Array("bo")
    vEndingsAlter.Add 1, Array()
    
    

    Dim strNewWord As String
    Dim intAddLen As Integer
    boolExitFor = False
    For i = 6 To 1 Step -1
        For Each ending In vEndingsStrip(i)
            If right(strWord, i) = ending Then
                If Len(strWord) - 2 >= i Then
                    strVerb = left(strWord, Len(strWord) - i)
                Else
                    strVerb = strWord
                End If
                boolExitFor = True
                Exit For
            End If
        Next
        
        If boolExitFor Then Exit For
        For Each ending In vEndingsAlter(i)
            If right(strWord, i) = ending Then
                If right(strWord, i) = "iuntur" Or right(strWord, i) = "erunt" Or right(strWord, i) = "untur" Or right(strWord, i) = "iunt" Or right(strWord, i) = "unt" Then
                    strNewWord = left(strWord, Len(strWord) - i) & "i"
                    intAddLen = 1
                ElseIf right(strWord, i) = "beris" Or right(strWord, i) = "bor" Or right(strWord, i) = "bo" Then
                    strNewWord = left(strWord, Len(strWord) - i) & "bi"
                    intAddLen = 2
                Else
                    strNewWord = left(strWord, Len(strWord) - i) & "eri"
                    intAddLen = 3
                End If
                
                If Len(strNewWord) >= 2 + intAddLen Then
                    strVerb = strNewWord
                Else
                    strVerb = strWord
                End If
                boolExitFor = True
            End If
        Next
        If boolExitFor Then Exit For
        
    Next
    
    
    Schinke = strNoun & "," & strVerb
    
End Function
