Attribute VB_Name = "Encode_Caverphone"
Option Explicit
Const strLCaseSet = "abcdefghijklmnopqrstuvwxyz"
Function Caverphone(strWord As String, Optional intVersion As Integer = 2) As String
    strWord = LCase$(strWord)
    Dim strWordAlpha As String
    Dim i As Long
    
    For i = 1 To Len(strWord)
        If InStr(strLCaseSet, Mid$(strWord, i, 1)) > 0 Then
            strWordAlpha = strWordAlpha & Mid$(strWord, i, 1)
        End If
    Next
    
    If intVersion <> 1 And Right(strWordAlpha, 1) = "e" Then
        strWordAlpha = Left$(strWordAlpha, Len(strWordAlpha) - 1)
    End If
    
    If Left$(strWordAlpha, 5) = "cough" Then
        Mid$(strWordAlpha, 1, 5) = "cou2f"
    End If
    If Left$(strWordAlpha, 5) = "rough" Then
        Mid$(strWordAlpha, 1, 5) = "rou2f"
    End If
    If Left$(strWordAlpha, 5) = "tough" Then
        Mid$(strWordAlpha, 1, 5) = "tou2f"
    End If
    If Left$(strWordAlpha, 6) = "enough" Then
        Mid$(strWordAlpha, 1, 6) = "enou2f"
    End If
    If intVersion <> 1 And Left$(strWordAlpha, 6) = "trough" Then
        Mid$(strWordAlpha, 1, 6) = "trou2f"
    End If
    If Left$(strWordAlpha, 2) = "gn" Then
        Mid$(strWordAlpha, 1, 2) = "2n"
    End If
    If Right(strWordAlpha, 2) = "mb" Then
    strWordAlpha = Left$(strWordAlpha, Len(strWordAlpha) - 1) & "2"
    End If
    
    strWordAlpha = Replace(strWordAlpha, "cq", "2q")
    strWordAlpha = Replace(strWordAlpha, "ci", "si")
    strWordAlpha = Replace(strWordAlpha, "ce", "se")
    strWordAlpha = Replace(strWordAlpha, "cy", "sy")
    strWordAlpha = Replace(strWordAlpha, "tch", "2ch")
    strWordAlpha = Replace(strWordAlpha, "c", "k")
    strWordAlpha = Replace(strWordAlpha, "q", "k")
    strWordAlpha = Replace(strWordAlpha, "x", "k")
    strWordAlpha = Replace(strWordAlpha, "v", "f")
    strWordAlpha = Replace(strWordAlpha, "dg", "2g")
    strWordAlpha = Replace(strWordAlpha, "tio", "sio")
    strWordAlpha = Replace(strWordAlpha, "tia", "sia")
    strWordAlpha = Replace(strWordAlpha, "d", "t")
    strWordAlpha = Replace(strWordAlpha, "ph", "fh")
    strWordAlpha = Replace(strWordAlpha, "b", "p")
    strWordAlpha = Replace(strWordAlpha, "sh", "s2")
    strWordAlpha = Replace(strWordAlpha, "z", "s")
    
    If Left$(strWordAlpha, 1) Like "[aeiou]" Then
        strWordAlpha = "A" & Mid$(strWordAlpha, 2)
    End If
    
    strWordAlpha = Replace(strWordAlpha, "a", "3")
    strWordAlpha = Replace(strWordAlpha, "e", "3")
    strWordAlpha = Replace(strWordAlpha, "i", "3")
    strWordAlpha = Replace(strWordAlpha, "o", "3")
    strWordAlpha = Replace(strWordAlpha, "u", "3")
    
    If intVersion <> 1 Then
        strWordAlpha = Replace(strWordAlpha, "j", "y")
        
        If Left$(strWordAlpha, 2) = "y3" Then
            strWordAlpha = "Y3" & Mid$(strWordAlpha, 3)
        End If
        If Left$(strWordAlpha, 1) = "y" Then
            strWordAlpha = "A" & Mid$(strWordAlpha, 2)
        End If
        strWordAlpha = Replace(strWordAlpha, "y", "3")
    End If
    
    strWordAlpha = Replace(strWordAlpha, "3gh3", "3kh3")
    strWordAlpha = Replace(strWordAlpha, "gh", "22")
    strWordAlpha = Replace(strWordAlpha, "g", "k")
    
    Dim strChar As Variant
    For Each strChar In Split("s,t,p,k,f,m,n", ",")
        While InStr(strWordAlpha, strChar & strChar) > 0
            strWordAlpha = Replace(strWordAlpha, strChar & strChar, strChar)
        Wend
        strWordAlpha = Replace(strWordAlpha, strChar, UCase(strChar))
    Next
    
    strWordAlpha = Replace(strWordAlpha, "w3", "W3")
    
    If intVersion = 1 Then
        strWordAlpha = Replace(strWordAlpha, "wy", "Wy")
    End If
    
    strWordAlpha = Replace(strWordAlpha, "wh3", "Wh3")
    If intVersion = 1 Then
        strWordAlpha = Replace(strWordAlpha, "why", "Why")
    End If
    
    If intVersion <> 1 And Right(strWordAlpha, 1) = "w" Then
        strWordAlpha = Left$(strWordAlpha, Len(strWordAlpha) - 1) & "3"
    End If
    strWordAlpha = Replace(strWordAlpha, "w", "2")
    
    If Left$(strWordAlpha, 1) = "h" Then
        strWordAlpha = "A" & Mid$(strWordAlpha, 2)
    End If
    strWordAlpha = Replace(strWordAlpha, "h", "2")
    strWordAlpha = Replace(strWordAlpha, "r3", "R3")
    
    If intVersion = 1 Then
        strWordAlpha = Replace(strWordAlpha, "ry", "Ry")
    End If
    
    If intVersion <> 1 And Right(strWordAlpha, 1) = "r" Then
        strWordAlpha = Left$(strWordAlpha, Len(strWordAlpha) - 1) & "3"
    End If
    strWordAlpha = Replace(strWordAlpha, "r", "2")
    strWordAlpha = Replace(strWordAlpha, "l3", "L3")
    
    If intVersion = 1 Then
        strWordAlpha = Replace(strWordAlpha, "ly", "Ly")
    End If
    
    If intVersion <> 1 And Right(strWordAlpha, 1) = "l" Then
        strWordAlpha = Left$(strWordAlpha, Len(strWordAlpha) - 1) & "3"
    End If
    strWordAlpha = Replace(strWordAlpha, "l", "2")
    
    If intVersion = 1 Then
        strWordAlpha = Replace(strWordAlpha, "j", "y")
        strWordAlpha = Replace(strWordAlpha, "y3", "Y3")
        strWordAlpha = Replace(strWordAlpha, "y", "2")
    End If

    strWordAlpha = Replace(strWordAlpha, "2", "")
    If intVersion <> 1 And Right(strWordAlpha, 1) = "3" Then
             strWordAlpha = Left$(strWordAlpha, Len(strWordAlpha) - 1) & "A"
    End If
    strWordAlpha = Replace(strWordAlpha, "3", "")
        
    strWordAlpha = strWordAlpha & String(10, "1")
    If intVersion <> 1 Then
        strWordAlpha = Left$(strWordAlpha, 10)
    Else
        strWordAlpha = Left$(strWordAlpha, 6)
    End If

    Caverphone = strWordAlpha
End Function
