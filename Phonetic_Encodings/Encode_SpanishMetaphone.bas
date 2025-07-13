Attribute VB_Name = "Encode_SpanishMetaphone"
Option Explicit

Function SpanishMetaphone(strInput As String, Optional intMaxLength As Integer = 6, Optional boolModified As Boolean = False) As String
    strInput = UCase$(strInput)
        
    If boolModified Then
        strInput = Replace(strInput, "MB", "NB")
        strInput = Replace(strInput, "MP", "NP")
        strInput = Replace(strInput, "BS", "S")
        If left(strInput, 2) = "PS" Then strInput = Mid$(strInput, 2)
    End If
    
    If InStr(strInput, "") > 0 Then strInput = Replace(strInput, "Á", "A")
    If InStr(strInput, "CH") > 0 Then strInput = Replace(strInput, "CH", "X")
    If InStr(strInput, "Ç") > 0 Then strInput = Replace(strInput, "Ç", "S")
    If InStr(strInput, "É") > 0 Then strInput = Replace(strInput, "É", "E")
    If InStr(strInput, "Í") > 0 Then strInput = Replace(strInput, "Í", "I")
    If InStr(strInput, "Ó") > 0 Then strInput = Replace(strInput, "Ó", "O")
    If InStr(strInput, "Ú") > 0 Then strInput = Replace(strInput, "Ú", "U")
    If InStr(strInput, "Ñ") > 0 Then strInput = Replace(strInput, "Ñ", "NY")
    If InStr(strInput, "GÜ") > 0 Then strInput = Replace(strInput, "GÜ", "W")
    If InStr(strInput, "Ü") > 0 Then strInput = Replace(strInput, "Ü", "U")
    If InStr(strInput, "B") > 0 Then strInput = Replace(strInput, "B", "V")
    If InStr(strInput, "LL") > 0 Then strInput = Replace(strInput, "LL", "Y")
    
    Dim strMetaKey As String
    Dim strChar As String
    Dim intPos As Integer
    intPos = 1
    Do While Len(strMetaKey) < intMaxLength And intPos <= Len(strInput)
        strChar = Mid(strInput, intPos, 1)
        If intPos = 1 And strChar Like "[AEIOUY]" Then
            strMetaKey = strChar
            intPos = intPos + 1
        Else
            Select Case strChar
               Case "D", "F", "J", "K", "M", "N", "P", "T", "V", "L", "Y":
                    strMetaKey = strMetaKey & strChar
                    'Skip doubled consonants
                    If Mid(strInput, intPos + 1, 1) = strChar Then
                        intPos = intPos + 2
                    Else
                        intPos = intPos + 1
                    End If
                Case "C"
                    If Mid(strInput, intPos + 1, 1) = "C" Then
                        strMetaKey = strMetaKey & "X"
                        intPos = intPos + 2
                    ElseIf Mid(strInput, intPos + 1, 1) Like "[EI]" Then
                        strMetaKey = strMetaKey & "Z"
                        intPos = intPos + 2
                    Else
                        strMetaKey = strMetaKey & "K"
                        intPos = intPos + 1
                    End If
                Case "G"
                    If Mid(strInput, intPos + 1, 1) Like "[EI]" Then
                        strMetaKey = strMetaKey & "J"
                        intPos = intPos + 2
                    Else
                        strMetaKey = strMetaKey & "G"
                        intPos = intPos + 1
                    End If
                Case "H":
                    If Mid(strInput, intPos + 1, 1) Like "[AEIOU]" Then
                        strMetaKey = strMetaKey & Mid(strInput, intPos + 1, 1)
                        intPos = intPos + 2
                    Else
                        strMetaKey = strMetaKey & "H"
                        intPos = intPos + 1
                    End If
                Case "Q"
                    If Mid(strInput, intPos + 1, 1) = "U" Then
                        intPos = intPos + 2
                    Else
                        intPos = intPos + 1
                    End If
                    strMetaKey = strMetaKey & "K"
                Case "W":
                    strMetaKey = strMetaKey & "U"
                    intPos = intPos + 1
                Case "Z", "R":
                    strMetaKey = strMetaKey & strChar
                    intPos = intPos + 1
                Case "S"
                    If Not Mid(strInput, intPos + 1, 1) Like "[AEIOU]" And intPos = 1 Then
                        strMetaKey = strMetaKey & "ES"
                    Else
                        strMetaKey = strMetaKey & "S"
                    End If
                    intPos = intPos + 1
                Case "X"
                    If Len(strInput) > 1 And intPos = 1 And Not Mid(strInput, intPos + 1, 1) Like "[AEIOU]" Then
                        strMetaKey = strMetaKey & "EX"
                    Else
                        strMetaKey = strMetaKey & "X"
                    End If
                    intPos = intPos + 1
                Case Else:
                    intPos = intPos + 1
            End Select
        End If
    Loop

    If boolModified Then strMetaKey = Replace(strMetaKey, "S", "Z")
    
    SpanishMetaphone = strMetaKey
End Function
