Attribute VB_Name = "Encode_Haase"
Option Explicit
Const vowels As String = "AEIJOUY"
Const ucSet As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ"
Function Haase(strWord As String, Optional boolPrimaryOnly As Boolean = False) As String
    strWord = UCase$(strWord)
    strWord = Replace(strWord, "Ä", "AE")
    strWord = Replace(strWord, "Ö", "OE")
    strWord = Replace(strWord, "Ü", "UE")
    Dim cleanWord As String
    Dim i As Long
    For i = 1 To Len(strWord)
        If InStr(ucSet, Mid(strWord, i, 1)) > 0 Then cleanWord = cleanWord & Mid(strWord, i, 1)
    Next
    
    Dim variants As New Collection
    
    If boolPrimaryOnly Then
        variants.Add strWord
    Else
        Dim intPos As Integer
        intPos = 1
        Dim intCartesianCount As Integer
        
        If left(strWord, 2) = "CH" Then
            variants.Add Array("CH", "SCH"): intCartesianCount = intCartesianCount + 1
            intPos = intPos + 2
        End If
        
        Dim len3Vars As New Dictionary
        len3Vars.Add "OWN", "AUN": len3Vars.Add "WSK", "RSK": len3Vars.Add "SCH", "CH": len3Vars.Add "GLI", "LI": len3Vars.Add "AUX", "0": len3Vars.Add "EUX", "0"
        
        
        Do While intPos <= Len(strWord)
            If Mid(strWord, intPos, 4) = "ILLE" Then
                variants.Add Array("ILLE", "I"): intCartesianCount = intCartesianCount + 1
                intPos = intPos + 4
            ElseIf len3Vars.Exists(Mid(strWord, intPos, 3)) Then
                variants.Add Array(Mid(strWord, intPos, 3), len3Vars(Mid(strWord, intPos, 3))): intCartesianCount = intCartesianCount + 1
                intPos = intPos + 3
            ElseIf Mid(strWord, intPos, 2) = "RB" Then
                variants.Add Array("RB", "RW"): intCartesianCount = intCartesianCount + 1
                intPos = intPos + 2
            ElseIf Len(Mid(strWord, intPos)) = 3 And Mid(strWord, intPos) = "EAU" Then
                variants.Add Array("EAU", "O"): intCartesianCount = intCartesianCount + 1
                intPos = intPos + 3
            ElseIf Len(Mid(strWord, intPos)) = 1 And Mid(strWord, intPos) Like "[AO]" Then
                If Mid(strWord, intPos) = "O" Then
                    variants.Add Array("O", "OW"): intCartesianCount = intCartesianCount + 1
                Else
                    variants.Add Array("A", "AR"): intCartesianCount = intCartesianCount + 1
                End If
                intPos = intPos + 1
            Else
                variants.Add Array(Mid(strWord, intPos, 1), "")
                
                intPos = intPos + 1
            End If
        Loop
    End If
    
    Dim j As Integer

    Dim curVariantsList As New Collection
    Dim newVariantsList As Collection
    If variants(1)(1) = "" Then
        curVariantsList.Add variants(1)(0)
    Else
        curVariantsList.Add variants(1)(0)
        curVariantsList.Add variants(1)(1)
    End If

    For i = 2 To variants.Count
        Set newVariantsList = New Collection
        
        If variants(i)(1) = "" Then
            For j = 1 To curVariantsList.Count
                newVariantsList.Add curVariantsList(j) & variants(i)(0)
            Next
        Else
            For j = 1 To curVariantsList.Count
                newVariantsList.Add curVariantsList(j) & variants(i)(0)
                newVariantsList.Add curVariantsList(j) & variants(i)(1)
            Next
        End If
        
        Set curVariantsList = newVariantsList
    Next
    
    Dim encodedList As New Collection

    On Error Resume Next 'get uniques
    For i = 1 To curVariantsList.Count
        encodedList.Add HaaseCode(curVariantsList.item(i)), HaaseCode(curVariantsList.item(i))
    Next
    
    Dim strEncodedList() As String
    ReDim strEncodedList(1 To encodedList.Count)
    
    For i = 1 To encodedList.Count
        strEncodedList(i) = encodedList.item(i)
    Next
    
    If UBound(encodedListString) > 1 Then
        Haase = Join(strEncodedList, ",")
    Else
        Haase = strEncodedList(1)
    End If
End Function
Function After(strWord As String, intPos As Integer, strLetters As String) As Boolean
    If intPos > 1 Then
        After = Mid(strWord, intPos - 1, 1) Like "[" & strLetters & "]"
    Else
        After = False
    End If
End Function
Function Before(strWord, intPos As Integer, strLetters As String) As Boolean
    If intPos + 1 <= Len(strWord) Then
        Before = Mid(strWord, intPos + 1, 1) Like "[" & strLetters & "]"
    Else
        Before = False
    End If
End Function
Function HaaseCode(strWord As String) As String
    Dim strSoundex As String
    Dim i As Integer
    
    For i = 1 To Len(strWord)
        If Mid(strWord, i, 1) Like "[AEIJOUY]" Then
            strSoundex = strSoundex & "9"
        ElseIf Mid(strWord, i, 1) = "B" Then
            strSoundex = strSoundex & "1"
        ElseIf Mid(strWord, i, 1) = "P" Then
            If Before(strWord, i, "H") Then
                strSoundex = strSoundex & "3"
            Else
                strSoundex = strSoundex & "1"
            End If
        ElseIf Mid(strWord, i, 1) Like "[DT]" Then
            If Before(strWord, i, "CSZ") Then
                strSoundex = strSoundex & "8"
            Else
                strSoundex = strSoundex & "2"
            End If
        ElseIf Mid(strWord, i, 1) Like "[FVW]" Then
            strSoundex = strSoundex & "3"
        ElseIf Mid(strWord, i, 1) Like "[GKQ]" Then
            strSoundex = strSoundex & "4"
        ElseIf Mid(strWord, i, 1) = "C" Then
            If After(strWord, i, "SZ") Then
                strSoundex = strSoundex & "8"
            ElseIf i = 1 Then
                If Before(strWord, i, "AHKLOQRUX") Then
                    strSoundex = strSoundex & "4"
                Else
                    strSoundex = strSoundex & "8"
                End If
            ElseIf Before(strWord, i, "AHKOQUX") Then
                strSoundex = strSoundex & "4"
            Else
                strSoundex = strSoundex & "8"
            End If
        ElseIf Mid(strWord, i, 1) = "X" Then
            If After(strWord, i, "CKQ") Then
                 strSoundex = strSoundex & "8"
            Else
                strSoundex = strSoundex & "48"
            End If
        ElseIf Mid(strWord, i, 1) = "L" Then
            strSoundex = strSoundex & "5"
        ElseIf Mid(strWord, i, 1) Like "[MN]" Then
            strSoundex = strSoundex & "6"
        ElseIf Mid(strWord, i, 1) = "R" Then
            strSoundex = strSoundex & "7"
        ElseIf Mid(strWord, i, 1) Like "[SZ]" Then
            strSoundex = strSoundex & "8"
        End If
    
    Next
    strSoundex = DeleteConsecutiveRepeats(strSoundex)
    HaaseCode = strSoundex
End Function
