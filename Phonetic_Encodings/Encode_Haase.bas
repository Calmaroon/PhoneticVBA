Attribute VB_Name = "Encode_Haase"
Option Explicit

Function Haase(strWord As String, Optional boolPrimaryOnly As Boolean = False) As String
    Const vowels As String = "AEIJOUY"
    Const ucSet As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZÄÖÜ"

    Dim i As Long, c As String
    Dim variants As Collection
    Set variants = New Collection

    ' Step 1: Normalize input
    strWord = UCase$(strWord)
    strWord = Replace(strWord, "Ä", "AE")
    strWord = Replace(strWord, "Ö", "OE")
    strWord = Replace(strWord, "Ü", "UE")

    ' Strip any non-A-Z characters
    Dim cleanWord As String
    For i = 1 To Len(strWord)
        c = Mid(strWord, i, 1)
        If InStr(ucSet, c) > 0 Then cleanWord = cleanWord & c
    Next
    strWord = cleanWord

    If Len(strWord) = 0 Then
        Haase = ""
        Exit Function
    End If

    ' Step 2: Generate variants
    Dim pos As Long: pos = 1
    Dim vParts() As Variant
    ReDim vParts(0 To 0)
    vParts(0) = Array()

    If boolPrimaryOnly Then
        ReDim vParts(0)
        vParts(0) = Array(strWord)
    Else
        Dim chunk As String
        Do While pos <= Len(strWord)
            Dim part As Variant
            chunk = Mid(strWord, pos, 4)
            Select Case True
                Case Left(chunk, 4) = "ILLE"
                    part = Array("ILLE", "I")
                    pos = pos + 4
                Case Left(chunk, 3) = "OWN"
                    part = Array("OWN", "AUN")
                    pos = pos + 3
                Case Left(chunk, 3) = "WSK"
                    part = Array("WSK", "RSK")
                    pos = pos + 3
                Case Left(chunk, 3) = "SCH"
                    part = Array("SCH", "CH")
                    pos = pos + 3
                Case Left(chunk, 3) = "GLI"
                    part = Array("GLI", "LI")
                    pos = pos + 3
                Case Left(chunk, 3) = "AUX"
                    part = Array("AUX", "O")
                    pos = pos + 3
                Case Left(chunk, 3) = "EUX"
                    part = Array("EUX", "O")
                    pos = pos + 3
                Case Left(chunk, 2) = "RB"
                    part = Array("RB", "RW")
                    pos = pos + 2
                Case Mid(strWord, pos) = "EAU"
                    part = Array("EAU", "O")
                    pos = pos + 3
                Case Mid(strWord, pos) = "O"
                    part = Array("O", "OW")
                    pos = pos + 1
                Case Mid(strWord, pos) = "A"
                    part = Array("A", "AR")
                    pos = pos + 1
                Case Else
                    part = Array(Mid(strWord, pos, 1))
                    pos = pos + 1
            End Select
            ' Expand combinations
            Dim newParts() As Variant, vi As Long, pj As Long
            ReDim newParts(0 To (UBound(vParts) + 1) * (UBound(part) + 1) - 1)
            Dim np As Long: np = 0
            For vi = 0 To UBound(vParts)
                For pj = 0 To UBound(part)
                    newParts(np) = vParts(vi) & part(pj)
                    np = np + 1
                Next pj
            Next vi
            vParts = newParts
        Loop
    End If

    ' Step 3: Encode variants
    Dim encoded As Collection: Set encoded = New Collection
    Dim encSet As Object: Set encSet = CreateObject("Scripting.Dictionary")
    Dim variantWord As String, encodedWord As String
    Dim idx As Long
    For idx = 0 To UBound(vParts)
        If boolPrimaryOnly Then
            variantWord = strWord
        Else
            variantWord = vParts(idx)
        End If

        encodedWord = ""
        For i = 1 To Len(variantWord)
            c = Mid(variantWord, i, 1)
            Select Case True
                Case InStr(vowels, c) > 0: encodedWord = encodedWord & "9"
                Case c = "B": encodedWord = encodedWord & "1"
                Case c = "P": If i < Len(variantWord) And Mid(variantWord, i + 1, 1) = "H" Then encodedWord = encodedWord & "3" Else encodedWord = encodedWord & "1"
                Case c = "D" Or c = "T": If i < Len(variantWord) And Mid(variantWord, i + 1, 1) Like "[CSZ]" Then encodedWord = encodedWord & "8" Else encodedWord = encodedWord & "2"
                Case c Like "[FVW]": encodedWord = encodedWord & "3"
                Case c Like "[GKQ]": encodedWord = encodedWord & "4"
                Case c = "C":
                    If i > 1 And Mid(variantWord, i - 1, 1) Like "[SZ]" Then
                        encodedWord = encodedWord & "8"
                    ElseIf i = 1 Then
                        If i < Len(variantWord) And Mid(variantWord, i + 1, 1) Like "[AHKLOQRUX]" Then
                            encodedWord = encodedWord & "4"
                        Else
                            encodedWord = encodedWord & "8"
                        End If
                    ElseIf i < Len(variantWord) And Mid(variantWord, i + 1, 1) Like "[AHKOQUX]" Then
                        encodedWord = encodedWord & "4"
                    Else
                        encodedWord = encodedWord & "8"
                    End If
                Case c = "X":
                    If i > 1 And Mid(variantWord, i - 1, 1) Like "[CKQ]" Then
                        encodedWord = encodedWord & "8"
                    Else
                        encodedWord = encodedWord & "48"
                    End If
                Case c = "L": encodedWord = encodedWord & "5"
                Case c Like "[MN]": encodedWord = encodedWord & "6"
                Case c = "R": encodedWord = encodedWord & "7"
                Case c Like "[SZ]": encodedWord = encodedWord & "8"
            End Select
        Next

        ' Deduplicate digits
        encodedWord = DeleteConsecutiveRepeats(encodedWord)

        If Not encSet.Exists(encodedWord) Then
            encSet.Add encodedWord, True
            encoded.Add encodedWord
        End If
    Next

    ' Return result
    If encoded.Count = 1 Then
        Haase = encoded(1)
    Else
        Dim result As String
        For i = 1 To encoded.Count
            result = result & IIf(i > 1, ",", "") & encoded(i)
        Next
        Haase = result
    End If
End Function


