Attribute VB_Name = "Phonic_Encode"
Option Explicit
Const strTranscodeIn = "DTNMRLJCKGQXFVBPSZ"
Const strTranscodeOut = "112345677777889900"
Function PHONIC(strWord As String, Optional intMaxLength As Integer = 5, Optional boolZeroPad As Boolean = True, Optional boolExtended As Boolean = False) As String
    Dim dTrans2 As New Dictionary
    dTrans2.Add "CH", "6"
    dTrans2.Add "SH", "6"
    dTrans2.Add "PH", "8"
    dTrans2.Add "CE", "0"
    dTrans2.Add "CI", "0"
    dTrans2.Add "CY", "0"
    
    If intMaxLength > 0 Then
        If intMaxLength < 5 Then intMaxLength = 5
        If intMaxLength > 64 Then intMaxLength = 64
    Else
        intMaxLength = 64
    End If
    
    strWord = UCase(strWord)
    
    Dim strFirst As String
    strFirst = Left(strWord, 1)
    
    Dim strCode As String
    Dim i As Long
    i = 1
    While i < Len(strWord)
        If dTrans2.Exists(Mid(strWord, i, 2)) Then
            strCode = strCode & dTrans2(Mid(strWord, i, 2))
            i = i + 1
        ElseIf InStr(strTranscodeIn, Mid(strWord, i, 1)) > 0 Then
            strCode = strCode & Mid(strTranscodeOut, InStr(strTranscodeIn, Mid(strWord, i, 1)), 1)
        Else
            strCode = strCode & "."
        End If
        i = i + 1
    Wend
    
    strCode = Replace(PhoneticFunctions.DeleteConsecutiveRepeats(strCode), ".", "")
    
    If boolZeroPad Then
        strCode = strCode & String(intMaxLength, "0")
    End If
    
    If Not boolExtended Then
        strCode = strFirst & strCode
    End If
    
    PHONIC = Left(strCode, intMaxLength)
End Function
