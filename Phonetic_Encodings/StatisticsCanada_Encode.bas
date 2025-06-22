Attribute VB_Name = "StatisticsCanada_Encode"
Option Explicit
Function StatisticsCanada(strWord As String, Optional intMaxLength As Integer = 4) As String
    If intMaxLength < 1 Then intMaxLength = 1
    strWord = UnicodeFunctions.UnicodeStrip(strWord)
    strWord = PhoneticFunctions.GetAlphaOnly(UCase$(strWord))

    If strWord = "" Then
        StatisticsCanada = ""
        Exit Function
    End If

    Dim strCode As String
    strCode = Mid(strWord, 2)
    
    strCode = Replace(strCode, "A", "")
    strCode = Replace(strCode, "E", "")
    strCode = Replace(strCode, "I", "")
    strCode = Replace(strCode, "O", "")
    strCode = Replace(strCode, "U", "")
    
    strCode = Left(strWord, 1) & strCode
    
    strCode = PhoneticFunctions.DeleteConsecutiveRepeats(strCode)
    strCode = Replace(strCode, " ", "")
    
    StatisticsCanada = Left(strCode, intMaxLength)

End Function
