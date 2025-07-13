Attribute VB_Name = "Encode_RussellIndex"
Option Explicit
Const strAllowed As String = "ABCDEFGIKLMNOPQRSTUVXYZ"
Const StrTranslate As String = "12341231356712383412313"
Function RussellIndex(strInput As String)
    Dim strEncoding As String
    Dim strCode As String
    Dim i As Long
    
    strInput = UCase$(strInput)
    strInput = Replace$(strInput, "GH", "")
    If right$(strInput, 1) Like "[SZ]" Then strInput = left$(strInput, Len(strInput) - 1)
    
    Dim strSoundex As String
    Dim strChar As String
    For i = 1 To Len(strInput)
        strChar = Mid$(strInput, i, 1)
        If InStr(strAllowed, strChar) > 0 Then strSoundex = strSoundex & Mid$(StrTranslate, InStr(strAllowed, strChar), 1)
    Next
    
    'Remove any 1 after the first 1
    Dim intOne As Integer
    intOne = InStr(strSoundex, "1")
    If intOne > 0 Then strSoundex = left(strSoundex, intOne) & Replace(Mid(strSoundex, intOne), "1", "")
    
    strSoundex = DeleteConsecutiveRepeats(strSoundex)
    
    RussellIndex = strSoundex
End Function

