Attribute VB_Name = "Encode_Phonem"
Option Explicit
Const strTransIn = "ZKGQÇÑßFWPTÁÀÂÃÅÄÆÉÈÊËIJÌÍÎÏÜÝ§ÚÙÛÔÒÓÕØ"
Const strTransOut = "CCCCCNSVVBDAAAAAEEEEEEYYYYYYYYUUUUOOOOÖ"
Const strUCSet = "ABCDLMNORSUVWXYÖ"
Function Phonem(strInput As String) As String
    strInput = UCase(strInput)
    strInput = Replace(strInput, "SC", "C")
    strInput = Replace(strInput, "SZ", "C")
    strInput = Replace(strInput, "CZ", "C")
    strInput = Replace(strInput, "TZ", "C")
    strInput = Replace(strInput, "TS", "C")
    strInput = Replace(strInput, "KS", "X")
    strInput = Replace(strInput, "PF", "V")
    strInput = Replace(strInput, "QU", "KW")
    strInput = Replace(strInput, "PH", "V")
    strInput = Replace(strInput, "UE", "Y")
    strInput = Replace(strInput, "AE", "E")
    strInput = Replace(strInput, "OE", "Ö")
    strInput = Replace(strInput, "EI", "AY")
    strInput = Replace(strInput, "EY", "AY")
    strInput = Replace(strInput, "EU", "OY")
    strInput = Replace(strInput, "AE", "A§")
    strInput = Replace(strInput, "OU", "§")
    
    Dim i As Long
    
    For i = 1 To Len(strInput)
        If InStr(strTransIn, Mid(strInput, i, 1)) > 0 Then
            Mid(strInput, i, 1) = Mid(strTransOut, InStr(strTransIn, Mid(strInput, i, 1)), 1)
        End If
    Next
    
    For i = 1 To Len(strInput)
        If InStr(strUCSet, Mid(strInput, i, 1)) = 0 Then
             Mid(strInput, i, 1) = "9"
        End If
    Next
    
    strInput = Replace(strInput, "9", "")
    strInput = DeleteConsecutiveRepeats(strInput)
    
    Phonem = strInput
End Function
