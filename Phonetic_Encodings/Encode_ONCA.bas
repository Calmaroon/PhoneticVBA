Attribute VB_Name = "Encode_ONCA"
Option Explicit
Function ONCA(strWord As String, Optional intMaxLength As Integer = 5, Optional boolZeroPad As Boolean = True) As String
    ONCA = Soundex(NYSIIS(strWord, intMaxLength, boolZeroPad))
End Function
