Attribute VB_Name = "Encode_MetaSoundex"
Option Explicit
Const strTransIn As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const strTransOut As String = "07430755015866075943077514"
Enum Lang
    strEn = 1
    strEs = 2
End Enum

Function MetaSoundex(strWord As String, Optional strLang As Lang = strEn) As String
    If strLang = 1 Then 'English
        strWord = Soundex(Metaphone(strWord))
        MetaSoundex = Mid(strTransOut, InStr(strTransIn, Left(strWord, 1)), 1) & Mid(strWord, 2)
    Else
        MetaSoundex = PhoneticSpanish(SpanishMetaphone(strWord))
    End If
End Function
