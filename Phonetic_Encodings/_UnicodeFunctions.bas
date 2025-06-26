Attribute VB_Name = "_UnicodeFunctions"
Public Function UnicodeStrip(inputText As String) As String
    Dim result As String
    Dim i As Integer
    Dim char As String
    
    result = inputText
    
    ' Create arrays for search and replace
    Dim accented As Variant
    Dim unaccented As Variant
    
    accented = Array("�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", _
                     "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", _
                     "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", _
                     "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�", "�")
    
    unaccented = Array("a", "a", "a", "a", "a", "a", "ae", "c", "e", "e", "e", "e", "i", "i", "i", "i", _
                       "d", "n", "o", "o", "o", "o", "o", "o", "u", "u", "u", "u", "y", "th", "y", _
                       "A", "A", "A", "A", "A", "A", "AE", "C", "E", "E", "E", "E", "I", "I", "I", "I", _
                       "D", "N", "O", "O", "O", "O", "O", "O", "U", "U", "U", "U", "Y", "TH", "Y")
    
    For i = 0 To UBound(accented)
        result = Replace(result, accented(i), unaccented(i))
    Next i
    
    UnicodeStrip = result
End Function

