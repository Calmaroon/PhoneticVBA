Attribute VB_Name = "Encode_AlphaSIS"
Option Explicit
Function AlphaSIS(strWord As String, Optional intMaxLength As Integer = 14) As String
    If intMaxLength < 0 Then intMaxLength = 64
    
    Dim alphaSISInitials As New Collection
        alphaSISInitials.Add Array("GF", "08")
        alphaSISInitials.Add Array("GM", "03")
        alphaSISInitials.Add Array("GN", "02")
        alphaSISInitials.Add Array("KN", "02")
        alphaSISInitials.Add Array("PF", "08")
        alphaSISInitials.Add Array("PN", "02")
        alphaSISInitials.Add Array("PS", "00")
        alphaSISInitials.Add Array("WR", "04")
        alphaSISInitials.Add Array("A", "1")
        alphaSISInitials.Add Array("E", "1")
        alphaSISInitials.Add Array("H", "2")
        alphaSISInitials.Add Array("I", "1")
        alphaSISInitials.Add Array("J", "3")
        alphaSISInitials.Add Array("O", "1")
        alphaSISInitials.Add Array("U", "1")
        alphaSISInitials.Add Array("W", "4")
        alphaSISInitials.Add Array("Y", "5")
        
    Dim alphaSISBasic As New Collection
        alphaSISBasic.Add Array("SCH", "6")
        alphaSISBasic.Add Array("CZ", Array("70", "6", "0"))
        alphaSISBasic.Add Array("CH", Array("6", "70", "0"))
        alphaSISBasic.Add Array("CK", Array("7", "6"))
        alphaSISBasic.Add Array("DS", Array("0", "10"))
        alphaSISBasic.Add Array("DZ", Array("0", "10"))
        alphaSISBasic.Add Array("TS", Array("0", "10"))
        alphaSISBasic.Add Array("TZ", Array("0", "10"))
        alphaSISBasic.Add Array("CI", "0")
        alphaSISBasic.Add Array("CY", "0")
        alphaSISBasic.Add Array("CE", "0")
        alphaSISBasic.Add Array("SH", "6")
        alphaSISBasic.Add Array("DG", "7")
        alphaSISBasic.Add Array("PH", "8")
        alphaSISBasic.Add Array("C", Array("7", "6"))
        alphaSISBasic.Add Array("K", Array("7", "6"))
        alphaSISBasic.Add Array("Z", "0")
        alphaSISBasic.Add Array("S", "0")
        alphaSISBasic.Add Array("D", "1")
        alphaSISBasic.Add Array("T", "1")
        alphaSISBasic.Add Array("N", "2")
        alphaSISBasic.Add Array("M", "3")
        alphaSISBasic.Add Array("R", "4")
        alphaSISBasic.Add Array("L", "5")
        alphaSISBasic.Add Array("J", "6")
        alphaSISBasic.Add Array("G", "7")
        alphaSISBasic.Add Array("Q", "7")
        alphaSISBasic.Add Array("X", "7")
        alphaSISBasic.Add Array("F", "8")
        alphaSISBasic.Add Array("V", "8")
        alphaSISBasic.Add Array("B", "9")
        alphaSISBasic.Add Array("P", "9")

    Dim alphaCollection As New Collection
    
    Dim intPos As Integer
    intPos = 1
    
    strWord = UCase(strWord)
    strWord = GetAlphaOnly(strWord)
    
    Dim i As Integer
    For i = 1 To alphaSISInitials.Count
        If Left(strWord, Len(alphaSISInitials.item(i)(0))) = alphaSISInitials.item(i)(0) Then
            alphaCollection.Add alphaSISInitials.item(i)(1)
            intPos = intPos + Len(alphaSISInitials.item(i)(0))
            Exit For
        End If
    Next

    If alphaCollection.Count = 0 Then alphaCollection.Add "0"

    Dim intOrigPos  As Integer
    Dim tmpAlphaCollection As Collection
    
    Dim d As Long
    Dim t As Long
    Do While intPos <= Len(strWord)
        intOrigPos = intPos
        
        For i = 1 To alphaSISBasic.Count
            If Left$(Mid(strWord, intPos), Len(alphaSISBasic(i)(0))) = alphaSISBasic(i)(0) Then
                Set tmpAlphaCollection = New Collection
                
                If IsArray(alphaSISBasic(i)(1)) Then
                    For t = 1 To alphaCollection.Count
                         For d = 0 To UBound(alphaSISBasic(i)(1))
                            tmpAlphaCollection.Add alphaCollection.item(t) & alphaSISBasic(i)(1)(d)
                         Next
                    Next
                Else
                    For t = 1 To alphaCollection.Count
                        tmpAlphaCollection.Add alphaCollection(t) & alphaSISBasic(i)(1)
                    Next
                End If
                intPos = intPos + Len(alphaSISBasic(i)(0))
                
                Set alphaCollection = tmpAlphaCollection
                Exit For
            End If
        Next
        
        'handle when it doesnt find anything and put in an underscore
        If intPos = intOrigPos Then
            Set tmpAlphaCollection = New Collection
            For i = 1 To alphaCollection.Count
                tmpAlphaCollection.Add alphaCollection.item(i) & "_"
            Next
            
            Set alphaCollection = tmpAlphaCollection
            intPos = intPos + 1
        End If
    Loop
    
    Dim strAlphaSIS() As String
    ReDim strAlphaSIS(1 To alphaCollection.Count)
    'Go thru each item and remove any repeat consecutives
    For i = 1 To alphaCollection.Count
        intPos = 2
        strAlphaSIS(i) = alphaCollection(i)
        Do While intPos <= Len(strAlphaSIS(i))
            If Mid(strAlphaSIS(i), intPos, 1) = Mid$(strAlphaSIS(i), intPos - 1, 1) Then
                strAlphaSIS(i) = Left$(strAlphaSIS(i), intPos - 1) & Mid(strAlphaSIS(i), intPos + 1)
            End If
            intPos = intPos + 1
        Loop
        strAlphaSIS(i) = Left(Replace(strAlphaSIS(i), "_", "") & String(intMaxLength, "0"), intMaxLength)
    Next
    
    
    AlphaSIS = Join(strAlphaSIS, ",")
    
End Function
