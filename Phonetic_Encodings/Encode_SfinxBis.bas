Attribute VB_Name = "Encode_SfinxBis"
Option Explicit
Const strTransIn As String = "BCDFGHJKLMNPQRSTVZAOU≈EIYƒ÷"
Const strTransOut As String = "123729224551268378999999999"
Function SfinxBis(strWord As String, Optional intMaxLength As Integer = -1) As String
    'Step 1
    strWord = UCase(strWord)
    strWord = Replace(strWord, "-", " ")
    
    'Step 2
    Dim adelstitel As Variant
    For Each adelstitel In Array(" DE LA ", " DE LAS ", " DE LOS ", " VAN DE ", " VAN DEN ", "VAN DER ", " VON DEM ", " VON DER ", " AF ", _
    " AV ", " DA ", " DE ", "DEL ", "DEN", " DES ", " DI ", " DO", " DON ", " DOS", " DU ", " E ", " IN ", " LA ", " LE ", " MAC ", " MC ", " VAN ", " VON ", " Y ", " S:T ")
        If InStr(strWord, adelstitel) > 0 Then
            strWord = Replace(strWord, adelstitel, " ")
        End If
        
        If Left(strWord, Len(Mid(adelstitel, 2))) = Mid(adelstitel, 2) Then
            strWord = Trim(Mid(strWord, Len(Mid(adelstitel, 2))))
        End If
    Next
    
    Dim strOrdLista() As String
    strOrdLista = Split(strWord, " ")
    
    'Step 3
    Dim i As Integer
    Dim t As Integer
    For i = 0 To UBound(strOrdLista)
        strOrdLista(i) = DeleteConsecutiveRepeats(strOrdLista(i))
    Next
    
    'Step 4
    For i = 0 To UBound(strOrdLista)
        strOrdLista(i) = foersvensker(strOrdLista(i))
    Next
    
    'step 5 for each item only allow letters in the uc set
    For i = 0 To UBound(strOrdLista)
        For t = 1 To Len(strOrdLista(i))
            If Not Mid(strOrdLista(i), t, 1) Like "[A-Zƒ≈÷]" Then
                Mid(strOrdLista(i), t, 1) = " "
            End If
        Next
        strOrdLista(i) = Replace(strOrdLista(i), " ", "")
    Next
    
    'step 6
    For i = 0 To UBound(strOrdLista)
        strOrdLista(i) = kodaFoerstaljudet(strOrdLista(i))
    Next
    
    'step 7
    Dim strRest() As String: ReDim strRest(UBound(strOrdLista))
    For i = 0 To UBound(strRest)
        strRest(i) = Mid(strOrdLista(i), 2)
    Next
    
    'Step 8
    For i = 0 To UBound(strRest)
        strRest(i) = Replace(strRest(i), "DT", "T")
        strRest(i) = Replace(strRest(i), "X", "KS")
    Next
    
    'Step 9
    Dim vokal As Variant
    For Each vokal In Array("E", "I", "Y", "ƒ", "÷")
        For i = 0 To UBound(strRest)
            strRest(i) = Replace(strRest(i), "C" & vokal, "8" & vokal)
        Next
    Next
    
    
    For i = 0 To UBound(strRest)
        For t = 1 To Len(strRest(i))
            Mid(strRest(i), t, 1) = Mid(strTransOut, InStr(strTransIn, Mid(strRest(i), t, 1)), 1)
        Next
    Next
    
    'Step 10
    For i = 0 To UBound(strRest)
        strRest(i) = DeleteConsecutiveRepeats(strRest(i))
    Next
    
    'Step 11
    For i = 0 To UBound(strRest)
        strRest(i) = Replace(strRest(i), "9", "")
    Next
    
    'step12
    For i = 0 To UBound(strOrdLista)
        strOrdLista(i) = Left(strOrdLista(i), 1) & strRest(i)
    Next
    SfinxBis = Join(strOrdLista, ",")
End Function
Function foersvensker(strLokalOrdet As String) As String
    strLokalOrdet = Replace(strLokalOrdet, "STIERN", "STJƒRN")
    strLokalOrdet = Replace(strLokalOrdet, "HIE", "HJ")
    strLokalOrdet = Replace(strLokalOrdet, "SI÷", "SJ÷")
    strLokalOrdet = Replace(strLokalOrdet, "SCH", "SH")
    strLokalOrdet = Replace(strLokalOrdet, "QU", "KV")
    strLokalOrdet = Replace(strLokalOrdet, "IO", "JO")
    strLokalOrdet = Replace(strLokalOrdet, "PH", "F")
    
    Dim vokaler As Variant
    For Each vokaler In Array("A", "O", "U", "≈") 'harde vokaler
        strLokalOrdet = Replace(strLokalOrdet, vokaler & "‹", vokaler & "J")
        strLokalOrdet = Replace(strLokalOrdet, vokaler & "Y ", vokaler & "J")
        strLokalOrdet = Replace(strLokalOrdet, vokaler & "I", vokaler & "J")
    Next
    
    For Each vokaler In Array("E", "I", "Y", "ƒ", "÷") 'mjuka vokaler
        strLokalOrdet = Replace(strLokalOrdet, vokaler & "‹", vokaler & "J")
        strLokalOrdet = Replace(strLokalOrdet, vokaler & "Y ", vokaler & "J")
        strLokalOrdet = Replace(strLokalOrdet, vokaler & "I", vokaler & "J")
    Next
    
    If InStr(strLokalOrdet, "H") > 0 Then
        For Each vokaler In Array("B", "C", "D", "F", "G", "H", "J", "K", "L", "M", "N", "P", "Q", "R", "S", "T", "V", "W", "X", "Z")
            strLokalOrdet = Replace(strLokalOrdet, "H" & vokaler, vokaler)
        Next
    End If
    
    Dim i As Integer
    For i = 1 To Len(strLokalOrdet)
        If Mid(strLokalOrdet, i, 1) Like "[WZ¿¡¬√∆«»… ÀÃÕŒœ—“”‘’ÿŸ⁄€‹›]" Then
            Mid(strLokalOrdet, i, 1) = Mid("VSAAAAƒCEEEEIIIINOOOO÷UUUYY", InStr("WZ¿¡¬√∆«»… ÀÃÕŒœ—“”‘’ÿŸ⁄€‹›", Mid(strLokalOrdet, i, 1)), 1)
        End If
    Next
    
    strLokalOrdet = Replace(strLokalOrdet, "–", "ETH")
    strLokalOrdet = Replace(strLokalOrdet, "ﬁ", "TH")
    
    foersvensker = strLokalOrdet
End Function
Function kodaFoerstaljudet(strLokalOrdet As String)
    If Left(strLokalOrdet, 1) Like "[AOU≈EIYƒ÷]" Then
        Mid(strLokalOrdet, 1, 1) = "$"
    ElseIf Left(strLokalOrdet, 2) Like "[DGHL]J" Then
        strLokalOrdet = "J" & Mid(strLokalOrdet, 3)
    ElseIf Left(strLokalOrdet, 1) = "G" And Mid(strLokalOrdet, 2, 1) Like "[EIYƒ÷]" Then
        strLokalOrdet = "J" & Mid(strLokalOrdet, 2)
    ElseIf Left(strLokalOrdet, 1) = "Q" Then
        Mid(strLokalOrdet, 1, 1) = "K"
    ElseIf Left(strLokalOrdet, 2) = "CH" And Mid(strLokalOrdet, 3, 1) Like "[AOU≈EIYƒ÷]" Then
        strLokalOrdet = "#" & Mid(strLokalOrdet, 3)
    ElseIf Left(strLokalOrdet, 1) = "C" And Mid(strLokalOrdet, 2, 1) Like "[AOU≈]" Then
        strLokalOrdet = "K" & Mid(strLokalOrdet, 2)
    ElseIf Left(strLokalOrdet, 1) = "C" And Mid(strLokalOrdet, 2, 1) Like "[BCDFGHJKLMNPQRSTVWXZ]" Then
        strLokalOrdet = "K" & Mid(strLokalOrdet, 2)
    ElseIf Left(strLokalOrdet, 1) = "X" Then
        strLokalOrdet = "S" & Mid(strLokalOrdet, 2)
    ElseIf Left(strLokalOrdet, 1) = "C" And Mid(strLokalOrdet, 2, 1) Like "[EIYƒ÷]" Then
        strLokalOrdet = "S" & Mid(strLokalOrdet, 2)
    ElseIf Left(strLokalOrdet, 3) = "SKJ" Or Left(strLokalOrdet, 3) = "STJ" Or Left(strLokalOrdet, 3) = "SCH" Then
        strLokalOrdet = "#" & Mid(strLokalOrdet, 4)
    ElseIf Left(strLokalOrdet, 2) = "SH" Or Left(strLokalOrdet, 2) Like "[KTS]J" Then
        strLokalOrdet = "#" & Mid(strLokalOrdet, 3)
    ElseIf Left(strLokalOrdet, 2) = "SK" And Mid(strLokalOrdet, 3, 1) Like "[EIYƒ÷]" Then 'mjuka
        strLokalOrdet = "#" & Mid(strLokalOrdet, 3)
    ElseIf Left(strLokalOrdet, 1) = "K" And Mid(strLokalOrdet, 2, 1) Like "[EIYƒ÷]" Then 'mjuka
        strLokalOrdet = "#" & Mid(strLokalOrdet, 2)
    End If
    kodaFoerstaljudet = strLokalOrdet
End Function
