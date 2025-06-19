Attribute VB_Name = "DaitchMokotoff_Encode"
Option Explicit

Function DaitchMokotoff(strWord As String, Optional intMaxLength As Integer = 6, Optional boolZeroPad As Boolean = True)
    If intMaxLength <> -1 Then
        If intMaxLength > 64 Then intMaxLength = 64
        If intMaxLength < 6 Then intMaxLength = 6
    Else
        intMaxLength = 64
    End If
    
    Dim dms_table As New Dictionary
    dms_table.Add "STCH", Array(2, 4, 4)
    dms_table.Add "DRZ", Array(4, 4, 4)
    dms_table.Add "ZH", Array(4, 4, 4)
    dms_table.Add "ZHDZH", Array(2, 4, 4)
    dms_table.Add "DZH", Array(4, 4, 4)
    dms_table.Add "DRS", Array(4, 4, 4)
    dms_table.Add "DZS", Array(4, 4, 4)
    dms_table.Add "SCHTCH", Array(2, 4, 4)
    dms_table.Add "SHTSH", Array(2, 4, 4)
    dms_table.Add "SZCZ", Array(2, 4, 4)
    dms_table.Add "TZS", Array(4, 4, 4)
    dms_table.Add "SZCS", Array(2, 4, 4)
    dms_table.Add "STSH", Array(2, 4, 4)
    dms_table.Add "SHCH", Array(2, 4, 4)
    dms_table.Add "D", Array(3, 3, 3)
    dms_table.Add "H", Array(5, 5, "_")
    dms_table.Add "TTSCH", Array(4, 4, 4)
    dms_table.Add "THS", Array(4, 4, 4)
    dms_table.Add "L", Array(8, 8, 8)
    dms_table.Add "P", Array(7, 7, 7)
    dms_table.Add "CHS", Array(5, 54, 54)
    dms_table.Add "T", Array(3, 3, 3)
    dms_table.Add "X", Array(5, 54, 54)
    dms_table.Add "OJ", Array(0, 1, "_")
    dms_table.Add "OI", Array(0, 1, "_")
    dms_table.Add "SCHTSH", Array(2, 4, 4)
    dms_table.Add "OY", Array(0, 1, "_")
    dms_table.Add "Y", Array(1, "_", "_")
    dms_table.Add "TSH", Array(4, 4, 4)
    dms_table.Add "ZDZ", Array(2, 4, 4)
    dms_table.Add "TSZ", Array(4, 4, 4)
    dms_table.Add "SHT", Array(2, 43, 43)
    dms_table.Add "SCHTSCH", Array(2, 4, 4)
    dms_table.Add "TTSZ", Array(4, 4, 4)
    dms_table.Add "TTZ", Array(4, 4, 4)
    dms_table.Add "SCH", Array(4, 4, 4)
    dms_table.Add "TTS", Array(4, 4, 4)
    dms_table.Add "SZD", Array(2, 43, 43)
    dms_table.Add "AI", Array(0, 1, "_")
    dms_table.Add "PF", Array(7, 7, 7)
    dms_table.Add "TCH", Array(4, 4, 4)
    dms_table.Add "PH", Array(7, 7, 7)
    dms_table.Add "TTCH", Array(4, 4, 4)
    dms_table.Add "SZT", Array(2, 43, 43)
    dms_table.Add "ZDZH", Array(2, 4, 4)
    dms_table.Add "EI", Array(0, 1, "_")
    dms_table.Add "G", Array(5, 5, 5)
    dms_table.Add "EJ", Array(0, 1, "_")
    dms_table.Add "ZD", Array(2, 43, 43)
    dms_table.Add "IU", Array(1, "_", "_")
    dms_table.Add "K", Array(5, 5, 5)
    dms_table.Add "O", Array(0, "_", "_")
    dms_table.Add "SHTCH", Array(2, 4, 4)
    dms_table.Add "S", Array(4, 4, 4)
    dms_table.Add "TRZ", Array(4, 4, 4)
    dms_table.Add "SHD", Array(2, 43, 43)
    dms_table.Add "DSH", Array(4, 4, 4)
    dms_table.Add "CSZ", Array(4, 4, 4)
    dms_table.Add "EU", Array(1, 1, "_")
    dms_table.Add "TRS", Array(4, 4, 4)
    dms_table.Add "ZS", Array(4, 4, 4)
    dms_table.Add "STRZ", Array(2, 4, 4)
    dms_table.Add "UY", Array(0, 1, "_")
    dms_table.Add "STRS", Array(2, 4, 4)
    dms_table.Add "CZS", Array(4, 4, 4)
    dms_table.Add "MN", Array("6_6", "6_6", "6_6")
    dms_table.Add "UI", Array(0, 1, "_")
    dms_table.Add "UJ", Array(0, 1, "_")
    dms_table.Add "UE", Array(0, "_", "_")
    dms_table.Add "EY", Array(0, 1, "_")
    dms_table.Add "W", Array(7, 7, 7)
    dms_table.Add "IA", Array(1, "_", "_")
    dms_table.Add "FB", Array(7, 7, 7)
    dms_table.Add "STSCH", Array(2, 4, 4)
    dms_table.Add "SCHT", Array(2, 43, 43)
    dms_table.Add "NM", Array("6_6", "6_6", "6_6")
    dms_table.Add "SCHD", Array(2, 43, 43)
    dms_table.Add "B", Array(7, 7, 7)
    dms_table.Add "DSZ", Array(4, 4, 4)
    dms_table.Add "F", Array(7, 7, 7)
    dms_table.Add "N", Array(6, 6, 6)
    dms_table.Add "CZ", Array(4, 4, 4)
    dms_table.Add "R", Array(9, 9, 9)
    dms_table.Add "U", Array(0, "_", "_")
    dms_table.Add "V", Array(7, 7, 7)
    dms_table.Add "CS", Array(4, 4, 4)
    dms_table.Add "Z", Array(4, 4, 4)
    dms_table.Add "SZ", Array(4, 4, 4)
    dms_table.Add "TSCH", Array(4, 4, 4)
    dms_table.Add "KH", Array(5, 5, 5)
    dms_table.Add "ST", Array(2, 43, 43)
    dms_table.Add "KS", Array(5, 54, 54)
    dms_table.Add "SH", Array(4, 4, 4)
    dms_table.Add "SC", Array(2, 4, 4)
    dms_table.Add "SD", Array(2, 43, 43)
    dms_table.Add "DZ", Array(4, 4, 4)
    dms_table.Add "ZHD", Array(2, 43, 43)
    dms_table.Add "DT", Array(3, 3, 3)
    dms_table.Add "ZSH", Array(4, 4, 4)
    dms_table.Add "DS", Array(4, 4, 4)
    dms_table.Add "TZ", Array(4, 4, 4)
    dms_table.Add "TS", Array(4, 4, 4)
    dms_table.Add "TH", Array(3, 3, 3)
    dms_table.Add "TC", Array(4, 4, 4)
    dms_table.Add "A", Array(0, "_", "_")
    dms_table.Add "E", Array(0, "_", "_")
    dms_table.Add "I", Array(0, "_", "_")
    dms_table.Add "AJ", Array(0, 1, "_")
    dms_table.Add "M", Array(6, 6, 6)
    dms_table.Add "Q", Array(5, 5, 5)
    dms_table.Add "AU", Array(0, 7, "_")
    dms_table.Add "IO", Array(1, "_", "_")
    dms_table.Add "AY", Array(0, 1, "_")
    dms_table.Add "IE", Array(1, "_", "_")
    dms_table.Add "ZSCH", Array(4, 4, 4)
    dms_table.Add "CH", Array(Array(5, 4), Array(5, 4), Array(5, 4))
    dms_table.Add "CK", Array(Array(5, 45), Array(5, 45), Array(5, 45))
    dms_table.Add "C", Array(Array(5, 4), Array(5, 4), Array(5, 4))
    dms_table.Add "J", Array(Array(1, 4), Array("_", 4), Array("_", 4))
    dms_table.Add "RZ", Array(Array(94, 4), Array(94, 4), Array(94, 4))
    dms_table.Add "RS", Array(Array(94, 4), Array(94, 4), Array(94, 4))
    
    Dim dms_order As New Dictionary
    dms_order.Add "A", Array("AI", "AJ", "AU", "AY", "A")
    dms_order.Add "B", Array("B")
    dms_order.Add "C", Array("CHS", "CSZ", "CZS", "CH", "CK", "CS", "CZ", "C")
    dms_order.Add "D", Array("DRS", "DRZ", "DSH", "DSZ", "DZH", "DZS", "DS", "DT", "DZ", "D")
    dms_order.Add "E", Array("EI", "EJ", "EU", "EY", "E")
    dms_order.Add "F", Array("FB", "F")
    dms_order.Add "G", Array("G")
    dms_order.Add "H", Array("H")
    dms_order.Add "I", Array("IA", "IE", "IO", "IU", "I")
    dms_order.Add "J", Array("J")
    dms_order.Add "K", Array("KH", "KS", "K")
    dms_order.Add "L", Array("L")
    dms_order.Add "M", Array("MN", "M")
    dms_order.Add "N", Array("NM", "N")
    dms_order.Add "O", Array("OI", "OJ", "OY", "O")
    dms_order.Add "P", Array("PF", "PH", "P")
    dms_order.Add "Q", Array("Q")
    dms_order.Add "R", Array("RS", "RZ", "R")
    dms_order.Add "S", Array("SCHTSCH", "SCHTCH", "SCHTSH", "SHTCH", "SHTSH", "STSCH", "SCHD", "SCHT", "SHCH", "STCH", "STRS", "STRZ", "STSH", "SZCS", "SZCZ", "SCH", "SHD", "SHT", "SZD", "SZT", "SC", "SD", "SH", "ST", "SZ", "S")
    dms_order.Add "T", Array("TTSCH", "TSCH", "TTCH", "TTSZ", "TCH", "THS", "TRS", "TRZ", "TSH", "TSZ", "TTS", "TTZ", "TZS", "TC", "TH", "TS", "TZ", "T")
    dms_order.Add "U", Array("UE", "UI", "UJ", "UY", "U")
    dms_order.Add "V", Array("V")
    dms_order.Add "W", Array("W")
    dms_order.Add "X", Array("X")
    dms_order.Add "Y", Array("Y")
    dms_order.Add "Z", Array("ZHDZH", "ZDZH", "ZSCH", "ZDZ", "ZHD", "ZSH", "ZD", "ZH", "ZS", "Z")

    strWord = PhoneticFunctions.GetAlphaOnly(strWord)
    strWord = UCase$(strWord)
    
    Dim sstr As Variant

    
    Dim dm_tup As Variant
    Dim dm_val As Variant
    
    Dim dms As New Collection
    dms.Add ""
    Dim tmp As Variant
    Dim i As Long
    Dim tmpCount As Integer
    Dim pos As Long
    pos = 1
     While pos <= Len(strWord)
        For Each sstr In dms_order(Mid(strWord, pos, 1))
            If Mid(strWord, pos, Len(sstr)) = sstr Then
                dm_tup = dms_table(sstr)

                If pos = 1 Then
                    dm_val = dm_tup(0)
                ElseIf pos + Len(sstr) <= Len(strWord) And InStr(1, "AEIJOUY", Mid(strWord, pos + Len(sstr), 1)) > 0 Then 'Added J and Y here
                    dm_val = dm_tup(1)
                Else
                    dm_val = dm_tup(2)
                End If
                
                If IsArray(dm_val) Then
                    If dms.Count = 1 And pos = 1 Then
                        dms.Remove 1
                        dms.Add dm_val(0)
                        dms.Add dm_val(1)
                    Else
                        tmpCount = dms.Count
                        For i = 1 To dms.Count
                            dms.Add dms(i) & dm_val(0)
                            dms.Add dms(i) & dm_val(1)
                        Next i
                        For i = 1 To tmpCount
                            dms.Remove 1
                        Next
                        
                    End If
                Else
                    tmpCount = dms.Count
                    For i = 1 To tmpCount
                        tmp = dms(1) & dm_val
                        dms.Add tmp
                        dms.Remove 1
                    Next
                End If
                pos = pos + Len(sstr)
                Exit For
            End If
        Next sstr
    Wend
    
    Dim filteredDMS As Collection
    Set filteredDMS = New Collection
    Dim tempStr As String
    Dim code As Variant
    For Each code In dms
        tempStr = PhoneticFunctions.DeleteConsecutiveRepeats(code)
        tempStr = Replace(tempStr, "_", "")
        If boolZeroPad Then tempStr = tempStr & String(intMaxLength, "0")
        filteredDMS.Add Left(tempStr, intMaxLength)
    Next code

    'sort
    Dim j As Long
    For i = 1 To filteredDMS.Count
        For j = i + 1 To filteredDMS.Count
            If CLng(filteredDMS(i)) > CLng(filteredDMS(j)) Then
                filteredDMS.Add filteredDMS(j), , , i
                filteredDMS.Add filteredDMS(i), , , j + 1
                filteredDMS.Remove i
                filteredDMS.Remove j
            End If
        Next
    Next
    
    Dim strResult As String
    strResult = ""
    For Each code In filteredDMS
        If InStr(strResult, code) = 0 Then
            If strResult <> "" Then
                strResult = strResult & ","
            End If
            strResult = strResult & code
        End If
    Next code
    
    DaitchMokotoff = strResult
End Function
