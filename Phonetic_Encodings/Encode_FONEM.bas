Attribute VB_Name = "Encode_FONEM"
Option Explicit
Function FONEM(strWord As String) As String
    Dim ruleOrder As New Collection
    Dim ruleTable As New Dictionary
    ruleOrder.Add "V-14":    ruleTable.Add "V-14", Array("([AEIOUY])(?=\1)", "")
    ruleOrder.Add "C-28":    ruleTable.Add "C-28", Array("([BDFGHJKMNPQRTVWXZ])\1", "$1")
    ruleOrder.Add "C-28a":   ruleTable.Add "C-28a", Array("CC(?=[BCDFGHJKLMNPQRSTVWXZ]|$)", "C")
    ruleOrder.Add "C-28b":   ruleTable.Add "C-28b", Array("(^|[BCDFGHJKLMNPQRSTVWXZ])SS", "$1S")
    ruleOrder.Add "C-28bb":  ruleTable.Add "C-28bb", Array("SS(?=[BCDFGHJKLMNPQRSTVWXZ]|$)", "S")
    ruleOrder.Add "C-28c":   ruleTable.Add "C-28c", Array("(^|[^I])LL", "$1L")
    ruleOrder.Add "C-28d":   ruleTable.Add "C-28d", Array("ILE$", "ILLE")
    ruleOrder.Add "C-12":    ruleTable.Add "C-12", Array("GE(O|AU)", "JO")
    ruleOrder.Add "C-8":     ruleTable.Add "C-8", Array("CC(?=[AOU])", "K")
    ruleOrder.Add "C-9":     ruleTable.Add "C-9", Array("CC(?=[EIY])", "X")
    ruleOrder.Add "C-10":    ruleTable.Add "C-10", Array("G(?=[EIY])", "J")
    ruleOrder.Add "C-16":    ruleTable.Add "C-16", Array("^MAC(?=[BCDFGHJKLMNPQRSTVWXZ])", "MA#")
    ruleOrder.Add "C-17":    ruleTable.Add "C-17", Array("^MC", "MA#")
    ruleOrder.Add "C-2":     ruleTable.Add "C-2", Array("([AEIOUY])C(?=[EIY])", "$1SS")
    
    ruleOrder.Add "C-3":     ruleTable.Add "C-3", Array("([BDFGHJKLMNPQRSTVWZ])C(?=[EIY])", "$1S")
    
    ruleOrder.Add "C-7":     ruleTable.Add "C-7", Array("C(?=[BDFGJKLMNPQRSTVWXZ])", "K")
    ruleOrder.Add "V-2,5":   ruleTable.Add "V-2,5", Array("(E?AU|O)L[TX]$", "O")
    ruleOrder.Add "V-3,4":   ruleTable.Add "V-3,4", Array("E?AU[TX]$", "O")
    ruleOrder.Add "V-6":     ruleTable.Add "V-6", Array("E?AUL?D$", "O")
    ruleOrder.Add "V-1":     ruleTable.Add "V-1", Array("E?AU", "O")
    ruleOrder.Add "C-14":    ruleTable.Add "C-14", Array("(^|[^PCS])H", "$1")
    ruleOrder.Add "C-31,33": ruleTable.Add "C-31,33", Array("^(SAINTE|STE)-?", "STE-")
    ruleOrder.Add "C-30,32": ruleTable.Add "C-30,32", Array("^(SA?INT?|SEI[NM]|CINQ?|ST)(?!E)-?", "ST-")
    ruleOrder.Add "C-11":    ruleTable.Add "C-11", Array("GA(?=I?[MN])", "G#")
    ruleOrder.Add "V-15":    ruleTable.Add "V-15", Array("[AE]M(?=[BCDFGHJKLMPQRSTVWXZ])(?!$)", "EN")
    ruleOrder.Add "V-17":    ruleTable.Add "V-17", Array("AN(?=[BCDFGHJKLMNPQRSTVWXZ])", "EN")
    ruleOrder.Add "V-18":    ruleTable.Add "V-18", Array("(AI[MN]|EIN)(?=[BCDFGHJKLMNPQRSTVWXZ]|$)", "IN")
    ruleOrder.Add "V-7":     ruleTable.Add "V-7", Array("([^G])AY$", "$1E")
    ruleOrder.Add "V-8":     ruleTable.Add "V-8", Array("EUX$", "eu")
    ruleOrder.Add "V-9":     ruleTable.Add "V-9", Array("EY(?=$|[BCDFGHJKLMNPQRSTVWXZ])", "E")
    
    ruleOrder.Add "V-10":     ruleTable.Add "V-10", Array("Y", "I")
    ruleOrder.Add "V-11":     ruleTable.Add "V-11", Array("([AEIOUY])I(?=[AEIOUY])", "$1Y")
    ruleOrder.Add "V-12":     ruleTable.Add "V-12", Array("([AEIOUY])ILL", "$1Y")
    ruleOrder.Add "V-13":     ruleTable.Add "V-13", Array("OU(?=[AEOU]|I(?!LL))", "W")
    ruleOrder.Add "V-16":     ruleTable.Add "V-16", Array("OM(?=[BCDFGHJKLMPQRSTVWXZ])", "ON")
    ruleOrder.Add "V-19":     ruleTable.Add "V-19", Array("B(O|U|OU)RNE?$", "BURN")
    ruleOrder.Add "V-20":     ruleTable.Add "V-20", Array("^IM|([BCDFGHJKLMNPQRSTVWXZ])IM(?=[BCDFGHJKLMPQRSTVWXZ])", "$1IN")
    ruleOrder.Add "C-1":     ruleTable.Add "C-1", Array("BV", "V")
    ruleOrder.Add "C-4":     ruleTable.Add "C-4", Array("^C(?=[EIY])", "S")
    ruleOrder.Add "C-5":     ruleTable.Add "C-5", Array("^C(?=[OUA])", "K")
    ruleOrder.Add "C-6":     ruleTable.Add "C-6", Array("([AEIOUY])C$", "$1K")
    ruleOrder.Add "C-13":     ruleTable.Add "C-13", Array("GNI(?=[AEIOUY])", "GN")
    ruleOrder.Add "C-15":     ruleTable.Add "C-15", Array("JEA", "JA")
    ruleOrder.Add "C-18":     ruleTable.Add "C-18", Array("PH", "F")
    ruleOrder.Add "C-19":     ruleTable.Add "C-19", Array("QU", "K")
    ruleOrder.Add "C-20":     ruleTable.Add "C-20", Array("^SC(?=[EIY])", "S")
    ruleOrder.Add "C-21":     ruleTable.Add "C-21", Array("(.)SC(?=[EIY])", "$1SS")
    ruleOrder.Add "C-22":     ruleTable.Add "C-22", Array("(.)SC(?=[AOU])", "$1SK")
    ruleOrder.Add "C-23":     ruleTable.Add "C-23", Array("SH", "CH")
    ruleOrder.Add "C-24":     ruleTable.Add "C-24", Array("TIA$", "SSIA")
    ruleOrder.Add "C-25":     ruleTable.Add "C-25", Array("([AIOUY])W", "$1")
    ruleOrder.Add "C-26":     ruleTable.Add "C-26", Array("X[CSZ]", "X")
    
    'Had to split to handle backlook
    ruleOrder.Add "C-27a":     ruleTable.Add "C-27a", Array("([BCDFGHJKLMNPQRSTVWXZ])Z(?=[BCDFGHJKLMNPQRSTVWXZ])", "$1S")
    ruleOrder.Add "C-27b":     ruleTable.Add "C-27b", Array("([AEIOUY])Z", "$1S")
    
    ruleOrder.Add "C-29":     ruleTable.Add "C-29", Array("(ILS|[CS]H|[MN]P|R[CFKLNSX])$|([BCDFGHJKLMNPQRSTVWXZ])[BCDFGHJKLMNPQRSTVWXZ]$", "$1$2")
    
    
    ruleOrder.Add "V-14"
    ruleOrder.Add "C-28"
    ruleOrder.Add "C-28a"
    ruleOrder.Add "C-28b"
    ruleOrder.Add "C-28bb"
    ruleOrder.Add "C-28c"
    ruleOrder.Add "C-28d"
    ruleOrder.Add "C-34":     ruleTable.Add "C-34", Array("G\#", "GA")
    ruleOrder.Add "C-35":     ruleTable.Add "C-35", Array("MA\#", "MAC")
    
    Dim regex As New VBScript_RegExp_55.RegExp
    regex.Global = True
    regex.IgnoreCase = True
    
    strWord = UCase(strWord)
    strWord = Replace(strWord, "Æ", "AE")
    strWord = Replace(strWord, "Œ", "OE")
    
    Dim strTranslate As String
    
    Dim i As Long
    For i = 1 To Len(strWord)
        If Mid(strWord, i, 1) Like "[A-Z-]" Then
            strTranslate = strTranslate & Mid(strWord, i, 1)
        End If
    Next

    strWord = strTranslate
    
    Dim rule As Variant
    For Each rule In ruleOrder
        regex.pattern = ruleTable(rule)(0)
        
        If regex.test(strWord) Then
            strWord = regex.Replace(strWord, ruleTable(rule)(1))
        End If
    Next

    FONEM = strWord
End Function
