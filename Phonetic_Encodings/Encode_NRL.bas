Attribute VB_Name = "Encode_NRL"
Option Explicit
Dim rules As Dictionary
Dim re_right As New RegExp
Dim re_left As New RegExp
Sub buildRules()
    Set rules = New Dictionary
    
    Dim c As Collection
    
    Set c = New Collection
        c.Add Array("", " ", "", " ")
        c.Add Array("", "-", "", "")
        c.Add Array(".", Chr(34) & "S", "", "z")
        c.Add Array("#:.E", Chr(34) & "S", "", "z")
        c.Add Array("#", Chr(34) & "S", "", "z")
        c.Add Array("", Chr(34), "", "")
        c.Add Array("", ",", "", " ")
        c.Add Array("", ".", "", " ")
        c.Add Array("", "?", "", " ")
        c.Add Array("", "!", "", " ")
    rules.Add " ", c
    
    Set c = New Collection
        c.Add Array("", "A", " ", "AX")
        c.Add Array(" ", "ARE", " ", "AAr")
        c.Add Array(" ", "AR", "O", "AXr")
        c.Add Array("", "AR", "#", "EHr")
        c.Add Array("^", "AS", "#", "EYs")
        c.Add Array("", "A", "WA", "AX")
        c.Add Array("", "AW", "", "AO")
        c.Add Array(" :", "ANY", "", "EHnIY")
        c.Add Array("", "A", "^+#", "EY")
        c.Add Array("#:", "ALLY", "", "AXlIY")
        c.Add Array(" ", "AL", "#", "AXl")
        c.Add Array("", "AGAIN", "", "AXgEHn")
        c.Add Array("#:", "AG", "E", "IHj")
        c.Add Array("", "A", "^+:#", "AE")
        c.Add Array(" :", "A", "^+ ", "EY")
        c.Add Array("", "A", "^%", "EY")
        c.Add Array(" ", "ARR", "", "AXr")
        c.Add Array("", "ARR", "", "AEr")
        c.Add Array(" :", "AR", " ", "AAr")
        c.Add Array("", "AR", " ", "ER")
        c.Add Array("", "AR", "", "AAr")
        c.Add Array("", "AIR", "", "EHr")
        c.Add Array("", "AI", "", "EY")
        c.Add Array("", "AY", "", "EY")
        c.Add Array("", "AU", "", "AO")
        c.Add Array("#:", "AL", " ", "AXl")
        c.Add Array("#:", "ALS", " ", "AXlz")
        c.Add Array("", "ALK", "", "AOk")
        c.Add Array("", "AL", "^", "AOl")
        c.Add Array(" :", "ABLE", "", "EYbAXl")
        c.Add Array("", "ABLE", "", "AXbAXl")
        c.Add Array("", "ANG", "+", "EYnj")
        c.Add Array("", "A", "", "AE")
    rules.Add "A", c
    
    Set c = New Collection
        c.Add Array(" ", "BE", "^#", "bIH")
            c.Add Array("", "BEING", "", "bIYIHNG")
            c.Add Array(" ", "BOTH", " ", "bOWTH")
            c.Add Array(" ", "BUS", "#", "bIHz")
            c.Add Array("", "BUIL", "", "bIHl")
            c.Add Array("", "B", "", "b")
    rules.Add "B", c
    
    Set c = New Collection
        c.Add Array(" ", "CH", "^", "k")
        c.Add Array("^E", "CH", "", "k")
        c.Add Array("", "CH", "", "CH")
        c.Add Array(" S", "CI", "#", "sAY")
        c.Add Array("", "CI", "A", "SH")
        c.Add Array("", "CI", "O", "SH")
        c.Add Array("", "CI", "EN", "SH")
        c.Add Array("", "C", "+", "s")
        c.Add Array("", "CK", "", "k")
        c.Add Array("", "COM", "%", "kAHm")
        c.Add Array("", "C", "", "k")
    rules.Add "C", c
    
    Set c = New Collection
        c.Add Array("#:", "DED", " ", "dIHd")
        c.Add Array(".E", "D", " ", "d")
        c.Add Array("#:^E", "D", " ", "t")
        c.Add Array(" ", "DE", "^#", "dIH")
        c.Add Array(" ", "DO", " ", "dUW")
        c.Add Array(" ", "DOES", "", "dAHz")
        c.Add Array(" ", "DOING", "", "dUWIHNG")
        c.Add Array(" ", "DOW", "", "dAW")
        c.Add Array("", "DU", "A", "jUW")
        c.Add Array("", "D", "", "d")
    rules.Add "D", c
    
    Set c = New Collection
        c.Add Array("#:", "E", " ", "")
        c.Add Array("':^", "E", " ", "")
        c.Add Array(" :", "E", " ", "IY")
        c.Add Array("#", "ED", " ", "d")
        c.Add Array("#:", "E", "D ", "")
        c.Add Array("", "EV", "ER", "EHv")
        c.Add Array("", "E", "^%", "IY")
        c.Add Array("", "ERI", "#", "IYrIY")
        c.Add Array("", "ERI", "", "EHrIH")
        c.Add Array("#:", "ER", "#", "ER")
        c.Add Array("", "ER", "#", "EHr")
        c.Add Array("", "ER", "", "ER")
        c.Add Array(" ", "EVEN", "", "IYvEHn")
        c.Add Array("#:", "E", "W", "")
        c.Add Array("T", "EW", "", "UW")
        c.Add Array("S", "EW", "", "UW")
        c.Add Array("R", "EW", "", "UW")
        c.Add Array("D", "EW", "", "UW")
        c.Add Array("L", "EW", "", "UW")
        c.Add Array("Z", "EW", "", "UW")
        c.Add Array("N", "EW", "", "UW")
        c.Add Array("J", "EW", "", "UW")
        c.Add Array("TH", "EW", "", "UW")
        c.Add Array("CH", "EW", "", "UW")
        c.Add Array("SH", "EW", "", "UW")
        c.Add Array("", "EW", "", "yUW")
        c.Add Array("", "E", "O", "IY")
        c.Add Array("#:S", "ES", " ", "IHz")
        c.Add Array("#:C", "ES", " ", "IHz")
        c.Add Array("#:G", "ES", " ", "IHz")
        c.Add Array("#:Z", "ES", " ", "IHz")
        c.Add Array("#:X", "ES", " ", "IHz")
        c.Add Array("#:J", "ES", " ", "IHz")
        c.Add Array("#:CH", "ES", " ", "IHz")
        c.Add Array("#:SH", "ES", " ", "IHz")
        c.Add Array("#:", "E", "S ", "")
        c.Add Array("#:", "ELY", " ", "lIY")
        c.Add Array("#:", "EMENT", "", "mEHnt")
        c.Add Array("", "EFUL", "", "fUHl")
        c.Add Array("", "EE", "", "IY")
        c.Add Array("", "EARN", "", "ERn")
        c.Add Array(" ", "EAR", "^", "ER")
        c.Add Array("", "EAD", "", "EHd")
        c.Add Array("#:", "EA", " ", "IYAX")
        c.Add Array("", "EA", "SU", "EH")
        c.Add Array("", "EA", "", "IY")
        c.Add Array("", "EIGH", "", "EY")
        c.Add Array("", "EI", "", "IY")
        c.Add Array(" ", "EYE", "", "AY")
        c.Add Array("", "EY", "", "IY")
        c.Add Array("", "EU", "", "yUW")
        c.Add Array("", "E", "", "EH")
    rules.Add "E", c

    Set c = New Collection
        c.Add Array("", "FUL", "", "fUHl")
        c.Add Array("", "F", "", "f")
    rules.Add "F", c
    
    Set c = New Collection
        c.Add Array("", "GIV", "", "gIHv")
        c.Add Array(" ", "G", "I^", "g")
        c.Add Array("", "GE", "T", "gEH")
        c.Add Array("SU", "GGES", "", "gjEHs")
        c.Add Array("", "GG", "", "g")
        c.Add Array(" B#", "G", "", "g")
        c.Add Array("", "G", "+", "j")
        c.Add Array("", "GREAT", "", "grEYt")
        c.Add Array("#", "GH", "", "")
        c.Add Array("", "G", "", "g")
    rules.Add "G", c
    
    Set c = New Collection
        c.Add Array(" ", "HAV", "", "hAEv")
        c.Add Array(" ", "HERE", "", "hIYr")
        c.Add Array(" ", "HOUR", "", "AWER")
        c.Add Array("", "HOW", "", "hAW")
        c.Add Array("", "H", "#", "h")
        c.Add Array("", "H", "", "")
    rules.Add "H", c
    
    Set c = New Collection
        c.Add Array(" ", "IN", "", "IHn")
        c.Add Array(" ", "I", " ", "AY")
        c.Add Array("", "IN", "D", "AYn")
        c.Add Array("", "IER", "", "IYER")
        c.Add Array("#:R", "IED", "", "IYd")
        c.Add Array("", "IED", " ", "AYd")
        c.Add Array("", "IEN", "", "IYEHn")
        c.Add Array("", "IE", "T", "AYEH")
        c.Add Array(" :", "I", "%", "AY")
        c.Add Array("", "I", "%", "IY")
        c.Add Array("", "IE", "", "IY")
        c.Add Array("", "I", "^+:#", "IH")
        c.Add Array("", "IR", "#", "AYr")
        c.Add Array("", "IZ", "%", "AYz")
        c.Add Array("", "IS", "%", "AYz")
        c.Add Array("", "I", "D%", "AY")
        c.Add Array("+^", "I", "^+", "IH")
        c.Add Array("", "I", "T%", "AY")
        c.Add Array("#:^", "I", "^+", "IH")
        c.Add Array("", "I", "^+", "AY")
        c.Add Array("", "IR", "", "ER")
        c.Add Array("", "IGH", "", "AY")
        c.Add Array("", "ILD", "", "AYld")
        c.Add Array("", "IGN", " ", "AYn")
        c.Add Array("", "IGN", "^", "AYn")
        c.Add Array("", "IGN", "%", "AYn")
        c.Add Array("", "IQUE", "", "IYk")
        c.Add Array("", "I", "", "IH")
    rules.Add "I", c
    
    Set c = New Collection
        c.Add Array("", "J", "", "j")
    rules.Add "J", c
    
    Set c = New Collection
        c.Add Array(" ", "K", "N", "")
        c.Add Array("", "K", "", "k")
    rules.Add "K", c
    
    Set c = New Collection
        c.Add Array("", "LO", "C#", "lOW")
        c.Add Array("L", "L", "", "")
        c.Add Array("#:^", "L", "%", "AXl")
        c.Add Array("", "LEAD", "", "lIYd")
        c.Add Array("", "L", "", "l")
    rules.Add "L", c
    
    Set c = New Collection
        c.Add Array("", "MOV", "", "mUWv")
        c.Add Array("", "M", "", "m")
    rules.Add "M", c
    
    Set c = New Collection
        c.Add Array("E", "NG", "+", "nj")
        c.Add Array("", "NG", "R", "NGg")
        c.Add Array("", "NG", "#", "NGg")
        c.Add Array("", "NGL", "%", "NGgAXl")
        c.Add Array("", "NG", "", "NG")
        c.Add Array("", "NK", "", "NGk")
        c.Add Array(" ", "NOW", " ", "nAW")
        c.Add Array("", "N", "", "n")
    rules.Add "N", c
    
    Set c = New Collection
        c.Add Array("", "OF", " ", "AXv")
        c.Add Array("", "OROUGH", "", "EROW")
        c.Add Array("#:", "OR", " ", "ER")
        c.Add Array("#:", "ORS", " ", "ERz")
        c.Add Array("", "OR", "", "AOr")
        c.Add Array(" ", "ONE", "", "wAHn")
        c.Add Array("", "OW", "", "OW")
        c.Add Array(" ", "OVER", "", "OWvER")
        c.Add Array("", "OV", "", "AHv")
        c.Add Array("", "O", "^%", "OW")
        c.Add Array("", "O", "^EN", "OW")
        c.Add Array("", "O", "^I#", "OW")
        c.Add Array("", "OL", "D", "OWl")
        c.Add Array("", "OUGHT", "", "AOt")
        c.Add Array("", "OUGH", "", "AHf")
        c.Add Array(" ", "OU", "", "AW")
        c.Add Array("H", "OU", "S#", "AW")
        c.Add Array("", "OUS", "", "AXs")
        c.Add Array("", "OUR", "", "AOr")
        c.Add Array("", "OULD", "", "UHd")
        c.Add Array("^", "OU", "^L", "AH")
        c.Add Array("", "OUP", "", "UWp")
        c.Add Array("", "OU", "", "AW")
        c.Add Array("", "OY", "", "OY")
        c.Add Array("", "OING", "", "OWIHNG")
        c.Add Array("", "OI", "", "OY")
        c.Add Array("", "OOR", "", "AOr")
        c.Add Array("", "OOK", "", "UHk")
        c.Add Array("", "OOD", "", "UHd")
        c.Add Array("", "OO", "", "UW")
        c.Add Array("", "O", "E", "OW")
        c.Add Array("", "O", " ", "OW")
        c.Add Array("", "OA", "", "OW")
        c.Add Array(" ", "ONLY", "", "OWnlIY")
        c.Add Array(" ", "ONCE", "", "wAHns")
        c.Add Array("", "ON'T", "", "OWnt")
        c.Add Array("C", "O", "N", "AA")
        c.Add Array("", "O", "NG", "AO")
        c.Add Array(" :^", "O", "N", "AH")
        c.Add Array("I", "ON", "", "AXn")
        c.Add Array("#:", "ON", " ", "AXn")
        c.Add Array("#^", "ON", "", "AXn")
        c.Add Array("", "O", "ST ", "OW")
        c.Add Array("", "OF", "^", "AOf")
        c.Add Array("", "OTHER", "", "AHDHER")
        c.Add Array("", "OSS", " ", "AOs")
        c.Add Array("#:^", "OM", "", "AHm")
        c.Add Array("", "O", "", "AA")
    rules.Add "O", c
    
    Set c = New Collection
        c.Add Array("", "PH", "", "f")
        c.Add Array("", "PEOP", "", "pIYp")
        c.Add Array("", "POW", "", "pAW")
        c.Add Array("", "PUT", " ", "pUHt")
        c.Add Array("", "P", "", "p")
    rules.Add "P", c
    
    Set c = New Collection
        c.Add Array("", "QUAR", "", "kwAOr")
        c.Add Array("", "QU", "", "kw")
        c.Add Array("", "Q", "", "k")
    rules.Add "Q", c
    
    Set c = New Collection
        c.Add Array(" ", "RE", "^#", "rIY")
        c.Add Array("", "R", "", "r")
    rules.Add "R", c
    
    Set c = New Collection
        c.Add Array("", "SH", "", "SH")
        c.Add Array("#", "SION", "", "ZHAXn")
        c.Add Array("", "SOME", "", "sAHm")
        c.Add Array("#", "SUR", "#", "ZHER")
        c.Add Array("", "SUR", "#", "SHER")
        c.Add Array("#", "SU", "#", "ZHUW")
        c.Add Array("#", "SSU", "#", "SHUW")
        c.Add Array("#", "SED", " ", "zd")
        c.Add Array("#", "S", "#", "z")
        c.Add Array("", "SAID", "", "sEHd")
        c.Add Array("^", "SION", "", "SHAXn")
        c.Add Array("", "S", "S", "")
        c.Add Array(".", "S", " ", "z")
        c.Add Array("#:.E", "S", " ", "z")
        c.Add Array("#:^##", "S", " ", "z")
        c.Add Array("#:^#", "S", " ", "s")
        c.Add Array("U", "S", " ", "s")
        c.Add Array(" :#", "S", " ", "z")
        c.Add Array(" ", "SCH", "", "sk")
        c.Add Array("", "S", "C+", "")
        c.Add Array("#", "SM", "", "zm")
        c.Add Array("#", "SN", "'", "zAXn")
        c.Add Array("", "S", "", "s")
    rules.Add "S", c
    
    Set c = New Collection
        c.Add Array(" ", "THE", " ", "DHAX")
        c.Add Array("", "TO", " ", "tUW")
        c.Add Array("", "THAT", " ", "DHAEt")
        c.Add Array(" ", "THIS", " ", "DHIHs")
        c.Add Array(" ", "THEY", "", "DHEY")
        c.Add Array(" ", "THERE", "", "DHEHr")
        c.Add Array("", "THER", "", "DHER")
        c.Add Array("", "THEIR", "", "DHEHr")
        c.Add Array(" ", "THAN", " ", "DHAEn")
        c.Add Array(" ", "THEM", " ", "DHEHm")
        c.Add Array("", "THESE", " ", "DHIYz")
        c.Add Array(" ", "THEN", "", "DHEHn")
        c.Add Array("", "THROUGH", "", "THrUW")
        c.Add Array("", "THOSE", "", "DHOWz")
        c.Add Array("", "THOUGH", " ", "DHOW")
        c.Add Array(" ", "THUS", "", "DHAHs")
        c.Add Array("", "TH", "", "TH")
        c.Add Array("#:", "TED", " ", "tIHd")
        c.Add Array("S", "TI", "#N", "CH")
        c.Add Array("", "TI", "O", "SH")
        c.Add Array("", "TI", "A", "SH")
        c.Add Array("", "TIEN", "", "SHAXn")
        c.Add Array("", "TUR", "#", "CHER")
        c.Add Array("", "TU", "A", "CHUW")
        c.Add Array(" ", "TWO", "", "tUW")
        c.Add Array("", "T", "", "t")
    rules.Add "T", c
    
    Set c = New Collection
        c.Add Array(" ", "UN", "I", "yUWn")
        c.Add Array(" ", "UN", "", "AHn")
        c.Add Array(" ", "UPON", "", "AXpAOn")
        c.Add Array("T", "UR", "#", "UHr")
        c.Add Array("S", "UR", "#", "UHr")
        c.Add Array("R", "UR", "#", "UHr")
        c.Add Array("D", "UR", "#", "UHr")
        c.Add Array("L", "UR", "#", "UHr")
        c.Add Array("Z", "UR", "#", "UHr")
        c.Add Array("N", "UR", "#", "UHr")
        c.Add Array("J", "UR", "#", "UHr")
        c.Add Array("TH", "UR", "#", "UHr")
        c.Add Array("CH", "UR", "#", "UHr")
        c.Add Array("SH", "UR", "#", "UHr")
        c.Add Array("", "UR", "#", "yUHr")
        c.Add Array("", "UR", "", "ER")
        c.Add Array("", "U", "^ ", "AH")
        c.Add Array("", "U", "^^", "AH")
        c.Add Array("", "UY", "", "AY")
        c.Add Array(" G", "U", "#", "")
        c.Add Array("G", "U", "%", "")
        c.Add Array("G", "U", "#", "w")
        c.Add Array("#N", "U", "", "yUW")
        c.Add Array("T", "U", "", "UW")
        c.Add Array("S", "U", "", "UW")
        c.Add Array("R", "U", "", "UW")
        c.Add Array("D", "U", "", "UW")
        c.Add Array("L", "U", "", "UW")
        c.Add Array("Z", "U", "", "UW")
        c.Add Array("N", "U", "", "UW")
        c.Add Array("J", "U", "", "UW")
        c.Add Array("TH", "U", "", "UW")
        c.Add Array("CH", "U", "", "UW")
        c.Add Array("SH", "U", "", "UW")
        c.Add Array("", "U", "", "yUW")
    rules.Add "U", c
    
    Set c = New Collection
        c.Add Array("", "VIEW", "", "vyUW")
        c.Add Array("", "V", "", "v")
    rules.Add "V", c
    
    Set c = New Collection
        c.Add Array(" ", "WERE", "", "wER")
        c.Add Array("", "WA", "S", "wAA")
        c.Add Array("", "WA", "T", "wAA")
        c.Add Array("", "WHERE", "", "WHEHr")
        c.Add Array("", "WHAT", "", "WHAAt")
        c.Add Array("", "WHOL", "", "hOWl")
        c.Add Array("", "WHO", "", "hUW")
        c.Add Array("", "WH", "", "WH")
        c.Add Array("", "WAR", "", "wAOr")
        c.Add Array("", "WOR", "^", "wER")
        c.Add Array("", "WR", "", "r")
        c.Add Array("", "W", "", "w")
    rules.Add "W", c
    
    Set c = New Collection
        c.Add Array("", "X", "", "ks")
    rules.Add "X", c
    
    Set c = New Collection
        c.Add Array("", "YOUNG", "", "yAHNG")
        c.Add Array(" ", "YOU", "", "yUW")
        c.Add Array(" ", "YES", "", "yEHs")
        c.Add Array(" ", "Y", "", "y")
        c.Add Array("#:^", "Y", " ", "IY")
        c.Add Array("#:^", "Y", "I", "IY")
        c.Add Array(" :", "Y", " ", "AY")
        c.Add Array(" :", "Y", "#", "AY")
        c.Add Array(" :", "Y", "^+:#", "IH")
        c.Add Array(" :", "Y", "^#", "AY")
        c.Add Array("", "Y", "", "IH")
    rules.Add "Y", c
    
    Set c = New Collection
        c.Add Array("", "Z", "", "z")
    rules.Add "Z", c
    
End Sub
Function NRL(StrInput As String) As String
    If rules.Count = 0 Then Call buildRules
    
    StrInput = UCase$(StrInput)
    
    Dim intPos As Long
    Dim strPronounciation As String
    Dim strLeftOrig As String, strRightOrig As String
    Dim strFirst As String
    Dim rule As Variant
    Dim strLeft As String, strMatch As String, strRight As String
    
    intPos = 1
    While intPos <= Len(StrInput)
        strLeftOrig = left$(StrInput, intPos - 1)
        strRightOrig = Mid$(StrInput, intPos)
    
        strFirst = Mid$(StrInput, intPos, 1)
        If rules.Exists(strFirst) Then
            For Each rule In rules(strFirst)
                strLeft = rule(0): strMatch = rule(1): strRight = rule(2)
                re_left.Pattern = vbNullString: re_right.Pattern = vbNullString

                If left(strRightOrig, Len(strMatch)) = strMatch Then
                    If Len(strLeft) > 0 Then re_left.Pattern = toRegex(strLeft, True)
                    If Len(strRight) > 0 Then re_right.Pattern = toRegex(strRight, False)
                    
                    If (Len(strLeft) = 0 Or re_left.test(strLeftOrig)) And (Len(strRight) = 0 Or re_right.test(Mid$(strRightOrig, Len(strMatch) + 1))) Then
                        strPronounciation = strPronounciation & rule(3)
                        intPos = intPos + Len(strMatch)
                        Exit For
                    End If
                End If
            Next rule
        Else
            intPos = intPos + 1
        End If
    Wend
    NRL = strPronounciation
End Function
Function toRegex(StrInput As String, boolLeft As Boolean) As String
    Dim strNewPattern As String
    Dim i As Integer
    For i = 1 To Len(StrInput)
        Select Case Mid(StrInput, i, 1)
            Case "#": strNewPattern = strNewPattern & "[AEIOU]+"
            Case ":": strNewPattern = strNewPattern & "[BCDFGHJKLMNPQRSTVWXYZ]*"
            Case "^": strNewPattern = strNewPattern & "[BCDFGHJKLMNPQRSTVWXYZ]"
            Case ".": strNewPattern = strNewPattern & "[BDVGJLMNTWZ]"
            Case "%": strNewPattern = strNewPattern & "(ER|E|ES|ED|ING|ELY)"
            Case "+": strNewPattern = strNewPattern & "[EIY]"
            Case " ": strNewPattern = strNewPattern & "^"
            Case Else: strNewPattern = strNewPattern & Mid(StrInput, i, 1)
        End Select
    Next
    
    If boolLeft Then
        strNewPattern = strNewPattern & "$"
        If InStr(StrInput, "^") = 0 Then strNewPattern = "^.*" & strNewPattern
        
        'Differing from Python re_match, have to add ^ to front if not there
        If left$(strNewPattern, 1) <> "^" Then strNewPattern = "^" & strNewPattern
    Else
        strNewPattern = "^" & Replace$(strNewPattern, "^", "$")
        If InStr(strNewPattern, "$") = 0 Then strNewPattern = strNewPattern & ".*$"
    End If
    
    toRegex = strNewPattern
End Function
