Attribute VB_Name = "Encode_Ainsworth"
Option Explicit
Dim regex As New VBScript_RegExp_55.RegExp
Dim boolSetUpComplete As Boolean
Dim lookBehindRegex As New VBScript_RegExp_55.RegExp
Dim regexList As New Collection
Const strSuffixes = "(able|ance|ence|less|ment|ness|ship|sion|tion|age|ant|ate|ent|ery|ful|ify|ise|ism|ity|ive|ize|ous|al|cy|en|er|es|fy|ry|s|y)"
Sub SetUpRegExList()
    Set regexList = Nothing
    regexList.Add Array("^a$", ChrW(601), 1)
    regexList.Add Array("^are", ChrW(593), 3)
    regexList.Add Array("a(?=[ei])", ChrW(400) & "i", 1)
    regexList.Add Array("ar", ChrW(593), 2)
    regexList.Add Array("a(?=sk)", ChrW(593), 1)
    regexList.Add Array("a(?=st)", ChrW(593), 1)
    regexList.Add Array("a(?=th)", ChrW(593), 1)
    regexList.Add Array("a(?=ft)", ChrW(593), 1)
    regexList.Add Array("ai", ChrW(400) & "i", 2)
    regexList.Add Array("ay", ChrW(400) & "i", 2)
    regexList.Add Array("aw", ChrW(596), 2)
    regexList.Add Array("au", ChrW(596), 2)
    regexList.Add Array("al(?=l)", ChrW(596), 2)
    regexList.Add Array("a(?=ble)", ChrW(400) & "i", 1)
    regexList.Add Array("a(?=ng" & strSuffixes & ")", ChrW(400) & "i", 1)
    regexList.Add Array("a", ChrW(230), 1)

    regexList.Add Array("b", "b", 1)
    
    regexList.Add Array("ch", "t" & ChrW(643), 2)
    regexList.Add Array("ck", "k", 2)
    regexList.Add Array("c(?=y)", "s", 1)
    regexList.Add Array("c(?=e)", "s", 1)
    regexList.Add Array("c(?=i)", "s", 1)
    regexList.Add Array("c", "k", 1)
    regexList.Add Array("d", "d", 1)
    
    regexList.Add Array("[aeiou][bcdfghjklmnpqrstvwxyz]e$", "", 1, "[aeiou][bcdfghjklmnpqrstvwxyz]", 2)
    regexList.Add Array("the$", ChrW(601), 1, "th", 2)
    regexList.Add Array("^[bcdfghjklmnpqrstvwxyz]e$", "i", 1, "[bcdfghjklmnpqrstvwxyz]", 1)
    regexList.Add Array("^[bcdfghjklmnpqrstvwxyz]e(?=d)$", ChrW(603), 1, "[bcdfghjklmnpqrstvwxyz]", 1)
    
    
    regexList.Add Array("o(?=ld)", ChrW(601) & ChrW(650), 1)
    regexList.Add Array("oy", ChrW(596) & "i", 2)
    regexList.Add Array("o(?=ing)", ChrW(601) & ChrW(650), 1)
    regexList.Add Array("oi", ChrW(596) & "i", 2)
    regexList.Add Array("you", "u", 2, "y", 1)
    
    regexList.Add Array("ou(?=s)", ChrW(652), 2)
    regexList.Add Array("ough(?=t)", ChrW(596), 4)
    
    regexList.Add Array("bough", ChrW(593) & ChrW(650), 4, "b", 1)
    regexList.Add Array("tough", ChrW(652) & "f", 4, "t", 1)
    regexList.Add Array("cough", "of", 4, "c", 1)
    regexList.Add Array("^rough", ChrW(593) & "f", 4, "r", 1)
    regexList.Add Array("rough", ChrW(650), 4, "r", 1)
    
    regexList.Add Array("ough", ChrW(601) & ChrW(650), 4)
    regexList.Add Array("oul(?=d)", ChrW(650), 3)
    regexList.Add Array("ou", "a" & ChrW(650), 2)
    regexList.Add Array("oor", ChrW(596), 3)
    regexList.Add Array("oo(?=k)", ChrW(650), 2)
    regexList.Add Array("foo(?=d)", "u", 2, "f", 1)
    regexList.Add Array("oo(?=d)", ChrW(650), 2)
    regexList.Add Array("foo(?=t)", ChrW(650), 2, "f", 1)
    regexList.Add Array("soo(?=t)", ChrW(650), 2, "s", 1)
    
    regexList.Add Array("woo", ChrW(650), 2, "w", 1)
    regexList.Add Array("oo", "u", 2)
    regexList.Add Array("shoe", "u", 2, "sh", 2)
    regexList.Add Array("oe", ChrW(601) & ChrW(650), 2)
    regexList.Add Array("[aeiou][bcdfghjklmnpqrstvwxyz]de(?=d)$", ChrW(601), 1, "[aeiou][bcdfghjklmnpqrstvwxyz]d", 3)
    regexList.Add Array("[aeiou][bcdfghjklmnpqrstvwxyz]te(?=d)$", ChrW(601), 1, "[aeiou][bcdfghjklmnpqrstvwxyz]t", 3)
    regexList.Add Array("[aeiou][bcdfghjklmnpqrstvwxyz]e(?=d)$", "", 1, "[aeiou][bcdfghjklmnpqrstvwxyz]", 2)
    regexList.Add Array("e(?=r)$", ChrW(601), 1)
    regexList.Add Array("where", ChrW(603) & ChrW(601), 3, "wh", 2)
    regexList.Add Array("here", "i" & ChrW(601), 3, "h", 1)
    regexList.Add Array("were", ChrW(604), 3, "w", 1)
    regexList.Add Array("ere", "ir", 3)
    regexList.Add Array("ee", "i", 2)
    regexList.Add Array("ear", "ir", 3)
    regexList.Add Array("ea", "i", 2)
    regexList.Add Array("e(?=ver)", ChrW(603), 1)
    regexList.Add Array("eye", ChrW(593) & "i", 3)
    regexList.Add Array("e(?=[ei])", "i", 1)
    regexList.Add Array("cei", "i", 2, "c", 1)
    regexList.Add Array("ei", ChrW(593) & "i", 2)
    regexList.Add Array("e(?=r)", ChrW(604), 1)
    regexList.Add Array("eo", "i", 2)
    regexList.Add Array("ew", "ju", 2)
    regexList.Add Array("e(?=u)", "", 1)
    regexList.Add Array("e", ChrW(603), 1)
    
    regexList.Add Array("f$", "v", 1)
    regexList.Add Array("f", "f", 1)
    regexList.Add Array("g(?=e)$", "d" & ChrW(439), 1)
    regexList.Add Array("g(?=es)$", "d" & ChrW(439), 1)
    regexList.Add Array("g(?=" & strSuffixes & ")", "g", 1)
    regexList.Add Array("g(?=i)", "d" & ChrW(439), 1)
    regexList.Add Array("g(?=et)", "g", 1)
    
    regexList.Add Array("cow", ChrW(593) & ChrW(650), 2, "c", 1)
    regexList.Add Array("how", ChrW(593) & ChrW(650), 2, "h", 1)
    regexList.Add Array("now", ChrW(593) & ChrW(650), 2, "n", 1)
    regexList.Add Array("vow", ChrW(593) & ChrW(650), 2, "v", 1)
    regexList.Add Array("row", ChrW(593) & ChrW(650), 2, "r", 1)
    regexList.Add Array("ow", ChrW(601) & ChrW(650), 2)
    regexList.Add Array("go$", ChrW(601) & ChrW(650), 2, "g", 1)
    regexList.Add Array("no$", ChrW(601) & ChrW(650), 2, "n", 1)
    regexList.Add Array("so$", ChrW(601) & ChrW(650), 2, "s", 1)
    regexList.Add Array("o$", "u", 1)
    regexList.Add Array("o", "o", 1)
    regexList.Add Array("ph", "f", 2)
    regexList.Add Array("psy", "s" & ChrW(593) & "i", 3)
    regexList.Add Array("p", "p", 1)
    regexList.Add Array("q", "kw", 1)
    regexList.Add Array("r$", "", 1)
    regexList.Add Array("rho", "r" & ChrW(399) & ChrW(433), 3)
    regexList.Add Array("r", "r", 1)
    regexList.Add Array("sh", ChrW(643), 2)
    regexList.Add Array("ss", "s", 2)
    regexList.Add Array("sch", "sk", 3)
    'regexList.Add Array("Xvs","z",1,"Xvs")
    
    regexList.Add Array("[aeiou]s$", "z", 1, "[aeiou]", 1)
    
    regexList.Add Array("s", "s", 1)
    regexList.Add Array("there", ChrW(240) & ChrW(603) & ChrW(601), 5)
    regexList.Add Array("g(?=e)", "d" & ChrW(658), 1)
    regexList.Add Array("gh", "g", 2)
    regexList.Add Array("g", "g", 1)
    regexList.Add Array("wh", "", 1, "w", 1)
    
    regexList.Add Array("ha(?=v)", "h" & ChrW(230), 2)
    regexList.Add Array("h", "h", 1)
    regexList.Add Array("^i$", ChrW(593) & "i", 1)
    regexList.Add Array("i(?=ty)", ChrW(618), 1)
    regexList.Add Array("i(?=[ei])", ChrW(593) & "i", 1)
    regexList.Add Array("ir", ChrW(604), 2)
    regexList.Add Array("igh", ChrW(593) & "i", 3)
    
    regexList.Add Array("tio(?=n)", ChrW(652), 2, "t", 1)
    
    regexList.Add Array("i(?=nd)", ChrW(593) & "i", 1)
    regexList.Add Array("i(?=ld)", ChrW(593) & "i", 1)
    regexList.Add Array("^[bcdfghjklmnpqrstvwxyz]ie", "ai", 2, "[bcdfghjklmnpqrstvwxyz]", 1)
    regexList.Add Array("[aeiou][bcdfghjklmnpqrstvwxyz]ie", "i", 2, "[aeiou][bcdfghjklmnpqrstvwxyz]", 2)
    
    regexList.Add Array("i", ChrW(618), 1)
    regexList.Add Array("j", "d" & ChrW(658), 1) '439=658??
    
    regexList.Add Array("^k(?=n)", "", 1)
    regexList.Add Array("k", "k", 1)
    
    regexList.Add Array("le$", ChrW(399) & "l", 2)
    regexList.Add Array("l", "l", 1)
    
    regexList.Add Array("m", "m", 1)
    
    regexList.Add Array("n(?=g)", ChrW(330), 1)
    regexList.Add Array("n", "n", 1)
    
    regexList.Add Array("or", ChrW(390), 2)
    regexList.Add Array("o(?=[ei])", ChrW(399) & ChrW(433), 1)
    regexList.Add Array("oa", ChrW(399) & ChrW(433), 2)
    
    regexList.Add Array("their", ChrW(208) & ChrW(400) & ChrW(399), 5)
    regexList.Add Array("th(?=r)", ChrW(920), 2)
    regexList.Add Array("th", ChrW(208), 2)
    regexList.Add Array("t(?=ion)", ChrW(643), 1)
    regexList.Add Array("t", "t", 1)
    
    regexList.Add Array("u(?=pon)", ChrW(652), 1)
    regexList.Add Array("u(?=[aeiou])", "u", 1)
    regexList.Add Array("u(?=[bcdfghjklmnpqrstvwxyz])$", ChrW(652), 1)
    
    regexList.Add Array("ru", "u", 1, "r", 1)
    regexList.Add Array("lu", "u", 1, "l", 1)
    regexList.Add Array("u", "ju", 1)
    regexList.Add Array("v", "v", 1)
    
    regexList.Add Array("w(?=r)", "", 1)
    regexList.Add Array("wh(?=o)", "h", 2)
    regexList.Add Array("wha(?=t)", "wo", 3)
    regexList.Add Array("wa", "wo", 2)
    regexList.Add Array("wo(?=r)", "w" & ChrW(604), 2)
    regexList.Add Array("w", "w", 1)
    
    regexList.Add Array("x", "ks", 1)
    regexList.Add Array("^y", "j", 1)
    regexList.Add Array("[aeiou][bcdfghjklmnpqrstvwxyz]y", ChrW(618), 1, "[aeiou][bcdfghjklmnpqrstvwxyz]", 2)
    regexList.Add Array("^[bcdfghjklmnpqrstvwxyz]y", "ai", 1, "[bcdfghjklmnpqrstvwxyz]", 1)
    regexList.Add Array("y(?=[ei])", ChrW(593) & "i", 1)
    regexList.Add Array("y", ChrW(618), 1)
    
    regexList.Add Array("z", "z", 1)
    boolSetUpComplete = True
End Sub
Function Ainsworth(strWord As String) As String
    If Not boolSetUpComplete Then Call SetUpRegExList
    
    strWord = LCase(strWord)
    Dim strPron As String
    Dim strFragment As String
    
    regex.IgnoreCase = True
    regex.Global = True
    
    lookBehindRegex.IgnoreCase = True
    
    Dim rule As Variant
    Dim matched As Boolean
    Dim intPos As Integer
    intPos = 1
    Do While intPos <= Len(strWord)
        matched = False
        strFragment = Mid(strWord, intPos)
        
        For Each rule In regexList
            If UBound(rule) >= 3 Then 'if it has a lookbehind
                If Left(rule(0), 1) = "^" And intPos > 1 Then
                
                Else
                    regex.pattern = "^" & rule(0)
                    lookBehindRegex.pattern = "^" & rule(3)
                    If (intPos - rule(4)) > 0 Then
                         If regex.test(Mid(strWord, intPos - rule(4))) Then
        
                            If lookBehindRegex.test(Mid(strWord, intPos - rule(4))) Then
                                strPron = strPron & rule(1)
                                intPos = intPos + rule(2)
                                matched = True
                                Exit For
                            End If
                        End If
                    End If
                End If
            Else
                If Left(rule(0), 1) = "^" And intPos > 1 Then
                
                Else
                    regex.pattern = "^" & rule(0)
                    If regex.test(strFragment) Then
                        strPron = strPron & rule(1)
                        intPos = intPos + rule(2)
                        matched = True
                        Exit For
                    End If
                End If
            End If
        Next
        
        If Not matched Then intPos = intPos + 1
    Loop
    DoEvents
    
    Ainsworth = strPron
End Function
