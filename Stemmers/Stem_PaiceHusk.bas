Attribute VB_Name = "Stem_PaiceHusk"
Option Explicit
Dim ruleTable As New Dictionary
Sub SetUpRules()
    ruleTable.RemoveAll
    Dim c As Collection
    Set c = New Collection
        c.Add Array("ifiabl", Array(False, 6, "", True))
        c.Add Array("plicat", Array(False, 4, "y", True))
    ruleTable.Add 6, c
    Set c = New Collection
        c.Add Array("guish", Array(False, 5, "ct", True))
        c.Add Array("sumpt", Array(False, 2, "", True))
        c.Add Array("istry", Array(False, 5, "", True))
    ruleTable.Add 5, c
    Set c = New Collection
        c.Add Array("ytic", Array(False, 3, "s", True))
        c.Add Array("ceed", Array(False, 2, "ss", True))
        c.Add Array("hood", Array(False, 4, "", False))
        c.Add Array("lief", Array(False, 1, "v", True))
        c.Add Array("verj", Array(False, 1, "t", True))
        c.Add Array("misj", Array(False, 2, "t", True))
        c.Add Array("iabl", Array(False, 4, "y", True))
        c.Add Array("iful", Array(False, 4, "y", True))
        c.Add Array("sion", Array(False, 4, "j", False))
        c.Add Array("xion", Array(False, 4, "ct", True))
        c.Add Array("ship", Array(False, 4, "", False))
        c.Add Array("ness", Array(False, 4, "", False))
        c.Add Array("ment", Array(False, 4, "", False))
        c.Add Array("ript", Array(False, 2, "b", True))
        c.Add Array("orpt", Array(False, 2, "b", True))
        c.Add Array("duct", Array(False, 1, "", True))
        c.Add Array("cept", Array(False, 2, "iv", True))
        c.Add Array("olut", Array(False, 2, "v", True))
        c.Add Array("sist", Array(False, 0, "", True))
    ruleTable.Add 4, c
    Set c = New Collection
        c.Add Array("ied", Array(False, 3, "y", False))
        c.Add Array("eed", Array(False, 1, "", True))
        c.Add Array("ing", Array(False, 3, "", False))
        c.Add Array("iag", Array(False, 3, "y", True))
        c.Add Array("ish", Array(False, 3, "", False))
        c.Add Array("fuj", Array(False, 1, "s", True))
        c.Add Array("hej", Array(False, 1, "r", True))
        c.Add Array("abl", Array(False, 3, "", False))
        c.Add Array("ibl", Array(False, 3, "", True))
        c.Add Array("bil", Array(False, 2, "l", False))
        c.Add Array("ful", Array(False, 3, "", False))
        c.Add Array("ial", Array(False, 3, "", False))
        c.Add Array("ual", Array(False, 3, "", False))
        c.Add Array("ium", Array(False, 3, "", True))
        c.Add Array("ism", Array(False, 3, "", False))
        c.Add Array("ion", Array(False, 3, "", False))
        c.Add Array("ian", Array(False, 3, "", False))
        c.Add Array("een", Array(False, 0, "", True))
        c.Add Array("ear", Array(False, 0, "", True))
        c.Add Array("ier", Array(False, 3, "y", False))
        c.Add Array("ies", Array(False, 3, "y", False))
        c.Add Array("sis", Array(False, 2, "", True))
        c.Add Array("ous", Array(False, 3, "", False))
        c.Add Array("ent", Array(False, 3, "", False))
        c.Add Array("ant", Array(False, 3, "", False))
        c.Add Array("ist", Array(False, 3, "", False))
        c.Add Array("iqu", Array(False, 3, "", True))
        c.Add Array("ogu", Array(False, 1, "", True))
        c.Add Array("siv", Array(False, 3, "j", False))
        c.Add Array("eiv", Array(False, 0, "", True))
        c.Add Array("bly", Array(False, 1, "", False))
        c.Add Array("ily", Array(False, 3, "y", False))
        c.Add Array("ply", Array(False, 0, "", True))
        c.Add Array("ogy", Array(False, 1, "", True))
        c.Add Array("phy", Array(False, 1, "", True))
        c.Add Array("omy", Array(False, 1, "", True))
        c.Add Array("opy", Array(False, 1, "", True))
        c.Add Array("ity", Array(False, 3, "", False))
        c.Add Array("ety", Array(False, 3, "", False))
        c.Add Array("lty", Array(False, 2, "", True))
        c.Add Array("ary", Array(False, 3, "", False))
        c.Add Array("ory", Array(False, 3, "", False))
        c.Add Array("ify", Array(False, 3, "", True))
        c.Add Array("ncy", Array(False, 2, "t", False))
        c.Add Array("acy", Array(False, 3, "", False))
    ruleTable.Add 3, c
    Set c = New Collection
        c.Add Array("ia", Array(True, 2, "", True))
        c.Add Array("bb", Array(False, 1, "", True))
        c.Add Array("ic", Array(False, 2, "", False))
        c.Add Array("nc", Array(False, 1, "t", False))
        c.Add Array("dd", Array(False, 1, "", True))
        c.Add Array("ed", Array(False, 2, "", False))
        c.Add Array("if", Array(False, 2, "", False))
        c.Add Array("ag", Array(False, 2, "", False))
        c.Add Array("gg", Array(False, 1, "", True))
        c.Add Array("th", Array(True, 2, "", True))
        c.Add Array("ij", Array(False, 1, "d", True))
        c.Add Array("uj", Array(False, 1, "d", True))
        c.Add Array("oj", Array(False, 1, "d", True))
        c.Add Array("nj", Array(False, 1, "d", True))
        c.Add Array("cl", Array(False, 1, "", True))
        c.Add Array("ul", Array(False, 2, "", True))
        c.Add Array("al", Array(False, 2, "", False))
        c.Add Array("ll", Array(False, 1, "", True))
        c.Add Array("um", Array(True, 2, "", True))
        c.Add Array("mm", Array(False, 1, "", True))
        c.Add Array("an", Array(False, 2, "", False))
        c.Add Array("en", Array(False, 2, "", False))
        c.Add Array("nn", Array(False, 1, "", True))
        c.Add Array("pp", Array(False, 1, "", True))
        c.Add Array("er", Array(False, 2, "", False))
        c.Add Array("ar", Array(False, 2, "", True))
        c.Add Array("or", Array(False, 2, "", False))
        c.Add Array("ur", Array(False, 2, "", False))
        c.Add Array("rr", Array(False, 1, "", True))
        c.Add Array("tr", Array(False, 1, "", False))
        c.Add Array("is", Array(False, 2, "", False))
        c.Add Array("ss", Array(False, 0, "", True))
        c.Add Array("us", Array(True, 2, "", True))
        c.Add Array("at", Array(False, 2, "", False))
        c.Add Array("tt", Array(False, 1, "", True))
        c.Add Array("iv", Array(False, 2, "", False))
        c.Add Array("ly", Array(False, 2, "", False))
        c.Add Array("iz", Array(False, 2, "", False))
        c.Add Array("yz", Array(False, 1, "s", True))
    ruleTable.Add 2, c
    Set c = New Collection
        c.Add Array("a", Array(True, 1, "", True))
        c.Add Array("e", Array(False, 1, "", False))
        'c.Add Array("i", Array(True, 1, "", True))
        c.Add Array("j", Array(False, 1, "s", True))
        'c.Add Array("s", Array(True, 1, "", False))
    ruleTable.Add 1, c
End Sub
Function PaiceHusk(strWord As String) As String
    'Initialize Rule Set
    If ruleTable.Count = 0 Then Call SetUpRules
    Dim boolTerminate As Boolean
    Dim boolIntact As Boolean: boolIntact = True
    Dim boolAccept As Boolean
    Dim i As Integer
    Dim item As Variant
    Dim rule As Variant
    Dim result As Variant
    Dim j As Integer
    Do While Not boolTerminate And j < 10
        For i = 6 To 0 Step -1
            If i = 1 And right(strWord, 1) Like "[is]" Then
                boolAccept = False
                If right(strWord, i) = "i" Then
                    result = applyRule(strWord, Array(True, 1, "", True), boolIntact, boolTerminate)
                    strWord = result(0): boolAccept = result(1): boolIntact = result(2): boolTerminate = result(3)
                    If boolAccept Then Exit For
                    result = applyRule(strWord, Array(False, 1, "y", False), boolIntact, boolTerminate)
                    strWord = result(0): boolAccept = result(1): boolIntact = result(2): boolTerminate = result(3)
                    If boolAccept Then Exit For
                Else
                    result = applyRule(strWord, Array(True, 1, "", False), boolIntact, boolTerminate)
                    strWord = result(0): boolAccept = result(1): boolIntact = result(2): boolTerminate = result(3)
                    If boolAccept Then Exit For
                    result = applyRule(strWord, Array(False, 0, "", True), boolIntact, boolTerminate)
                    strWord = result(0): boolAccept = result(1): boolIntact = result(2): boolTerminate = result(3)
                    If boolAccept Then Exit For
                End If
            Else
                If ruleTable.Exists(i) Then
                    For Each item In ruleTable(i)
                        boolAccept = False
                        If right(strWord, i) = item(0) Then
                            result = applyRule(strWord, item(1), boolIntact, boolTerminate)
                            strWord = result(0): boolAccept = result(1): boolIntact = result(2): boolTerminate = result(3)
                            If boolAccept Then Exit For
                        End If
                    Next
                    If boolAccept Then Exit For
                End If
            End If
        Next
        j = j + 1
    Loop
    PaiceHusk = strWord
End Function
Function acceptable(strWord As String) As Boolean
    If Len(strWord) > 0 And left(strWord, 1) Like "[aeiou]" Then
        acceptable = Len(strWord) > 1
        Exit Function
    End If
    acceptable = Len(strWord) > 2 And Mid(strWord, 2) Like "*[aeiouy]*"
End Function
Function applyRule(strWord As String, rule As Variant, intact As Boolean, terminate As Boolean) As Variant
    Dim oldWord As String: oldWord = strWord
    
    Dim onlyIntact As Boolean: onlyIntact = rule(0)
    Dim del_len As Integer: del_len = rule(1)
    Dim add_Str As String: add_Str = rule(2)
    Dim set_terminate As Boolean: set_terminate = rule(3)
    
    Dim returnResult As Variant
    ReDim returnResult(0 To 3)
    
    If Not onlyIntact Or (intact And onlyIntact) Then
        If del_len > 0 Then
            strWord = left(strWord, Len(strWord) - del_len)
        End If
        If add_Str <> "" Then
            strWord = strWord & add_Str
        End If
    Else
        returnResult(0) = strWord: returnResult(1) = False: returnResult(2) = intact: returnResult(3) = terminate
        applyRule = returnResult
        Exit Function
    End If
    
    If acceptable(strWord) Then
        returnResult(0) = strWord: returnResult(1) = True: returnResult(2) = False: returnResult(3) = set_terminate
    Else
        returnResult(0) = oldWord: returnResult(1) = False: returnResult(2) = intact: returnResult(3) = terminate
    End If
    applyRule = returnResult
End Function
