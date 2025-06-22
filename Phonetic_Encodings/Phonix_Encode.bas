Attribute VB_Name = "Phonix_Encode"
Option Explicit
Const strTranscodeIn = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const strTranscodeOut = "01230720022455012683070808"
Dim cRegexSubstitutions As New Collection
Dim regexPhonixReplace As New VBScript_RegExp_55.RegExp
Function Phonix(strWord As String, Optional intMaxLength As Integer = 4, Optional boolZeroPad As Boolean = True) As String
    If intMaxLength <> -1 Then
        If intMaxLength < 4 Then intMaxLength = 4
        If intMaxLength > 64 Then intMaxLength = 64
    Else
        intMaxLength = 64
    End If
    
    If cRegexSubstitutions Is Nothing Or cRegexSubstitutions.Count = 0 Then
        Call PhonixSetup
    End If
    
    strWord = UnicodeFunctions.UnicodeStrip(strWord)
    strWord = PhoneticFunctions.GetAlphaOnly(UCase$(strWord))
    
    Dim strSoundex As String
    Dim vTrans As Variant
    If strWord <> "" Then
        For Each vTrans In cRegexSubstitutions
            regexPhonixReplace.pattern = vTrans(0)
            If regexPhonixReplace.test(strWord) Then
                strWord = regexPhonixReplace.Replace(strWord, vTrans(1))
            End If
        Next
        
        If Left$(strWord, 1) Like "[AEIOUY]" Then
            strSoundex = "v"
        Else
            strSoundex = Left$(strWord, 1)
        End If
    End If
    
    Dim i As Long
    For i = 2 To Len(strWord)
        strSoundex = strSoundex & Mid$(strTranscodeOut, InStr(strTranscodeIn, Mid$(strWord, i, 1)), 1)
    Next
    
    strSoundex = PhoneticFunctions.DeleteConsecutiveRepeats(strSoundex)
    strSoundex = Replace$(strSoundex, "0", "")
    
    If boolZeroPad Then
        strSoundex = strSoundex & String(intMaxLength, "0")
    End If
    
    If strSoundex = "" Then strSoundex = "0"
    
    Phonix = Left$(strSoundex, intMaxLength)
End Function
Sub PhonixSetup()
    regexPhonixReplace.Global = True
    regexPhonixReplace.IgnoreCase = True
    
    cRegexSubstitutions.Add Array("DG", "G")
    cRegexSubstitutions.Add Array("CO", "KO")
    cRegexSubstitutions.Add Array("CA", "KA")
    cRegexSubstitutions.Add Array("CU", "KU")
    cRegexSubstitutions.Add Array("C[IY]", "SI") 'CI/CY
    cRegexSubstitutions.Add Array("CE", "SE")
    cRegexSubstitutions.Add Array("^CL(?=[AEIOU])", "KL")
    cRegexSubstitutions.Add Array("(CK)|([GJ]C$)", "K")
    cRegexSubstitutions.Add Array("^CH?R(?=[AEIOU])", "KR") 'CHR/CR - 0
    cRegexSubstitutions.Add Array("^WR", "R")
    cRegexSubstitutions.Add Array("NC", "NK")
    cRegexSubstitutions.Add Array("CT", "KT")
    cRegexSubstitutions.Add Array("PH", "F")
    cRegexSubstitutions.Add Array("AA", "AR")
    cRegexSubstitutions.Add Array("SCH", "SH")
    cRegexSubstitutions.Add Array("BTL", "TL")
    cRegexSubstitutions.Add Array("GHT", "T")
    cRegexSubstitutions.Add Array("AUGH", "ARF")
    cRegexSubstitutions.Add Array("([AEIOU])LJ(?=[AEIOU])", "$1LD")
    cRegexSubstitutions.Add Array("LOUGH", "LOW")
    cRegexSubstitutions.Add Array("^Q", "KW")
    cRegexSubstitutions.Add Array("(^KN)|(GN$)|GHN|(GNE$)", "N")
    cRegexSubstitutions.Add Array("GHNE", "NE")
    cRegexSubstitutions.Add Array("GNES$", "NS")
    cRegexSubstitutions.Add Array("(^GN)|(GN(?=[^AEIOU])|(GN$))", "N")
    cRegexSubstitutions.Add Array("^P([ST])", "$1")
    cRegexSubstitutions.Add Array("^CZ", "C")
    cRegexSubstitutions.Add Array("([AEIOU])WZ(?=.)", "$1Z")
    cRegexSubstitutions.Add Array("(.)CZ(?=.)", "$1CH")
    cRegexSubstitutions.Add Array("([LR])Z", "$1SH") 'LZ/RZ -> LSH/RSH
    cRegexSubstitutions.Add Array("(.)Z(?=[AEIOU])", "$1S")
    cRegexSubstitutions.Add Array("ZZ", "TS")
    cRegexSubstitutions.Add Array("([^AEIOU])Z(?=.)", "$1TS")
    cRegexSubstitutions.Add Array("([AEIOU])Q(?=[AEIOU])", "$1KW")
    cRegexSubstitutions.Add Array("([AEIOU])J(?=[AEIOU])", "$1Y")
    cRegexSubstitutions.Add Array("HROUG", "REW")
    cRegexSubstitutions.Add Array("OUGH", "OF")
    cRegexSubstitutions.Add Array("^YJ(?=[AEIOU])", "Y")
    cRegexSubstitutions.Add Array("^GH", "G")
    cRegexSubstitutions.Add Array("([AEIOU])GH$", "$1E")
    cRegexSubstitutions.Add Array("^CY", "S")
    cRegexSubstitutions.Add Array("NX", "NKS")
    cRegexSubstitutions.Add Array("^PF", "F")
    cRegexSubstitutions.Add Array("DT$", "T")
    cRegexSubstitutions.Add Array("TL$", "TIL")
    cRegexSubstitutions.Add Array("DL$", "DIL")
    cRegexSubstitutions.Add Array("YTH", "ITH")
    cRegexSubstitutions.Add Array("^TJ(?=[AEIOU])", "CH")
    cRegexSubstitutions.Add Array("^TSJ(?=[AEIOU])", "CH")
    cRegexSubstitutions.Add Array("^TS(?=[AEIOU])", "T")
    cRegexSubstitutions.Add Array("TCH", "CH")
    cRegexSubstitutions.Add Array("([AEIOU])WSK(?=.)", "$1VSKIE")
    cRegexSubstitutions.Add Array("([AEIOU])WSK$", "$1VSKIE")
    cRegexSubstitutions.Add Array("^MN(?=[AEIOU])", "N")
    cRegexSubstitutions.Add Array("^PN(?=[AEIOU])", "N")
    cRegexSubstitutions.Add Array("([AEIOU])STL(?=.)", "$1SL")
    cRegexSubstitutions.Add Array("([AEIOU])STL$", "$1SL")
    cRegexSubstitutions.Add Array("TNT$", "ENT")
    cRegexSubstitutions.Add Array("EAUX$", "OH")
    cRegexSubstitutions.Add Array("EXCI|[X]", "ECS")
    cRegexSubstitutions.Add Array("NED$", "ND")
    cRegexSubstitutions.Add Array("JR", "DR")
    cRegexSubstitutions.Add Array("EE$", "EA")
    cRegexSubstitutions.Add Array("ZS", "S")
    cRegexSubstitutions.Add Array("([AEIOU])R(?=[^AEIOU])", "$1AH")
    cRegexSubstitutions.Add Array("([AEIOU])R$", "$1AH")
    cRegexSubstitutions.Add Array("([AEIOU])HR(?=[^AEIOU])", "$1AH")
    cRegexSubstitutions.Add Array("([AEIOU])HR$", "$1AH")
    cRegexSubstitutions.Add Array("RE$", "AR")
    cRegexSubstitutions.Add Array("([AEIOU])R$", "$1AH")
    cRegexSubstitutions.Add Array("LLE", "LE")
    cRegexSubstitutions.Add Array("([^AEIOU])LE$", "$1ILE")
    cRegexSubstitutions.Add Array("([^AEIOU])LES$", "$1ILES")
    cRegexSubstitutions.Add Array("E$", "")
    cRegexSubstitutions.Add Array("ES$", "S")
    cRegexSubstitutions.Add Array("([AEIOU])SS$", "$1AS")
    cRegexSubstitutions.Add Array("([AEIOU])MB$", "$1M")
    cRegexSubstitutions.Add Array("MPTS", "MPS")
    cRegexSubstitutions.Add Array("MPS", "MS")
    cRegexSubstitutions.Add Array("MPT", "MT")
End Sub
