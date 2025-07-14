Attribute VB_Name = "Encode_RogerRoot"
Option Explicit
Dim MedPatterns As Collection
Dim InitPatterns As Collection
Sub SetUpPatterns()
    Set InitPatterns = New Collection
    InitPatterns.Add Array("TSCH", "06")
    InitPatterns.Add Array("TSH", "06")
    InitPatterns.Add Array("SCH", "06")
    
    InitPatterns.Add Array("CE", "0*0")
    InitPatterns.Add Array("CH", "06")
    InitPatterns.Add Array("CI", "0*0")
    InitPatterns.Add Array("CY", "0*0")
    InitPatterns.Add Array("DG", "07")
    InitPatterns.Add Array("GF", "08")
    InitPatterns.Add Array("GM", "03")
    InitPatterns.Add Array("GN", "02")
    InitPatterns.Add Array("KN", "02")
    InitPatterns.Add Array("PF", "08")
    InitPatterns.Add Array("PH", "08")
    InitPatterns.Add Array("PN", "02")
    InitPatterns.Add Array("SH", "06")
    InitPatterns.Add Array("TS", "0*0")
    InitPatterns.Add Array("WR", "04")
    
    InitPatterns.Add Array("A", "1")
    InitPatterns.Add Array("B", "09")
    InitPatterns.Add Array("C", "07")
    InitPatterns.Add Array("D", "01")
    InitPatterns.Add Array("E", "1")
    InitPatterns.Add Array("F", "08")
    InitPatterns.Add Array("G", "07")
    InitPatterns.Add Array("H", "2")
    InitPatterns.Add Array("I", "1")
    InitPatterns.Add Array("J", "3")
    InitPatterns.Add Array("K", "07")
    InitPatterns.Add Array("L", "05")
    InitPatterns.Add Array("M", "03")
    InitPatterns.Add Array("N", "02")
    InitPatterns.Add Array("O", "1")
    InitPatterns.Add Array("P", "09")
    InitPatterns.Add Array("Q", "07")
    InitPatterns.Add Array("R", "04")
    InitPatterns.Add Array("S", "0*0")
    InitPatterns.Add Array("T", "01")
    InitPatterns.Add Array("U", "1")
    InitPatterns.Add Array("V", "08")
    InitPatterns.Add Array("W", "4")
    InitPatterns.Add Array("X", "07")
    InitPatterns.Add Array("Y", "5")
    InitPatterns.Add Array("Z", "0*0")

    Set MedPatterns = New Collection
    MedPatterns.Add Array("TSCH", "6")
    MedPatterns.Add Array("TSH", "6")
    MedPatterns.Add Array("SCH", "6")
    MedPatterns.Add Array("CE", "0")
    
    MedPatterns.Add Array("CH", "6")
    MedPatterns.Add Array("CI", "0")
    MedPatterns.Add Array("CY", "0")
    MedPatterns.Add Array("DG", "7")
    MedPatterns.Add Array("PH", "8")
    MedPatterns.Add Array("SH", "6")
    MedPatterns.Add Array("TS", "0")
    MedPatterns.Add Array("B", "9")
    MedPatterns.Add Array("C", "7")
    MedPatterns.Add Array("D", "1")
    MedPatterns.Add Array("F", "8")
    MedPatterns.Add Array("G", "7")
    MedPatterns.Add Array("J", "6")
    MedPatterns.Add Array("K", "7")
    MedPatterns.Add Array("L", "5")
    MedPatterns.Add Array("M", "3")
    MedPatterns.Add Array("N", "2")
    MedPatterns.Add Array("P", "9")
    MedPatterns.Add Array("Q", "7")
    MedPatterns.Add Array("R", "4")
    MedPatterns.Add Array("S", "0")
    MedPatterns.Add Array("T", "1")
    MedPatterns.Add Array("V", "8")
    MedPatterns.Add Array("X", "7")
    MedPatterns.Add Array("Z", "0")
    MedPatterns.Add Array("A", "*")
    MedPatterns.Add Array("E", "*")
    MedPatterns.Add Array("H", "*")
    MedPatterns.Add Array("I", "*")
    MedPatterns.Add Array("O", "*")
    MedPatterns.Add Array("U", "*")
    MedPatterns.Add Array("W", "*")
    MedPatterns.Add Array("Y", "*")
End Sub
Function RogerRoot(strInput As String, Optional intMaxLength As Integer = 5, Optional boolZeroPad As Boolean = True) As String
    If MedPatterns Is Nothing Or InitPatterns Is Nothing Then Call SetUpPatterns
    strInput = UCase$(strInput)
    Dim intPos As Integer
    Dim Pattern As Variant
    Dim strCode As String
    Dim strSubstr As String
    Dim boolMatched As Boolean
    
    intPos = 1
    For Each Pattern In InitPatterns
        If left$(strInput, Len(Pattern(0))) = Pattern(0) Then
            strCode = strCode & Pattern(1)
            intPos = intPos + Len(Pattern(0))
            Exit For
        End If
    Next
    
    Do While intPos <= Len(strInput)
        boolMatched = False
        strSubstr = Mid$(strInput, intPos)
        For Each Pattern In MedPatterns
            If left$(strSubstr, Len(Pattern(0))) = Pattern(0) Then
                strCode = strCode & Pattern(1)
                intPos = intPos + Len(Pattern(0))
                boolMatched = True
                Exit For
            End If
        Next
        
        If Not boolMatched Then
            strCode = strCode & Mid$(strInput, intPos, 1)
            intPos = intPos + 1
        End If
    Loop
    strCode = DeleteConsecutiveRepeats(strCode)
    strCode = Replace$(strCode, "*", "")
    
    If boolZeroPad Then strCode = strCode & String(intMaxLength, "0")
    
    RogerRoot = left$(strCode, intMaxLength)
End Function
