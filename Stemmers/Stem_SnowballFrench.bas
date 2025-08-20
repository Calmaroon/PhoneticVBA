Attribute VB_Name = "Stem_SnowballFrench"
Option Explicit
Dim StopWords As Dictionary
Dim Step1Suffixes As Collection
Dim Step2ASuffixes As Collection
Dim step2BSuffixes As Collection
Dim step4Suffixes As Collection
'This is a port from NLTK
Function SnowballFrench(strWord As String) As String
    If StopWords Is Nothing Or Step1Suffixes Is Nothing Or Step2ASuffixes Is Nothing Or step4Suffixes Is Nothing Then SetUpDictionaries
    
    strWord = LCase(strWord)
    
    If StopWords.Exists(strWord) Then
        SnowballFrench = strWord
        Exit Function
    End If
    
    Dim boolStep1Success As Boolean
    Dim boolRVEndingFound As Boolean
    Dim boolstep2ASuccess As Boolean
    Dim boolStep2BSuccess As Boolean
    
    Dim i As Integer
    If Len(strWord) > 1 Then
        For i = 2 To Len(strWord)
            If Mid(strWord, i - 1, 1) = "q" And Mid(strWord, i, 1) = "u" Then
                strWord = Left(strWord, i - 1) & "U" & Mid(strWord, i + 1)
            End If
        Next
    End If
    If Len(strWord) > 1 Then
        For i = 2 To Len(strWord)
            If Mid(strWord, i - 1, 1) Like "[aeiouyâàëéêèïîôûù]" And Mid(strWord, i, 1) = "u" And Mid(strWord, i + 1, 1) Like "[aeiouyâàëéêèïîôûù]" Then
                strWord = Left(strWord, i - 1) & "U" & Mid(strWord, i + 1)
            ElseIf Mid(strWord, i - 1, 1) Like "[aeiouyâàëéêèïîôûù]" And Mid(strWord, i, 1) = "i" And Mid(strWord, i + 1, 1) Like "[aeiouyâàëéêèïîôûù]" Then
                strWord = Left(strWord, i - 1) & "I" & Mid(strWord, i + 1)
            End If
            
            If Mid(strWord, i - 1, 1) Like "[aeiouyâàëéêèïîôûù]" Or Mid(strWord, i + 1, 1) Like "[aeiouyâàëéêèïîôûù]" Then
                If Mid(strWord, i, 1) = "y" Then
                    strWord = Left(strWord, i - 1) & "Y" & Mid(strWord, i + 1)
                End If
            End If
            
            
        Next
    End If
    'make U and I and Y logic too

    
    Dim r1 As String: r1 = sbR1R2(strWord)(0)
    Dim r2 As String: r2 = sbR1R2(strWord)(1)
    Dim rv As String: rv = sbRV(strWord)
    
    'Step 1
    Dim suffix As Variant
    For Each suffix In Step1Suffixes
        If Right(strWord, Len(suffix)) = suffix Then
            If suffix = "eaux" Then
                strWord = Left(strWord, Len(strWord) - 1)
                boolStep1Success = True
            ElseIf suffix = "euse" Or suffix = "euses" Then
                If InStr(r2, suffix) > 0 Then
                    strWord = Left(strWord, Len(strWord) - Len(suffix))
                    boolStep1Success = True
                ElseIf InStr(r1, suffix) > 0 Then
                    strWord = Left(strWord, Len(strWord) - Len(suffix)) & "eux"
                    boolStep1Success = True
                End If
            ElseIf (suffix = "ement" Or suffix = "ements") And InStr(rv, suffix) > 0 Then
                strWord = Left(strWord, Len(strWord) - Len(suffix))
                boolStep1Success = True
                
                If Right(strWord, 2) = "iv" And InStr(r2, "iv") > 0 Then
                    strWord = Left(strWord, Len(strWord) - 2)
                    If Right(strWord, 2) = "at" And InStr(r2, "at") > 0 Then
                        strWord = Left(strWord, Len(strWord) - 2)
                    End If
                ElseIf Right(strWord, 3) = "eus" Then
                    If InStr(r2, "eus") > 0 Then
                        strWord = Left(strWord, Len(strWord) - 3)
                    ElseIf InStr(r1, "eus") Then
                        strWord = Left(strWord, Len(strWord) - 1) & "x"
                    End If
                ElseIf Right(strWord, 3) = "abl" Or Right(strWord, 3) = "iqU" Then
                    If InStr(r2, "abl") > 0 Or InStr(r2, "iqU") > 0 Then
                        strWord = Left(strWord, Len(strWord) - 3)
                    End If
                ElseIf Right(strWord, 3) Like "[iI]èr" Then
                    If InStr(rv, "ièr") > 0 Or InStr(rv, "Ièr") > 0 Then
                        strWord = Left(strWord, Len(strWord) - 3) & "i"
                    End If
                End If
            ElseIf suffix = "amment" And InStr(rv, suffix) > 0 Then
                strWord = Left(strWord, Len(strWord) - 6) & "ant"
                rv = Left(rv, Len(rv) - 6) & "ant"
                boolRVEndingFound = True
            ElseIf suffix = "emment" And InStr(rv, suffix) > 0 Then
                strWord = Left(strWord, Len(strWord) - 6) & "ent"
                boolRVEndingFound = True
            ElseIf (suffix = "ment" Or suffix = "ments") And InStr(rv, suffix) > 0 And Left(rv, Len(suffix)) <> suffix Then ' and the letter before the suffix in rv is not a vowel
                If Len(rv) > Len(suffix) Then
                    If Mid(rv, Len(rv) - Len(suffix), 1) Like "[aeiouyâàëéêèïîôûù]" Then
                        strWord = Left(strWord, Len(strWord) - Len(suffix))
                        rv = Left(rv, Len(rv) - Len(suffix))
                        boolRVEndingFound = True
                    End If
                End If
            ElseIf suffix = "aux" And InStr(r1, suffix) > 0 Then
                strWord = Left(strWord, Len(strWord) - 2) & "l"
                boolStep1Success = True
            ElseIf (suffix = "issement" Or suffix = "issements") And InStr(r1, suffix) > 0 Then
                If Not Mid(strWord, Len(strWord) - Len(suffix), 1) Like "[aeiouyâàëéêèïîôûù]" Then
                    strWord = Left(strWord, Len(strWord) - Len(suffix))
                    boolStep1Success = True
                End If
            ElseIf (suffix = "ance" Or suffix = "iqUe" Or suffix = "isme" Or suffix = "able" Or suffix = "iste" Or suffix = "eux" Or suffix = "ances" Or suffix = "iqUes" Or suffix = "ismes" Or suffix = "ables" Or suffix = "istes") And InStr(r2, suffix) > 0 Then
                strWord = Left(strWord, Len(strWord) - Len(suffix))
                boolStep1Success = True
            ElseIf (suffix = "atrice" Or suffix = "ateur" Or suffix = "ation" Or suffix = "atrices" Or suffix = "ateurs" Or suffix = "ations") And InStr(r2, suffix) > 0 Then
                strWord = Left(strWord, Len(strWord) - Len(suffix))
                boolStep1Success = True
                
                If Right(strWord, 2) = "ic" Then
                    If InStr(r2, "ic") Then
                        strWord = Left(strWord, Len(strWord) - 2)
                    Else
                        strWord = Left(strWord, Len(strWord) - 2) & "iqU"
                    End If
                End If
            ElseIf (suffix = "logie" Or suffix = "logies") And InStr(r2, suffix) > 0 Then
                strWord = Left(strWord, Len(strWord) - Len(suffix)) & "log"
                boolStep1Success = True
            ElseIf (suffix = "usion" Or suffix = "ution" Or suffix = "usions" Or suffix = "utions") And InStr(r2, suffix) > 0 Then
                strWord = Left(strWord, Len(strWord) - Len(suffix)) & "u"
                boolStep1Success = True
            ElseIf (suffix = "ence" Or suffix = "ences") And InStr(r2, suffix) > 0 Then
                strWord = Left(strWord, Len(strWord) - Len(suffix)) & "ent"
                boolStep1Success = True
            ElseIf (suffix = "ité" Or suffix = "ités") And InStr(r2, suffix) > 0 Then
                strWord = Left(strWord, Len(strWord) - Len(suffix))
                boolStep1Success = True
                If Right(strWord, 4) = "abil" Then
                    If InStr(r2, "abil") > 0 Then
                        strWord = Left(strWord, Len(strWord) - 4)
                    Else
                        strWord = Left(strWord, Len(strWord) - 2) & "l"
                    End If
                ElseIf Right(strWord, 2) = "ic" Then
                    If InStr(r2, "ic") > 0 Then
                        strWord = Left(strWord, Len(strWord) - 2)
                    Else
                        strWord = Left(strWord, Len(strWord) - 2) & "iqU"
                    End If
                ElseIf Right(strWord, 2) = "iv" Then
                    If InStr(r2, "iv") > 0 Then
                        strWord = Left(strWord, Len(strWord) - 2)
                    End If
                End If
            ElseIf (suffix = "if" Or suffix = "ive" Or suffix = "ifs" Or suffix = "ives") And InStr(r2, suffix) > 0 Then
                strWord = Left(strWord, Len(strWord) - Len(suffix))
                boolStep1Success = True
                If Right(strWord, 2) = "at" And InStr(r2, "at") > 0 Then
                    strWord = Left(strWord, Len(strWord) - 2)
                    
                    If Right(strWord, 2) = "ic" Then
                        If InStr(r2, "ic") > 0 Then
                            strWord = Left(strWord, Len(strWord) - 2)
                        Else
                            strWord = Left(strWord, Len(strWord) - 2) & "iqU"
                        End If
                    End If
                End If
            End If
            Exit For
        End If
    Next
    
    'Step 2A
    If Not boolStep1Success Or boolRVEndingFound Then
        For Each suffix In Step2ASuffixes
            If Right(strWord, Len(suffix)) = suffix Then
                If InStr(rv, suffix) > 0 And Len(rv) > Len(suffix) Then
                    If Not Mid(rv, Len(rv) - Len(suffix), 1) Like "[aeiouyâàëéêèïîôûù]" Then
                        strWord = Left(strWord, Len(strWord) - Len(suffix))
                        boolstep2ASuccess = True
                    End If
                End If
                Exit For
            End If
        Next
        
        
        If Not boolstep2ASuccess Then
            For Each suffix In step2BSuffixes
                If Right(rv, Len(suffix)) = suffix Then
                    If suffix = "ions" And InStr(r2, "ions") > 0 Then
                        strWord = Left(strWord, Len(strWord) - 4)
                        boolStep2BSuccess = True
                    ElseIf suffix = "eraIent" Or suffix = "erions" Or suffix = "èrent" Or suffix = "erais" Or suffix = "erait" Or suffix = "eriez" Or suffix = "erons" Or suffix = "eront" _
                        Or suffix = "erai" Or suffix = "eras" Or suffix = "erez" Or suffix = "ées" Or suffix = "era" Or suffix = "iez" Or suffix = "ée" Or suffix = "és" Or suffix = "er" Or suffix = "ez" Or suffix = "é" Then 'IN BIG LIST
                        strWord = Left(strWord, Len(strWord) - Len(suffix))
                        boolStep2BSuccess = True
                    
                    ElseIf suffix = "assions" Or suffix = "assent" Or suffix = "assiez" Or suffix = "aIent" Or suffix = "antes" Or suffix = "asses" Or suffix = "âmes" Or suffix = "âtes" Or suffix = "ante" Or suffix = "ants" Or suffix = "asse" Or suffix = "ais" _
                        Or suffix = "ait" Or suffix = "ant" Or suffix = "ât" Or suffix = "ai" Or suffix = "as" Or suffix = "a" Then
                        strWord = Left(strWord, Len(strWord) - Len(suffix))
                        If Len(rv) >= Len(suffix) Then rv = Left(rv, Len(rv) - Len(suffix))
                        
                        boolStep2BSuccess = True
                        If Right(rv, 1) = "e" Then
                            strWord = Left(strWord, Len(strWord) - 1)
                        End If
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    
    
    'step 3
    If boolStep1Success Or boolstep2ASuccess Or boolStep2BSuccess Then
        If Right(strWord, 1) = "Y" Then
            strWord = Left(strWord, Len(strWord) - 1) & "i"
        ElseIf Right(strWord, 1) = "ç" Then
            strWord = Left(strWord, Len(strWord) - 1) & "c"
        End If
    Else
        'Step 4 Residual Suffixes
        If Len(strWord) > 1 Then
            If Len(strWord) >= 2 And Right(strWord, 1) = "s" And Not Mid(strWord, Len(strWord) - 1, 1) Like "[aiouès]" Then
                strWord = Left(strWord, Len(strWord) - 1)
            End If
        End If
        
        For Each suffix In step4Suffixes
            If Right(strWord, Len(suffix)) = suffix Then
                If InStr(rv, suffix) > 0 Then
                    If suffix = "ion" And InStr(r2, suffix) > 0 And Len(rv) >= 4 Then
                        If Mid(rv, Len(rv) - 3, 1) Like "[st]" Then
                            strWord = Left(strWord, Len(strWord) - 3)
                        End If
                    ElseIf suffix = "ier" Or suffix = "Ier" Or suffix = "ière" Or suffix = "Ière" Then
                        strWord = Left(strWord, Len(strWord) - Len(suffix)) & "i"
                    ElseIf suffix = "e" Then
                        strWord = Left(strWord, Len(strWord) - 1)
                    ElseIf suffix = "ë" And Len(strWord) >= 3 Then
                        If Mid(strWord, Len(strWord) - 2, 2) = "gu" Then
                            strWord = Left(strWord, Len(strWord) - 1)
                        End If
                    End If
                    Exit For
                End If
            End If
        Next
        
    End If
    
    'step 5: Undouble
    If Right(strWord, 3) = "enn" Or Right(strWord, 3) = "onn" Or Right(strWord, 3) = "ett" Or Right(strWord, 3) = "ell" Or Right(strWord, 4) = "eill" Then
        strWord = Left(strWord, Len(strWord) - 1)
    End If
    
    'step 6 unaccent
    For i = Len(strWord) To 1 Step -1
        If Mid(strWord, i, 1) Like "[aeiouyâàëéêèïîôûù]" Then
            If i <> Len(strWord) And (Mid(strWord, i, 1) = "é" Or Mid(strWord, i, 1) = "è") And i <> 1 Then
                strWord = Left(strWord, i - 1) & "e" & Mid(strWord, i + 1)
            End If
            Exit For
        End If
    Next

    SnowballFrench = Replace(Replace(Replace(strWord, "I", "i"), "U", "u"), "Y", "y")
End Function
Sub SetUpDictionaries()
    Set StopWords = New Dictionary
        StopWords.Add "à", ""
        StopWords.Add "ai", ""
        StopWords.Add "aie", ""
        StopWords.Add "aient", ""
        StopWords.Add "aies", ""
        StopWords.Add "ait", ""
        StopWords.Add "as", ""
        StopWords.Add "au", ""
        StopWords.Add "aura", ""
        StopWords.Add "aurai", ""
        StopWords.Add "auraient", ""
        StopWords.Add "aurais", ""
        StopWords.Add "aurait", ""
        StopWords.Add "auras", ""
        StopWords.Add "aurez", ""
        StopWords.Add "auriez", ""
        StopWords.Add "aurions", ""
        StopWords.Add "aurons", ""
        StopWords.Add "auront", ""
        StopWords.Add "aux", ""
        StopWords.Add "avaient", ""
        StopWords.Add "avais", ""
        StopWords.Add "avait", ""
        StopWords.Add "avec", ""
        StopWords.Add "avez", ""
        StopWords.Add "aviez", ""
        StopWords.Add "avions", ""
        StopWords.Add "avons", ""
        StopWords.Add "ayant", ""
        StopWords.Add "ayante", ""
        StopWords.Add "ayantes", ""
        StopWords.Add "ayants", ""
        StopWords.Add "ayez", ""
        StopWords.Add "ayons", ""
        StopWords.Add "c", ""
        StopWords.Add "ce", ""
        StopWords.Add "ces", ""
        StopWords.Add "d", ""
        StopWords.Add "dans", ""
        StopWords.Add "de", ""
        StopWords.Add "des", ""
        StopWords.Add "du", ""
        StopWords.Add "elle", ""
        StopWords.Add "en", ""
        StopWords.Add "es", ""
        StopWords.Add "est", ""
        StopWords.Add "et", ""
        StopWords.Add "étaient", ""
        StopWords.Add "étais", ""
        StopWords.Add "était", ""
        StopWords.Add "étant", ""
        StopWords.Add "étante", ""
        StopWords.Add "étantes", ""
        StopWords.Add "étants", ""
        StopWords.Add "été", ""
        StopWords.Add "étée", ""
        StopWords.Add "étées", ""
        StopWords.Add "étés", ""
        StopWords.Add "êtes", ""
        StopWords.Add "étiez", ""
        StopWords.Add "étions", ""
        StopWords.Add "eu", ""
        StopWords.Add "eue", ""
        StopWords.Add "eues", ""
        StopWords.Add "eûmes", ""
        StopWords.Add "eurent", ""
        StopWords.Add "eus", ""
        StopWords.Add "eusse", ""
        StopWords.Add "eussent", ""
        StopWords.Add "eusses", ""
        StopWords.Add "eussiez", ""
        StopWords.Add "eussions", ""
        StopWords.Add "eut", ""
        StopWords.Add "eût", ""
        StopWords.Add "eûtes", ""
        StopWords.Add "eux", ""
        StopWords.Add "fûmes", ""
        StopWords.Add "furent", ""
        StopWords.Add "fus", ""
        StopWords.Add "fusse", ""
        StopWords.Add "fussent", ""
        StopWords.Add "fusses", ""
        StopWords.Add "fussiez", ""
        StopWords.Add "fussions", ""
        StopWords.Add "fut", ""
        StopWords.Add "fût", ""
        StopWords.Add "fûtes", ""
        StopWords.Add "il", ""
        StopWords.Add "ils", ""
        StopWords.Add "j", ""
        StopWords.Add "je", ""
        StopWords.Add "l", ""
        StopWords.Add "la", ""
        StopWords.Add "le", ""
        StopWords.Add "les", ""
        StopWords.Add "leur", ""
        StopWords.Add "lui", ""
        StopWords.Add "m", ""
        StopWords.Add "ma", ""
        StopWords.Add "mais", ""
        StopWords.Add "me", ""
        StopWords.Add "même", ""
        StopWords.Add "mes", ""
        StopWords.Add "moi", ""
        StopWords.Add "mon", ""
        StopWords.Add "n", ""
        StopWords.Add "ne", ""
        StopWords.Add "nos", ""
        StopWords.Add "notre", ""
        StopWords.Add "nous", ""
        StopWords.Add "on", ""
        StopWords.Add "ont", ""
        StopWords.Add "ou", ""
        StopWords.Add "par", ""
        StopWords.Add "pas", ""
        StopWords.Add "pour", ""
        StopWords.Add "qu", ""
        StopWords.Add "que", ""
        StopWords.Add "qui", ""
        StopWords.Add "s", ""
        StopWords.Add "sa", ""
        StopWords.Add "se", ""
        StopWords.Add "sera", ""
        StopWords.Add "serai", ""
        StopWords.Add "seraient", ""
        StopWords.Add "serais", ""
        StopWords.Add "serait", ""
        StopWords.Add "seras", ""
        StopWords.Add "serez", ""
        StopWords.Add "seriez", ""
        StopWords.Add "serions", ""
        StopWords.Add "serons", ""
        StopWords.Add "seront", ""
        StopWords.Add "ses", ""
        StopWords.Add "soient", ""
        StopWords.Add "sois", ""
        StopWords.Add "soit", ""
        StopWords.Add "sommes", ""
        StopWords.Add "son", ""
        StopWords.Add "sont", ""
        StopWords.Add "soyez", ""
        StopWords.Add "soyons", ""
        StopWords.Add "suis", ""
        StopWords.Add "sur", ""
        StopWords.Add "t", ""
        StopWords.Add "ta", ""
        StopWords.Add "te", ""
        StopWords.Add "tes", ""
        StopWords.Add "toi", ""
        StopWords.Add "ton", ""
        StopWords.Add "tu", ""
        StopWords.Add "un", ""
        StopWords.Add "une", ""
        StopWords.Add "vos", ""
        StopWords.Add "votre", ""
        StopWords.Add "vous", ""
        StopWords.Add "y", ""

        
    Set Step1Suffixes = New Collection
        Step1Suffixes.Add "issements"
        Step1Suffixes.Add "issement"
        Step1Suffixes.Add "atrices"
        Step1Suffixes.Add "atrice"
        Step1Suffixes.Add "ateurs"
        Step1Suffixes.Add "ations"
        Step1Suffixes.Add "logies"
        Step1Suffixes.Add "usions"
        Step1Suffixes.Add "utions"
        Step1Suffixes.Add "ements"
        Step1Suffixes.Add "amment"
        Step1Suffixes.Add "emment"
        Step1Suffixes.Add "ances"
        Step1Suffixes.Add "iqUes"
        Step1Suffixes.Add "ismes"
        Step1Suffixes.Add "ables"
        Step1Suffixes.Add "istes"
        Step1Suffixes.Add "ateur"
        Step1Suffixes.Add "ation"
        Step1Suffixes.Add "logie"
        Step1Suffixes.Add "usion"
        Step1Suffixes.Add "ution"
        Step1Suffixes.Add "ences"
        Step1Suffixes.Add "ement"
        Step1Suffixes.Add "euses"
        Step1Suffixes.Add "ments"
        Step1Suffixes.Add "ance"
        Step1Suffixes.Add "iqUe"
        Step1Suffixes.Add "isme"
        Step1Suffixes.Add "able"
        Step1Suffixes.Add "iste"
        Step1Suffixes.Add "ence"
        Step1Suffixes.Add "ités"
        Step1Suffixes.Add "ives"
        Step1Suffixes.Add "eaux"
        Step1Suffixes.Add "euse"
        Step1Suffixes.Add "ment"
        Step1Suffixes.Add "eux"
        Step1Suffixes.Add "ité"
        Step1Suffixes.Add "ive"
        Step1Suffixes.Add "ifs"
        Step1Suffixes.Add "aux"
        Step1Suffixes.Add "if"
    
    Set Step2ASuffixes = New Collection
        Step2ASuffixes.Add "issaIent"
        Step2ASuffixes.Add "issantes"
        Step2ASuffixes.Add "iraIent"
        Step2ASuffixes.Add "issante"
        Step2ASuffixes.Add "issants"
        Step2ASuffixes.Add "issions"
        Step2ASuffixes.Add "irions"
        Step2ASuffixes.Add "issais"
        Step2ASuffixes.Add "issait"
        Step2ASuffixes.Add "issant"
        Step2ASuffixes.Add "issent"
        Step2ASuffixes.Add "issiez"
        Step2ASuffixes.Add "issons"
        Step2ASuffixes.Add "irais"
        Step2ASuffixes.Add "irait"
        Step2ASuffixes.Add "irent"
        Step2ASuffixes.Add "iriez"
        Step2ASuffixes.Add "irons"
        Step2ASuffixes.Add "iront"
        Step2ASuffixes.Add "isses"
        Step2ASuffixes.Add "issez"
        Step2ASuffixes.Add "îmes"
        Step2ASuffixes.Add "îtes"
        Step2ASuffixes.Add "irai"
        Step2ASuffixes.Add "iras"
        Step2ASuffixes.Add "irez"
        Step2ASuffixes.Add "isse"
        Step2ASuffixes.Add "ies"
        Step2ASuffixes.Add "ira"
        Step2ASuffixes.Add "ît"
        Step2ASuffixes.Add "ie"
        Step2ASuffixes.Add "ir"
        Step2ASuffixes.Add "is"
        Step2ASuffixes.Add "it"
        Step2ASuffixes.Add "i"

    Set step2BSuffixes = New Collection
        step2BSuffixes.Add "eraIent"
        step2BSuffixes.Add "assions"
        step2BSuffixes.Add "erions"
        step2BSuffixes.Add "assent"
        step2BSuffixes.Add "assiez"
        step2BSuffixes.Add "èrent"
        step2BSuffixes.Add "erais"
        step2BSuffixes.Add "erait"
        step2BSuffixes.Add "eriez"
        step2BSuffixes.Add "erons"
        step2BSuffixes.Add "eront"
        step2BSuffixes.Add "aIent"
        step2BSuffixes.Add "antes"
        step2BSuffixes.Add "asses"
        step2BSuffixes.Add "ions"
        step2BSuffixes.Add "erai"
        step2BSuffixes.Add "eras"
        step2BSuffixes.Add "erez"
        step2BSuffixes.Add "âmes"
        step2BSuffixes.Add "âtes"
        step2BSuffixes.Add "ante"
        step2BSuffixes.Add "ants"
        step2BSuffixes.Add "asse"
        step2BSuffixes.Add "ées"
        step2BSuffixes.Add "era"
        step2BSuffixes.Add "iez"
        step2BSuffixes.Add "ais"
        step2BSuffixes.Add "ait"
        step2BSuffixes.Add "ant"
        step2BSuffixes.Add "ée"
        step2BSuffixes.Add "és"
        step2BSuffixes.Add "er"
        step2BSuffixes.Add "ez"
        step2BSuffixes.Add "ât"
        step2BSuffixes.Add "ai"
        step2BSuffixes.Add "as"
        step2BSuffixes.Add "é"
        step2BSuffixes.Add "a"
        
    Set step4Suffixes = New Collection
        step4Suffixes.Add "ière"
        step4Suffixes.Add "Ière"
        step4Suffixes.Add "ion"
        step4Suffixes.Add "ier"
        step4Suffixes.Add "Ier"
        step4Suffixes.Add "e"
        step4Suffixes.Add "ë"
        
End Sub
Function sbR1R2(strWord As String) As String()
    Dim strResult() As String
    ReDim strResult(0 To 1)
    
    Dim strR1 As String
    Dim strR2 As String
    
    Dim i As Integer
    For i = 2 To Len(strWord)
        If Not Mid(strWord, i, 1) Like "[aeiouyâàëéêèïîôûù]" And Mid(strWord, i - 1, 1) Like "[aeiouyâàëéêèïîôûù]" Then
            strR1 = Mid(strWord, i + 1)
            Exit For
        End If
    Next
    
    For i = 2 To Len(strR1)
        If Not Mid(strR1, i, 1) Like "[aeiouyâàëéêèïîôûù]" And Mid(strR1, i - 1, 1) Like "[aeiouyâàëéêèïîôûù]" Then
            strR2 = Mid(strR1, i + 1)
            Exit For
        End If
    Next
    strResult(0) = strR1: strResult(1) = strR2
    sbR1R2 = strResult()
End Function
Function sbRV(strWord As String) As String
    Dim rv As String
    If Len(strWord) >= 2 Then
        If (Left(strWord, 3) = "par" Or Left(strWord, 3) = "col" Or Left(strWord, 3) = "tap") Or (Left(strWord, 1) Like "[aeiouyâàëéêèïîôûù]" And Mid(strWord, 2, 1) Like "[aeiouyâàëéêèïîôûù]") Then
            rv = Mid(strWord, 4)
        Else
            Dim i As Integer
            For i = 2 To Len(strWord)
                If Mid(strWord, i, 1) Like "[aeiouyâàëéêèïîôûù]" Then
                    rv = Mid(strWord, i + 1)
                    Exit For
                End If
            Next
        End If
    End If
    
    sbRV = rv
End Function




