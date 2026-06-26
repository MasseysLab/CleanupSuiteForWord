Option Explicit

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub

Private Sub UserForm_Initialize()
    ApplyMSFormTitleStrategy Me, True

    optAll.GroupName = "CapMode"
    optSentence.GroupName = "CapMode"
    optTitle.GroupName = "CapMode"
    optUpper.GroupName = "CapMode"
    optLower.GroupName = "CapMode"
    optCustom.GroupName = "CapMode"
    optScopeDocument.GroupName = "CapScope"
    optScopeSelection.GroupName = "CapScope"

    optAll.Caption = "Conservative"
    optSentence.Caption = "Balanced (coming soon)"
    optTitle.Caption = "Aggressive (coming soon)"
    optCustom.Caption = "Custom"

    optAll.Visible = True
    optSentence.Visible = True
    optTitle.Visible = True
    optUpper.Visible = False
    optLower.Visible = False
    optCustom.Visible = True
    optSentence.Enabled = False
    optTitle.Enabled = False

    optAll.Value = True

    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If

    lblIntro.Caption = "Choose how strongly capitalization should be repaired."
    chkSentence.Caption = "Use abbreviation lists"
    chkTitle.Caption = "Protect common acronyms"
    chkUpper.Caption = "Protect names and brands"
    chkLower.Caption = "Treat likely headings carefully"
    chkSmartSentences.Caption = "Use quote, bullet, and boundary context"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"

    ApplyModeDefaults
    UpdateAdvancedVisibility
    UpdateModeSummary
    LayoutCleanupToolForm Me
End Sub

Private Sub UpdateAdvancedVisibility()
    fraCustom.Visible = False
    optUpper.Visible = False
    optLower.Visible = False
    chkSentence.Visible = optCustom.Value
    chkTitle.Visible = optCustom.Value
    chkUpper.Visible = optCustom.Value
    chkLower.Visible = optCustom.Value
    chkSmartSentences.Visible = optCustom.Value
    cmdSelectAll.Visible = optCustom.Value
    cmdDeselectAll.Visible = optCustom.Value
End Sub

Private Sub ApplyModeDefaults()
    Select Case ActiveRepairModeName(False)
        Case "Conservative"
            chkSentence.Value = True
            chkTitle.Value = True
            chkUpper.Value = True
            chkLower.Value = True
            chkSmartSentences.Value = True
        Case Else
            If Not (chkSentence.Value Or chkTitle.Value Or chkUpper.Value Or chkLower.Value Or chkSmartSentences.Value) Then
                chkSentence.Value = True
                chkTitle.Value = True
                chkUpper.Value = True
                chkLower.Value = True
                chkSmartSentences.Value = True
            End If
    End Select
End Sub

Private Sub UpdateModeSummary()
    Select Case ActiveRepairModeName(False)
        Case "Conservative"
            lblModeSummary.Caption = "Fixes only obvious capitalization damage and leaves already-reasonable wording alone."
        Case Else
            lblModeSummary.Caption = "Custom uses only the protections you select below."
    End Select
End Sub

Private Function ActiveRepairModeName(Optional ByVal unused As Boolean = False) As String
    If optAll.Value Then
        ActiveRepairModeName = "Conservative"
        Exit Function
    End If
    If optCustom.Value Then
        ActiveRepairModeName = "Custom"
        Exit Function
    End If
    ActiveRepairModeName = "Conservative"
End Function

Private Sub RefreshCapitalizationChoices()
    ApplyModeDefaults
    UpdateAdvancedVisibility
    UpdateModeSummary
    LayoutCleanupToolForm Me
End Sub

Private Sub optAll_Click(): RefreshCapitalizationChoices: End Sub
Private Sub optSentence_Click(): RefreshCapitalizationChoices: End Sub
Private Sub optTitle_Click(): RefreshCapitalizationChoices: End Sub
Private Sub optUpper_Click(): RefreshCapitalizationChoices: End Sub
Private Sub optLower_Click(): RefreshCapitalizationChoices: End Sub
Private Sub optCustom_Click(): RefreshCapitalizationChoices: End Sub

Private Sub cmdSelectAll_Click()
    optCustom.Value = True
    UpdateAdvancedVisibility
    UpdateModeSummary
    chkSentence.Value = True
    chkTitle.Value = True
    chkUpper.Value = True
    chkLower.Value = True
    chkSmartSentences.Value = True
    LayoutCleanupToolForm Me
End Sub

Private Sub cmdDeselectAll_Click()
    optCustom.Value = True
    UpdateAdvancedVisibility
    UpdateModeSummary
    chkSentence.Value = False
    chkTitle.Value = False
    chkUpper.Value = False
    chkLower.Value = False
    chkSmartSentences.Value = False
    LayoutCleanupToolForm Me
End Sub

Private Sub chkSentence_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkTitle_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkUpper_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkLower_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkSmartSentences_Click(): LayoutCleanupToolForm Me: End Sub

Private Sub cmdPreview_Click()
    PreviewFromPanel
End Sub

Public Sub PreviewFromPanel()
    chkPreviewOnly.Value = True
    cmdRun_Click
    chkPreviewOnly.Value = False
End Sub

Private Sub cmdReset_Click()
    UserForm_Initialize
    chkPreviewOnly.Value = False
End Sub

Public Sub RunAfterPreview()
    chkPreviewOnly.Value = False
    cmdRun_Click
End Sub

Private Sub cmdRun_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected. Please remove protection before running cleanup.", vbExclamation, "Document Protected"
        Exit Sub
    End If
    If Not GuardBeforeCleanup("Capitalization Fixer") Then
        Unload Me
        Exit Sub
    End If

    Dim previewOnly As Boolean
    previewOnly = chkPreviewOnly.Value

    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If

    Dim modeName As String
    modeName = ActiveRepairModeName(False)

    Dim useAbbreviations As Boolean
    Dim protectAcronyms As Boolean
    Dim protectNames As Boolean
    Dim useHeadingHeuristics As Boolean
    Dim useContextRules As Boolean
    useAbbreviations = chkSentence.Value
    protectAcronyms = chkTitle.Value
    protectNames = chkUpper.Value
    useHeadingHeuristics = chkLower.Value
    useContextRules = chkSmartSentences.Value
    If modeName = "Conservative" Then
        useAbbreviations = True
        protectAcronyms = True
        protectNames = True
        useHeadingHeuristics = True
        useContextRules = True
    End If
    If modeName = "Custom" Then
        If Not (useAbbreviations Or protectAcronyms Or protectNames Or useHeadingHeuristics Or useContextRules) Then
            MsgBox "No custom repair protections are selected.", vbInformation, "Capitalization Fixer"
            Exit Sub
        End If
    End If

    Dim p As Paragraph
    Dim original As String
    Dim transformed As String
    Dim changedCount As Long
    changedCount = 0

    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
            original = p.Range.Text
            If Right$(original, 1) = vbCr Then original = Left$(original, Len(original) - 1)
            transformed = SmartRepairParagraph(original, modeName, useAbbreviations, protectAcronyms, protectNames, useHeadingHeuristics, useContextRules)
            If transformed <> original Then
                p.Range.HighlightColorIndex = wdYellow
                changedCount = changedCount + 1
            End If
        End If
    Next p

    If previewOnly Then
        ShowPreviewActions Me, "Capitalization Fixer", "Preview complete. " & changedCount & " paragraphs highlighted."
        Exit Sub
    End If

    MarkCleanupStart "Capitalization Fixer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Capitalization Fixer"
    On Error GoTo RunErr

    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
            original = p.Range.Text
            If Right$(original, 1) = vbCr Then original = Left$(original, Len(original) - 1)
            transformed = SmartRepairParagraph(original, modeName, useAbbreviations, protectAcronyms, protectNames, useHeadingHeuristics, useContextRules)
            If transformed <> original Then p.Range.Text = transformed & vbCr
        End If
    Next p

    RemoveAllHighlighting targetRange
    Dim results As Collection
    Set results = New Collection
    results.Add "Mode: " & modeName
    results.Add "Paragraphs changed: " & changedCount
    ShowCleanupReport "Capitalization Cleanup", results

    undoRec.EndCustomRecord
    MarkCleanupEnd
    MarkCleanupToolApplied
    Unload Me
    Exit Sub

RunErr:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    On Error Resume Next
    undoRec.EndCustomRecord
    On Error GoTo 0
    MarkCleanupEnd
    Dim errMsg As String
    errMsg = "An unexpected error occurred: " & originalErrNumber & " - " & originalErrDescription
    errMsg = errMsg & vbCrLf & vbCrLf & _
             "Some changes may have been made to the document." & vbCrLf & _
             "Use Word's Undo command (Ctrl+Z) after closing this message if you want to roll them back."
    MsgBox errMsg, vbCritical, "Cleanup Error"
End Sub

Private Function SmartRepairParagraph(ByVal textValue As String, ByVal modeName As String, ByVal useAbbreviations As Boolean, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean, ByVal useHeadingHeuristics As Boolean, ByVal useContextRules As Boolean) As String
    If Len(Trim$(textValue)) = 0 Then
        SmartRepairParagraph = textValue
        Exit Function
    End If

    Dim headingLike As Boolean
    headingLike = useHeadingHeuristics And IsLikelyHeadingText(textValue)

    Dim working As String
    working = textValue

    If ShouldNormalizeEntireParagraph(textValue, modeName, headingLike) Then
        working = LCase$(working)
    End If

    working = NormalizeLongUppercaseWords(working, modeName, protectAcronyms, protectNames, headingLike)
    working = ApplySentenceRepairs(working, modeName, useAbbreviations, useContextRules, headingLike)
    working = RestoreProtectedTerms(working, protectAcronyms, protectNames)

    SmartRepairParagraph = working
End Function

Private Function ShouldNormalizeEntireParagraph(ByVal textValue As String, ByVal modeName As String, ByVal headingLike As Boolean) As Boolean
    If headingLike Then Exit Function

    Dim lowerCount As Long
    Dim upperCount As Long
    Dim letterCount As Long
    CountLetterCase textValue, lowerCount, upperCount, letterCount

    If letterCount < 6 Then Exit Function

    Dim lowerRatio As Double
    Dim upperRatio As Double
    lowerRatio = lowerCount / letterCount
    upperRatio = upperCount / letterCount

    Select Case modeName
        Case "Balanced"
            If lowerRatio >= 0.92 And upperCount <= 2 Then ShouldNormalizeEntireParagraph = True
            If upperRatio >= 0.9 And CountWords(textValue) >= 4 Then ShouldNormalizeEntireParagraph = True
        Case "Aggressive"
            If lowerRatio >= 0.78 Or upperRatio >= 0.78 Then ShouldNormalizeEntireParagraph = True
            If CountAllUpperWords(textValue, 4) >= 2 Then ShouldNormalizeEntireParagraph = True
    End Select
End Function

Private Sub CountLetterCase(ByVal textValue As String, ByRef lowerCount As Long, ByRef upperCount As Long, ByRef letterCount As Long)
    Dim i As Long
    Dim currentChar As String
    For i = 1 To Len(textValue)
        currentChar = Mid$(textValue, i, 1)
        If currentChar Like "[A-Za-z]" Then
            letterCount = letterCount + 1
            If currentChar = LCase$(currentChar) Then
                lowerCount = lowerCount + 1
            Else
                upperCount = upperCount + 1
            End If
        End If
    Next i
End Sub

Private Function CountWords(ByVal textValue As String) As Long
    Dim i As Long
    Dim inWord As Boolean
    Dim currentChar As String
    For i = 1 To Len(textValue)
        currentChar = Mid$(textValue, i, 1)
        If IsTokenWordChar(currentChar) Then
            If Not inWord Then
                CountWords = CountWords + 1
                inWord = True
            End If
        Else
            inWord = False
        End If
    Next i
End Function

Private Function CountAllUpperWords(ByVal textValue As String, ByVal minLetters As Long) As Long
    Dim i As Long
    Dim token As String
    Dim currentChar As String
    For i = 1 To Len(textValue) + 1
        If i <= Len(textValue) Then
            currentChar = Mid$(textValue, i, 1)
        Else
            currentChar = " "
        End If

        If IsTokenWordChar(currentChar) Then
            token = token & currentChar
        ElseIf Len(token) > 0 Then
            If IsLongUppercaseWord(token, minLetters) Then CountAllUpperWords = CountAllUpperWords + 1
            token = ""
        End If
    Next i
End Function

Private Function NormalizeLongUppercaseWords(ByVal textValue As String, ByVal modeName As String, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean, ByVal headingLike As Boolean) As String
    If headingLike Then
        NormalizeLongUppercaseWords = textValue
        Exit Function
    End If

    Dim threshold As Long
    Select Case modeName
        Case "Conservative": threshold = 8
        Case "Balanced": threshold = 5
        Case Else: threshold = 4
    End Select

    Dim resultText As String
    Dim token As String
    Dim i As Long
    Dim currentChar As String
    For i = 1 To Len(textValue) + 1
        If i <= Len(textValue) Then
            currentChar = Mid$(textValue, i, 1)
        Else
            currentChar = " "
        End If

        If IsTokenWordChar(currentChar) Then
            token = token & currentChar
        Else
            If Len(token) > 0 Then
                resultText = resultText & NormalizeUppercaseToken(token, threshold, protectAcronyms, protectNames)
                token = ""
            End If
            If i <= Len(textValue) Then resultText = resultText & currentChar
        End If
    Next i
    NormalizeLongUppercaseWords = resultText
End Function

Private Function NormalizeUppercaseToken(ByVal tokenText As String, ByVal minLetters As Long, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean) As String
    If Not IsLongUppercaseWord(tokenText, minLetters) Then
        NormalizeUppercaseToken = tokenText
        Exit Function
    End If

    Dim canonical As String
    canonical = CanonicalProtectedToken(tokenText, protectAcronyms, protectNames)
    If Len(canonical) > 0 Then
        NormalizeUppercaseToken = canonical
    Else
        NormalizeUppercaseToken = ProperWordTokenCase(tokenText)
    End If
End Function

Private Function IsLongUppercaseWord(ByVal tokenText As String, ByVal minLetters As Long) As Boolean
    Dim lettersOnly As String
    lettersOnly = TokenLettersOnly(tokenText)
    If Len(lettersOnly) < minLetters Then Exit Function
    If lettersOnly <> UCase$(lettersOnly) Then Exit Function
    IsLongUppercaseWord = True
End Function

Private Function ApplySentenceRepairs(ByVal textValue As String, ByVal modeName As String, ByVal useAbbreviations As Boolean, ByVal useContextRules As Boolean, ByVal headingLike As Boolean) As String
    Dim resultText As String
    Dim capNext As Boolean
    capNext = True

    Dim currentWord As String
    Dim lastWord As String
    Dim i As Long
    Dim currentChar As String

    For i = 1 To Len(textValue)
        currentChar = Mid$(textValue, i, 1)

        If IsTokenWordChar(currentChar) Then
            currentWord = currentWord & currentChar
            If capNext And currentChar Like "[a-z]" Then
                currentChar = UCase$(currentChar)
                capNext = False
            ElseIf capNext And currentChar Like "[A-Z]" Then
                capNext = False
            End If
            resultText = resultText & currentChar
        Else
            If Len(currentWord) > 0 Then
                lastWord = currentWord
                currentWord = ""
            End If

            resultText = resultText & currentChar

            Select Case currentChar
                Case ".", "?", "!"
                    If currentChar = "." And useAbbreviations And IsCapitalizationAbbreviation(lastWord & ".") Then
                        capNext = False
                    ElseIf headingLike And modeName = "Conservative" Then
                        capNext = False
                    Else
                        capNext = True
                    End If
                Case ":", ";"
                    If modeName = "Aggressive" And Not headingLike Then capNext = True
                Case vbCr, vbLf
                    capNext = True
                Case Else
                    If capNext Then
                        If useContextRules Then
                            If Not IsContextSkippableWhileAwaitingSentence(currentChar) Then capNext = False
                        ElseIf currentChar <> " " And currentChar <> vbTab Then
                            capNext = False
                        End If
                    End If
            End Select
        End If
    Next i

    ApplySentenceRepairs = resultText
End Function

Private Function IsContextSkippableWhileAwaitingSentence(ByVal currentChar As String) As Boolean
    Select Case currentChar
        Case " ", vbTab, """", "'", "(", "[", "{", ChrW$(147), ChrW$(148), "-", ChrW$(8211), ChrW$(8212), "*", ChrW$(8226)
            IsContextSkippableWhileAwaitingSentence = True
    End Select
End Function

Private Function RestoreProtectedTerms(ByVal textValue As String, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean) As String
    Dim resultText As String
    Dim token As String
    Dim i As Long
    Dim currentChar As String

    For i = 1 To Len(textValue) + 1
        If i <= Len(textValue) Then
            currentChar = Mid$(textValue, i, 1)
        Else
            currentChar = " "
        End If

        If IsTokenWordChar(currentChar) Then
            token = token & currentChar
        Else
            If Len(token) > 0 Then
                resultText = resultText & RestoreProtectedToken(token, protectAcronyms, protectNames)
                token = ""
            End If
            If i <= Len(textValue) Then resultText = resultText & currentChar
        End If
    Next i

    RestoreProtectedTerms = resultText
End Function

Private Function RestoreProtectedToken(ByVal tokenText As String, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean) As String
    Dim canonical As String
    canonical = CanonicalProtectedToken(tokenText, protectAcronyms, protectNames)
    If Len(canonical) > 0 Then
        RestoreProtectedToken = canonical
    Else
        RestoreProtectedToken = tokenText
    End If
End Function

Private Function CanonicalProtectedToken(ByVal tokenText As String, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean) As String
    If InStr(tokenText, "-") > 0 Then
        CanonicalProtectedToken = CanonicalizeDelimitedToken(tokenText, "-", protectAcronyms, protectNames)
        If Len(CanonicalProtectedToken) > 0 Then Exit Function
    End If

    If protectNames Then
        Dim nameKey As String
        nameKey = LCase$(tokenText)
        If GetCapitalizationProtectedTermDict().Exists(nameKey) Then
            CanonicalProtectedToken = CStr(GetCapitalizationProtectedTermDict()(nameKey))
            Exit Function
        End If
        CanonicalProtectedToken = CanonicalizeNameByHeuristic(tokenText)
        If Len(CanonicalProtectedToken) > 0 Then Exit Function
    End If

    If protectAcronyms Then
        Dim acronymKey As String
        acronymKey = UCase$(TokenLettersOnly(tokenText))
        If GetCapitalizationAcronymDict().Exists(acronymKey) Then
            CanonicalProtectedToken = CStr(GetCapitalizationAcronymDict()(acronymKey))
        End If
    End If
End Function

Private Function CanonicalizeDelimitedToken(ByVal tokenText As String, ByVal delimiterText As String, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean) As String
    Dim parts As Variant
    parts = Split(tokenText, delimiterText)

    Dim i As Long
    Dim rebuilt As String
    For i = LBound(parts) To UBound(parts)
        Dim tokenPart As String
        tokenPart = CStr(parts(i))
        Dim canonicalPart As String
        canonicalPart = CanonicalProtectedToken(tokenPart, protectAcronyms, protectNames)
        If Len(canonicalPart) = 0 Then canonicalPart = tokenPart
        If i > LBound(parts) Then rebuilt = rebuilt & delimiterText
        rebuilt = rebuilt & canonicalPart
    Next i

    If rebuilt <> tokenText Then CanonicalizeDelimitedToken = rebuilt
End Function

Private Function CanonicalizeNameByHeuristic(ByVal tokenText As String) As String
    Dim lowerToken As String
    lowerToken = LCase$(tokenText)

    If Left$(lowerToken, 2) = "mc" And Len(lowerToken) > 2 Then
        CanonicalizeNameByHeuristic = "Mc" & UCase$(Mid$(lowerToken, 3, 1)) & Mid$(lowerToken, 4)
        Exit Function
    End If

    If Left$(lowerToken, 2) = "o'" And Len(lowerToken) > 2 Then
        CanonicalizeNameByHeuristic = "O'" & UCase$(Mid$(lowerToken, 3, 1)) & Mid$(lowerToken, 4)
    End If
End Function

Private Function ProperWordTokenCase(ByVal tokenText As String) As String
    Dim lowerToken As String
    lowerToken = LCase$(tokenText)

    Dim i As Long
    Dim currentChar As String
    Dim makeUpper As Boolean
    makeUpper = True

    For i = 1 To Len(lowerToken)
        currentChar = Mid$(lowerToken, i, 1)
        If currentChar Like "[A-Za-z]" Then
            If makeUpper Then
                ProperWordTokenCase = ProperWordTokenCase & UCase$(currentChar)
                makeUpper = False
            Else
                ProperWordTokenCase = ProperWordTokenCase & currentChar
            End If
        Else
            ProperWordTokenCase = ProperWordTokenCase & currentChar
            makeUpper = (currentChar = "-")
        End If
    Next i
End Function

Private Function TokenLettersOnly(ByVal tokenText As String) As String
    Dim i As Long
    Dim currentChar As String
    For i = 1 To Len(tokenText)
        currentChar = Mid$(tokenText, i, 1)
        If currentChar Like "[A-Za-z0-9]" Then TokenLettersOnly = TokenLettersOnly & currentChar
    Next i
End Function

Private Function IsTokenWordChar(ByVal currentChar As String) As Boolean
    If currentChar Like "[A-Za-z0-9]" Then
        IsTokenWordChar = True
    ElseIf currentChar = "'" Or currentChar = "-" Then
        IsTokenWordChar = True
    End If
End Function

Private Function IsLikelyHeadingText(ByVal textValue As String) As Boolean
    Dim trimmedText As String
    trimmedText = Trim$(textValue)
    If Len(trimmedText) = 0 Then Exit Function
    If Len(trimmedText) > 90 Then Exit Function
    If Right$(trimmedText, 1) = "." Or Right$(trimmedText, 1) = "?" Or Right$(trimmedText, 1) = "!" Then Exit Function
    If CountWords(trimmedText) > 10 Then Exit Function

    Dim lowerCount As Long
    Dim upperCount As Long
    Dim letterCount As Long
    CountLetterCase trimmedText, lowerCount, upperCount, letterCount
    If letterCount = 0 Then Exit Function

    If upperCount / letterCount >= 0.65 Then
        IsLikelyHeadingText = True
        Exit Function
    End If

    Dim titleWords As Long
    Dim words As Variant
    words = Split(Replace(trimmedText, vbTab, " "), " ")
    Dim i As Long
    For i = LBound(words) To UBound(words)
        Dim currentWord As String
        currentWord = CStr(words(i))
        If Len(TokenLettersOnly(currentWord)) > 0 Then
            If Left$(currentWord, 1) = UCase$(Left$(currentWord, 1)) And Mid$(currentWord, 2) = LCase$(Mid$(currentWord, 2)) Then
                titleWords = titleWords + 1
            End If
        End If
    Next i

    If titleWords >= 2 And titleWords >= CountWords(trimmedText) - 1 Then IsLikelyHeadingText = True
End Function

Private Function IsCapitalizationAbbreviation(ByVal tokenText As String) As Boolean
    Dim normalized As String
    normalized = LCase$(Trim$(tokenText))
    If Len(normalized) = 2 And Right$(normalized, 1) = "." Then
        IsCapitalizationAbbreviation = True
        Exit Function
    End If
    IsCapitalizationAbbreviation = GetCapitalizationAbbreviationDict().Exists(normalized)
End Function

Private Function GetCapitalizationAbbreviationDict() As Object
    Static dict As Object
    If dict Is Nothing Then
        Set dict = CreateObject("Scripting.Dictionary")
        AddLookupValues dict, Array("dr.", "mr.", "mrs.", "ms.", "prof.", "sr.", "jr.", "st.", "mt.", "vs.", "etc.", "inc.", "ltd.", "co.", "corp.", "dept.", "fig.", "no.", "nos.", "est.", "misc.", "approx.", "jan.", "feb.", "mar.", "apr.", "jun.", "jul.", "aug.", "sep.", "sept.", "oct.", "nov.", "dec.")
    End If
    Set GetCapitalizationAbbreviationDict = dict
End Function

Private Function GetCapitalizationAcronymDict() As Object
    Static dict As Object
    If dict Is Nothing Then
        Set dict = CreateObject("Scripting.Dictionary")
        AddLookupPairs dict, Array("API|API", "CEO|CEO", "CFO|CFO", "CIA|CIA", "COO|COO", "FAQ|FAQ", "FBI|FBI", "HTML|HTML", "OCR|OCR", "PDF|PDF", "PR|PR", "SQL|SQL", "URL|URL", "USA|USA", "VBA|VBA", "NASA|NASA")
    End If
    Set GetCapitalizationAcronymDict = dict
End Function

Private Function GetCapitalizationProtectedTermDict() As Object
    Static dict As Object
    If dict Is Nothing Then
        Set dict = CreateObject("Scripting.Dictionary")
        AddLookupPairs dict, Array("ebay|eBay", "github|GitHub", "iphone|iPhone", "ipad|iPad", "linkedin|LinkedIn", "macbook|MacBook", "mcdonald|McDonald", "oreilly|O'Reilly", "powerpoint|PowerPoint", "youtube|YouTube")
    End If
    Set GetCapitalizationProtectedTermDict = dict
End Function

Private Sub AddLookupValues(ByVal dict As Object, ByVal values As Variant)
    Dim i As Long
    For i = LBound(values) To UBound(values)
        dict(CStr(values(i))) = True
    Next i
End Sub

Private Sub AddLookupPairs(ByVal dict As Object, ByVal values As Variant)
    Dim i As Long
    For i = LBound(values) To UBound(values)
        Dim parts As Variant
        parts = Split(CStr(values(i)), "|")
        dict(CStr(parts(0))) = CStr(parts(1))
    Next i
End Sub
