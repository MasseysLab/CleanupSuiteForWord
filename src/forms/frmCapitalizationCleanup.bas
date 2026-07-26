Option Explicit

Private Const CAPITALIZATION_EXCEPTIONS_PROPERTY As String = "CleanupSuiteCapitalizationExceptions"
Private Const CAPITALIZATION_PROPERTY_TYPE_STRING As Long = 4

Private mCapitalizationCacheKey As String
Private mCapitalizationAbbreviationDict As Object
Private mCapitalizationAcronymDict As Object
Private mCapitalizationProtectedTermDict As Object

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub

Private Sub UserForm_Initialize()
    ApplyMSFormTitleStrategy Me, True

    optAll.GroupName = "CapMode"
    optCustom.GroupName = "CapMode"
    optScopeDocument.GroupName = "CapScope"
    optScopeSelection.GroupName = "CapScope"

    optAll.Caption = "Recommended repair"
    optCustom.Caption = "Custom repair"

    optAll.Visible = True
    optCustom.Visible = True

    optAll.Value = True

    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    RefreshScopeSelectionState Me, True

    lblIntro.Caption = "Use the recommended repair, or customize its protections."
    chkSentence.Caption = "Recognize common abbreviations"
    chkTitle.Caption = "Protect common acronyms"
    chkUpper.Caption = "Protect names and brands"
    chkLower.Caption = "Repair likely headings in title case"
    chkHeadingParentheses.Caption = "Also title-case words inside parentheses"
    chkHeadingParentheses.Value = False
    chkSmartSentences.Caption = "Recognize quotes, bullets, and sentence boundaries"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    cmdEditExceptions.Caption = "Edit custom exceptions"

    ApplyModeDefaults
    UpdateAdvancedVisibility
    UpdateModeSummary
    LayoutCapitalizationForm
End Sub
Private Sub LayoutCapitalizationForm()
    LayoutCleanupToolForm Me
End Sub
Private Sub cmdRiskPlacementCapitalizationRecommended_Click(): ShowToolRiskChoiceExplanation "Recommended repair", "Text caution", "Repairs sentence starts, obvious all-caps damage, headings, acronyms, names, and brands using every protection.": End Sub
Private Sub cmdRiskPlacementCapitalizationCustom_Click(): ShowToolRiskChoiceExplanation "Custom capitalization", "Inspect first", "Lets you choose which protections guide visible word repair. Preview names, acronyms, and headings before applying.": End Sub

Private Sub UpdateAdvancedVisibility()
    chkSentence.Visible = optCustom.Value
    chkTitle.Visible = optCustom.Value
    chkUpper.Visible = optCustom.Value
    chkLower.Visible = optCustom.Value
    chkHeadingParentheses.Visible = optCustom.Value And chkLower.Value
    chkHeadingParentheses.Enabled = optCustom.Value And chkLower.Value
    chkSmartSentences.Visible = optCustom.Value
    cmdSelectAll.Visible = optCustom.Value
    cmdDeselectAll.Visible = optCustom.Value
    cmdEditExceptions.Visible = optCustom.Value
End Sub

Private Sub ApplyModeDefaults()
    If optAll.Value Or Not (chkSentence.Value Or chkTitle.Value Or chkUpper.Value Or chkLower.Value Or chkSmartSentences.Value) Then
        chkSentence.Value = True
        chkTitle.Value = True
        chkUpper.Value = True
        chkLower.Value = True
        chkSmartSentences.Value = True
    End If
    If optAll.Value Then chkHeadingParentheses.Value = False
End Sub

Private Sub UpdateModeSummary()
    If optAll.Value Then
        lblModeSummary.Caption = "Repairs sentence starts, obvious all-caps damage, headings, acronyms, names, and brands."
    Else
        lblModeSummary.Caption = "Custom repair uses only the protections selected below."
    End If
End Sub

Private Sub RefreshCapitalizationChoices()
    ApplyModeDefaults
    UpdateAdvancedVisibility
    UpdateModeSummary
    LayoutCapitalizationForm
End Sub

Private Sub optAll_Click(): RefreshCapitalizationChoices: End Sub
Private Sub optCustom_Click(): RefreshCapitalizationChoices: End Sub

Private Sub cmdSelectAll_Click()
    optCustom.Value = True
    UpdateAdvancedVisibility
    UpdateModeSummary
    chkSentence.Value = True
    chkTitle.Value = True
    chkUpper.Value = True
    chkLower.Value = True
    chkHeadingParentheses.Value = True
    chkSmartSentences.Value = True
    UpdateAdvancedVisibility
    LayoutCapitalizationForm
End Sub

Private Sub cmdDeselectAll_Click()
    optCustom.Value = True
    UpdateAdvancedVisibility
    UpdateModeSummary
    chkSentence.Value = False
    chkTitle.Value = False
    chkUpper.Value = False
    chkLower.Value = False
    chkHeadingParentheses.Value = False
    chkSmartSentences.Value = False
    UpdateAdvancedVisibility
    LayoutCapitalizationForm
End Sub

Private Sub chkSentence_Click(): LayoutCapitalizationForm: End Sub
Private Sub chkTitle_Click(): LayoutCapitalizationForm: End Sub
Private Sub chkUpper_Click(): LayoutCapitalizationForm: End Sub
Private Sub chkLower_Click()
    If Not chkLower.Value Then chkHeadingParentheses.Value = False
    UpdateAdvancedVisibility
    LayoutCapitalizationForm
End Sub
Private Sub chkHeadingParentheses_Click(): LayoutCapitalizationForm: End Sub
Private Sub chkSmartSentences_Click(): LayoutCapitalizationForm: End Sub

Private Sub cmdEditExceptions_Click()
    PromptForCapitalizationCustomExceptions
End Sub

Private Sub cmdPreview_Click()
    PreviewFromPanel
End Sub

Public Sub PreviewFromPanel()
    chkPreviewOnly.Value = True
    BeginPreviewActionIndicator Me
    cmdRun_Click
    EndPreviewActionIndicator
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

Public Function SmokeSmartRepairText(ByVal textValue As String) As String
    PrepareCapitalizationLookups
    SmokeSmartRepairText = SmartRepairParagraph(textValue, True, True, True, True, False, True, False)
End Function

Public Function SmokeSmartRepairHeadingText(ByVal textValue As String) As String
    PrepareCapitalizationLookups
    SmokeSmartRepairHeadingText = SmartRepairParagraph(textValue, True, True, True, True, False, True, True)
End Function

Public Function SmokeSmartRepairHeadingParenthesesText(ByVal textValue As String) As String
    PrepareCapitalizationLookups
    SmokeSmartRepairHeadingParenthesesText = SmartRepairParagraph(textValue, True, True, True, True, True, True, True)
End Function

Public Function SmokeSmartRepairWithoutHeadingsText(ByVal textValue As String, Optional ByVal styleHeadingLike As Boolean = False) As String
    PrepareCapitalizationLookups
    SmokeSmartRepairWithoutHeadingsText = SmartRepairParagraph(textValue, True, True, True, False, False, True, styleHeadingLike)
End Function

Public Function SmokeFriendlyCapitalizationExceptions(ByVal exceptionText As String) As String
    SmokeFriendlyCapitalizationExceptions = NormalizeCapitalizationExceptionList(exceptionText)
End Function

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

    Dim useAbbreviations As Boolean
    Dim protectAcronyms As Boolean
    Dim protectNames As Boolean
    Dim useHeadingHeuristics As Boolean
    Dim titleCaseParentheticalText As Boolean
    Dim useContextRules As Boolean
    useAbbreviations = chkSentence.Value
    protectAcronyms = chkTitle.Value
    protectNames = chkUpper.Value
    useHeadingHeuristics = chkLower.Value
    titleCaseParentheticalText = useHeadingHeuristics And chkHeadingParentheses.Value
    useContextRules = chkSmartSentences.Value
    If optAll.Value Then
        useAbbreviations = True
        protectAcronyms = True
        protectNames = True
        useHeadingHeuristics = True
        titleCaseParentheticalText = False
        useContextRules = True
    End If
    If optCustom.Value Then
        If Not (useAbbreviations Or protectAcronyms Or protectNames Or useHeadingHeuristics Or useContextRules) Then
            MsgBox "No custom repair protections are selected.", vbInformation, "Capitalization Fixer"
            Exit Sub
        End If
    End If

    ' Prepare document-specific lookups once before entering the word-processing hot path.
    PrepareCapitalizationLookups

    Dim p As Paragraph
    Dim original As String
    Dim transformed As String
    Dim styleHeadingLike As Boolean
    Dim changedCount As Long
    changedCount = 0

    If previewOnly Then
        On Error GoTo PreviewErr
        BeginCleanupProgress "Capitalization Fixer", "Analyzing preview"
        Dim previewProcessed As Long
        Dim previewTotal As Long
        previewTotal = targetRange.Paragraphs.Count

        For Each p In targetRange.Paragraphs
            If ParagraphContainedInRange(p, targetRange) Then
                original = ParagraphBodyText(p)
                styleHeadingLike = ParagraphUsesHeadingStyle(p)
                transformed = SmartRepairParagraph(original, useAbbreviations, protectAcronyms, protectNames, useHeadingHeuristics, titleCaseParentheticalText, useContextRules, styleHeadingLike)
                If transformed <> original Then
                    HighlightCapitalizationDifferenceRuns p, original, transformed
                    changedCount = changedCount + 1
                End If
            End If
            previewProcessed = previewProcessed + 1
            If previewProcessed Mod 50 = 0 Or previewProcessed = previewTotal Then
                UpdateCleanupProgress "Analyzing preview", previewProcessed, previewTotal
            End If
        Next p

        EndCleanupProgress
        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Paragraphs", changedCount
        ShowPreviewActionsSummary Me, "Capitalization Fixer", previewRows
        Exit Sub
    End If

    MarkCleanupStart "Capitalization Fixer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Capitalization Fixer"
    On Error GoTo RunErr

    Dim applyProcessed As Long
    Dim applyTotal As Long
    applyTotal = targetRange.Paragraphs.Count
    For Each p In targetRange.Paragraphs
        If ParagraphContainedInRange(p, targetRange) Then
            original = ParagraphBodyText(p)
            styleHeadingLike = ParagraphUsesHeadingStyle(p)
            transformed = SmartRepairParagraph(original, useAbbreviations, protectAcronyms, protectNames, useHeadingHeuristics, titleCaseParentheticalText, useContextRules, styleHeadingLike)
            If transformed <> original Then ReplaceParagraphBodyText p, transformed
        End If
        applyProcessed = applyProcessed + 1
        If applyProcessed Mod 50 = 0 Or applyProcessed = applyTotal Then
            UpdateCleanupProgress "Applying changes", applyProcessed, applyTotal
        End If
    Next p

    RemoveAllHighlighting targetRange

    undoRec.EndCustomRecord
    MarkCleanupEnd
    MarkCleanupToolApplied
    Unload Me
    Exit Sub

PreviewErr:
    Dim previewErrNumber As Long
    Dim previewErrDescription As String
    previewErrNumber = Err.Number
    previewErrDescription = Err.Description
    RestoreCleanupSuiteTransientState
    MsgBox "Preview could not be completed: " & previewErrNumber & " - " & previewErrDescription, _
           vbCritical, "Capitalization Preview Error"
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

Private Sub HighlightCapitalizationDifferenceRuns(ByVal p As Paragraph, ByVal originalText As String, ByVal transformedText As String)
    Dim i As Long
    Dim compareLength As Long
    Dim runStart As Long
    Dim paragraphStart As Long
    Dim changedRange As Range
    compareLength = Len(originalText)
    If Len(transformedText) < compareLength Then compareLength = Len(transformedText)

    paragraphStart = p.Range.Start
    i = 1
    Do While i <= compareLength
        If Mid$(originalText, i, 1) = Mid$(transformedText, i, 1) Then
            i = i + 1
        Else
            runStart = i
            Do While i <= compareLength
                If Mid$(originalText, i, 1) = Mid$(transformedText, i, 1) Then Exit Do
                i = i + 1
            Loop
            Set changedRange = ActiveDocument.Range(paragraphStart + runStart - 1, paragraphStart + i - 1)
            ApplyPreviewHighlight changedRange
        End If
    Loop

    If Len(originalText) <> Len(transformedText) Then ApplyPreviewHighlight p.Range
End Sub

Private Function SmartRepairParagraph(ByVal textValue As String, ByVal useAbbreviations As Boolean, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean, ByVal useHeadingHeuristics As Boolean, ByVal titleCaseParentheticalText As Boolean, ByVal useContextRules As Boolean, Optional ByVal styleHeadingLike As Boolean = False) As String
    If Len(Trim$(textValue)) = 0 Then
        SmartRepairParagraph = textValue
        Exit Function
    End If

    Dim headingLike As Boolean
    headingLike = styleHeadingLike Or IsLikelyHeadingText(textValue)

    If headingLike And Not useHeadingHeuristics Then
        SmartRepairParagraph = textValue
        Exit Function
    End If

    Dim working As String
    working = textValue

    If headingLike Then
        working = ApplyHeadingCase(working, titleCaseParentheticalText)
    Else
        working = NormalizeLongUppercaseWords(working, protectAcronyms, protectNames)
        working = ApplySentenceRepairs(working, useAbbreviations, useContextRules)
    End If
    working = RestoreProtectedTerms(working, useAbbreviations, protectAcronyms, protectNames)
    If headingLike And Not titleCaseParentheticalText Then
        working = RestoreOriginalParentheticalText(textValue, working)
    End If

    SmartRepairParagraph = working
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

Private Function CountHeadingWords(ByVal textValue As String) As Long
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
        Else
            If Len(token) > 0 Then
                If TokenContainsLetter(token) Then CountHeadingWords = CountHeadingWords + 1
                token = ""
            End If
        End If
    Next i
End Function

Private Function TokenContainsLetter(ByVal tokenText As String) As Boolean
    Dim i As Long
    For i = 1 To Len(tokenText)
        If Mid$(tokenText, i, 1) Like "[A-Za-z]" Then
            TokenContainsLetter = True
            Exit Function
        End If
    Next i
End Function

Private Function NormalizeLongUppercaseWords(ByVal textValue As String, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean) As String
    Const LONG_UPPERCASE_THRESHOLD As Long = 8
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
                resultText = resultText & NormalizeUppercaseToken(token, LONG_UPPERCASE_THRESHOLD, protectAcronyms, protectNames)
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
    canonical = CanonicalProtectedToken(tokenText, False, protectAcronyms, protectNames)
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

Private Function ApplySentenceRepairs(ByVal textValue As String, ByVal useAbbreviations As Boolean, ByVal useContextRules As Boolean) As String
    Dim resultText As String
    Dim capNext As Boolean
    Dim lowerDotSegment As Boolean
    capNext = True

    Dim currentWord As String
    Dim lastWord As String
    Dim i As Long
    Dim currentChar As String

    For i = 1 To Len(textValue)
        currentChar = Mid$(textValue, i, 1)

        If useContextRules Then
            If IsOpeningQuoteAfterComma(textValue, i) Then capNext = True
        End If

        If IsTokenWordChar(currentChar) Then
            If lowerDotSegment And currentChar Like "[A-Z]" Then currentChar = LCase$(currentChar)
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
                    If currentChar = "." And PeriodStartsLowercaseDotSegment(textValue, i) Then
                        capNext = False
                        lowerDotSegment = True
                    ElseIf currentChar = "." And useAbbreviations And IsCapitalizationAbbreviation(lastWord & ".") Then
                        capNext = False
                        lowerDotSegment = False
                    Else
                        capNext = True
                        lowerDotSegment = False
                    End If
                Case vbCr, vbLf
                    capNext = True
                    lowerDotSegment = False
                Case Else
                    lowerDotSegment = False
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

Private Function IsOpeningQuoteAfterComma(ByVal textValue As String, ByVal quoteIndex As Long) As Boolean
    Dim quoteChar As String
    Dim i As Long
    quoteChar = Mid$(textValue, quoteIndex, 1)

    Select Case quoteChar
        Case """", "'", ChrW$(&H2018), ChrW$(&H201C)
        Case Else
            Exit Function
    End Select

    For i = quoteIndex - 1 To 1 Step -1
        Select Case Mid$(textValue, i, 1)
            Case " ", vbTab
            Case ","
                IsOpeningQuoteAfterComma = True
                Exit Function
            Case Else
                Exit Function
        End Select
    Next i
End Function

Private Function PeriodStartsLowercaseDotSegment(ByVal textValue As String, ByVal periodIndex As Long) As Boolean
    Dim previousChar As String
    Dim nextChar As String
    Dim segmentLength As Long
    Dim i As Long

    If periodIndex <= 1 Or periodIndex >= Len(textValue) Then Exit Function
    previousChar = Mid$(textValue, periodIndex - 1, 1)
    nextChar = Mid$(textValue, periodIndex + 1, 1)
    If Not previousChar Like "[A-Za-z0-9]" Then Exit Function
    If Not nextChar Like "[A-Za-z0-9]" Then Exit Function
    If PeriodFollowsInitialChain(textValue, periodIndex) Then Exit Function

    For i = periodIndex + 1 To Len(textValue)
        nextChar = Mid$(textValue, i, 1)
        If nextChar Like "[A-Za-z0-9-]" Then
            segmentLength = segmentLength + 1
        Else
            Exit For
        End If
    Next i

    PeriodStartsLowercaseDotSegment = (segmentLength >= 2)
End Function

Private Function PeriodFollowsInitialChain(ByVal textValue As String, ByVal periodIndex As Long) As Boolean
    If periodIndex < 4 Then Exit Function
    If Mid$(textValue, periodIndex - 2, 1) <> "." Then Exit Function
    PeriodFollowsInitialChain = (Mid$(textValue, periodIndex - 3, 1) Like "[A-Za-z]")
End Function

Private Function IsContextSkippableWhileAwaitingSentence(ByVal currentChar As String) As Boolean
    Select Case currentChar
        Case " ", vbTab, """", "'", "(", "[", "{", _
             ChrW$(&H2018), ChrW$(&H2019), ChrW$(&H201C), ChrW$(&H201D), _
             "-", ChrW$(8211), ChrW$(8212), "*", ChrW$(8226)
            IsContextSkippableWhileAwaitingSentence = True
    End Select
End Function

Private Function ApplyHeadingCase(ByVal textValue As String, ByVal titleCaseParentheticalText As Boolean) As String
    Dim resultText As String
    resultText = ApplyOuterHeadingCase(textValue)
    If titleCaseParentheticalText Then
        resultText = ApplyBalancedParentheticalHeadingCase(resultText)
    End If
    ApplyHeadingCase = resultText
End Function

Private Function ApplyOuterHeadingCase(ByVal textValue As String) As String
    Dim resultText As String
    Dim token As String
    Dim i As Long
    Dim wordIndex As Long
    Dim currentChar As String
    Dim startsLowerDotSegment As Boolean
    Dim lowerDotSegment As Boolean
    Dim parenthesisDepth As Long
    Dim startsHeadingSegment As Boolean
    startsHeadingSegment = True
    For i = 1 To Len(textValue) + 1
        If i <= Len(textValue) Then currentChar = Mid$(textValue, i, 1) Else currentChar = " "
        If i <= Len(textValue) And currentChar = "(" Then
            If Len(token) > 0 Then
                If TokenContainsLetter(token) Then
                    wordIndex = wordIndex + 1
                    If IsAlwaysLowercaseHeadingWord(token) Then
                        token = LCase$(token)
                    ElseIf startsHeadingSegment Or wordIndex = 1 Or IsFinalHeadingWord(textValue, i) Or Not IsHeadingMinorWord(token) Then
                        token = ApplyHeadingMajorWordCase(token)
                    Else
                        token = LCase$(token)
                    End If
                    startsHeadingSegment = False
                End If
                resultText = resultText & token
                token = ""
            End If
            parenthesisDepth = parenthesisDepth + 1
            resultText = resultText & currentChar
        ElseIf i <= Len(textValue) And currentChar = ")" And parenthesisDepth > 0 Then
            parenthesisDepth = parenthesisDepth - 1
            resultText = resultText & currentChar
        ElseIf parenthesisDepth > 0 Then
            If i <= Len(textValue) Then resultText = resultText & currentChar
        ElseIf IsTokenWordChar(currentChar) Then
            token = token & currentChar
        Else
            startsLowerDotSegment = False
            If i <= Len(textValue) Then
                If currentChar = "." Then startsLowerDotSegment = PeriodStartsLowercaseDotSegment(textValue, i)
            End If
            If Len(token) > 0 Then
                If TokenContainsLetter(token) Then
                    If lowerDotSegment Then
                        token = LCase$(token)
                    Else
                        wordIndex = wordIndex + 1
                        If Not startsLowerDotSegment Then
                            If IsAlwaysLowercaseHeadingWord(token) Then
                                token = LCase$(token)
                            ElseIf startsHeadingSegment Or wordIndex = 1 Or IsFinalHeadingWord(textValue, i) Or Not IsHeadingMinorWord(token) Then
                                token = ApplyHeadingMajorWordCase(token)
                            Else
                                token = LCase$(token)
                            End If
                        End If
                        startsHeadingSegment = False
                    End If
                End If
                resultText = resultText & token
                token = ""
            End If
            If i <= Len(textValue) Then resultText = resultText & currentChar
            If currentChar = ":" Then startsHeadingSegment = True
            lowerDotSegment = startsLowerDotSegment
        End If
    Next i
    ApplyOuterHeadingCase = resultText
End Function

Private Function IsFinalHeadingWord(ByVal textValue As String, ByVal afterIndex As Long) As Boolean
    Dim i As Long
    For i = afterIndex + 1 To Len(textValue)
        If Mid$(textValue, i, 1) Like "[A-Za-z0-9]" Then Exit Function
    Next i
    IsFinalHeadingWord = True
End Function

Private Function ApplyBalancedParentheticalHeadingCase(ByVal textValue As String) As String
    Dim resultText As String
    Dim openIndex As Long
    Dim closeIndex As Long
    Dim innerText As String
    Dim repairedInnerText As String
    resultText = textValue
    openIndex = 1

    Do While openIndex <= Len(resultText)
        openIndex = InStr(openIndex, resultText, "(", vbBinaryCompare)
        If openIndex = 0 Then Exit Do
        closeIndex = MatchingClosingParenthesis(resultText, openIndex)
        If closeIndex = 0 Then Exit Do
        innerText = Mid$(resultText, openIndex + 1, closeIndex - openIndex - 1)
        If Len(innerText) > 0 Then
            repairedInnerText = ApplyHeadingCase(innerText, True)
            Mid$(resultText, openIndex + 1, Len(innerText)) = repairedInnerText
        End If
        openIndex = closeIndex + 1
    Loop

    ApplyBalancedParentheticalHeadingCase = resultText
End Function

Private Function MatchingClosingParenthesis(ByVal textValue As String, ByVal openIndex As Long) As Long
    Dim depth As Long
    Dim i As Long
    Dim currentChar As String
    For i = openIndex To Len(textValue)
        currentChar = Mid$(textValue, i, 1)
        If currentChar = "(" Then
            depth = depth + 1
        ElseIf currentChar = ")" Then
            depth = depth - 1
            If depth = 0 Then
                MatchingClosingParenthesis = i
                Exit Function
            End If
        End If
    Next i
End Function

Private Function RestoreOriginalParentheticalText(ByVal originalText As String, ByVal repairedText As String) As String
    Dim resultText As String
    Dim parenthesisDepth As Long
    Dim i As Long
    Dim currentChar As String
    resultText = repairedText
    If Len(resultText) <> Len(originalText) Then
        RestoreOriginalParentheticalText = repairedText
        Exit Function
    End If

    For i = 1 To Len(originalText)
        currentChar = Mid$(originalText, i, 1)
        If currentChar = "(" Then
            parenthesisDepth = parenthesisDepth + 1
        ElseIf currentChar = ")" And parenthesisDepth > 0 Then
            parenthesisDepth = parenthesisDepth - 1
        ElseIf parenthesisDepth > 0 Then
            Mid$(resultText, i, 1) = currentChar
        End If
    Next i

    RestoreOriginalParentheticalText = resultText
End Function

Private Function ApplyHeadingMajorWordCase(ByVal tokenText As String) As String
    Dim resultText As String
    Dim i As Long
    Dim currentChar As String

    resultText = LCase$(tokenText)
    For i = 1 To Len(resultText)
        currentChar = Mid$(resultText, i, 1)
        If currentChar Like "[A-Za-z]" Then
            Mid$(resultText, i, 1) = UCase$(currentChar)
            Exit For
        End If
    Next i

    ApplyHeadingMajorWordCase = resultText
End Function

Private Function IsHeadingMinorWord(ByVal tokenText As String) As Boolean
    Select Case LCase$(tokenText)
        Case "a", "an", "and", "as", "at", "but", "by", "for", "from", "in", "nor", "of", "on", "or", "the", "to", "up", "yet"
            IsHeadingMinorWord = True
    End Select
End Function

Private Function IsAlwaysLowercaseHeadingWord(ByVal tokenText As String) As Boolean
    IsAlwaysLowercaseHeadingWord = (LCase$(tokenText) = "etc")
End Function
Private Function RestoreProtectedTerms(ByVal textValue As String, ByVal useAbbreviations As Boolean, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean) As String
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
                resultText = resultText & RestoreProtectedToken(token, useAbbreviations, protectAcronyms, protectNames)
                token = ""
            End If
            If i <= Len(textValue) Then resultText = resultText & currentChar
        End If
    Next i

    RestoreProtectedTerms = resultText
End Function

Private Function RestoreProtectedToken(ByVal tokenText As String, ByVal useAbbreviations As Boolean, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean) As String
    Dim canonical As String
    canonical = CanonicalProtectedToken(tokenText, useAbbreviations, protectAcronyms, protectNames)
    If Len(canonical) > 0 Then
        RestoreProtectedToken = canonical
    Else
        RestoreProtectedToken = tokenText
    End If
End Function

Private Function CanonicalProtectedToken(ByVal tokenText As String, ByVal useAbbreviations As Boolean, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean) As String
    If InStr(tokenText, "-") > 0 Then
        CanonicalProtectedToken = CanonicalizeDelimitedToken(tokenText, "-", useAbbreviations, protectAcronyms, protectNames)
        If Len(CanonicalProtectedToken) > 0 Then Exit Function
    End If

    If useAbbreviations Then
        Select Case LCase$(tokenText)
            Case "dr": CanonicalProtectedToken = "Dr"
            Case "mr": CanonicalProtectedToken = "Mr"
            Case "mrs": CanonicalProtectedToken = "Mrs"
            Case "ms": CanonicalProtectedToken = "Ms"
            Case "prof": CanonicalProtectedToken = "Prof"
            Case "rev": CanonicalProtectedToken = "Rev"
            Case "hon": CanonicalProtectedToken = "Hon"
            Case "jr": CanonicalProtectedToken = "Jr"
            Case "sr": CanonicalProtectedToken = "Sr"
        End Select
        If Len(CanonicalProtectedToken) > 0 Then Exit Function
    End If

    If protectNames Then
        Dim nameKey As String
        nameKey = LCase$(tokenText)
        If mCapitalizationProtectedTermDict.Exists(nameKey) Then
            CanonicalProtectedToken = CStr(mCapitalizationProtectedTermDict(nameKey))
            Exit Function
        End If
        CanonicalProtectedToken = CanonicalizeNameByHeuristic(tokenText)
        If Len(CanonicalProtectedToken) > 0 Then Exit Function
    End If

    If protectAcronyms Then
        Dim acronymKey As String
        acronymKey = UCase$(TokenLettersOnly(tokenText))
        If mCapitalizationAcronymDict.Exists(acronymKey) Then
            CanonicalProtectedToken = CStr(mCapitalizationAcronymDict(acronymKey))
        End If
    End If
End Function

Private Function CanonicalizeDelimitedToken(ByVal tokenText As String, ByVal delimiterText As String, ByVal useAbbreviations As Boolean, ByVal protectAcronyms As Boolean, ByVal protectNames As Boolean) As String
    Dim parts As Variant
    parts = Split(tokenText, delimiterText)

    Dim i As Long
    Dim rebuilt As String
    For i = LBound(parts) To UBound(parts)
        Dim tokenPart As String
        tokenPart = CStr(parts(i))
        Dim canonicalPart As String
        canonicalPart = CanonicalProtectedToken(tokenPart, useAbbreviations, protectAcronyms, protectNames)
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
    ElseIf currentChar = "'" Or currentChar = ChrW$(8217) Or currentChar = "-" Then
        IsTokenWordChar = True
    End If
End Function

Private Function IsLikelyHeadingText(ByVal textValue As String) As Boolean
    Dim trimmedText As String
    Dim headingCandidate As String
    Dim headingWordCount As Long
    trimmedText = Trim$(textValue)
    If Len(trimmedText) = 0 Then Exit Function
    headingCandidate = Trim$(TextOutsideParentheses(trimmedText))
    If Len(headingCandidate) = 0 Then Exit Function
    If Len(headingCandidate) > 90 Then Exit Function
    If Right$(headingCandidate, 1) = "." Or Right$(headingCandidate, 1) = "?" Or Right$(headingCandidate, 1) = "!" Then Exit Function
    headingWordCount = CountHeadingWords(headingCandidate)
    If headingWordCount = 0 Then Exit Function
    If headingWordCount > 10 Then Exit Function

    Dim lowerCount As Long
    Dim upperCount As Long
    Dim letterCount As Long
    CountLetterCase headingCandidate, lowerCount, upperCount, letterCount
    If letterCount = 0 Then Exit Function

    If upperCount / letterCount >= 0.65 Then
        IsLikelyHeadingText = True
        Exit Function
    End If

    Dim titleWords As Long
    Dim uppercaseWords As Long
    CountHeadingCaseWords headingCandidate, titleWords, uppercaseWords

    If titleWords >= 2 And titleWords >= headingWordCount - 1 Then IsLikelyHeadingText = True
    If uppercaseWords >= 2 And titleWords >= 2 And titleWords + uppercaseWords >= headingWordCount - 1 Then IsLikelyHeadingText = True
End Function

Private Sub CountHeadingCaseWords(ByVal textValue As String, ByRef titleWords As Long, ByRef uppercaseWords As Long)
    Dim i As Long
    Dim token As String
    Dim lettersOnly As String
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
            lettersOnly = TokenLettersOnly(token)
            If Len(lettersOnly) > 0 Then
                If lettersOnly = UCase$(lettersOnly) And lettersOnly <> LCase$(lettersOnly) Then
                    uppercaseWords = uppercaseWords + 1
                ElseIf Left$(lettersOnly, 1) = UCase$(Left$(lettersOnly, 1)) And Mid$(lettersOnly, 2) = LCase$(Mid$(lettersOnly, 2)) Then
                    titleWords = titleWords + 1
                End If
            End If
            token = ""
        End If
    Next i
End Sub

Private Function TextOutsideParentheses(ByVal textValue As String) As String
    Dim resultText As String
    Dim parenthesisDepth As Long
    Dim i As Long
    Dim currentChar As String

    For i = 1 To Len(textValue)
        currentChar = Mid$(textValue, i, 1)
        If currentChar = "(" Then
            parenthesisDepth = parenthesisDepth + 1
            resultText = resultText & " "
        ElseIf currentChar = ")" And parenthesisDepth > 0 Then
            parenthesisDepth = parenthesisDepth - 1
            resultText = resultText & " "
        ElseIf parenthesisDepth > 0 Then
            resultText = resultText & " "
        Else
            resultText = resultText & currentChar
        End If
    Next i

    TextOutsideParentheses = resultText
End Function

Private Function ParagraphUsesHeadingStyle(ByVal p As Paragraph) As Boolean
    Dim styleName As String

    On Error Resume Next
    styleName = LCase$(CStr(p.Style.NameLocal))
    On Error GoTo 0

    If Left$(styleName, 7) = "heading" Then
        ParagraphUsesHeadingStyle = True
        Exit Function
    End If

    ParagraphUsesHeadingStyle = (InStr(1, styleName, " title", vbTextCompare) > 0 Or styleName = "title")
End Function

Private Function IsCapitalizationAbbreviation(ByVal tokenText As String) As Boolean
    Dim normalized As String
    normalized = LCase$(Trim$(tokenText))
    If Len(normalized) = 2 And Right$(normalized, 1) = "." Then
        IsCapitalizationAbbreviation = True
        Exit Function
    End If
    IsCapitalizationAbbreviation = mCapitalizationAbbreviationDict.Exists(normalized)
End Function

Private Sub PrepareCapitalizationLookups()
    Dim cacheKey As String
    cacheKey = CapitalizationCacheKey()

    If cacheKey = mCapitalizationCacheKey Then
        If Not mCapitalizationAbbreviationDict Is Nothing Then
            If Not mCapitalizationAcronymDict Is Nothing Then
                If Not mCapitalizationProtectedTermDict Is Nothing Then Exit Sub
            End If
        End If
    End If

    Set mCapitalizationAbbreviationDict = CreateObject("Scripting.Dictionary")
    Set mCapitalizationAcronymDict = CreateObject("Scripting.Dictionary")
    Set mCapitalizationProtectedTermDict = CreateObject("Scripting.Dictionary")
    LoadCapitalizationCompiledDataPacks
    LoadCapitalizationCustomExceptions GetCapitalizationCustomExceptions()
    mCapitalizationCacheKey = cacheKey
End Sub

Private Function CapitalizationCacheKey() As String
    Dim docKey As String

    On Error Resume Next
    docKey = ActiveDocument.FullName
    If Len(docKey) = 0 Then docKey = ActiveDocument.Name
    On Error GoTo 0

    CapitalizationCacheKey = docKey & "|" & GetCapitalizationCustomExceptions()
End Function

Private Sub InvalidateCapitalizationCache()
    mCapitalizationCacheKey = ""
    Set mCapitalizationAbbreviationDict = Nothing
    Set mCapitalizationAcronymDict = Nothing
    Set mCapitalizationProtectedTermDict = Nothing
End Sub

Private Sub LoadCapitalizationCompiledDataPacks()
    AddLookupValues mCapitalizationAbbreviationDict, Array("acct.", "admin.", "approx.", "appt.", "assoc.", "attn.", "ave.", "avg.", "bldg.", "blvd.", "bros.", "capt.", "cf.", "ch.", "chs.", "cir.", "co.", "col.", "corp.", "ct.", "dept.", "dist.", "div.", "dr.", "ed.", "eds.", "e.g.", "est.", "etc.", "fig.", "gen.", "gov.", "hon.", "hwy.", "i.e.", "inc.", "intro.", "jan.", "feb.", "mar.", "apr.", "jun.", "jul.", "aug.", "sep.", "sept.", "oct.", "nov.", "dec.", "jr.", "ln.", "ltd.", "maj.", "mgmt.", "misc.", "mr.", "mrs.", "ms.", "mt.", "no.", "nos.", "pkwy.", "pl.", "pp.", "prof.", "rd.", "ref.", "rev.", "sec.", "sq.", "sr.", "st.", "ste.", "supt.", "tel.", "terr.", "trans.", "univ.", "vol.", "vols.", "vs.")

    AddLookupPairs mCapitalizationAcronymDict, Array("AI|AI", "API|API", "BIOS|BIOS", "CEO|CEO", "CFO|CFO", "CIA|CIA", "COO|COO", "CPU|CPU", "CRM|CRM", "CSV|CSV", "DOCX|DOCX", "FAQ|FAQ", "FBI|FBI", "GPU|GPU", "HTML|HTML", "HTTP|HTTP", "HTTPS|HTTPS", "JSON|JSON", "LLC|LLC", "NASA|NASA", "OCR|OCR", "PDF|PDF", "PR|PR", "QA|QA", "RAM|RAM", "ROI|ROI", "SQL|SQL", "SSN|SSN", "UI|UI", "URL|URL", "USA|USA", "USB|USB", "UX|UX", "VBA|VBA", "XML|XML")
    AddLookupPairs mCapitalizationAcronymDict, Array("ADP|ADP", "AED|AED", "AIDS|AIDS", "ASCII|ASCII", "ATP|ATP", "BCE|BCE", "CDC|CDC", "CE|CE", "CO2|CO2", "COVID|COVID", "CPR|CPR", "DNA|DNA", "ECG|ECG", "EKG|EKG", "EPA|EPA", "FDA|FDA", "FEMA|FEMA", "H2O|H2O", "HIPAA|HIPAA", "HIV|HIV", "HMO|HMO", "ICU|ICU", "MRNA|mRNA", "MRI|MRI", "NATO|NATO", "NIH|NIH", "RNA|RNA", "SARS|SARS", "STEM|STEM", "UNESCO|UNESCO")
    AddLookupPairs mCapitalizationAcronymDict, Array("CSS|CSS", "DOI|DOI", "DVD|DVD", "GIF|GIF", "ISBN|ISBN", "ISSN|ISSN", "JPEG|JPEG", "JPG|JPG", "PC|PC", "PNG|PNG", "PPT|PPT", "PPTX|PPTX", "RTF|RTF", "SVG|SVG", "TIFF|TIFF", "TSV|TSV", "TXT|TXT", "XLS|XLS", "XLSX|XLSX", "ZIP|ZIP")
    AddLookupPairs mCapitalizationAcronymDict, Array("DNS|DNS", "FTP|FTP", "GDPR|GDPR", "GIS|GIS", "GPA|GPA", "GPS|GPS", "LAN|LAN", "LDAP|LDAP", "SDK|SDK", "SEO|SEO", "SMTP|SMTP", "TCP|TCP", "TLS|TLS", "UDP|UDP", "UUID|UUID", "VPN|VPN", "WAN|WAN")

    AddLookupPairs mCapitalizationProtectedTermDict, Array("adobe|Adobe", "airbnb|Airbnb", "chatgpt|ChatGPT", "cleanupsuite|CleanupSuite", "ebay|eBay", "excel|Excel", "github|GitHub", "google|Google", "iphone|iPhone", "ipad|iPad", "javascript|JavaScript", "linkedin|LinkedIn", "macbook|MacBook", "mcdonald|McDonald", "metadata|MetaData", "metadatasuite|MetaDataSuite", "microsoft|Microsoft", "onedrive|OneDrive", "onenote|OneNote", "oreilly|O'Reilly", "powerpoint|PowerPoint", "sharepoint|SharePoint", "youtube|YouTube")
    AddLookupPairs mCapitalizationProtectedTermDict, Array("android|Android", "airpods|AirPods", "bitbucket|Bitbucket", "dropbox|Dropbox", "facebook|Facebook", "gitlab|GitLab", "icloud|iCloud", "imac|iMac", "instagram|Instagram", "ios|iOS", "jquery|jQuery", "linux|Linux", "macos|macOS", "mongodb|MongoDB", "mysql|MySQL", "openai|OpenAI", "playstation|PlayStation", "postgresql|PostgreSQL", "powershell|PowerShell", "python|Python", "sqlite|SQLite", "tiktok|TikTok", "typescript|TypeScript", "ubuntu|Ubuntu", "whatsapp|WhatsApp", "wordpress|WordPress", "xbox|Xbox")
End Sub

Private Function GetCapitalizationCustomExceptions() As String
    On Error Resume Next
    GetCapitalizationCustomExceptions = CStr(ActiveDocument.CustomDocumentProperties(CAPITALIZATION_EXCEPTIONS_PROPERTY).Value)
    On Error GoTo 0
End Function

Private Sub SetCapitalizationCustomExceptions(ByVal exceptionText As String)
    exceptionText = NormalizeCapitalizationExceptionList(exceptionText)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties(CAPITALIZATION_EXCEPTIONS_PROPERTY).Value = exceptionText
    If Err.Number <> 0 Then
        Err.Clear
        ActiveDocument.CustomDocumentProperties.Add Name:=CAPITALIZATION_EXCEPTIONS_PROPERTY, LinkToContent:=False, Type:=CAPITALIZATION_PROPERTY_TYPE_STRING, Value:=exceptionText
    End If
    On Error GoTo 0
    InvalidateCapitalizationCache
End Sub

Private Sub PromptForCapitalizationCustomExceptions()
    Dim response As String
    response = InputBox("Enter words exactly as you want them capitalized, separated by semicolons." & vbCrLf & _
                        "Example: MetaDataSuite; CleanUpSuite; JW", _
                        "Capitalization Custom Exceptions", _
                        NormalizeCapitalizationExceptionList(GetCapitalizationCustomExceptions()))
    SetCapitalizationCustomExceptions response
End Sub

Private Function NormalizeCapitalizationExceptionList(ByVal exceptionText As String) As String
    Dim entries As Variant
    Dim i As Long
    Dim normalizedEntry As String

    If Len(Trim$(exceptionText)) = 0 Then Exit Function

    entries = Split(exceptionText, ";")
    For i = LBound(entries) To UBound(entries)
        normalizedEntry = Trim$(CStr(entries(i)))
        If Len(normalizedEntry) > 0 Then
            If Len(NormalizeCapitalizationExceptionList) > 0 Then _
                NormalizeCapitalizationExceptionList = NormalizeCapitalizationExceptionList & "; "
            NormalizeCapitalizationExceptionList = NormalizeCapitalizationExceptionList & normalizedEntry
        End If
    Next i
End Function

Private Sub LoadCapitalizationCustomExceptions(ByVal exceptionText As String)
    Dim entries As Variant
    Dim i As Long

    If Len(Trim$(exceptionText)) = 0 Then Exit Sub

    entries = Split(exceptionText, ";")
    For i = LBound(entries) To UBound(entries)
        AddCustomCapitalizationException CStr(entries(i))
    Next i
End Sub

Private Sub AddCustomCapitalizationException(ByVal entryText As String)
    Dim canonicalText As String

    canonicalText = Trim$(entryText)
    If Len(canonicalText) = 0 Then Exit Sub

    mCapitalizationProtectedTermDict(LCase$(canonicalText)) = canonicalText
    If canonicalText = UCase$(canonicalText) And Len(TokenLettersOnly(canonicalText)) > 1 Then
        mCapitalizationAcronymDict(UCase$(TokenLettersOnly(canonicalText))) = canonicalText
    End If
End Sub

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
