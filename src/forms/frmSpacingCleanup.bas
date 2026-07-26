Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Spacing Fixer
'  Corrects spacing inconsistencies throughout the document: collapses
'  double (and multiple) spaces, trims leading and trailing spaces from
'  paragraphs, removes spaces before punctuation, normalizes spaces after
'  punctuation to a single space, and collapses runs of excess blank lines.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optAll.GroupName = "SpacMode"
    optDoubleSpaces.GroupName = "SpacMode"
    optTrim.GroupName = "SpacMode"
    optCustom.GroupName = "SpacMode"
    optScopeDocument.GroupName = "SpacScope"
    optScopeSelection.GroupName = "SpacScope"
    optAll.Value = True: UpdateCustomVisibility: ClearCustomChecks
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    RefreshScopeSelectionState Me, True
    optAll.Caption = "All spacing fixes"
    optDoubleSpaces.Caption = "Double spaces only"
    optTrim.Caption = "Trim leading / trailing spaces"
    optCustom.Caption = "Custom"
    chkDoubleSpaces.Caption = "Collapse double (and multiple) spaces"
    chkTrimSpaces.Caption = "Remove leading and trailing spaces"
    chkSpaceBeforePunct.Caption = "Remove space before punctuation  ( , . : ; ! ? )"
    chkNormalizeAfterPunct.Caption = "Normalize spacing after punctuation"
    chkExtraBlankLines.Caption = "Collapse extra blank lines between paragraphs"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutSpacingForm
End Sub
Private Sub LayoutSpacingForm()
    LayoutCleanupToolForm Me
End Sub
Private Sub cmdRiskPlacementSpacingAll_Click(): ShowToolRiskChoiceExplanation "All spacing fixes", "Safe cleanup", "Runs the common spacing repairs together while preserving the visible words.": End Sub
Private Sub cmdRiskPlacementSpacingDouble_Click(): ShowToolRiskChoiceExplanation "Double spaces only", "Safe cleanup", "Collapses repeated spaces between words to one regular space.": End Sub
Private Sub cmdRiskPlacementSpacingTrim_Click(): ShowToolRiskChoiceExplanation "Trim leading / trailing spaces", "Safe cleanup", "Removes stray spaces at paragraph edges without changing the words.": End Sub
Private Sub cmdRiskPlacementSpacingCustom_Click(): ShowToolRiskChoiceExplanation "Custom spacing cleanup", "Inspect first", "Lets you combine individual spacing repairs, including blank-line cleanup. Preview first for structured documents.": End Sub
Private Sub UpdateCustomVisibility(): fraCustom.Visible = False: End Sub
Private Sub optAll_Click(): UpdateCustomVisibility: LayoutSpacingForm: End Sub
Private Sub optDoubleSpaces_Click(): UpdateCustomVisibility: LayoutSpacingForm: End Sub
Private Sub optTrim_Click(): UpdateCustomVisibility: LayoutSpacingForm: End Sub
Private Sub optCustom_Click(): UpdateCustomVisibility: LayoutSpacingForm: End Sub
Private Sub ClearCustomChecks()
    chkDoubleSpaces.Value = False: chkTrimSpaces.Value = False: chkSpaceBeforePunct.Value = False
    chkNormalizeAfterPunct.Value = False: chkExtraBlankLines.Value = False
End Sub
Private Sub cmdSelectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    chkDoubleSpaces.Value = True
    chkTrimSpaces.Value = True
    chkSpaceBeforePunct.Value = True
    chkNormalizeAfterPunct.Value = True
    chkExtraBlankLines.Value = True
    LayoutSpacingForm
End Sub
Private Sub cmdDeselectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    ClearCustomChecks
    LayoutSpacingForm
End Sub
Private Sub chkDoubleSpaces_Click(): LayoutSpacingForm: End Sub
Private Sub chkTrimSpaces_Click(): LayoutSpacingForm: End Sub
Private Sub chkSpaceBeforePunct_Click(): LayoutSpacingForm: End Sub
Private Sub chkNormalizeAfterPunct_Click(): LayoutSpacingForm: End Sub
Private Sub chkExtraBlankLines_Click(): LayoutSpacingForm: End Sub
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
Private Sub cmdRun_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected. Please remove protection before running cleanup.", vbExclamation, "Document Protected": Exit Sub
    End If
    If Not GuardBeforeCleanup("Spacing Fixer") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doDouble As Boolean, doTrim As Boolean, doBeforePunct As Boolean, doAfterPunct As Boolean, doBlank As Boolean
    If optAll.Value Then doDouble = True: doTrim = True: doBeforePunct = True: doAfterPunct = True: doBlank = True
    If optDoubleSpaces.Value Then doDouble = True
    If optTrim.Value Then doTrim = True
    If optCustom.Value Then
        doDouble = chkDoubleSpaces.Value: doTrim = chkTrimSpaces.Value: doBeforePunct = chkSpaceBeforePunct.Value
        doAfterPunct = chkNormalizeAfterPunct.Value: doBlank = chkExtraBlankLines.Value
    End If
    If Not (doDouble Or doTrim Or doBeforePunct Or doAfterPunct Or doBlank) Then MsgBox "No spacing options selected.", vbInformation: Exit Sub
    If previewOnly Then
        Dim previewRows As Collection
        Dim previewDouble As Long
        Dim previewTrim As Long
        Dim previewBefore As Long
        Dim previewAfter As Long
        Dim previewBlank As Long
        Dim previewBlankStarts As Collection
        Dim previewProtectedCells As Long
        Dim previewProtectedRows As Long
        Dim previewProtectedFinal As Long
        Dim previewSkipped As Long
        If doDouble Then previewDouble = CountPreviewFindMatches(targetRange, "  ")
        If doTrim Then
            Dim previewPara As Paragraph
            For Each previewPara In targetRange.Paragraphs
                If ParagraphContainedInRange(previewPara, targetRange) Then
                    Dim previewParagraphText As String
                    previewParagraphText = ParagraphBodyText(previewPara)
                    If Trim$(previewParagraphText) <> previewParagraphText Then previewTrim = previewTrim + 1
                End If
            Next previewPara
        End If
        If doBeforePunct Then
            previewBefore = CountPreviewFindMatches(targetRange, " .") + CountPreviewFindMatches(targetRange, " ,") + CountPreviewFindMatches(targetRange, " ;") + CountPreviewFindMatches(targetRange, " :") + CountPreviewFindMatches(targetRange, " !") + CountPreviewFindMatches(targetRange, " ?")
        End If
        If doAfterPunct Then
            previewAfter = CountPreviewFindMatches(targetRange, ".  ") + CountPreviewFindMatches(targetRange, "?  ") + CountPreviewFindMatches(targetRange, "!  ") + CountPreviewFindMatches(targetRange, ",  ") + CountPreviewFindMatches(targetRange, ";  ") + CountPreviewFindMatches(targetRange, ":  ")
        End If
        If doBlank Then
            Set previewBlankStarts = CollectCollapsibleBlankParagraphStarts(targetRange, False, previewProtectedCells, previewProtectedRows, previewProtectedFinal, previewSkipped)
            previewBlank = previewBlankStarts.Count
            Dim previewBlankIndex As Long
            For previewBlankIndex = 1 To previewBlankStarts.Count
                ApplyPreviewMinimalMarker ActiveDocument.Range(previewBlankStarts(previewBlankIndex), previewBlankStarts(previewBlankIndex))
            Next previewBlankIndex
        End If

        ' Highlight preview matches.
        If doDouble Then
            Call HighlightPreviewFindMatches(targetRange, "  ")
        End If
        If doBeforePunct Then
            Dim previewBeforePunctPatterns
            Dim previewBeforePunct
            previewBeforePunctPatterns = Array(" .", " ,", " ;", " :", " !", " ?")
            For Each previewBeforePunct In previewBeforePunctPatterns
                Call HighlightPreviewFindMatches(targetRange, CStr(previewBeforePunct))
            Next previewBeforePunct
        End If
        If doAfterPunct Then
            Dim punctPatterns
            Dim punct
            punctPatterns = Array(".  ", "?  ", "!  ", ",  ", ";  ", ":  ")
            For Each punct In punctPatterns
                Call HighlightPreviewFindMatches(targetRange, CStr(punct))
            Next punct
        End If
        If doTrim Then
            Call HighlightPreviewFindMatches(targetRange, "^p ")
            Call HighlightPreviewFindMatches(targetRange, " ^p")
        End If

        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Double spaces", previewDouble, False, (Not doDouble)
        AddPreviewSummaryRow previewRows, "Leading/trailing spaces", previewTrim, False, (Not doTrim)
        AddPreviewSummaryRow previewRows, "Spaces before punctuation", previewBefore, False, (Not doBeforePunct)
        AddPreviewSummaryRow previewRows, "Spaces after punctuation", previewAfter, False, (Not doAfterPunct)
        AddPreviewSummaryRow previewRows, "Extra blank lines", previewBlank, False, (Not doBlank)
        ShowPreviewActionsSummary Me, "Spacing Fixer", previewRows
        Exit Sub
    End If
    MarkCleanupStart "Spacing Fixer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Spacing Fixer"
    On Error GoTo RunErr
    ' Perform replacements with counting
    Dim cntDouble As Long, cntTrim As Long, cntBefore As Long, cntAfter As Long, cntBlank As Long
    Dim replacedThisPass As Long
    If doTrim Then
        Dim para As Paragraph
        Dim paraIndex As Long
        Dim trimParagraphCount As Long
        trimParagraphCount = targetRange.Paragraphs.Count
        For paraIndex = trimParagraphCount To 1 Step -1
            Set para = targetRange.Paragraphs(paraIndex)
            If ParagraphContainedInRange(para, targetRange) Then
                Dim paraText As String
                Dim trimmedText As String
                paraText = ParagraphBodyText(para)
                trimmedText = Trim$(paraText)
                If trimmedText <> paraText Then
                    ReplaceParagraphBodyText para, trimmedText
                    cntTrim = cntTrim + 1
                End If
            End If
        Next paraIndex
    End If
    If doBeforePunct Then
        Dim beforePunctPatterns: beforePunctPatterns = Array(" .", " ,", " ;", " :", " !", " ?")
        Dim beforePunct
        For Each beforePunct In beforePunctPatterns
            cntBefore = cntBefore + ReplaceSpacingLiteralMatches(targetRange, CStr(beforePunct), Right$(CStr(beforePunct), 1))
        Next beforePunct
    End If
    If doAfterPunct Then
        Dim applyPunctPatterns: applyPunctPatterns = Array(".  ","?  ","!  ",",  ",";  ",":  ")
        Dim pattern
        For Each pattern In applyPunctPatterns
            Do
                replacedThisPass = ReplaceSpacingLiteralMatches(targetRange, CStr(pattern), Left$(CStr(pattern), 1) & " ")
                cntAfter = cntAfter + replacedThisPass
            Loop While replacedThisPass > 0
        Next pattern
    End If
    If doDouble Then
        Do
            replacedThisPass = ReplaceSpacingLiteralMatches(targetRange, "  ", " ")
            cntDouble = cntDouble + replacedThisPass
        Loop While replacedThisPass > 0
    End If
    If doBlank Then
        Dim blankStarts As Collection
        Dim protectedCells As Long
        Dim protectedRows As Long
        Dim protectedFinal As Long
        Dim skippedStructures As Long
        Set blankStarts = CollectCollapsibleBlankParagraphStarts(targetRange, False, protectedCells, protectedRows, protectedFinal, skippedStructures)
        Dim blankIndex As Long
        For blankIndex = blankStarts.Count To 1 Step -1
            If RemoveCleanupBlankParagraphAtStart(ActiveDocument, CLng(blankStarts(blankIndex)), cbpkRemovableBody) Then
                cntBlank = cntBlank + 1
            End If
        Next blankIndex
    End If
    RemoveAllHighlighting ActiveDocument.Content
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
    On Error Resume Next: undoRec.EndCustomRecord: On Error GoTo 0
    MarkCleanupEnd
    Dim errMsg As String
    errMsg = "An unexpected error occurred: " & originalErrNumber & " - " & originalErrDescription
    errMsg = errMsg & vbCrLf & vbCrLf & _
             "Some changes may have been made to the document." & vbCrLf & _
             "Use Word's Undo command (Ctrl+Z) after closing this message if you want to roll them back."
    MsgBox errMsg, vbCritical, "Cleanup Error"
End Sub

Private Function ReplaceSpacingLiteralMatches(ByVal searchRange As Range, ByVal findText As String, ByVal replacementText As String) As Long
    Dim matchRange As Range
    Set matchRange = searchRange.Duplicate

    With matchRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False

        Do While .Execute
            matchRange.Text = replacementText
            ReplaceSpacingLiteralMatches = ReplaceSpacingLiteralMatches + 1
            matchRange.Collapse wdCollapseEnd
        Loop
    End With
End Function
