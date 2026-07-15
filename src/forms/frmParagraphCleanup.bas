Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Paragraph Structure Fixer
'  Cleans up paragraph structure: collapses repeated empty paragraphs,
'  converts soft returns to paragraphs, resets
'  SpaceBefore/SpaceAfter to zero for paragraphs with direct spacing
'  overrides, and resets left indent to zero.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optAll.GroupName = "ParaMode"
    optRemoveEmpty.GroupName = "ParaMode"
    optNormalizeSpacing.GroupName = "ParaMode"
    optCustom.GroupName = "ParaMode"
    optScopeDocument.GroupName = "ParaScope"
    optScopeSelection.GroupName = "ParaScope"
    optAll.Value = True: UpdateCustomVisibility: ClearCustomChecks
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    RefreshScopeSelectionState Me, True
    optAll.Caption = "All paragraph fixes"
    optRemoveEmpty.Caption = "Remove empty paragraphs"
    optNormalizeSpacing.Caption = "Normalize paragraph spacing"
    optCustom.Caption = "Custom"
    chkRemoveEmpty.Caption = "Collapse multiple consecutive empty paragraphs"
    chkCollapseBreaks.Caption = "Convert soft returns (Shift+Enter) to paragraph marks"
    chkNormalizeParaSpacing.Caption = "Normalize Space Before / Space After settings"
    chkFixIndent.Caption = "Remove manual tab / space indentation at paragraph start"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutParagraphForm
End Sub
Private Sub LayoutParagraphForm()
    LayoutCleanupToolForm Me
End Sub
Private Sub UpdateCustomVisibility(): fraCustom.Visible = False: End Sub
Private Sub optAll_Click(): UpdateCustomVisibility: LayoutParagraphForm: End Sub
Private Sub optRemoveEmpty_Click(): UpdateCustomVisibility: LayoutParagraphForm: End Sub
Private Sub optNormalizeSpacing_Click(): UpdateCustomVisibility: LayoutParagraphForm: End Sub
Private Sub optCustom_Click(): UpdateCustomVisibility: LayoutParagraphForm: End Sub
Private Sub ClearCustomChecks(): chkRemoveEmpty.Value = False: chkCollapseBreaks.Value = False: chkNormalizeParaSpacing.Value = False: chkFixIndent.Value = False: End Sub
Private Sub cmdSelectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    chkRemoveEmpty.Value = True
    chkCollapseBreaks.Value = True
    chkNormalizeParaSpacing.Value = True
    chkFixIndent.Value = True
    LayoutParagraphForm
End Sub
Private Sub cmdDeselectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    ClearCustomChecks
    LayoutParagraphForm
End Sub
Private Sub chkRemoveEmpty_Click(): LayoutParagraphForm: End Sub
Private Sub chkCollapseBreaks_Click(): LayoutParagraphForm: End Sub
Private Sub chkNormalizeParaSpacing_Click(): LayoutParagraphForm: End Sub
Private Sub chkFixIndent_Click(): LayoutParagraphForm: End Sub
Private Sub cmdRiskPlacementParagraphAll_Click(): ShowToolRiskChoiceExplanation "All paragraph fixes", "Structure change", "Runs the full paragraph structure cleanup set." : End Sub
Private Sub cmdRiskPlacementParagraphRemoveEmpty_Click(): ShowToolRiskChoiceExplanation "Remove empty paragraphs", "Safe cleanup", "Collapses repeated empty paragraphs while preserving document text." : End Sub
Private Sub cmdRiskPlacementParagraphSpacing_Click(): ShowToolRiskChoiceExplanation "Normalize paragraph spacing", "Formatting change", "Clears direct Space Before and Space After overrides." : End Sub
Private Sub cmdRiskPlacementParagraphCustom_Click(): ShowToolRiskChoiceExplanation "Custom paragraph cleanup", "Inspect first", "Lets you combine individual paragraph structure repairs." : End Sub
Private Sub cmdRiskPlacementParagraphChkRemoveEmpty_Click(): ShowToolRiskChoiceExplanation "Collapse empty paragraphs", "Safe cleanup", "Reduces repeated blank paragraphs." : End Sub
Private Sub cmdRiskPlacementParagraphChkBreaks_Click(): ShowToolRiskChoiceExplanation "Soft returns to paragraphs", "Structure change", "Turns manual line breaks into paragraph marks." : End Sub
Private Sub cmdRiskPlacementParagraphChkSpacing_Click(): ShowToolRiskChoiceExplanation "Normalize spacing", "Formatting change", "Clears direct paragraph spacing overrides." : End Sub
Private Sub cmdRiskPlacementParagraphChkIndent_Click(): ShowToolRiskChoiceExplanation "Fix manual indents", "Structure change", "Removes leading manual tab or space indentation." : End Sub
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
    If Not GuardBeforeCleanup("Paragraph Fixer") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Set targetRange = ParagraphCleanupTargetRange(targetRange)
    Dim doRemove As Boolean, doCollapse As Boolean, doNormalize As Boolean, doIndent As Boolean
    If optAll.Value Then doRemove = True: doCollapse = True: doNormalize = True: doIndent = True
    If optRemoveEmpty.Value Then doRemove = True
    If optNormalizeSpacing.Value Then doNormalize = True
    If optCustom.Value Then doRemove = chkRemoveEmpty.Value: doCollapse = chkCollapseBreaks.Value: doNormalize = chkNormalizeParaSpacing.Value: doIndent = chkFixIndent.Value
    If Not (doRemove Or doCollapse Or doNormalize Or doIndent) Then MsgBox "No paragraph options selected.", vbInformation: Exit Sub
    Dim p As Paragraph
    Dim cntRemove As Long, cntCollapse As Long, cntNormalize As Long, cntIndent As Long
    If previewOnly Then
        If doRemove Then cntRemove = HighlightExtraEmptyParagraphs(targetRange)
        If doCollapse Then cntCollapse = CountPreviewFindMatches(targetRange, "^l")
        If doNormalize Then
            For Each p In targetRange.Paragraphs
                If ParagraphOverlapsRange(p, targetRange) Then
                    If p.Format.SpaceBefore <> 0 Or p.Format.SpaceAfter <> 0 Then
                        ApplyPreviewShading p.Range
                        cntNormalize = cntNormalize + 1
                    End If
                End If
            Next p
        End If
        If doIndent Then
            For Each p In targetRange.Paragraphs
                If ParagraphOverlapsRange(p, targetRange) Then
                    Dim leadingIndentLength As Long
                    leadingIndentLength = ParagraphLeadingIndentLength(p)
                    If p.Format.LeftIndent <> 0 Or leadingIndentLength > 0 Then
                        If leadingIndentLength > 0 Then
                            ActiveDocument.Range(p.Range.Start, p.Range.Start + leadingIndentLength).HighlightColorIndex = wdBrightGreen
                        Else
                            ApplyPreviewShading p.Range
                        End If
                        cntIndent = cntIndent + 1
                    End If
                End If
            Next p
        End If

        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Empty paragraphs", cntRemove, False, (Not doRemove)
        AddPreviewSummaryRow previewRows, "Soft returns", cntCollapse, True, (Not doCollapse)
        AddPreviewSummaryRow previewRows, "Paragraph spacing", cntNormalize, False, (Not doNormalize)
        AddPreviewSummaryRow previewRows, "Indents", cntIndent, False, (Not doIndent)
        ShowPreviewActionsSummary Me, "Paragraph Fixer", previewRows
        Exit Sub
    End If
    MarkCleanupStart "Paragraph Fixer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Paragraph Fixer"
    On Error GoTo RunErr
    ' Perform paragraph cleanup with counting
    cntRemove = 0: cntCollapse = 0: cntNormalize = 0: cntIndent = 0
    Dim replacedThisPass As Long
    If doRemove Then
        UpdateCleanupProgress "Removing extra empty paragraphs"
        Do
            replacedThisPass = ReplaceParagraphFindMatches(targetRange, "^p^p^p", "^p^p")
            cntRemove = cntRemove + replacedThisPass
        Loop While replacedThisPass > 0
    End If
    If doCollapse Then
        UpdateCleanupProgress "Converting soft returns"
        Do
            replacedThisPass = ReplaceParagraphFindMatches(targetRange, "^l", "^p")
            cntCollapse = cntCollapse + replacedThisPass
        Loop While replacedThisPass > 0
    End If
    ' Normalize paragraph spacing
    If doNormalize Then
        UpdateCleanupProgress "Normalizing paragraph spacing"
        For Each p In targetRange.Paragraphs
            If ParagraphOverlapsRange(p, targetRange) Then
                With p.Format
                    If .SpaceBefore <> 0 Or .SpaceAfter <> 0 Then
                        .SpaceBefore = 0
                        .SpaceAfter = 0
                        cntNormalize = cntNormalize + 1
                    End If
                End With
            End If
        Next p
    End If
    ' Fix indent
    If doIndent Then
        UpdateCleanupProgress "Fixing manual indents"
        Dim paragraphIndex As Long
        For paragraphIndex = targetRange.Paragraphs.Count To 1 Step -1
            Set p = targetRange.Paragraphs(paragraphIndex)
            If ParagraphOverlapsRange(p, targetRange) Then
                leadingIndentLength = ParagraphLeadingIndentLength(p)
                If p.Format.LeftIndent <> 0 Or leadingIndentLength > 0 Then
                    p.Format.LeftIndent = 0
                    If leadingIndentLength > 0 Then ActiveDocument.Range(p.Range.Start, p.Range.Start + leadingIndentLength).Delete
                    cntIndent = cntIndent + 1
                End If
            End If
        Next paragraphIndex
    End If
    RemoveAllHighlighting targetRange
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
Private Function ParagraphLeadingIndentLength(ByVal p As Paragraph) As Long
    Dim textValue As String
    Dim i As Long
    textValue = p.Range.Text
    For i = 1 To Len(textValue)
        If Mid$(textValue, i, 1) = " " Or Mid$(textValue, i, 1) = vbTab Then
            ParagraphLeadingIndentLength = ParagraphLeadingIndentLength + 1
        Else
            Exit For
        End If
    Next i
End Function

Private Function ParagraphCleanupTargetRange(ByVal sourceRange As Range) As Range
    Dim expanded As Range
    Set expanded = sourceRange.Duplicate

    If expanded.Paragraphs.Count > 0 Then
        expanded.Start = expanded.Paragraphs(1).Range.Start
        expanded.End = expanded.Paragraphs(expanded.Paragraphs.Count).Range.End
    End If

    Set ParagraphCleanupTargetRange = expanded
End Function

Private Function HighlightParagraphFindMatches(ByVal searchRange As Range, ByVal findText As String) As Long
    Dim matchRange As Range
    Dim searchEnd As Long
    searchEnd = searchRange.End
    Set matchRange = searchRange.Duplicate

    Do
        With matchRange.Find
            .ClearFormatting
            .Text = findText
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = False
            If Not .Execute Then Exit Do
        End With

        matchRange.HighlightColorIndex = wdYellow
        HighlightParagraphFindMatches = HighlightParagraphFindMatches + 1
        matchRange.Collapse wdCollapseEnd
        If matchRange.Start >= searchEnd Then Exit Do
        matchRange.End = searchEnd
    Loop
End Function
Private Function HighlightExtraEmptyParagraphs(ByVal searchRange As Range) As Long
    Dim p As Paragraph
    Dim blankRun As Long
    For Each p In searchRange.Paragraphs
        If ParagraphOverlapsRange(p, searchRange) Then
            If Len(Trim$(Replace(Replace(p.Range.Text, vbCr, ""), Chr$(7), ""))) = 0 Then
                blankRun = blankRun + 1
                If blankRun > 1 Then
                    ApplyPreviewShading p.Range
                    HighlightExtraEmptyParagraphs = HighlightExtraEmptyParagraphs + 1
                End If
            Else
                blankRun = 0
            End If
        Else
            blankRun = 0
        End If
    Next p
End Function

Private Function ReplaceParagraphFindMatches(ByVal searchRange As Range, ByVal findText As String, ByVal replacementText As String) As Long
    Dim matchRange As Range
    Set matchRange = searchRange.Duplicate

    With matchRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = replacementText
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False

        Do While .Execute(Replace:=wdReplaceOne)
            ReplaceParagraphFindMatches = ReplaceParagraphFindMatches + 1
        Loop
    End With
End Function
