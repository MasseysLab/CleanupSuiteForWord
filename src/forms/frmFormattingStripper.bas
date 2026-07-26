Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Direct Formatting Stripper
'  Removes manual (direct) character and paragraph formatting while
'  PRESERVING the paragraph style.  Unlike Word built-in Clear Formatting,
'  which resets text to the Normal style, this tool returns each paragraph
'  to exactly what its applied style specifies.  Three emphasis modes
'  balance completeness against speed.  Drop caps are preserved by default;
'  untick 'Preserve drop caps' to reset them too.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    chkResetChar.Value = True
    chkResetPara.Value = True
    chkPreserveHighlight.Value = True
    chkPreserveDropCaps.Value = True
    chkPreserveHighlight.Caption = "Preserve highlighting"
    chkPreserveDropCaps.Caption = "Preserve drop caps"
    optEmphQuick.GroupName = "StripEmphasis"
    optEmphThorough.GroupName = "StripEmphasis"
    optEmphStrip.GroupName = "StripEmphasis"
    optEmphQuick.Value = True
    lblSpeedWarning.Caption = "Quick mode is fast and preserves emphasis that applies to whole paragraphs."
    optScopeDocument.GroupName = "StripScope"
    optScopeSelection.GroupName = "StripScope"
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    RefreshScopeSelectionState Me, True
    chkResetChar.Caption = "Strip direct character formatting"
    chkResetPara.Caption = "Strip direct paragraph formatting"
    optEmphQuick.Caption = "Quick  (preserve paragraph-level emphasis)"
    optEmphThorough.Caption = "Thorough  (preserve word-level emphasis)"
    optEmphStrip.Caption = "Strip  (remove all direct formatting)"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub optEmphQuick_Click()
    lblSpeedWarning.Caption = "Quick mode is fast and preserves emphasis that applies to whole paragraphs."
    LayoutCleanupToolForm Me
End Sub
Private Sub optEmphThorough_Click()
    lblSpeedWarning.Caption = "Thorough mode examines every run individually and preserves word-level emphasis.  On a long document this may take a minute or more."
    LayoutCleanupToolForm Me
End Sub
Private Sub optEmphStrip_Click()
    lblSpeedWarning.Caption = "Strip mode removes all direct formatting, including bold and italic emphasis."
    LayoutCleanupToolForm Me
End Sub
Private Sub chkResetChar_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkResetPara_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkPreserveHighlight_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkPreserveDropCaps_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub cmdRiskPlacementFormattingChar_Click(): ShowToolRiskChoiceExplanation "Strip character formatting", "Formatting change", "Removes direct font-level formatting while leaving styles in charge." : End Sub
Private Sub cmdRiskPlacementFormattingPara_Click(): ShowToolRiskChoiceExplanation "Strip paragraph formatting", "Formatting change", "Removes direct paragraph spacing, indents, and alignment." : End Sub
Private Sub cmdRiskPlacementFormattingQuick_Click(): ShowToolRiskChoiceExplanation "Quick emphasis preservation", "Inspect first", "Preserves emphasis that applies to whole paragraphs." : End Sub
Private Sub cmdRiskPlacementFormattingThorough_Click(): ShowToolRiskChoiceExplanation "Thorough emphasis preservation", "Inspect first", "Checks smaller text runs and may take longer." : End Sub
Private Sub cmdRiskPlacementFormattingStrip_Click(): ShowToolRiskChoiceExplanation "Strip all direct emphasis", "Formatting change", "Removes manual emphasis such as direct bold and italic." : End Sub
Private Sub cmdRiskPlacementFormattingHighlight_Click(): ShowToolRiskChoiceExplanation "Preserve highlighting", "Safe cleanup", "Leaves highlighting in place while other direct formatting is stripped." : End Sub
Private Sub cmdRiskPlacementFormattingDropCaps_Click(): ShowToolRiskChoiceExplanation "Preserve drop caps", "Safe cleanup", "Leaves decorative drop caps in place." : End Sub
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
    If Not GuardBeforeCleanup("Formatting Stripper") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doResetChar As Boolean: doResetChar = chkResetChar.Value
    Dim doResetPara As Boolean: doResetPara = chkResetPara.Value
    Dim preserveHL As Boolean: preserveHL = chkPreserveHighlight.Value
    Dim preserveDC As Boolean: preserveDC = chkPreserveDropCaps.Value
    If Not (doResetChar Or doResetPara) Then MsgBox "Nothing selected to reset. Tick at least one reset option.", vbInformation: Exit Sub
    Dim p As Paragraph
    Dim previewChar As Long
    Dim previewPara As Long
    If previewOnly Then
        For Each p In targetRange.Paragraphs
            If ParagraphContainedInRange(p, targetRange) Then
                Dim hasDirectCharacterFormatting As Boolean
                Dim hasDirectParagraphFormatting As Boolean
                GetDirectFormattingFlags p, doResetChar, doResetPara, _
                                         hasDirectCharacterFormatting, hasDirectParagraphFormatting
                If hasDirectCharacterFormatting Then previewChar = previewChar + 1
                If hasDirectParagraphFormatting Then previewPara = previewPara + 1
                If hasDirectCharacterFormatting Or hasDirectParagraphFormatting Then
                    ApplyPreviewHighlight p.Range, wdYellow
                End If
            End If
        Next p
        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Character formatting", previewChar, False, (Not doResetChar)
        AddPreviewSummaryRow previewRows, "Paragraph formatting", previewPara, False, (Not doResetPara)
        ShowPreviewActionsSummary Me, "Formatting Stripper", previewRows
        Exit Sub
    End If
    Dim emphMode As Integer: emphMode = 0
    If optEmphThorough.Value Then emphMode = 1
    If optEmphStrip.Value Then emphMode = 2
    If emphMode = 1 Then
        Dim totalChars As Long: totalChars = 0
        For Each p In targetRange.Paragraphs
            If ParagraphContainedInRange(p, targetRange) Then
                totalChars = totalChars + p.Range.Characters.Count
            End If
        Next p
        If totalChars > 20000 Then
            If MsgBox("This document has " & totalChars & " characters in scope." & vbCrLf & "Thorough cleaning may take a while. Continue?", vbQuestion + vbYesNo, "Thorough Clean") = vbNo Then Exit Sub
        End If
    End If
    MarkCleanupStart "Formatting Stripper"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Formatting Stripper"
    On Error GoTo RunErr
    Application.ScreenUpdating = False
    Dim cntReset As Long: cntReset = 0
    Dim processed As Long: processed = 0
    Dim paragraphsToProcess As Long
    paragraphsToProcess = targetRange.Paragraphs.Count
    For Each p In targetRange.Paragraphs
        If ParagraphContainedInRange(p, targetRange) Then
            ApplyStripToParagraph p, doResetChar, doResetPara, emphMode, preserveHL, preserveDC
            cntReset = cntReset + 1
            processed = processed + 1
            If emphMode = 1 And (processed Mod 25 = 0 Or processed = paragraphsToProcess) Then
                UpdateCleanupProgress "Cleaning formatting", processed, paragraphsToProcess
            End If
        End If
    Next p
    Application.ScreenUpdating = True
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
    Application.ScreenUpdating = True
    On Error Resume Next: undoRec.EndCustomRecord: On Error GoTo 0
    MarkCleanupEnd
    Dim errMsg As String
    errMsg = "An unexpected error occurred: " & originalErrNumber & " - " & originalErrDescription
    errMsg = errMsg & vbCrLf & vbCrLf & _
             "Some changes may have been made to the document." & vbCrLf & _
             "Use Word's Undo command (Ctrl+Z) after closing this message if you want to roll them back."
    MsgBox errMsg, vbCritical, "Cleanup Error"
End Sub
Private Sub GetDirectFormattingFlags(ByVal p As Paragraph, _
                                     ByVal checkCharacterFormatting As Boolean, _
                                     ByVal checkParagraphFormatting As Boolean, _
                                     ByRef hasCharacterFormatting As Boolean, _
                                     ByRef hasParagraphFormatting As Boolean)
    Dim paragraphRange As Range
    Dim paragraphFormat As ParagraphFormat
    Dim styleItem As Style

    hasCharacterFormatting = False
    hasParagraphFormatting = False
    On Error GoTo SafeExit

    Set paragraphRange = p.Range
    Set paragraphFormat = paragraphRange.ParagraphFormat
    Set styleItem = ActiveDocument.Styles(p.Style)
    If styleItem Is Nothing Then Exit Sub

    If checkCharacterFormatting Then
        If paragraphRange.Font.Name <> styleItem.Font.Name Then hasCharacterFormatting = True
        If paragraphRange.Font.Size <> styleItem.Font.Size And paragraphRange.Font.Size > 0 Then hasCharacterFormatting = True
        If paragraphRange.Font.Bold <> styleItem.Font.Bold Then hasCharacterFormatting = True
        If paragraphRange.Font.Italic <> styleItem.Font.Italic Then hasCharacterFormatting = True
        If paragraphRange.Font.Color <> styleItem.Font.Color Then hasCharacterFormatting = True
    End If

    If checkParagraphFormatting Then
        If paragraphFormat.LeftIndent <> styleItem.ParagraphFormat.LeftIndent Then hasParagraphFormatting = True
        If paragraphFormat.SpaceBefore <> styleItem.ParagraphFormat.SpaceBefore Then hasParagraphFormatting = True
        If paragraphFormat.SpaceAfter <> styleItem.ParagraphFormat.SpaceAfter Then hasParagraphFormatting = True
        If paragraphFormat.Alignment <> styleItem.ParagraphFormat.Alignment Then hasParagraphFormatting = True
    End If

SafeExit:
End Sub
Private Sub ApplyStripToParagraph(p As Paragraph, doChar As Boolean, doPara As Boolean, emphMode As Integer, preserveHL As Boolean, preserveDropCaps As Boolean)
    If Not preserveDropCaps Then
        On Error Resume Next
        p.DropCap.Position = wdDropNone
        On Error GoTo 0
    End If
    If Not doChar Then
        If doPara Then p.Range.ParagraphFormat.Reset
        Exit Sub
    End If
    If emphMode = 2 Then
        p.Range.Font.Reset
        If doPara Then p.Range.ParagraphFormat.Reset
    ElseIf emphMode = 0 Then
        Dim wasBold As Variant, wasItalic As Variant, savedHL As Variant
        wasBold = p.Range.Font.Bold
        wasItalic = p.Range.Font.Italic
        savedHL = p.Range.HighlightColorIndex
        p.Range.Font.Reset
        If doPara Then p.Range.ParagraphFormat.Reset
        If wasBold = True Then p.Range.Font.Bold = True
        If wasItalic = True Then p.Range.Font.Italic = True
        If preserveHL And savedHL <> wdUndefined And savedHL <> wdNoHighlight Then p.Range.HighlightColorIndex = savedHL
    Else
        Dim n As Long: n = p.Range.Characters.Count
        If n < 1 Then Exit Sub
        Dim bArr() As Boolean, iArr() As Boolean, hArr() As Long
        ReDim bArr(1 To n)
        ReDim iArr(1 To n)
        ReDim hArr(1 To n)
        Dim ch As Range, idx As Long: idx = 0
        For Each ch In p.Range.Characters
            idx = idx + 1
            If idx <= n Then
                bArr(idx) = (ch.Font.Bold = True)
                iArr(idx) = (ch.Font.Italic = True)
                hArr(idx) = ch.HighlightColorIndex
            End If
        Next ch
        Dim baseStart As Long: baseStart = p.Range.Start
        p.Range.Font.Reset
        If doPara Then p.Range.ParagraphFormat.Reset
        ApplyRuns baseStart, bArr, iArr, hArr, n, preserveHL
    End If
End Sub
Private Sub ApplyRuns(baseStart As Long, bArr() As Boolean, iArr() As Boolean, hArr() As Long, n As Long, preserveHL As Boolean)
    Dim i As Long
    i = 1
    Do While i <= n
        If bArr(i) Then
            Dim jb As Long: jb = i
            Do While jb < n
                If Not bArr(jb + 1) Then Exit Do
                jb = jb + 1
            Loop
            ActiveDocument.Range(baseStart + i - 1, baseStart + jb).Font.Bold = True
            i = jb + 1
        Else
            i = i + 1
        End If
    Loop
    i = 1
    Do While i <= n
        If iArr(i) Then
            Dim ji As Long: ji = i
            Do While ji < n
                If Not iArr(ji + 1) Then Exit Do
                ji = ji + 1
            Loop
            ActiveDocument.Range(baseStart + i - 1, baseStart + ji).Font.Italic = True
            i = ji + 1
        Else
            i = i + 1
        End If
    Loop
    If preserveHL Then
        i = 1
        Do While i <= n
            If hArr(i) <> wdNoHighlight And hArr(i) <> wdUndefined Then
                Dim jh As Long: jh = i
                Do While jh < n
                    If hArr(jh + 1) <> hArr(i) Then Exit Do
                    jh = jh + 1
                Loop
                ActiveDocument.Range(baseStart + i - 1, baseStart + jh).HighlightColorIndex = hArr(i)
                i = jh + 1
            Else
                i = i + 1
            End If
        Loop
    End If
End Sub
