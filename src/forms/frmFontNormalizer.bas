Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Font Normalizer
'  Resets direct font overrides back to what the paragraph style defines,
'  eliminating the font chaos that accumulates from copy-paste.  Face and
'  size are reset by default; bold, italic, and colour are opt-in since
'  those properties are more likely to be intentional.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    chkFontFace.Value = True
    chkFontSize.Value = True
    chkBold.Value = False
    chkItalic.Value = False
    chkFontColor.Value = False
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    RefreshScopeSelectionState Me, True
    chkFontFace.Caption = "Normalize font face to document default"
    chkFontSize.Caption = "Normalize font size to style-defined size"
    chkBold.Caption = "Normalize direct bold overrides"
    chkItalic.Caption = "Normalize direct italic overrides"
    chkFontColor.Caption = "Remove direct font color overrides"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub chkFontFace_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkFontSize_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkBold_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkItalic_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkFontColor_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub cmdRiskPlacementFontFace_Click(): ShowToolRiskChoiceExplanation "Normalize font face", "Formatting change", "Removes direct font-name overrides so styles can control font face." : End Sub
Private Sub cmdRiskPlacementFontSize_Click(): ShowToolRiskChoiceExplanation "Normalize font size", "Formatting change", "Removes direct size overrides so styles can control size." : End Sub
Private Sub cmdRiskPlacementFontBold_Click(): ShowToolRiskChoiceExplanation "Normalize bold", "Formatting change", "Removes manual bolding not supplied by the style." : End Sub
Private Sub cmdRiskPlacementFontItalic_Click(): ShowToolRiskChoiceExplanation "Normalize italic", "Formatting change", "Removes manual italics not supplied by the style." : End Sub
Private Sub cmdRiskPlacementFontColor_Click(): ShowToolRiskChoiceExplanation "Remove font color overrides", "Formatting change", "Clears direct font colors so style color can lead." : End Sub
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
    If Not GuardBeforeCleanup("Font Normalizer") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doFace As Boolean: doFace = chkFontFace.Value
    Dim doSize As Boolean: doSize = chkFontSize.Value
    Dim doBold As Boolean: doBold = chkBold.Value
    Dim doItalic As Boolean: doItalic = chkItalic.Value
    Dim doColor As Boolean: doColor = chkFontColor.Value
    If Not (doFace Or doSize Or doBold Or doItalic Or doColor) Then MsgBox "No font properties selected.", vbInformation: Exit Sub
    Dim p As Paragraph
    Dim styleFont As Font
    Dim cnt As Long
    If previewOnly Then
        Dim previewFace As Long
        Dim previewSize As Long
        Dim previewBold As Long
        Dim previewItalic As Long
        Dim previewColor As Long
        For Each p In targetRange.Paragraphs
            If ParagraphContainedInRange(p, targetRange) Then
                Set styleFont = Nothing
                On Error Resume Next
                Set styleFont = ActiveDocument.Styles(p.Style).Font
                On Error GoTo 0
                If Not styleFont Is Nothing Then
                    Dim hasDirect As Boolean
                    hasDirect = False
                    If doFace And p.Range.Font.Name <> styleFont.Name Then hasDirect = True: previewFace = previewFace + 1
                    If doSize And p.Range.Font.Size <> styleFont.Size And p.Range.Font.Size > 0 Then hasDirect = True: previewSize = previewSize + 1
                    If doBold And p.Range.Font.Bold <> styleFont.Bold Then hasDirect = True: previewBold = previewBold + 1
                    If doItalic And p.Range.Font.Italic <> styleFont.Italic Then hasDirect = True: previewItalic = previewItalic + 1
                    If doColor And p.Range.Font.Color <> styleFont.Color Then hasDirect = True: previewColor = previewColor + 1
                    If hasDirect Then
                        ApplyPreviewHighlight p.Range, wdYellow
                        cnt = cnt + 1
                    End If
                End If
            End If
        Next p

        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Font face", previewFace, False, (Not doFace)
        AddPreviewSummaryRow previewRows, "Font size", previewSize, False, (Not doSize)
        AddPreviewSummaryRow previewRows, "Bold", previewBold, False, (Not doBold)
        AddPreviewSummaryRow previewRows, "Italic", previewItalic, False, (Not doItalic)
        AddPreviewSummaryRow previewRows, "Font color", previewColor, False, (Not doColor)
        ShowPreviewActionsSummary Me, "Font Normalizer", previewRows
        Exit Sub
    End If
    MarkCleanupStart "Font Normalizer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Font Normalizer"
    On Error GoTo RunErr
    Dim cntFixed As Long: cntFixed = 0
    For Each p In targetRange.Paragraphs
        If ParagraphContainedInRange(p, targetRange) Then
            Set styleFont = Nothing
            On Error Resume Next
            Set styleFont = ActiveDocument.Styles(p.Style).Font
            On Error GoTo 0
            If Not styleFont Is Nothing Then
                Dim changed As Boolean: changed = False
                If doFace Then p.Range.Font.Name = styleFont.Name: changed = True
                If doSize Then p.Range.Font.Size = styleFont.Size: changed = True
                If doBold Then p.Range.Font.Bold = styleFont.Bold: changed = True
                If doItalic Then p.Range.Font.Italic = styleFont.Italic: changed = True
                If doColor Then p.Range.Font.Color = styleFont.Color: changed = True
                If changed Then cntFixed = cntFixed + 1
            End If
        End If
    Next p
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
