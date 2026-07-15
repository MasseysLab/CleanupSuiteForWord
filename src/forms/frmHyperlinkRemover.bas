Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Hyperlink Cleaner
'  Either hides visible hyperlink styling while keeping the link target, or
'  removes hyperlink targets while keeping the visible text.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optHideLinks.GroupName = "HLAction"
    optRemoveLinks.GroupName = "HLAction"
    optHideLinks.Value = True
    optScopeDocument.GroupName = "HLScope"
    optScopeSelection.GroupName = "HLScope"
    optHideLinks.Caption = "Hide hyperlink styling, keep links"
    optRemoveLinks.Caption = "Remove hyperlinks, keep visible text"
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    RefreshScopeSelectionState Me, True
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutHyperlinkForm
End Sub
Private Sub LayoutHyperlinkForm()
    LayoutCleanupToolForm Me
End Sub
Private Sub optHideLinks_Click(): LayoutHyperlinkForm: End Sub
Private Sub optRemoveLinks_Click(): LayoutHyperlinkForm: End Sub
Private Sub cmdRiskPlacementHyperlinkHide_Click()
    ShowToolRiskChoiceExplanation "Hide hyperlink styling", "Formatting change", "This keeps the hyperlink target but clears the visible blue underline styling so the text looks normal until someone hovers over or opens the link."
End Sub
Private Sub cmdRiskPlacementHyperlinkRemove_Click()
    ShowToolRiskChoiceExplanation "Remove hyperlinks", "Removes content", "This removes the hyperlink target from the document while keeping the visible text."
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
Private Sub cmdRun_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected. Please remove protection before running cleanup.", vbExclamation, "Document Protected": Exit Sub
    End If
    If Not GuardBeforeCleanup("Hyperlink Cleaner") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim removeLinks As Boolean: removeLinks = optRemoveLinks.Value
    Dim i As Long, cnt As Long: cnt = 0
    If previewOnly Then
        For i = 1 To ActiveDocument.Hyperlinks.Count
            Dim hp As Hyperlink: Set hp = ActiveDocument.Hyperlinks(i)
            If hp.Range.Start >= targetRange.Start And hp.Range.End <= targetRange.End Then hp.Range.HighlightColorIndex = wdYellow: cnt = cnt + 1
        Next i
        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Hyperlinks", cnt
        ShowPreviewActionsSummary Me, "Hyperlink Cleaner", previewRows
        Exit Sub
    End If
    MarkCleanupStart "Hyperlink Cleaner"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Hyperlink Cleaner"
    On Error GoTo RunErr
    For i = ActiveDocument.Hyperlinks.Count To 1 Step -1
        Dim hl As Hyperlink: Set hl = ActiveDocument.Hyperlinks(i)
        If hl.Range.Start >= targetRange.Start And hl.Range.End <= targetRange.End Then
            If removeLinks Then
                Dim rg As Range: Set rg = hl.Range.Duplicate
                ClearHyperlinkVisibleStyle rg
                hl.Delete
                cnt = cnt + 1
            Else
                ClearHyperlinkVisibleStyle hl.Range
                cnt = cnt + 1
            End If
        End If
    Next i
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
Private Sub ClearHyperlinkVisibleStyle(ByVal target As Range)
    On Error Resume Next
    target.Font.Underline = wdUnderlineNone
    target.Font.ColorIndex = wdAuto
    On Error GoTo 0
End Sub
