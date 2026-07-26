Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Break Normalizer
'  Collapses consecutive section and page breaks -- a common problem in
'  documents assembled from multiple files -- and optionally converts all
'  section breaks to a chosen type: Next Page, Continuous, Even Page,
'  or Odd Page.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optConvertNextPage.GroupName = "BreakMode"
    optConvertContinuous.GroupName = "BreakMode"
    optConvertEvenPage.GroupName = "BreakMode"
    optConvertOddPage.GroupName = "BreakMode"
    optScopeDocument.GroupName = "BreakScope"
    optScopeSelection.GroupName = "BreakScope"
    chkCollapseSectionBreaks.Value = False
    chkCollapsePageBreaks.Value = False
    chkConvertSectionBreaks.Value = False
    fraConvertTo.Enabled = False
    optConvertContinuous.Value = True
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    RefreshScopeSelectionState Me, True
    chkCollapseSectionBreaks.Caption = "Collapse consecutive section breaks"
    chkCollapsePageBreaks.Caption = "Collapse consecutive page breaks"
    chkConvertSectionBreaks.Caption = "Convert section breaks to page breaks"
    optConvertNextPage.Caption = "Next Page"
    optConvertContinuous.Caption = "Continuous"
    optConvertEvenPage.Caption = "Even Page"
    optConvertOddPage.Caption = "Odd Page"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub chkConvertSectionBreaks_Click()
    fraConvertTo.Enabled = chkConvertSectionBreaks.Value
    LayoutCleanupToolForm Me
End Sub
Private Sub chkCollapseSectionBreaks_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkCollapsePageBreaks_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub optConvertNextPage_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub optConvertContinuous_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub optConvertEvenPage_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub optConvertOddPage_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub cmdRiskPlacementBreakSections_Click(): ShowToolRiskChoiceExplanation "Collapse consecutive section breaks", "Structure change", "Merges repeated section boundaries and can affect section-level layout." : End Sub
Private Sub cmdRiskPlacementBreakPages_Click(): ShowToolRiskChoiceExplanation "Collapse consecutive page breaks", "Structure change", "Merges repeated manual page breaks and can change pagination." : End Sub
Private Sub cmdRiskPlacementBreakConvert_Click(): ShowToolRiskChoiceExplanation "Convert section breaks", "Inspect first", "Changes section breaks to page-break behavior and may affect section formatting." : End Sub
Private Sub cmdRiskPlacementBreakNextPage_Click(): ShowToolRiskChoiceExplanation "Next Page conversion", "Structure change", "Starts converted sections on a new page." : End Sub
Private Sub cmdRiskPlacementBreakContinuous_Click(): ShowToolRiskChoiceExplanation "Continuous conversion", "Structure change", "Keeps converted breaks in the current page flow." : End Sub
Private Sub cmdRiskPlacementBreakEvenPage_Click(): ShowToolRiskChoiceExplanation "Even Page conversion", "Structure change", "Forces converted sections to begin on an even page." : End Sub
Private Sub cmdRiskPlacementBreakOddPage_Click(): ShowToolRiskChoiceExplanation "Odd Page conversion", "Structure change", "Forces converted sections to begin on an odd page." : End Sub
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
    If Not GuardBeforeCleanup("Break Normalizer") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doCollSect As Boolean: doCollSect = chkCollapseSectionBreaks.Value
    Dim doCollPage As Boolean: doCollPage = chkCollapsePageBreaks.Value
    Dim doConvert As Boolean: doConvert = chkConvertSectionBreaks.Value
    If Not (doCollSect Or doCollPage Or doConvert) Then MsgBox "No break options selected.", vbInformation: Exit Sub
    ' Determine target section start type for conversion
    Dim targetSectStart As Long: targetSectStart = wdSectionContinuous
    If optConvertNextPage.Value Then targetSectStart = wdSectionNewPage
    If optConvertContinuous.Value Then targetSectStart = wdSectionContinuous
    If optConvertEvenPage.Value Then targetSectStart = wdSectionEvenPage
    If optConvertOddPage.Value Then targetSectStart = wdSectionOddPage
    If previewOnly Then
        Dim previewSectionBreaks As Long
        Dim previewPageBreaks As Long
        Dim previewConversions As Long
        If doCollSect Then previewSectionBreaks = HighlightPreviewFindBoundaryMatches(targetRange, "^b^b")
        If doCollPage Then previewPageBreaks = HighlightPreviewFindBoundaryMatches(targetRange, "^m^m")
        If doConvert Then
            Dim previewSec As Section
            For Each previewSec In ActiveDocument.Sections
                If previewSec.Range.Start >= targetRange.Start And previewSec.Range.End <= targetRange.End Then
                    If previewSec.Index > 1 Then
                        previewConversions = previewConversions + 1
                        ApplyPreviewMinimalMarker previewSec.Range, True
                    End If
                End If
            Next previewSec
        End If

        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Consecutive section breaks", previewSectionBreaks, False, (Not doCollSect)
        AddPreviewSummaryRow previewRows, "Consecutive page breaks", previewPageBreaks, False, (Not doCollPage)
        AddPreviewSummaryRow previewRows, "Section break conversions", previewConversions, False, (Not doConvert)
        ShowPreviewActionsSummary Me, "Break Normalizer", previewRows
        Exit Sub
    End If
    MarkCleanupStart "Break Normalizer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Break Normalizer"
    On Error GoTo RunErr
    Dim cntSect As Long, cntPage As Long, cntConv As Long
    If doCollSect Then
        UpdateCleanupProgress "Collapsing section breaks"
        cntSect = CollapseRepeatedBreaks(targetRange, "^b^b")
    End If
    If doCollPage Then
        UpdateCleanupProgress "Collapsing page breaks"
        cntPage = CollapseRepeatedBreaks(targetRange, "^m^m")
    End If
    If doConvert Then
        UpdateCleanupProgress "Converting section breaks"
        Dim sec As Section
        For Each sec In ActiveDocument.Sections
            If sec.Range.Start >= targetRange.Start And sec.Range.End <= targetRange.End Then
                If sec.Index > 1 Then
                    sec.PageSetup.SectionStart = targetSectStart
                    cntConv = cntConv + 1
                End If
            End If
        Next sec
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
Private Function CollapseRepeatedBreaks(ByVal searchRange As Range, ByVal findText As String) As Long
    Dim matchRange As Range
    Dim deleteRange As Range
    Do
        Set matchRange = searchRange.Duplicate
        With matchRange.Find
            .ClearFormatting
            .Text = findText
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = False
            If Not .Execute Then Exit Do
        End With
        Set deleteRange = matchRange.Duplicate
        deleteRange.Start = deleteRange.End - 1
        deleteRange.Delete
        CollapseRepeatedBreaks = CollapseRepeatedBreaks + 1
    Loop
End Function
