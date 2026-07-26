Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Document Trim
'  Removes trailing empty paragraphs from the end of the document,
'  leaving exactly one final paragraph (which Word requires).  Always
'  operates on the full document -- "end of document" is inherently
'  document-level, so scope selection is disabled for this tool.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    ' Scope controls disabled -- this tool always operates on the full document
    fraScopeSelection.Enabled = False
    optScopeDocument.Value = True
    optScopeDocument.Caption = "Entire document (always)"
    optScopeSelection.Enabled = False
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
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
    If Not GuardBeforeCleanup("Document Trim") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim trailingStarts As Collection
    Set trailingStarts = CollectTrailingRemovableBlankStarts(ActiveDocument)
    Dim toRemove As Long: toRemove = trailingStarts.Count
    Dim i As Long
    If previewOnly Then
        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Trailing empty paragraphs", IIf(toRemove > 0, toRemove, 0)
        If toRemove > 0 Then
            For i = 1 To trailingStarts.Count
                ApplyPreviewMinimalMarker ActiveDocument.Range(trailingStarts(i), trailingStarts(i)), False, wdYellow
            Next i
        End If
        ShowPreviewActionsSummary Me, "Document Trim", previewRows
        Exit Sub
    End If
    If toRemove <= 0 Then
        MarkCleanupToolApplied
        MsgBox "No trailing empty paragraphs found.", vbInformation
        Unload Me
        Exit Sub
    End If
    MarkCleanupStart "Document Trim"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Document Trim"
    On Error GoTo RunErr
    Dim removed As Long: removed = 0
    For i = 1 To trailingStarts.Count
        If RemoveCleanupBlankParagraphAtStart(ActiveDocument, CLng(trailingStarts(i)), cbpkRemovableBody) Then
            removed = removed + 1
        End If
    Next i
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

Private Function CollectTrailingRemovableBlankStarts(ByVal sourceDocument As Document) As Collection
    Dim starts As Collection
    Dim paragraphIndex As Long
    Dim paragraphItem As Paragraph
    Dim blankKind As CleanupBlankParagraphKind

    Set starts = New Collection
    For paragraphIndex = sourceDocument.Paragraphs.Count To 1 Step -1
        Set paragraphItem = sourceDocument.Paragraphs(paragraphIndex)
        blankKind = ClassifyCleanupBlankParagraph(paragraphItem)
        If blankKind = cbpkProtectedFinalDocument Then
            ' Word's required final paragraph is never a removal candidate.
        ElseIf blankKind = cbpkRemovableBody Then
            starts.Add paragraphItem.Range.Start
        Else
            Exit For
        End If
    Next paragraphIndex
    Set CollectTrailingRemovableBlankStarts = starts
End Function
