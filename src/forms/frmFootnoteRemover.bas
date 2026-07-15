Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Footnote / Endnote Remover
'  Deletes footnotes and/or endnotes (reference mark and note text).
'  Optionally keeps each note's text inline, in square brackets, at the
'  reference point.  Word renumbers any remaining notes automatically.
'  Works on the whole document or the current selection.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    chkFootnotes.Value = True
    chkEndnotes.Value = True
    chkKeepTextInline.Value = False
    chkFootnotes.Caption = "Remove footnotes"
    chkEndnotes.Caption = "Remove endnotes"
    chkKeepTextInline.Caption = "Keep note text inline [in brackets] instead of deleting it"
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    RefreshScopeSelectionState Me, True
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Function InScope(r As Range, targetRange As Range) As Boolean
    InScope = (r.Start >= targetRange.Start And r.Start <= targetRange.End)
End Function
Private Sub chkFootnotes_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkEndnotes_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkKeepTextInline_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub cmdRiskPlacementFootnoteFootnotes_Click(): ShowToolRiskChoiceExplanation "Remove footnotes", "Removes content", "Deletes footnotes and their reference marks." : End Sub
Private Sub cmdRiskPlacementFootnoteEndnotes_Click(): ShowToolRiskChoiceExplanation "Remove endnotes", "Removes content", "Deletes endnotes and their reference marks." : End Sub
Private Sub cmdRiskPlacementFootnoteKeepText_Click(): ShowToolRiskChoiceExplanation "Keep note text inline", "Text caution", "Copies note wording into the document before removing the note itself." : End Sub
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
    If Not GuardBeforeCleanup("Footnote / Endnote Remover") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doFoot As Boolean: doFoot = chkFootnotes.Value
    Dim doEnd As Boolean: doEnd = chkEndnotes.Value
    Dim keepText As Boolean: keepText = chkKeepTextInline.Value
    If Not (doFoot Or doEnd) Then MsgBox "Select footnotes, endnotes, or both.", vbInformation: Exit Sub
    If previewOnly Then
        Dim pf As Long, pe As Long
        Dim fnP As Footnote, enP As Endnote
        If doFoot Then
            For Each fnP In ActiveDocument.Footnotes
                If InScope(fnP.Reference, targetRange) Then fnP.Reference.HighlightColorIndex = wdYellow: pf = pf + 1
            Next fnP
        End If
        If doEnd Then
            For Each enP In ActiveDocument.Endnotes
                If InScope(enP.Reference, targetRange) Then enP.Reference.HighlightColorIndex = wdYellow: pe = pe + 1
            Next enP
        End If
        Dim previewRows As Collection
        Dim keptInline As Long
        If keepText Then keptInline = pf + pe
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Footnotes", pf, False, (Not doFoot)
        AddPreviewSummaryRow previewRows, "Endnotes", pe, False, (Not doEnd)
        AddPreviewSummaryRow previewRows, "Note text kept inline", keptInline, True, (Not keepText)
        ShowPreviewActionsSummary Me, "Footnote / Endnote Remover", previewRows
        Exit Sub
    End If
    MarkCleanupStart "Footnote / Endnote Remover"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Footnote / Endnote Remover"
    On Error GoTo RunErr
    Dim cntFoot As Long, cntEnd As Long, i As Long
    If doFoot Then
        For i = ActiveDocument.Footnotes.Count To 1 Step -1
            Dim fn As Footnote: Set fn = ActiveDocument.Footnotes(i)
            If InScope(fn.Reference, targetRange) Then
                Dim footInsertAt As Long
                Dim footText As String
                footInsertAt = fn.Reference.Start
                If keepText Then footText = CleanInlineNoteText(fn.Range.Text)
                fn.Delete
                If keepText Then ActiveDocument.Range(footInsertAt, footInsertAt).InsertAfter " [" & footText & "]"
                cntFoot = cntFoot + 1
            End If
        Next i
    End If
    If doEnd Then
        For i = ActiveDocument.Endnotes.Count To 1 Step -1
            Dim en As Endnote: Set en = ActiveDocument.Endnotes(i)
            If InScope(en.Reference, targetRange) Then
                Dim endInsertAt As Long
                Dim endText As String
                endInsertAt = en.Reference.Start
                If keepText Then endText = CleanInlineNoteText(en.Range.Text)
                en.Delete
                If keepText Then ActiveDocument.Range(endInsertAt, endInsertAt).InsertAfter " [" & endText & "]"
                cntEnd = cntEnd + 1
            End If
        Next i
    End If
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
Private Function CleanInlineNoteText(ByVal noteText As String) As String
    noteText = Replace(noteText, vbCr, "")
    noteText = Replace(noteText, Chr$(2), "")
    CleanInlineNoteText = Trim$(noteText)
End Function
