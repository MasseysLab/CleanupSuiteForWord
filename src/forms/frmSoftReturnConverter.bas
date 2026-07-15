Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Soft Return Converter
'  Converts soft returns (Shift+Enter line breaks) to real paragraph marks,
'  or the reverse.  Imported text from PDFs and web pages is often riddled
'  with soft returns that look like paragraph breaks but break styles,
'  spell-check, and spacing.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optSoftToPara.GroupName = "SoftDir"
    optParaToSoft.GroupName = "SoftDir"
    optSoftToPara.Value = True
    optScopeDocument.GroupName = "SoftScope"
    optScopeSelection.GroupName = "SoftScope"
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    RefreshScopeSelectionState Me, True
    optSoftToPara.Caption = "Convert soft returns (Shift+Enter) to paragraph marks"
    optParaToSoft.Caption = "Convert paragraph marks to soft returns"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub optSoftToPara_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub optParaToSoft_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub cmdRiskPlacementSoftReturnToPara_Click(): ShowToolRiskChoiceExplanation "Soft returns to paragraphs", "Structure change", "Turns Shift+Enter line breaks into paragraph marks." : End Sub
Private Sub cmdRiskPlacementSoftReturnToSoft_Click(): ShowToolRiskChoiceExplanation "Paragraphs to soft returns", "Structure change", "Turns paragraph marks into soft returns in the selected scope." : End Sub
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
    If Not GuardBeforeCleanup("Soft Return Converter") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim softToPara As Boolean: softToPara = optSoftToPara.Value
    If previewOnly Then
        Dim pc As Long
        Dim previewRows As Collection
        RemoveAllHighlighting ActiveDocument.Content
        pc = CountSoftReturnConvertibleMarks(targetRange, softToPara)
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Soft returns", IIf(softToPara, pc, 0), True, (Not softToPara)
        AddPreviewSummaryRow previewRows, "Paragraph marks", IIf(softToPara, 0, pc), True, softToPara
        ShowPreviewActionsSummary Me, "Soft Return Converter", previewRows
        Exit Sub
    End If
    MarkCleanupStart "Soft Return Converter"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Soft Return Converter"
    On Error GoTo RunErr
    Dim cnt As Long: cnt = ReplaceSoftReturnMarks(targetRange, softToPara)
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
Private Function CountSoftReturnConvertibleMarks(ByVal searchRange As Range, ByVal softToPara As Boolean) As Long
    If softToPara Then
        CountSoftReturnConvertibleMarks = CountSoftReturnCharacters(searchRange.Text, Chr$(11))
    Else
        CountSoftReturnConvertibleMarks = CountSoftReturnCharacters(searchRange.Text, Chr$(13))
        If Right$(searchRange.Text, 1) = Chr$(13) Then
            If CountSoftReturnConvertibleMarks > 0 Then CountSoftReturnConvertibleMarks = CountSoftReturnConvertibleMarks - 1
        End If
    End If
End Function
Private Function CountSoftReturnCharacters(ByVal textValue As String, ByVal markChar As String) As Long
    CountSoftReturnCharacters = Len(textValue) - Len(Replace(textValue, markChar, ""))
End Function
Private Function ReplaceSoftReturnMarks(ByVal searchRange As Range, ByVal softToPara As Boolean) As Long
    Dim matchRange As Range
    Dim searchEnd As Long
    Dim nextStart As Long
    searchEnd = searchRange.End
    Set matchRange = ActiveDocument.Range(searchRange.Start, searchEnd)

    Do While matchRange.Start < searchEnd
        With matchRange.Find
            .ClearFormatting
            If softToPara Then
                .Text = "^l"
            Else
                .Text = "^p"
            End If
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = False
            If Not .Execute Then Exit Do
        End With

        ' A terminal selected paragraph mark belongs to the boundary with
        ' unselected content. Converting it would merge and reformat the
        ' paragraph after the selection.
        If (Not softToPara) And matchRange.End >= searchEnd Then Exit Do

        If softToPara Then
            matchRange.Text = Chr$(13)
        Else
            If matchRange.End >= ActiveDocument.Content.End Then Exit Do
            matchRange.Text = Chr$(11)
        End If
        ReplaceSoftReturnMarks = ReplaceSoftReturnMarks + 1
        nextStart = matchRange.End
        If nextStart >= searchEnd Then Exit Do
        Set matchRange = ActiveDocument.Range(nextStart, searchEnd)
    Loop
End Function
