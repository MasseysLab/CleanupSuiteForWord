Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Header / Footer Standardizer
'  Standardizes header/footer text across every section and slot
'  (primary, first-page, even-page), or clears all header/footer content.
'  Standardize can reset the font to the document default, remove paragraph
'  spacing, and set a chosen alignment.  Page numbers and fields are left
'  untouched.  Always operates on the whole document.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optStandardize.GroupName = "HFMode"
    optClearAll.GroupName = "HFMode"
    optAlignLeft.GroupName = "HFAlign"
    optAlignCenter.GroupName = "HFAlign"
    optAlignRight.GroupName = "HFAlign"
    optStandardize.Value = True
    chkHeaders.Value = True
    chkFooters.Value = True
    chkFont.Value = True
    chkSpacing.Value = True
    chkAlignment.Value = False
    optAlignLeft.Value = True
    chkBreakLinks.Value = False
    optStandardize.Caption = "Standardize formatting"
    optClearAll.Caption = "Clear all header/footer content"
    chkHeaders.Caption = "Include headers"
    chkFooters.Caption = "Include footers"
    chkFont.Caption = "Reset font to the document default"
    chkSpacing.Caption = "Remove paragraph spacing"
    chkAlignment.Caption = "Set alignment"
    optAlignLeft.Caption = "Left"
    optAlignCenter.Caption = "Center"
    optAlignRight.Caption = "Right"
    chkBreakLinks.Caption = "Unlink sections (set each independently)"
    fraScopeSelection.Enabled = False
    optScopeDocument.Value = True
    optScopeDocument.Caption = "Entire document (always)"
    optScopeSelection.Enabled = False
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub optStandardize_Click()
    LayoutCleanupToolForm Me
End Sub
Private Sub optClearAll_Click()
    LayoutCleanupToolForm Me
End Sub
Private Sub chkAlignment_Click()
    LayoutCleanupToolForm Me
End Sub
Private Sub ApplyToHF(hf As HeaderFooter, doClear As Boolean, doFont As Boolean, doSpacing As Boolean, doAlign As Boolean, alignVal As Long, baseName As String, baseSize As Single, breakLinks As Boolean, ByRef cntAreas As Long, ByRef cntCleared As Long)
    On Error Resume Next
    If breakLinks Then hf.LinkToPrevious = False
    If doClear Then
        If Len(Trim$(hf.Range.Text)) > 0 Then cntCleared = cntCleared + 1
        hf.Range.Delete
    Else
        If doFont Then
            hf.Range.Font.Name = baseName
            hf.Range.Font.Size = baseSize
        End If
        If doSpacing Then
            hf.Range.ParagraphFormat.SpaceBefore = 0
            hf.Range.ParagraphFormat.SpaceAfter = 0
        End If
        If doAlign Then
            hf.Range.ParagraphFormat.Alignment = alignVal
        End If
    End If
    cntAreas = cntAreas + 1
End Sub
Private Sub chkHeaders_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkFooters_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkFont_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkSpacing_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkBreakLinks_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub optAlignLeft_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub optAlignCenter_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub optAlignRight_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub cmdPreview_Click()
    PreviewFromPanel
End Sub
Public Sub PreviewFromPanel()
    chkPreviewOnly.Value = True
    cmdRun_Click
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
    If Not GuardBeforeCleanup("Header / Footer Standardizer") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim doClear As Boolean: doClear = optClearAll.Value
    Dim doHeaders As Boolean: doHeaders = chkHeaders.Value
    Dim doFooters As Boolean: doFooters = chkFooters.Value
    Dim doFont As Boolean: doFont = chkFont.Value
    Dim doSpacing As Boolean: doSpacing = chkSpacing.Value
    Dim doAlign As Boolean: doAlign = chkAlignment.Value
    Dim breakLinks As Boolean: breakLinks = chkBreakLinks.Value
    If Not (doHeaders Or doFooters) Then MsgBox "Select headers, footers, or both.", vbInformation: Exit Sub
    If Not doClear And Not (doFont Or doSpacing Or doAlign) Then MsgBox "Select at least one formatting option to standardize.", vbInformation: Exit Sub
    Dim alignVal As Long: alignVal = wdAlignParagraphLeft
    If optAlignCenter.Value Then alignVal = wdAlignParagraphCenter
    If optAlignRight.Value Then alignVal = wdAlignParagraphRight
    Dim baseName As String: baseName = ActiveDocument.Styles(wdStyleNormal).Font.Name
    Dim baseSize As Single: baseSize = ActiveDocument.Styles(wdStyleNormal).Font.Size
    Dim sec As Section, hf As HeaderFooter
    If previewOnly Then
        Dim pc As Long: pc = 0
        For Each sec In ActiveDocument.Sections
            If doHeaders Then
                For Each hf In sec.Headers
                    pc = pc + 1
                Next hf
            End If
            If doFooters Then
                For Each hf In sec.Footers
                    pc = pc + 1
                Next hf
            End If
        Next sec
        Dim verb As String: verb = "standardized"
        If doClear Then verb = "cleared"
        ShowPreviewActions Me, "Header / Footer Standardizer", "Preview complete." & vbCrLf & pc & " header/footer areas across " & ActiveDocument.Sections.Count & " sections would be " & verb & "."
        Exit Sub
    End If
    MarkCleanupStart "Header / Footer Standardizer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Header / Footer Standardizer"
    On Error GoTo RunErr
    Dim cntAreas As Long, cntCleared As Long
    For Each sec In ActiveDocument.Sections
        If doHeaders Then
            For Each hf In sec.Headers
                ApplyToHF hf, doClear, doFont, doSpacing, doAlign, alignVal, baseName, baseSize, breakLinks, cntAreas, cntCleared
            Next hf
        End If
        If doFooters Then
            For Each hf In sec.Footers
                ApplyToHF hf, doClear, doFont, doSpacing, doAlign, alignVal, baseName, baseSize, breakLinks, cntAreas, cntCleared
            Next hf
        End If
    Next sec
    Dim results As Collection: Set results = New Collection
    results.Add "Header/footer areas processed: " & cntAreas
    If doClear Then results.Add "Areas that had content cleared: " & cntCleared
    ShowCleanupReport "Header / Footer Standardizer", results
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
