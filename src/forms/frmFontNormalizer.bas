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
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
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
    ' Preview: highlight paragraphs with direct font overrides
    Dim p As Paragraph, styleFont As Font, cnt As Long: cnt = 0
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
            On Error Resume Next
            Set styleFont = ActiveDocument.Styles(p.Style).Font
            On Error GoTo 0
            If Not styleFont Is Nothing Then
                Dim hasDirect As Boolean: hasDirect = False
                If doFace And p.Range.Font.Name <> styleFont.Name Then hasDirect = True
                If doSize And p.Range.Font.Size <> styleFont.Size And p.Range.Font.Size > 0 Then hasDirect = True
                If hasDirect Then p.Range.HighlightColorIndex = wdYellow: cnt = cnt + 1
            End If
        End If
    Next p
    If previewOnly Then
        ShowPreviewActions Me, "Font Normalizer", "Preview complete. " & cnt & " paragraphs with direct font overrides highlighted."
        Exit Sub
    End If
    MarkCleanupStart "Font Normalizer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Font Normalizer"
    On Error GoTo RunErr
    Dim cntFixed As Long: cntFixed = 0
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
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
    Dim results As Collection: Set results = New Collection
    results.Add "Paragraphs normalized: " & cntFixed
    ShowCleanupReport "Font Normalizer", results
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
