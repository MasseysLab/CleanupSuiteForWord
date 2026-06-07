Option Explicit
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
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Font Normalizer" & vbCrLf & vbCrLf
    h = h & "Resets direct font overrides to match the style defined for each paragraph," & vbCrLf
    h = h & "eliminating the font chaos that accumulates from copying and pasting text" & vbCrLf
    h = h & "between documents, emails, and web pages." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS  (choose which properties to reset)" & vbCrLf
    h = h & "  Font face   --  Resets the typeface (e.g. Arial, Times New Roman) to the" & vbCrLf
    h = h & "                  style's defined font." & vbCrLf
    h = h & "  Font size   --  Resets the point size to the style's defined size." & vbCrLf
    h = h & "  Bold        --  Resets bold to the style's setting." & vbCrLf
    h = h & "  Italic      --  Resets italic to the style's setting." & vbCrLf
    h = h & "  Font colour --  Resets text colour to the style's defined colour." & vbCrLf
    h = h & vbCrLf
    h = h & "TIP: Start with Face and Size only.  Bold, Italic, and Colour are opt-in" & vbCrLf
    h = h & "because those properties are more likely to be intentional (e.g. a bolded" & vbCrLf
    h = h & "phrase or a deliberately coloured heading within a paragraph)." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights paragraphs with direct font overrides." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Font Normalizer"
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
    MsgBox "Preview complete. " & cnt & " paragraphs with direct font overrides highlighted.", vbInformation
    If previewOnly Then Unload Me: Exit Sub
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
    Unload Me
    Exit Sub
RunErr:
    On Error Resume Next: undoRec.EndCustomRecord: On Error GoTo 0
    MarkCleanupEnd
    Dim errMsg As String
    errMsg = "An unexpected error occurred: " & Err.Number & " - " & Err.Description
    errMsg = errMsg & vbCrLf & vbCrLf & "Some changes may have been made to the document." & vbCrLf & "Would you like to undo all changes made before the error?"
    If MsgBox(errMsg, vbCritical + vbYesNo, "Cleanup Error") = vbYes Then
        On Error Resume Next: Application.Undo: On Error GoTo 0
    End If
End Sub