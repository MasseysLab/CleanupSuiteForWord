Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then ReturnToMainAfterToolClose
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
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
    optSoftToPara.Caption = "Convert soft returns (Shift+Enter) to paragraph marks"
    optParaToSoft.Caption = "Convert paragraph marks to soft returns"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Soft Return Converter" & vbCrLf & vbCrLf
    h = h & "Converts soft returns (line breaks made with Shift+Enter) to real " & _
            "paragraph marks, or the reverse.  Documents imported from PDFs and web" & vbCrLf
    h = h & "pages are often full of soft returns that look like paragraph breaks but " & _
            "are not -- which quietly breaks spell-check, styles, and spacing." & vbCrLf
    h = h & vbCrLf
    h = h & "DIRECTION" & vbCrLf
    h = h & "  Soft returns to paragraphs  --  the usual fix for imported text." & vbCrLf
    h = h & "  Paragraphs to soft returns  --  the reverse (less common)." & vbCrLf
    h = h & vbCrLf
    h = h & "CAUTION: Converting paragraphs to soft returns merges the affected text " & _
            "into a single paragraph.  Use the reverse only on small selections." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights the breaks that would be converted." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Soft Return Converter"
End Sub
Private Sub cmdPreview_Click()
    PreviewFromPanel
End Sub
Public Sub PreviewFromPanel()
    chkPreviewOnly.Value = True
    cmdRun_Click
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
    Dim findText As String, replText As String, markChar As String
    If softToPara Then
        findText = "^l": replText = "^p": markChar = Chr(11)
    Else
        findText = "^p": replText = "^l": markChar = Chr(13)
    End If
    If previewOnly Then
        With targetRange.Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = findText: .Replacement.Text = "^&": .Replacement.Highlight = True
            .Forward = True: .Wrap = wdFindStop
            .Execute Replace:=wdReplaceAll
        End With
        Dim pt As String: pt = targetRange.Text
        Dim pc As Long: pc = Len(pt) - Len(Replace(pt, markChar, ""))
        ShowPreviewActions Me, "Soft Return Converter", "Preview complete. " & pc & " breaks highlighted."
        Exit Sub
    End If
    MarkCleanupStart "Soft Return Converter"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Soft Return Converter"
    On Error GoTo RunErr
    Dim txt As String: txt = targetRange.Text
    Dim cnt As Long: cnt = Len(txt) - Len(Replace(txt, markChar, ""))
    With targetRange.Find
        .ClearFormatting: .Replacement.ClearFormatting
        .Text = findText: .Replacement.Text = replText
        .Forward = True: .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
    Dim results As Collection: Set results = New Collection
    If softToPara Then
        results.Add "Soft returns converted to paragraphs: " & cnt
    Else
        results.Add "Paragraphs converted to soft returns: " & cnt
    End If
    ShowCleanupReport "Soft Return Converter", results
    undoRec.EndCustomRecord
    MarkCleanupEnd
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
