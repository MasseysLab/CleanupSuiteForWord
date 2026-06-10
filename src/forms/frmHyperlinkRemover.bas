Option Explicit
' --------------------------------------------------------------
'  Hyperlink Remover
'  Removes hyperlinks but keeps the visible text.  Optionally also clears
'  the blue underlined hyperlink formatting, returning the text to its
'  normal style appearance.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    chkRemoveFormat.Value = True
    optScopeDocument.GroupName = "HLScope"
    optScopeSelection.GroupName = "HLScope"
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
    chkRemoveFormat.Caption = "Remove hyperlink character style (blue underline)"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Hyperlink Remover" & vbCrLf & vbCrLf
    h = h & "Removes hyperlinks while keeping the visible text.  Pasting from web" & vbCrLf
    h = h & "pages embeds links throughout a document; this strips the links but" & vbCrLf
    h = h & "leaves the words exactly as they are." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Clear hyperlink formatting  --  also removes the blue underlined look" & vbCrLf
    h = h & "                                  (resets to the default character font)." & vbCrLf
    h = h & "                                  Turn off to keep the blue/underline." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights every hyperlink that would be removed." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Hyperlink Remover"
End Sub
Private Sub cmdRun_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected. Please remove protection before running cleanup.", vbExclamation, "Document Protected": Exit Sub
    End If
    If Not GuardBeforeCleanup("Hyperlink Remover") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doRemoveFormat As Boolean: doRemoveFormat = chkRemoveFormat.Value
    Dim i As Long, cnt As Long: cnt = 0
    If previewOnly Then
        For i = 1 To ActiveDocument.Hyperlinks.Count
            Dim hp As Hyperlink: Set hp = ActiveDocument.Hyperlinks(i)
            If hp.Range.Start >= targetRange.Start And hp.Range.End <= targetRange.End Then hp.Range.HighlightColorIndex = wdYellow: cnt = cnt + 1
        Next i
        MsgBox "Preview complete. " & cnt & " hyperlinks highlighted.", vbInformation
        Unload Me: Exit Sub
    End If
    MarkCleanupStart "Hyperlink Remover"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Hyperlink Remover"
    On Error GoTo RunErr
    For i = ActiveDocument.Hyperlinks.Count To 1 Step -1
        Dim hl As Hyperlink: Set hl = ActiveDocument.Hyperlinks(i)
        If hl.Range.Start >= targetRange.Start And hl.Range.End <= targetRange.End Then
            Dim rs As Long, rEnd As Long
            rs = hl.Range.Start: rEnd = hl.Range.End
            hl.Delete
            cnt = cnt + 1
            If doRemoveFormat Then
                Dim rg As Range: Set rg = ActiveDocument.Range(rs, rEnd)
                rg.Style = ActiveDocument.Styles(wdStyleDefaultParagraphFont)
                rg.Font.Underline = wdUnderlineNone
                rg.Font.ColorIndex = wdAuto
            End If
        End If
    Next i
    Dim results As Collection: Set results = New Collection
    results.Add "Hyperlinks removed: " & cnt
    If doRemoveFormat Then results.Add "Hyperlink formatting cleared"
    ShowCleanupReport "Hyperlink Remover", results
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
