Option Explicit
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
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Document Trim" & vbCrLf & vbCrLf
    h = h & "Removes trailing empty paragraphs from the end of the document, " & _
            "leaving exactly one final empty paragraph (which Word requires to exist)." & vbCrLf
    h = h & vbCrLf
    h = h & "WHY THIS MATTERS" & vbCrLf
    h = h & "  Documents assembled from multiple files, or saved from templates, " & _
            "  often end with many blank paragraphs.  These appear as unwanted blank" & vbCrLf
    h = h & "  space at the end of the printed document and can cause unexpected " & _
            "  behaviour with headers, footers, and page numbering in some layouts." & vbCrLf
    h = h & vbCrLf
    h = h & "NOTE: This tool always operates on the full document.  Scope selection " & _
            "is disabled because 'end of document' is inherently document-level --" & vbCrLf
    h = h & "trimming to the end of a selection would not be meaningful." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights the trailing paragraphs that would be removed." & vbCrLf
    h = h & "Run again without Preview to apply." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Document Trim"
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
    ' Find trailing empty paragraphs (walk backwards from end)
    Dim totalParas As Long: totalParas = ActiveDocument.Paragraphs.Count
    Dim trailingCount As Long: trailingCount = 0
    Dim i As Long
    For i = totalParas To 1 Step -1
        Dim p As Paragraph: Set p = ActiveDocument.Paragraphs(i)
        If Len(Trim$(p.Range.Text)) <= 1 Then
            trailingCount = trailingCount + 1
        Else
            Exit For
        End If
    Next i
    ' Always leave at least one paragraph at the end
    Dim toRemove As Long: toRemove = trailingCount - 1
    If toRemove <= 0 Then MsgBox "No trailing empty paragraphs found.", vbInformation: Unload Me: Exit Sub
    ' Highlight the ones that will be removed
    For i = totalParas To totalParas - toRemove + 1 Step -1
        ActiveDocument.Paragraphs(i).Range.HighlightColorIndex = wdYellow
    Next i
    If previewOnly Then
        ShowPreviewActions Me, "Document Trim", "Preview complete. " & toRemove & " trailing empty paragraphs highlighted."
        Exit Sub
    End If
    MarkCleanupStart "Document Trim"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Document Trim"
    On Error GoTo RunErr
    Dim removed As Long: removed = 0
    Dim currentTotal As Long: currentTotal = ActiveDocument.Paragraphs.Count
    For i = currentTotal To currentTotal - toRemove + 1 Step -1
        ActiveDocument.Paragraphs(i).Range.Delete
        removed = removed + 1
    Next i
    Dim results As Collection: Set results = New Collection
    results.Add "Trailing empty paragraphs removed: " & removed
    ShowCleanupReport "Document Trim", results
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
