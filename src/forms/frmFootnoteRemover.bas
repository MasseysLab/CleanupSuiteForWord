Option Explicit
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
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Footnote / Endnote Remover" & vbCrLf & vbCrLf
    h = h & "Deletes footnotes and/or endnotes -- both the reference mark in the body " & _
            "and the note text -- to produce a clean text extract." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Remove footnotes  --  deletes footnotes in the target scope." & vbCrLf
    h = h & "  Remove endnotes   --  deletes endnotes in the target scope." & vbCrLf
    h = h & "  Keep note text inline  --  inserts each note's text at the reference" & vbCrLf
    h = h & "                            point in [square brackets] before deleting it." & vbCrLf
    h = h & vbCrLf
    h = h & "Word renumbers any remaining notes automatically." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text first to remove only the notes referenced in the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights the reference marks that would be removed." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Footnote / Endnote Remover"
End Sub
Private Function InScope(r As Range, targetRange As Range) As Boolean
    InScope = (r.Start >= targetRange.Start And r.Start <= targetRange.End)
End Function
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
        MsgBox "Preview complete." & vbCrLf & "Footnotes to remove: " & pf & vbCrLf & "Endnotes to remove: " & pe, vbInformation, "Footnote / Endnote Preview"
        Unload Me: Exit Sub
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
                If keepText Then fn.Reference.InsertAfter " [" & Trim$(fn.Range.Text) & "]"
                fn.Delete
                cntFoot = cntFoot + 1
            End If
        Next i
    End If
    If doEnd Then
        For i = ActiveDocument.Endnotes.Count To 1 Step -1
            Dim en As Endnote: Set en = ActiveDocument.Endnotes(i)
            If InScope(en.Reference, targetRange) Then
                If keepText Then en.Reference.InsertAfter " [" & Trim$(en.Range.Text) & "]"
                en.Delete
                cntEnd = cntEnd + 1
            End If
        Next i
    End If
    Dim results As Collection: Set results = New Collection
    If doFoot Then results.Add "Footnotes removed: " & cntFoot
    If doEnd Then results.Add "Endnotes removed: " & cntEnd
    If keepText Then results.Add "Note text kept inline in [brackets]"
    ShowCleanupReport "Footnote / Endnote Remover", results
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
