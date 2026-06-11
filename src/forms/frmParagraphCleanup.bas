Option Explicit
' --------------------------------------------------------------
'  Paragraph Structure Fixer
'  Cleans up paragraph structure: removes empty paragraphs, collapses
'  runs of three or more paragraph breaks down to two, resets
'  SpaceBefore/SpaceAfter to zero for paragraphs with direct spacing
'  overrides, and resets left indent to zero.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optAll.GroupName = "ParaMode"
    optRemoveEmpty.GroupName = "ParaMode"
    optNormalizeSpacing.GroupName = "ParaMode"
    optCustom.GroupName = "ParaMode"
    optScopeDocument.GroupName = "ParaScope"
    optScopeSelection.GroupName = "ParaScope"
    optAll.Value = True: UpdateCustomVisibility: ClearCustomChecks
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
    optAll.Caption = "All paragraph fixes"
    optRemoveEmpty.Caption = "Remove empty paragraphs"
    optNormalizeSpacing.Caption = "Normalize paragraph spacing"
    optCustom.Caption = "Custom (choose below)"
    chkRemoveEmpty.Caption = "Collapse multiple consecutive empty paragraphs"
    chkCollapseBreaks.Caption = "Convert soft returns (Shift+Enter) to paragraph marks"
    chkNormalizeParaSpacing.Caption = "Normalize Space Before / Space After settings"
    chkFixIndent.Caption = "Remove manual tab / space indentation at paragraph start"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Paragraph Structure Fixer" & vbCrLf & vbCrLf
    h = h & "Cleans up paragraph structure and direct spacing overrides." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Remove empty paragraphs  --  Deletes completely empty paragraphs." & vbCrLf
    h = h & "  Collapse extra breaks    --  Reduces runs of 3+ consecutive paragraph breaks" & vbCrLf
    h = h & "                               to exactly 2 (a standard visual separator)." & vbCrLf
    h = h & "  Normalise para spacing   --  Resets SpaceBefore and SpaceAfter to zero for" & vbCrLf
    h = h & "                               paragraphs that have direct spacing overrides." & vbCrLf
    h = h & "                               Does not affect spacing set by the style itself." & vbCrLf
    h = h & "  Fix indent               --  Resets left indent to zero for paragraphs with" & vbCrLf
    h = h & "                               direct indent overrides." & vbCrLf
    h = h & vbCrLf
    h = h & "TIP: Run after assembling a document from multiple sources or templates, " & _
            "where inconsistent spacing between paragraphs is most common." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights empty paragraphs before any changes are made." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Paragraph Structure Fixer"
End Sub
Private Sub UpdateCustomVisibility(): fraCustom.Visible = optCustom.Value: End Sub
Private Sub ClearCustomChecks(): chkRemoveEmpty.Value = False: chkCollapseBreaks.Value = False: chkNormalizeParaSpacing.Value = False: chkFixIndent.Value = False: End Sub
Private Sub cmdSelectAll_Click()
    If fraCustom.Visible Then
        chkRemoveEmpty.Value = True
        chkCollapseBreaks.Value = True
        chkNormalizeParaSpacing.Value = True
        chkFixIndent.Value = True
    End If
End Sub
Private Sub cmdDeselectAll_Click()
    If fraCustom.Visible Then
        ClearCustomChecks
    End If
End Sub
Public Sub RunAfterPreview()
    chkPreviewOnly.Value = False
    cmdRun_Click
End Sub
Private Sub cmdRun_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected. Please remove protection before running cleanup.", vbExclamation, "Document Protected": Exit Sub
    End If
    If Not GuardBeforeCleanup("Paragraph Fixer") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doRemove As Boolean, doCollapse As Boolean, doNormalize As Boolean, doIndent As Boolean
    If optAll.Value Then doRemove = True: doCollapse = True: doNormalize = True: doIndent = True
    If optRemoveEmpty.Value Then doRemove = True
    If optNormalizeSpacing.Value Then doNormalize = True
    If optCustom.Value Then doRemove = chkRemoveEmpty.Value: doCollapse = chkCollapseBreaks.Value: doNormalize = chkNormalizeParaSpacing.Value: doIndent = chkFixIndent.Value
    If Not (doRemove Or doCollapse Or doNormalize Or doIndent) Then MsgBox "No paragraph options selected.", vbInformation: Exit Sub
    ' Preview: highlight empty paragraphs or triple breaks
    Dim p As Paragraph, cnt As Long: cnt = 0
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
            If (doRemove Or doCollapse) And Len(Trim$(p.Range.Text)) <= 1 Then p.Range.HighlightColorIndex = wdYellow: cnt = cnt + 1
        End If
    Next p
    If previewOnly Then
        ShowPreviewActions Me, "Paragraph Fixer", "Preview complete. " & cnt & " items highlighted."
        Exit Sub
    End If
    MarkCleanupStart "Paragraph Fixer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Paragraph Fixer"
    On Error GoTo RunErr
    ' Remove empty paragraphs
    Dim cntRemove As Long, cntCollapse As Long, cntNormalize As Long, cntIndent As Long
    If doRemove Then
        Dim i As Long
        For i = ActiveDocument.Paragraphs.Count To 1 Step -1
            Set p = ActiveDocument.Paragraphs(i)
            If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
                If Len(Trim$(p.Range.Text)) <= 1 Then p.Range.Delete: cntRemove = cntRemove + 1
            End If
        Next i
    End If
    ' Collapse extra breaks (3+ to 2)
    If doCollapse Then
        Do
            With targetRange.Find
                .ClearFormatting: .Replacement.ClearFormatting: .Text = "^p^p^p": .Replacement.Text = "^p^p"
                Do While .Execute(Replace:=wdReplaceOne): cntCollapse = cntCollapse + 1: Loop
            End With
        Loop While targetRange.Find.Execute(FindText:="^p^p^p", Forward:=True)
    End If
    ' Normalize paragraph spacing
    If doNormalize Then
        For Each p In ActiveDocument.Paragraphs
            If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
                With p.Format
                    If .SpaceBefore <> 0 Or .SpaceAfter <> 0 Then
                        .SpaceBefore = 0: .SpaceAfter = 0: cntNormalize = cntNormalize + 1
                    End If
                End With
            End If
        Next p
    End If
    ' Fix indent
    If doIndent Then
        For Each p In ActiveDocument.Paragraphs
            If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
                If p.LeftIndent <> 0 Then p.LeftIndent = 0: cntIndent = cntIndent + 1
            End If
        Next p
    End If
    RemoveAllHighlighting targetRange
    Dim results As Collection: Set results = New Collection
    If doRemove    Then results.Add "Empty paragraphs removed:        " & cntRemove
    If doCollapse  Then results.Add "Extra paragraph breaks collapsed: " & cntCollapse
    If doNormalize Then results.Add "Paragraph spacings normalized:    " & cntNormalize
    If doIndent    Then results.Add "Indents reset to zero:            " & cntIndent
    ShowCleanupReport "Paragraph Cleanup", results
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
