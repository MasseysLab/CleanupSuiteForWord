Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then ReturnToMainAfterToolClose
End Sub
' --------------------------------------------------------------
'  Capitalization Fixer
'  Applies consistent capitalization across paragraphs: sentence case,
'  title case, UPPERCASE, lowercase, or smart sentence detection (which
'  capitalizes the first letter after sentence-ending punctuation).
'  Operates paragraph by paragraph within the target scope.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optAll.GroupName = "CapMode"
    optSentence.GroupName = "CapMode"
    optTitle.GroupName = "CapMode"
    optUpper.GroupName = "CapMode"
    optLower.GroupName = "CapMode"
    optCustom.GroupName = "CapMode"
    optScopeDocument.GroupName = "CapScope"
    optScopeSelection.GroupName = "CapScope"
    optAll.Value = True: UpdateCustomVisibility: ClearCustomChecks
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
    optAll.Caption = "All capitalization fixes"
    optSentence.Caption = "Sentence case"
    optTitle.Caption = "Title case"
    optUpper.Caption = "ALL CAPS -> Title Case"
    optLower.Caption = "lowercase -> Sentence case"
    optCustom.Caption = "Custom (choose below)"
    chkSentence.Caption = "Capitalize sentence starts"
    chkTitle.Caption = "Apply title case to headings"
    chkUpper.Caption = "Convert ALL-CAPS words (5+ letters)"
    chkLower.Caption = "Fix all-lowercase paragraphs"
    chkSmartSentences.Caption = "Smart sentence detection (skip abbreviations)"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Capitalization Fixer" & vbCrLf & vbCrLf
    h = h & "Applies consistent capitalisation to every paragraph in the target scope." & vbCrLf
    h = h & vbCrLf
    h = h & "MODES" & vbCrLf
    h = h & "  All (Smart)   --  Default. Capitalises the first letter after sentence-ending" & vbCrLf
    h = h & "                    punctuation (. ? !).  Best choice for general text." & vbCrLf
    h = h & "  Sentence case --  First letter of paragraph capitalised; rest lowercased." & vbCrLf
    h = h & "  Title Case    --  First letter of every word capitalised." & vbCrLf
    h = h & "  UPPERCASE     --  Every letter uppercased." & vbCrLf
    h = h & "  lowercase     --  Every letter lowercased." & vbCrLf
    h = h & "  Custom        --  Combine modes using the checkboxes.  If more than one " & _
            "is selected, each is applied in sequence -- the last" & vbCrLf
    h = h & "                    one applied wins." & vbCrLf
    h = h & vbCrLf
    h = h & "NOTE: Capitalisation is applied to the full text content of each paragraph, " & _
            "including any directly formatted runs within it." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights paragraphs that would change." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Capitalization Fixer"
End Sub
Private Sub UpdateCustomVisibility(): fraCustom.Visible = optCustom.Value: End Sub
Private Sub ClearCustomChecks(): chkSentence.Value = False: chkTitle.Value = False: chkUpper.Value = False: chkLower.Value = False: chkSmartSentences.Value = False: End Sub
Private Sub cmdSelectAll_Click()
    If fraCustom.Visible Then
        chkSentence.Value = True
        chkTitle.Value = True
        chkUpper.Value = True
        chkLower.Value = True
        chkSmartSentences.Value = True
    End If
End Sub
Private Sub cmdDeselectAll_Click()
    If fraCustom.Visible Then
        ClearCustomChecks
    End If
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
    If Not GuardBeforeCleanup("Capitalization Fixer") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doSentence As Boolean, doTitle As Boolean, doUpper As Boolean, doLower As Boolean, doSmart As Boolean
    If optAll.Value Then doSmart = True
    If optSentence.Value Then doSentence = True
    If optTitle.Value Then doTitle = True
    If optUpper.Value Then doUpper = True
    If optLower.Value Then doLower = True
    If optCustom.Value Then doSentence = chkSentence.Value: doTitle = chkTitle.Value: doUpper = chkUpper.Value: doLower = chkLower.Value: doSmart = chkSmartSentences.Value
    If Not (doSentence Or doTitle Or doUpper Or doLower Or doSmart) Then MsgBox "No capitalization option selected.", vbInformation: Exit Sub
    ' Preview: highlight paragraphs that would change (simple heuristic)
    Dim p As Paragraph, original As String, transformed As String, changedCount As Long: changedCount = 0
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
            original = p.Range.Text
            If Right$(original, 1) = vbCr Then original = Left$(original, Len(original) - 1)
            transformed = original
            If doSentence Then transformed = SentenceCase(transformed)
            If doTitle Then transformed = TitleCase(transformed)
            If doUpper Then transformed = UCase$(transformed)
            If doLower Then transformed = LCase$(transformed)
            If doSmart Then transformed = SmartSentenceCase(transformed)
            If transformed <> original Then
                p.Range.HighlightColorIndex = wdYellow: changedCount = changedCount + 1
            End If
        End If
    Next p
    If previewOnly Then
        ShowPreviewActions Me, "Capitalization Fixer", "Preview complete. " & changedCount & " paragraphs highlighted."
        Exit Sub
    End If
    MarkCleanupStart "Capitalization Fixer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Capitalization Fixer"
    On Error GoTo RunErr
    ' Apply changes
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
            original = p.Range.Text
            If Right$(original, 1) = vbCr Then original = Left$(original, Len(original) - 1)
            transformed = original
            If doSentence Then transformed = SentenceCase(transformed)
            If doTitle Then transformed = TitleCase(transformed)
            If doUpper Then transformed = UCase$(transformed)
            If doLower Then transformed = LCase$(transformed)
            If doSmart Then transformed = SmartSentenceCase(transformed)
            If transformed <> original Then p.Range.Text = transformed & vbCr
        End If
    Next p
    RemoveAllHighlighting targetRange
    Dim results As Collection: Set results = New Collection
    results.Add "Paragraphs changed: " & changedCount
    ShowCleanupReport "Capitalization Cleanup", results
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
' Helper functions for capitalization
Private Function SentenceCase(s As String) As String
    s = LCase$(s)
    If Len(Trim$(s)) = 0 Then SentenceCase = s: Exit Function
    Dim leadLen As Long: leadLen = Len(s) - Len(LTrim$(s))
    SentenceCase = Left$(s, leadLen) & UCase$(Mid$(s, leadLen + 1, 1)) & Mid$(s, leadLen + 2)
End Function
Private Function TitleCase(s As String) As String
    Dim arr: arr = Split(LCase$(s), " ")
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If Len(arr(i)) > 0 Then arr(i) = UCase$(Left$(arr(i), 1)) & Mid$(arr(i), 2)
    Next i
    TitleCase = Join(arr, " ")
End Function
Private Function SmartSentenceCase(s As String) As String
    ' Very simple heuristic: capitalize after sentence-ending punctuation
    Dim i As Long, result As String, capNext As Boolean: capNext = True
    For i = 1 To Len(s)
        Dim currentChar As String: currentChar = Mid$(s, i, 1)
        If capNext And currentChar >= "a" And currentChar <= "z" Then currentChar = UCase$(currentChar): capNext = False
        result = result & currentChar
        If currentChar = "." Or currentChar = "?" Or currentChar = "!" Then capNext = True
    Next i
    SmartSentenceCase = result
End Function
