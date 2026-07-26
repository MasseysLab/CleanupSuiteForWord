Option Explicit
Private mHybridAnalysis As Object
Private mHybridDocument As Document
Private mHybridScopeStart As Long
Private mHybridScopeEnd As Long

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set mHybridAnalysis = Nothing
    Set mHybridDocument = Nothing
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Invisible Unicode Cleaner
'  Finds and removes invisible Unicode characters that accumulate from
'  web copying, PDF conversion, and cross-platform editing: non-breaking
'  spaces, zero-width joiners/non-joiners, byte-order marks, soft hyphens,
'  and non-breaking hyphens.  Invisible to the eye but harmful to search,
'  sort, and automated processing.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optAll.GroupName = "UniMode"
    optNBSP.GroupName = "UniMode"
    optZeroWidth.GroupName = "UniMode"
    optCustom.GroupName = "UniMode"
    optScopeDocument.GroupName = "UniScope"
    optScopeSelection.GroupName = "UniScope"
    optAll.Value = True
    UpdateCustomVisibility
    ClearCustomChecks
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    RefreshScopeSelectionState Me, True
    optAll.Caption = "All invisible / problem characters"
    optNBSP.Caption = "Non-breaking spaces only"
    optZeroWidth.Caption = "Zero-width characters only"
    optCustom.Caption = "Custom"
    chkNBSP.Caption = "Non-breaking spaces (U+00A0)"
    chkZWSP.Caption = "Zero-width spaces (U+200B)"
    chkZWNJ.Caption = "Zero-width non-joiners (U+200C)"
    chkZWJ.Caption = "Zero-width joiners (U+200D)"
    chkBOM.Caption = "Byte order marks (U+FEFF)"
    chkSoftHyphen.Caption = "Soft hyphens (U+00AD)"
    chkNBHyphen.Caption = "Non-breaking hyphens (U+2011)"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutUnicodeForm
End Sub
Private Sub LayoutUnicodeForm()
    LayoutCleanupToolForm Me
End Sub
Private Sub UpdateCustomVisibility()
    fraCustom.Visible = False
End Sub
Private Sub optAll_Click(): UpdateCustomVisibility: LayoutUnicodeForm: End Sub
Private Sub optNBSP_Click(): UpdateCustomVisibility: LayoutUnicodeForm: End Sub
Private Sub optZeroWidth_Click(): UpdateCustomVisibility: LayoutUnicodeForm: End Sub
Private Sub optCustom_Click(): UpdateCustomVisibility: LayoutUnicodeForm: End Sub
Private Sub cmdRiskPlacementUnicodeAll_Click(): ShowToolRiskChoiceExplanation "All invisible / problem characters", "Safe cleanup", "Sweeps the common invisible characters that interfere with search, wrapping, copying, and export.": End Sub
Private Sub cmdRiskPlacementUnicodeNBSP_Click(): ShowToolRiskChoiceExplanation "Non-breaking spaces only", "Safe cleanup", "Replaces non-breaking spaces with regular spaces while leaving other hidden characters alone.": End Sub
Private Sub cmdRiskPlacementUnicodeZeroWidth_Click(): ShowToolRiskChoiceExplanation "Zero-width characters only", "Safe cleanup", "Removes hidden zero-width characters that can split words and confuse search or cleanup.": End Sub
Private Sub cmdRiskPlacementUnicodeCustom_Click(): ShowToolRiskChoiceExplanation "Custom Unicode cleanup", "Inspect first", "Lets you combine individual invisible-character repairs. Preview first when you are not sure which hidden characters are present.": End Sub
Private Sub ClearCustomChecks()
    chkNBSP.Value = False: chkZWSP.Value = False: chkZWNJ.Value = False: chkZWJ.Value = False
    chkBOM.Value = False: chkSoftHyphen.Value = False: chkNBHyphen.Value = False
End Sub
Private Sub cmdSelectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    chkNBSP.Value = True
    chkZWSP.Value = True
    chkZWNJ.Value = True
    chkZWJ.Value = True
    chkBOM.Value = True
    chkSoftHyphen.Value = True
    chkNBHyphen.Value = True
    LayoutUnicodeForm
End Sub
Private Sub cmdDeselectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    ClearCustomChecks
    LayoutUnicodeForm
End Sub
Private Function UnicodeChar(ByVal codePoint As Long) As String
    If codePoint > &H7FFF Then
        UnicodeChar = ChrW$(codePoint - &H10000)
    Else
        UnicodeChar = ChrW$(codePoint)
    End If
End Function
Private Sub chkNBSP_Click(): LayoutUnicodeForm: End Sub
Private Sub chkZWSP_Click(): LayoutUnicodeForm: End Sub
Private Sub chkZWNJ_Click(): LayoutUnicodeForm: End Sub
Private Sub chkZWJ_Click(): LayoutUnicodeForm: End Sub
Private Sub chkBOM_Click(): LayoutUnicodeForm: End Sub
Private Sub chkSoftHyphen_Click(): LayoutUnicodeForm: End Sub
Private Sub chkNBHyphen_Click(): LayoutUnicodeForm: End Sub
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
    Set mHybridAnalysis = Nothing
    Set mHybridDocument = Nothing
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
    If Not GuardBeforeCleanup("Unicode Cleaner") Then Unload Me: Exit Sub
    Dim findList As Collection, replaceList As Collection, itemLabels As Collection
    Dim counts() As Long, idx As Long, i As Long
    Dim previewOnly As Boolean
    previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Set findList = New Collection: Set replaceList = New Collection: Set itemLabels = New Collection
    If optAll.Value Or optNBSP.Value Then findList.Add UnicodeChar(&HA0): replaceList.Add " " : itemLabels.Add "Non-breaking spaces (U+00A0)"
    If optAll.Value Or optZeroWidth.Value Then
        findList.Add UnicodeChar(&H200B): replaceList.Add "" : itemLabels.Add "Zero-width spaces (U+200B)"
        findList.Add UnicodeChar(&H200C): replaceList.Add "" : itemLabels.Add "Zero-width non-joiners (U+200C)"
        findList.Add UnicodeChar(&H200D): replaceList.Add "" : itemLabels.Add "Zero-width joiners (U+200D)"
        findList.Add UnicodeChar(&HFEFF): replaceList.Add "" : itemLabels.Add "BOM / zero-width no-break spaces (U+FEFF)"
    End If
    If optAll.Value Then findList.Add UnicodeChar(&HAD): replaceList.Add "-" : itemLabels.Add "Soft hyphens (U+00AD)" : findList.Add UnicodeChar(&H2011): replaceList.Add "-" : itemLabels.Add "Non-breaking hyphens (U+2011)"
    If optCustom.Value Then
        Set findList = New Collection: Set replaceList = New Collection: Set itemLabels = New Collection
        If chkNBSP.Value Then findList.Add UnicodeChar(&HA0): replaceList.Add " " : itemLabels.Add "Non-breaking spaces (U+00A0)"
        If chkZWSP.Value Then findList.Add UnicodeChar(&H200B): replaceList.Add "" : itemLabels.Add "Zero-width spaces (U+200B)"
        If chkZWNJ.Value Then findList.Add UnicodeChar(&H200C): replaceList.Add "" : itemLabels.Add "Zero-width non-joiners (U+200C)"
        If chkZWJ.Value Then findList.Add UnicodeChar(&H200D): replaceList.Add "" : itemLabels.Add "Zero-width joiners (U+200D)"
        If chkBOM.Value Then findList.Add UnicodeChar(&HFEFF): replaceList.Add "" : itemLabels.Add "BOM / zero-width no-break spaces (U+FEFF)"
        If chkSoftHyphen.Value Then findList.Add UnicodeChar(&HAD): replaceList.Add "-" : itemLabels.Add "Soft hyphens (U+00AD)"
        If chkNBHyphen.Value Then findList.Add UnicodeChar(&H2011): replaceList.Add "-" : itemLabels.Add "Non-breaking hyphens (U+2011)"
    End If
    If findList.Count = 0 Then MsgBox "No Unicode options selected.", vbInformation: Exit Sub
    ReDim counts(1 To findList.Count): For i = 1 To findList.Count: counts(i) = 0: Next i
    If previewOnly Then
        Dim previewRows As Collection
        Dim previewNBSP As Long
        Dim previewZWSP As Long
        Dim previewZWNJ As Long
        Dim previewZWJ As Long
        Dim previewBOM As Long
        Dim previewSoftHyphen As Long
        Dim previewNBHyphen As Long
        Dim includeNBSP As Boolean
        Dim includeZWSP As Boolean
        Dim includeZWNJ As Boolean
        Dim includeZWJ As Boolean
        Dim includeBOM As Boolean
        Dim includeSoftHyphen As Boolean
        Dim includeNBHyphen As Boolean
        includeNBSP = optAll.Value Or optNBSP.Value Or (optCustom.Value And chkNBSP.Value)
        includeZWSP = optAll.Value Or optZeroWidth.Value Or (optCustom.Value And chkZWSP.Value)
        includeZWNJ = optAll.Value Or optZeroWidth.Value Or (optCustom.Value And chkZWNJ.Value)
        includeZWJ = optAll.Value Or optZeroWidth.Value Or (optCustom.Value And chkZWJ.Value)
        includeBOM = optAll.Value Or optZeroWidth.Value Or (optCustom.Value And chkBOM.Value)
        includeSoftHyphen = optAll.Value Or (optCustom.Value And chkSoftHyphen.Value)
        includeNBHyphen = optAll.Value Or (optCustom.Value And chkNBHyphen.Value)

        Dim hybridFailure As String
        Dim hybridSucceeded As Boolean
        Set mHybridAnalysis = Nothing
        Set mHybridDocument = Nothing
        hybridSucceeded = HybridAnalyzeInvisibleUnicode( _
            targetRange, (optScopeSelection.Value And optScopeSelection.Enabled), _
            includeNBSP, includeZWSP, includeZWNJ, includeZWJ, _
            includeBOM, includeSoftHyphen, includeNBHyphen, _
            mHybridAnalysis, hybridFailure)
        If Not hybridSucceeded Then
            MsgBox hybridFailure, vbExclamation, "Unicode Analysis Unavailable"
            Exit Sub
        End If

        Set mHybridDocument = targetRange.Document
        mHybridScopeStart = targetRange.Start
        mHybridScopeEnd = targetRange.End
        Dim hybridCandidates As Collection
        Dim hybridCandidate As Object
        Dim hybridReason As String
        Dim hybridRange As Range
        Set hybridCandidates = mHybridAnalysis("candidates")
        For Each hybridCandidate In hybridCandidates
            hybridReason = CStr(hybridCandidate("reasonCode"))
            Select Case hybridReason
                Case "unicode.non-breaking-space": previewNBSP = previewNBSP + 1
                Case "unicode.zero-width-space": previewZWSP = previewZWSP + 1
                Case "unicode.zero-width-nonjoiner": previewZWNJ = previewZWNJ + 1
                Case "unicode.zero-width-joiner": previewZWJ = previewZWJ + 1
                Case "unicode.byte-order-mark": previewBOM = previewBOM + 1
                Case "unicode.soft-hyphen": previewSoftHyphen = previewSoftHyphen + 1
                Case "unicode.non-breaking-hyphen": previewNBHyphen = previewNBHyphen + 1
            End Select
            Set hybridRange = HybridCandidateAbsoluteRange(mHybridAnalysis, hybridCandidate, mHybridDocument)
            ApplyPreviewHighlight hybridRange
            If UnicodeHybridNeedsAdjacentFallback(hybridReason) Then HighlightUnicodeAdjacentCharacter hybridRange, targetRange
        Next hybridCandidate

        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Non-breaking spaces", previewNBSP, False, (Not includeNBSP)
        AddPreviewSummaryRow previewRows, "Zero-width spaces", previewZWSP, False, (Not includeZWSP)
        AddPreviewSummaryRow previewRows, "Zero-width non-joiners", previewZWNJ, False, (Not includeZWNJ)
        AddPreviewSummaryRow previewRows, "Zero-width joiners", previewZWJ, False, (Not includeZWJ)
        AddPreviewSummaryRow previewRows, "Byte order marks", previewBOM, False, (Not includeBOM)
        AddPreviewSummaryRow previewRows, "Soft hyphens", previewSoftHyphen, False, (Not includeSoftHyphen)
        AddPreviewSummaryRow previewRows, "Non-breaking hyphens", previewNBHyphen, False, (Not includeNBHyphen)
        ShowPreviewActionsSummary Me, "Unicode Cleaner", previewRows
        Exit Sub
    End If
    If mHybridAnalysis Is Nothing Or mHybridDocument Is Nothing Then
        MsgBox "Run Preview before applying Unicode cleanup.", vbInformation, "Preview Required"
        Exit Sub
    End If
    If Not ActiveDocument Is mHybridDocument Then
        MsgBox "The active document is no longer the document that was previewed. Run Preview again.", vbExclamation, "Preview Is Stale"
        Exit Sub
    End If
    Set targetRange = mHybridDocument.Range(mHybridScopeStart, mHybridScopeEnd)
    Dim applyFailure As String
    If Not HybridRevalidateInvisibleUnicode(mHybridAnalysis, targetRange, applyFailure) Then
        MsgBox applyFailure & vbCrLf & vbCrLf & "No changes were applied. Run Preview again.", vbExclamation, "Preview Is Stale"
        Set mHybridAnalysis = Nothing
        Exit Sub
    End If
    MarkCleanupStart "Unicode Cleaner"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Unicode Cleaner"
    On Error GoTo RunErr
    Dim applyCandidates As Collection
    Dim applyCandidate As Object
    Dim applyRange As Range
    Set applyCandidates = mHybridAnalysis("candidates")
    For idx = applyCandidates.Count To 1 Step -1
        Set applyCandidate = applyCandidates(idx)
        Set applyRange = HybridCandidateAbsoluteRange(mHybridAnalysis, applyCandidate, mHybridDocument)
        applyRange.Text = HybridCandidateReplacement(applyCandidate)
    Next idx
    RemoveAllHighlighting targetRange
    undoRec.EndCustomRecord
    MarkCleanupEnd
    MarkCleanupToolApplied
    Set mHybridAnalysis = Nothing
    Set mHybridDocument = Nothing
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

Private Function ReplaceUnicodeMatches(ByVal searchRange As Range, ByVal findText As String, ByVal replacementText As String) As Long
    Dim matchRange As Range
    Set matchRange = searchRange.Duplicate

    With matchRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False

        Do While .Execute
            matchRange.Text = replacementText
            ReplaceUnicodeMatches = ReplaceUnicodeMatches + 1
            matchRange.Collapse wdCollapseEnd
        Loop
    End With
End Function
Private Function UnicodePreviewNeedsAdjacentFallback(ByVal findText As String) As Boolean
    If findText = UnicodeChar(&H200C) Then
        UnicodePreviewNeedsAdjacentFallback = True
    ElseIf Not UnicodeFormattingMarksVisible() Then
        UnicodePreviewNeedsAdjacentFallback = (findText = UnicodeChar(&H200B) Or findText = UnicodeChar(&HFEFF))
    End If
End Function
Private Function UnicodeHybridNeedsAdjacentFallback(ByVal reasonCode As String) As Boolean
    If reasonCode = "unicode.zero-width-nonjoiner" Then
        UnicodeHybridNeedsAdjacentFallback = True
    ElseIf Not UnicodeFormattingMarksVisible() Then
        UnicodeHybridNeedsAdjacentFallback = _
            (reasonCode = "unicode.zero-width-space" Or _
             reasonCode = "unicode.byte-order-mark")
    End If
End Function
Private Function UnicodeFormattingMarksVisible() As Boolean
    On Error Resume Next
    UnicodeFormattingMarksVisible = ActiveWindow.View.ShowAll
    On Error GoTo 0
End Function
Private Sub HighlightUnicodeMatches(ByVal searchRange As Range, ByVal findText As String, ByVal useAdjacentFallback As Boolean)
    Dim matchRange As Range
    Dim searchEnd As Long
    searchEnd = searchRange.End
    Set matchRange = searchRange.Duplicate
    Do
        With matchRange.Find
            .ClearFormatting
            .Text = findText
            .Forward = True
            .Wrap = wdFindStop
            .MatchWildcards = False
            If Not .Execute Then Exit Do
        End With
        ApplyPreviewHighlight matchRange
        If useAdjacentFallback Then HighlightUnicodeAdjacentCharacter matchRange, searchRange
        matchRange.Collapse wdCollapseEnd
        If matchRange.Start >= searchEnd Then Exit Do
        matchRange.End = searchEnd
    Loop
End Sub
Private Sub HighlightUnicodeAdjacentCharacter(ByVal matchRange As Range, ByVal scopeRange As Range)
    Dim markerRange As Range
    If matchRange.End < scopeRange.End Then
        Set markerRange = ActiveDocument.Range(matchRange.End, matchRange.End + 1)
    ElseIf matchRange.Start > scopeRange.Start Then
        Set markerRange = ActiveDocument.Range(matchRange.Start - 1, matchRange.Start)
    Else
        Exit Sub
    End If
    ApplyPreviewHighlight markerRange
End Sub
