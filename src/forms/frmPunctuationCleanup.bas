Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Punctuation Normalizer
'  Replaces typographic / smart punctuation characters with plain-text
'  equivalents: curly double and single quotes, em and en dashes, and
'  the ellipsis character.  Essential before importing text into systems
'  that do not handle Unicode punctuation.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optAll.GroupName = "PunctMode"
    optQuotes.GroupName = "PunctMode"
    optDashes.GroupName = "PunctMode"
    optEllipses.GroupName = "PunctMode"
    optCustom.GroupName = "PunctMode"
    optScopeDocument.GroupName = "PunctScope"
    optScopeSelection.GroupName = "PunctScope"
    optAll.Value = True
    UpdateCustomVisibility
    ClearCustomChecks
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    RefreshScopeSelectionState Me, True
    optAll.Caption = "All punctuation types"
    optQuotes.Caption = "Quotes only"
    optDashes.Caption = "Dashes only"
    optEllipses.Caption = "Ellipses only"
    optCustom.Caption = "Custom"
    chkCurlyDouble.Caption = "Curly double quotes  -> straight double quotes"
    chkCurlySingle.Caption = "Curly single quotes / apostrophes  -> straight single quotes"
    chkEmDash.Caption = "Em dash  ->  hyphen-minus  -"
    chkEnDash.Caption = "En dash  ->  hyphen-minus  -"
    chkEllipses.Caption = "Ellipsis character  ->  three dots  ..."
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutPunctuationForm
End Sub
Private Sub LayoutPunctuationForm()
    LayoutCleanupToolForm Me
End Sub
Private Sub UpdateCustomVisibility()
    fraCustom.Visible = False
End Sub
Private Sub optAll_Click(): UpdateCustomVisibility: LayoutPunctuationForm: End Sub
Private Sub optQuotes_Click(): UpdateCustomVisibility: LayoutPunctuationForm: End Sub
Private Sub optDashes_Click(): UpdateCustomVisibility: LayoutPunctuationForm: End Sub
Private Sub optEllipses_Click(): UpdateCustomVisibility: LayoutPunctuationForm: End Sub
Private Sub optCustom_Click(): UpdateCustomVisibility: LayoutPunctuationForm: End Sub
Private Sub cmdRiskPlacementPunctuationAll_Click(): ShowToolRiskChoiceExplanation "All punctuation types", "Text caution", "Changes visible punctuation characters across quotes, dashes, and ellipses. Preview first for prose where typography matters.": End Sub
Private Sub cmdRiskPlacementPunctuationQuotes_Click(): ShowToolRiskChoiceExplanation "Quotes only", "Text caution", "Converts curly quotes and apostrophes to straight plain-text marks.": End Sub
Private Sub cmdRiskPlacementPunctuationDashes_Click(): ShowToolRiskChoiceExplanation "Dashes only", "Text caution", "Converts em and en dashes to plain hyphens, which changes visible punctuation style.": End Sub
Private Sub cmdRiskPlacementPunctuationEllipses_Click(): ShowToolRiskChoiceExplanation "Ellipses only", "Text caution", "Converts the single ellipsis character to three plain periods.": End Sub
Private Sub cmdRiskPlacementPunctuationCustom_Click(): ShowToolRiskChoiceExplanation "Custom punctuation cleanup", "Inspect first", "Lets you combine individual punctuation conversions. Preview first when preserving typography may matter.": End Sub
Private Sub ClearCustomChecks()
    chkCurlyDouble.Value = False
    chkCurlySingle.Value = False
    chkEmDash.Value = False
    chkEnDash.Value = False
    chkEllipses.Value = False
End Sub
Private Sub cmdSelectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    chkCurlyDouble.Value = True
    chkCurlySingle.Value = True
    chkEmDash.Value = True
    chkEnDash.Value = True
    chkEllipses.Value = True
    LayoutPunctuationForm
End Sub
Private Sub cmdDeselectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    ClearCustomChecks
    LayoutPunctuationForm
End Sub
Private Sub chkCurlyDouble_Click(): LayoutPunctuationForm: End Sub
Private Sub chkCurlySingle_Click(): LayoutPunctuationForm: End Sub
Private Sub chkEmDash_Click(): LayoutPunctuationForm: End Sub
Private Sub chkEnDash_Click(): LayoutPunctuationForm: End Sub
Private Sub chkEllipses_Click(): LayoutPunctuationForm: End Sub
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
    If Not GuardBeforeCleanup("Punctuation Normalizer") Then Unload Me: Exit Sub
    Dim findList As Collection, replaceList As Collection, itemLabels As Collection
    Dim idx As Long
    Dim previewOnly As Boolean
    Dim screenUpdatingWasOn As Boolean
    Dim autoFormatAsYouTypeQuotesWasOn As Boolean
    Dim autoFormatQuotesWasOn As Boolean
    previewOnly = chkPreviewOnly.Value
    screenUpdatingWasOn = Application.ScreenUpdating
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim includeCurlyDouble As Boolean
    Dim includeCurlySingle As Boolean
    Dim includeEmDash As Boolean
    Dim includeEnDash As Boolean
    Dim includeEllipses As Boolean
    includeCurlyDouble = optAll.Value Or optQuotes.Value Or (optCustom.Value And chkCurlyDouble.Value)
    includeCurlySingle = optAll.Value Or optQuotes.Value Or (optCustom.Value And chkCurlySingle.Value)
    includeEmDash = optAll.Value Or optDashes.Value Or (optCustom.Value And chkEmDash.Value)
    includeEnDash = optAll.Value Or optDashes.Value Or (optCustom.Value And chkEnDash.Value)
    includeEllipses = optAll.Value Or optEllipses.Value Or (optCustom.Value And chkEllipses.Value)
    Set findList = New Collection: Set replaceList = New Collection: Set itemLabels = New Collection
    If optAll.Value Or optQuotes.Value Then
        findList.Add ChrW(&H201C): replaceList.Add """"
        itemLabels.Add "Left curly double quotes (U+201C)"
        findList.Add ChrW(&H201D): replaceList.Add """"
        itemLabels.Add "Right curly double quotes (U+201D)"
        findList.Add ChrW(&H2018): replaceList.Add "'"
        itemLabels.Add "Left curly single quotes (U+2018)"
        findList.Add ChrW(&H2019): replaceList.Add "'"
        itemLabels.Add "Right curly single quotes / apostrophes (U+2019)"
    End If
    If optAll.Value Or optDashes.Value Then
        findList.Add ChrW(&H2014): replaceList.Add "-" : itemLabels.Add "Em dashes (U+2014)"
        findList.Add ChrW(&H2013): replaceList.Add "-" : itemLabels.Add "En dashes (U+2013)"
    End If
    If optAll.Value Or optEllipses.Value Then
        findList.Add ChrW(&H2026): replaceList.Add "..." : itemLabels.Add "Ellipses (U+2026)"
    End If
    If optCustom.Value Then
        Set findList = New Collection: Set replaceList = New Collection: Set itemLabels = New Collection
        If chkCurlyDouble.Value Then findList.Add ChrW(&H201C): replaceList.Add """" : itemLabels.Add "Left curly double quotes (U+201C)" : findList.Add ChrW(&H201D): replaceList.Add """" : itemLabels.Add "Right curly double quotes (U+201D)"
        If chkCurlySingle.Value Then findList.Add ChrW(&H2018): replaceList.Add "'" : itemLabels.Add "Left curly single quotes (U+2018)" : findList.Add ChrW(&H2019): replaceList.Add "'" : itemLabels.Add "Right curly single quotes / apostrophes (U+2019)"
        If chkEmDash.Value Then findList.Add ChrW(&H2014): replaceList.Add "-" : itemLabels.Add "Em dashes (U+2014)"
        If chkEnDash.Value Then findList.Add ChrW(&H2013): replaceList.Add "-" : itemLabels.Add "En dashes (U+2013)"
        If chkEllipses.Value Then findList.Add ChrW(&H2026): replaceList.Add "..." : itemLabels.Add "Ellipses (U+2026)"
    End If
    If findList.Count = 0 Then MsgBox "No punctuation options selected.", vbInformation: Exit Sub
    If previewOnly Then
        On Error GoTo PreviewErr
        Dim previewRows As Collection
        Dim previewCurlyDouble As Long
        Dim previewCurlySingle As Long
        Dim previewEmDash As Long
        Dim previewEnDash As Long
        Dim previewEllipses As Long
        If includeCurlyDouble Then previewCurlyDouble = CountPreviewFindMatches(targetRange, ChrW(&H201C), True) + CountPreviewFindMatches(targetRange, ChrW(&H201D), True)
        If includeCurlySingle Then previewCurlySingle = CountPreviewFindMatches(targetRange, ChrW(&H2018), True) + CountPreviewFindMatches(targetRange, ChrW(&H2019), True)
        If includeEmDash Then previewEmDash = CountPreviewFindMatches(targetRange, ChrW(&H2014), True)
        If includeEnDash Then previewEnDash = CountPreviewFindMatches(targetRange, ChrW(&H2013), True)
        If includeEllipses Then previewEllipses = CountPreviewFindMatches(targetRange, ChrW(&H2026), True)
        Dim previewCharacters As String
        For idx = 1 To findList.Count
            previewCharacters = previewCharacters & CStr(findList(idx))
        Next idx
        HighlightPreviewFindCharacters targetRange, previewCharacters
        On Error GoTo 0

        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Curly double quotes", previewCurlyDouble, False, (Not includeCurlyDouble)
        AddPreviewSummaryRow previewRows, "Curly single quotes", previewCurlySingle, False, (Not includeCurlySingle)
        AddPreviewSummaryRow previewRows, "Em dashes", previewEmDash, False, (Not includeEmDash)
        AddPreviewSummaryRow previewRows, "En dashes", previewEnDash, False, (Not includeEnDash)
        AddPreviewSummaryRow previewRows, "Ellipses", previewEllipses, False, (Not includeEllipses)
        ShowPreviewActionsSummary Me, "Punctuation Normalizer", previewRows
        Exit Sub
    End If
    MarkCleanupStart "Punctuation Normalizer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Punctuation Normalizer"
    On Error GoTo RunErr
    autoFormatAsYouTypeQuotesWasOn = Options.AutoFormatAsYouTypeReplaceQuotes
    autoFormatQuotesWasOn = Options.AutoFormatReplaceQuotes
    Options.AutoFormatAsYouTypeReplaceQuotes = False
    Options.AutoFormatReplaceQuotes = False
    Application.ScreenUpdating = False
    ' Replace each conversion family in one bulk pass.
    If includeCurlyDouble Then ReplacePunctuationCharacterSet targetRange, ChrW(&H201C) & ChrW(&H201D), """"
    If includeCurlySingle Then ReplacePunctuationCharacterSet targetRange, ChrW(&H2018) & ChrW(&H2019), "'"
    Dim dashCharacters As String
    If includeEmDash Then dashCharacters = dashCharacters & ChrW(&H2014)
    If includeEnDash Then dashCharacters = dashCharacters & ChrW(&H2013)
    If Len(dashCharacters) > 0 Then ReplacePunctuationCharacterSet targetRange, dashCharacters, "-"
    If includeEllipses Then ReplacePunctuationCharacterSet targetRange, ChrW(&H2026), "..."
    Application.ScreenUpdating = screenUpdatingWasOn
    Options.AutoFormatAsYouTypeReplaceQuotes = autoFormatAsYouTypeQuotesWasOn
    Options.AutoFormatReplaceQuotes = autoFormatQuotesWasOn
    ' Remove highlighting
    RemoveAllHighlighting targetRange
    undoRec.EndCustomRecord
    MarkCleanupEnd
    MarkCleanupToolApplied
    Unload Me
    Exit Sub
PreviewErr:
    Application.ScreenUpdating = screenUpdatingWasOn
    MsgBox "Could not build the punctuation preview: " & Err.Description, vbExclamation, "Punctuation Preview"
    Exit Sub
RunErr:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    Application.ScreenUpdating = screenUpdatingWasOn
    Options.AutoFormatAsYouTypeReplaceQuotes = autoFormatAsYouTypeQuotesWasOn
    Options.AutoFormatReplaceQuotes = autoFormatQuotesWasOn
    On Error Resume Next: undoRec.EndCustomRecord: On Error GoTo 0
    MarkCleanupEnd
    Dim errMsg As String
    errMsg = "An unexpected error occurred: " & originalErrNumber & " - " & originalErrDescription
    errMsg = errMsg & vbCrLf & vbCrLf & _
             "Some changes may have been made to the document." & vbCrLf & _
             "Use Word's Undo command (Ctrl+Z) after closing this message if you want to roll them back."
    MsgBox errMsg, vbCritical, "Cleanup Error"
End Sub

Private Function ReplacePunctuationCharacterSet(ByVal searchRange As Range, ByVal findCharacters As String, ByVal replacementText As String) As Long
    Dim replaceRange As Range
    Dim characterIndex As Long

    For characterIndex = 1 To Len(findCharacters)
        ReplacePunctuationCharacterSet = ReplacePunctuationCharacterSet + CountPreviewFindMatches(searchRange, Mid$(findCharacters, characterIndex, 1))
    Next characterIndex
    If ReplacePunctuationCharacterSet = 0 Then Exit Function
    Set replaceRange = searchRange.Duplicate

    With replaceRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "[" & findCharacters & "]"
        .Replacement.Text = replacementText
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .MatchWildcards = True
        .Execute Replace:=wdReplaceAll
    End With
End Function
