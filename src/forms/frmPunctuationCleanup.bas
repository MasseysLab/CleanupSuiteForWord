Option Explicit
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
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
    optAll.Caption = "All punctuation types"
    optQuotes.Caption = "Quotes only"
    optDashes.Caption = "Dashes only"
    optEllipses.Caption = "Ellipses only"
    optCustom.Caption = "Custom (choose below)"
    chkCurlyDouble.Caption = "Curly double quotes  -> straight double quotes"
    chkCurlySingle.Caption = "Curly single quotes / apostrophes  -> straight single quotes"
    chkEmDash.Caption = "Em dash  ->  hyphen-minus  -"
    chkEnDash.Caption = "En dash  ->  hyphen-minus  -"
    chkEllipses.Caption = "Ellipsis character  ->  three dots  ..."
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Punctuation Normalizer" & vbCrLf & vbCrLf
    h = h & "Replaces typographic ('smart') punctuation with plain-text equivalents." & vbCrLf
    h = h & "Use this before exporting to systems that do not support Unicode punctuation, " & _
            "or before running text-comparison or diff tools." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  All          --  Converts all punctuation types listed below." & vbCrLf
    h = h & "  Quotes only  --  Curly double and single quotes become straight" & vbCrLf
    h = h & "  Dashes only  --  Em dash and en dash  -->  hyphen (-)" & vbCrLf
    h = h & "  Ellipses     --  Ellipsis character (...)  -->  three dots (...)" & vbCrLf
    h = h & "  Custom       --  Tick individual characters below." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights matches in yellow without changing anything." & vbCrLf
    h = h & "Run again without Preview to apply." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Punctuation Normalizer"
End Sub
Private Sub UpdateCustomVisibility()
    fraCustom.Visible = optCustom.Value
End Sub
Private Sub ClearCustomChecks()
    chkCurlyDouble.Value = False
    chkCurlySingle.Value = False
    chkEmDash.Value = False
    chkEnDash.Value = False
    chkEllipses.Value = False
End Sub
Private Sub cmdSelectAll_Click()
    If fraCustom.Visible Then
        chkCurlyDouble.Value = True
        chkCurlySingle.Value = True
        chkEmDash.Value = True
        chkEnDash.Value = True
        chkEllipses.Value = True
    End If
End Sub
Private Sub cmdDeselectAll_Click()
    If fraCustom.Visible Then ClearCustomChecks
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
    If Not GuardBeforeCleanup("Punctuation Normalizer") Then Unload Me: Exit Sub
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
    ReDim counts(1 To findList.Count)
    For i = 1 To findList.Count: counts(i) = 0: Next i
    ' Highlight preview
    For idx = 1 To findList.Count
        With targetRange.Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = findList(idx): .Replacement.Text = "^&": .Replacement.Highlight = True
            .Forward = True: .Wrap = wdFindStop: .MatchWildcards = False: .Execute Replace:=wdReplaceAll
        End With
    Next idx
    If previewOnly Then
        ShowPreviewActions Me, "Punctuation Normalizer", "Preview complete. Matches highlighted."
        Exit Sub
    End If
    MarkCleanupStart "Punctuation Normalizer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Punctuation Normalizer"
    On Error GoTo RunErr
    ' Replace and count
    For idx = 1 To findList.Count
        With targetRange.Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = findList(idx): .Replacement.Text = replaceList(idx): .Forward = True: .Wrap = wdFindStop
            While .Execute(Replace:=wdReplaceOne)
                counts(idx) = counts(idx) + 1
            Wend
        End With
    Next idx
    ' Remove highlighting
    RemoveAllHighlighting targetRange
    ' Report
    Dim results As Collection: Set results = New Collection
    For idx = 1 To findList.Count: results.Add itemLabels(idx) & ": " & counts(idx) & " replaced": Next idx
    ShowCleanupReport "Punctuation Cleanup", results
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
