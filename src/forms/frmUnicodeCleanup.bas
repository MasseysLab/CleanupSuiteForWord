Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
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
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
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
    LayoutCleanupToolForm Me
End Sub
Private Sub UpdateCustomVisibility()
    fraCustom.Visible = False
End Sub
Private Sub optAll_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub optNBSP_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub optZeroWidth_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub optCustom_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
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
    LayoutCleanupToolForm Me
End Sub
Private Sub cmdDeselectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    ClearCustomChecks
    LayoutCleanupToolForm Me
End Sub
Private Function UnicodeChar(ByVal codePoint As Long) As String
    If codePoint > &H7FFF Then
        UnicodeChar = ChrW$(codePoint - &H10000)
    Else
        UnicodeChar = ChrW$(codePoint)
    End If
End Function
Private Sub chkNBSP_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkZWSP_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkZWNJ_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkZWJ_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkBOM_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkSoftHyphen_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkNBHyphen_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub cmdPreview_Click()
    PreviewFromPanel
End Sub
Public Sub PreviewFromPanel()
    chkPreviewOnly.Value = True
    cmdRun_Click
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
    For idx = 1 To findList.Count
        With targetRange.Find
            .ClearFormatting: .Replacement.ClearFormatting: .Text = findList(idx): .Replacement.Text = "^&": .Replacement.Highlight = True
            .Forward = True: .Wrap = wdFindStop: .MatchWildcards = False: .Execute Replace:=wdReplaceAll
        End With
    Next idx
    If previewOnly Then
        ShowPreviewActions Me, "Unicode Cleaner", "Preview complete. Matches highlighted."
        Exit Sub
    End If
    MarkCleanupStart "Unicode Cleaner"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Unicode Cleaner"
    On Error GoTo RunErr
    For idx = 1 To findList.Count
        With targetRange.Find
            .ClearFormatting: .Replacement.ClearFormatting: .Text = findList(idx): .Replacement.Text = replaceList(idx): .Forward = True: .Wrap = wdFindStop
            While .Execute(Replace:=wdReplaceOne): counts(idx) = counts(idx) + 1: Wend
        End With
    Next idx
    RemoveAllHighlighting targetRange
    Dim results As Collection: Set results = New Collection
    For idx = 1 To findList.Count: results.Add itemLabels(idx) & ": " & counts(idx) & " replaced": Next idx
    ShowCleanupReport "Unicode Cleanup", results
    undoRec.EndCustomRecord
    MarkCleanupEnd
    MarkCleanupToolApplied
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
