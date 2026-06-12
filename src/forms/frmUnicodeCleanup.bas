Option Explicit
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
    optCustom.Caption = "Custom (choose below)"
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
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Invisible Unicode Cleaner" & vbCrLf & vbCrLf
    h = h & "Removes Unicode characters that are invisible on screen but cause problems " & _
            "for search, sorting, comparison, and automated text processing." & vbCrLf
    h = h & "These characters accumulate from web copy-paste, PDF export, and " & _
            "cross-platform file exchange." & vbCrLf
    h = h & vbCrLf
    h = h & "CHARACTER TYPES" & vbCrLf
    h = h & "  NBSP  (U+00A0)  Non-Breaking Space: a space that prevents line breaks." & vbCrLf
    h = h & "  ZWSP  (U+200B)  Zero-Width Space: completely invisible; affects word break." & vbCrLf
    h = h & "  ZWNJ  (U+200C)  Zero-Width Non-Joiner: controls ligature/joining behaviour." & vbCrLf
    h = h & "  ZWJ   (U+200D)  Zero-Width Joiner: forces joined rendering." & vbCrLf
    h = h & "  BOM   (U+FEFF)  Byte Order Mark: encoding marker; never belongs in text." & vbCrLf
    h = h & "  SHY   (U+00AD)  Soft Hyphen: invisible optional hyphen." & vbCrLf
    h = h & "  NBHY  (U+2011)  Non-Breaking Hyphen: a hyphen that prevents line breaks." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights matches in yellow without changing anything." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Invisible Unicode Cleaner"
End Sub
Private Sub UpdateCustomVisibility()
    fraCustom.Visible = optCustom.Value
End Sub
Private Sub ClearCustomChecks()
    chkNBSP.Value = False: chkZWSP.Value = False: chkZWNJ.Value = False: chkZWJ.Value = False
    chkBOM.Value = False: chkSoftHyphen.Value = False: chkNBHyphen.Value = False
End Sub
Private Sub cmdSelectAll_Click()
    If fraCustom.Visible Then
        chkNBSP.Value = True: chkZWSP.Value = True: chkZWNJ.Value = True: chkZWJ.Value = True
        chkBOM.Value = True: chkSoftHyphen.Value = True: chkNBHyphen.Value = True
    End If
End Sub
Private Sub cmdDeselectAll_Click()
    If fraCustom.Visible Then ClearCustomChecks
End Sub
Private Function UnicodeChar(ByVal codePoint As Long) As String
    If codePoint > &H7FFF Then
        UnicodeChar = ChrW$(codePoint - &H10000)
    Else
        UnicodeChar = ChrW$(codePoint)
    End If
End Function
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
