Option Explicit
' Shared helper functions for cleanup forms
Public gPreviewActionPanel As Object

Public Sub ShowPreviewActions(sourceForm As Object, toolName As String, summaryText As String)
    On Error GoTo Fallback
    sourceForm.Hide
    Set gPreviewActionPanel = New frmPreviewActions
    gPreviewActionPanel.Configure sourceForm, toolName, summaryText
    gPreviewActionPanel.Show
    Exit Sub
Fallback:
    MsgBox summaryText & vbCrLf & vbCrLf & _
           "Preview is still visible in the document. Run again without Preview to apply, or clear highlighting manually.", _
           vbInformation, toolName & " Preview"
End Sub

Public Sub RemoveAllHighlighting(Optional scopeRange As Range = Nothing)
    Dim hlRange As Range
    If scopeRange Is Nothing Then
        Set hlRange = ActiveDocument.Content
    Else
        Set hlRange = scopeRange
    End If
    With hlRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Format = True
        .Highlight = True
        .Text = ""
        .Replacement.Text = ""
        .Replacement.Highlight = False
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Public Function DocumentHasVisibleContent(Optional scopeRange As Range = Nothing) As Boolean
    On Error GoTo SafeExit

    Dim textValue As String
    If scopeRange Is Nothing Then
        textValue = ActiveDocument.Content.Text
    Else
        textValue = scopeRange.Text
    End If

    textValue = Replace(textValue, vbCr, "")
    textValue = Replace(textValue, vbLf, "")
    textValue = Replace(textValue, vbTab, "")
    textValue = Replace(textValue, Chr$(11), "")
    textValue = Replace(textValue, Chr$(12), "")
    textValue = Replace(textValue, ChrW$(160), "")
    If Len(Trim$(textValue)) > 0 Then
        DocumentHasVisibleContent = True
        Exit Function
    End If

    If Not scopeRange Is Nothing Then Exit Function
    If ActiveDocument.Tables.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
    If ActiveDocument.InlineShapes.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
    If ActiveDocument.Shapes.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
    If ActiveDocument.Hyperlinks.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
    If ActiveDocument.Footnotes.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
    If ActiveDocument.Endnotes.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
    If ActiveDocument.Comments.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
SafeExit:
End Function
Public Function GuardBeforeCleanup(toolName As String) As Boolean
    GuardBeforeCleanup = False
    If ActiveDocument.Path = "" Then
        ' Document has never been saved -- Word has nothing to recover
        Dim msg As String
        msg = "This document has not been saved yet." & vbCrLf & vbCrLf & _
              "Saving first lets Word's built-in recovery protect you if anything goes wrong." & vbCrLf & vbCrLf & _
              "Save now before running " & toolName & "?"
        Dim ans As Integer
        ans = MsgBox(msg, vbExclamation + vbYesNoCancel, "Unsaved Document")
        If ans = vbCancel Then Exit Function
        If ans = vbYes Then
            On Error Resume Next
            ActiveDocument.Save
            If Err.Number <> 0 Then
                MsgBox "Save failed. Please save manually before running cleanup.", vbCritical, "Save Failed"
                Exit Function
            End If
            On Error GoTo 0
        End If
    ElseIf GetAutoSaveSetting() And Not IsRunningInMacroHostDocument() Then
        ' Document is saved and auto-save is ON -- save silently before proceeding
        On Error Resume Next
        ActiveDocument.Save
        On Error GoTo 0
    End If
    GuardBeforeCleanup = True
End Function
Public Function IsRunningInMacroHostDocument() As Boolean
    On Error Resume Next
    IsRunningInMacroHostDocument = (ActiveDocument.FullName = ThisDocument.FullName)
    On Error GoTo 0
End Function
Public Sub MarkCleanupStart(toolName As String)
    On Error Resume Next
    If IsRunningInMacroHostDocument() Then Exit Sub
    ActiveDocument.CustomDocumentProperties("CleanupSuiteInProgress").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteInProgress", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=toolName
    If Len(ActiveDocument.Path) > 0 Then ActiveDocument.Save
    On Error GoTo 0
End Sub
Public Sub MarkCleanupEnd()
    On Error Resume Next
    If IsRunningInMacroHostDocument() Then Exit Sub
    ActiveDocument.CustomDocumentProperties("CleanupSuiteInProgress").Delete
    On Error GoTo 0
End Sub
Public Function GetAutoSaveSetting() As Boolean
    ' Returns True (auto-save ON) by default; False only if explicitly disabled
    On Error Resume Next
    Dim prop As Object
    Set prop = ActiveDocument.CustomDocumentProperties("CleanupSuiteAutoSave")
    If Err.Number <> 0 Then
        GetAutoSaveSetting = True
        Err.Clear
    Else
        GetAutoSaveSetting = (prop.Value = "True")
    End If
    On Error GoTo 0
End Function
Public Sub SetAutoSaveSetting(enabled As Boolean)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("CleanupSuiteAutoSave").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteAutoSave", _
        LinkToContent:=False, Type:=msoPropertyTypeString, _
        Value:=IIf(enabled, "True", "False")
    On Error GoTo 0
End Sub
Public Function GetReturnToMainAfterApplySetting() As Boolean
    ' Returns True by default so Apply keeps the current workflow unless changed.
    On Error Resume Next
    Dim prop As Object
    Set prop = ActiveDocument.CustomDocumentProperties("CleanupSuiteReturnToMainAfterApply")
    If Err.Number <> 0 Then
        GetReturnToMainAfterApplySetting = True
        Err.Clear
    Else
        GetReturnToMainAfterApplySetting = (prop.Value = "True")
    End If
    On Error GoTo 0
End Function
Public Sub SetReturnToMainAfterApplySetting(enabled As Boolean)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("CleanupSuiteReturnToMainAfterApply").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteReturnToMainAfterApply", _
        LinkToContent:=False, Type:=msoPropertyTypeString, _
        Value:=IIf(enabled, "True", "False")
    On Error GoTo 0
End Sub
Public Sub ResetCleanupSuiteDefaults()
    On Error Resume Next
    RemoveAllHighlighting ActiveDocument.Content
    SetAutoSaveSetting True
    SetReturnToMainAfterApplySetting True
    If Not gPreviewActionPanel Is Nothing Then Unload gPreviewActionPanel
    Set gPreviewActionPanel = Nothing
    Unload frmPunctuationCleanup
    Unload frmUnicodeCleanup
    Unload frmSpacingCleanup
    Unload frmCapitalizationCleanup
    Unload frmListCleanup
    Unload frmParagraphCleanup
    Unload frmDuplicateDetector
    Unload frmFontNormalizer
    Unload frmTableCleaner
    Unload frmBreakNormalizer
    Unload frmDocumentTrim
    Unload frmFormattingStripper
    Unload frmHyperlinkRemover
    Unload frmSoftReturnConverter
    Unload frmMetadataScrubber
    Unload frmStyleCleanup
    Unload frmFootnoteRemover
    Unload frmHeaderFooterStandardizer
    Unload frmObjectRemover
    On Error GoTo 0
End Sub
Public Sub LayoutCleanupToolForm(toolForm As Object)
    On Error Resume Next
    Const M As Single = 12
    Const GAP As Single = 8
    Const BH As Single = 26
    Dim contentW As Single
    contentW = toolForm.Width - (2 * M) - 10
    If contentW < 320 Then contentW = 344

    toolForm.Controls("chkPreviewOnly").Visible = False
    toolForm.Controls("fraScopeSelection").Visible = False

    Dim bottomY As Single
    Dim ctl As Object
    For Each ctl In toolForm.Controls
        Select Case ctl.Name
            Case "cmdPreview", "cmdRun", "cmdReset", "cmdHelp", "chkPreviewOnly", "fraScopeSelection", "optScopeDocument", "optScopeSelection"
                ' laid out below the option area
            Case Else
                If ctl.Visible Then
                    If ctl.Top + ctl.Height > bottomY Then bottomY = ctl.Top + ctl.Height
                End If
        End Select
    Next ctl

    Dim y As Single
    y = bottomY + 12
    With toolForm.Controls("cmdPreview")
        .Move M, y, contentW, BH
        .Caption = "Preview"
        .Enabled = True
        .Visible = True
    End With
    y = y + BH + GAP

    Dim halfW As Single
    halfW = (contentW - GAP) / 2
    With toolForm.Controls("cmdRun")
        .Move M, y, halfW, BH
        .Caption = "Apply"
        .Visible = True
        .Enabled = True
    End With
    With toolForm.Controls("cmdReset")
        .Move M + halfW + GAP, y, halfW, BH
        .Caption = "Reset"
        .Visible = True
        .Enabled = True
    End With
    y = y + BH + 10

    With toolForm.Controls("optScopeDocument")
        .Move M + 8, y, contentW - 8, 18
        .Visible = True
    End With
    y = y + 24
    With toolForm.Controls("optScopeSelection")
        .Move M + 8, y, contentW - 8, 18
        .Visible = True
    End With
    y = y + 34

    With toolForm.Controls("cmdHelp")
        .Move M, y, contentW, BH
        .Caption = "Help"
        .Visible = True
    End With
    y = y + BH + 34
    If toolForm.Height < y Then toolForm.Height = y
End Sub
Public Function GetTargetRange() As Range
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        Set GetTargetRange = Selection.Range
    Else
        Set GetTargetRange = ActiveDocument.Content
    End If
End Function
Public Function ParagraphsInRange(scopeRange As Range) As Collection
    Dim col As Collection: Set col = New Collection
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= scopeRange.Start And p.Range.End <= scopeRange.End Then col.Add p
    Next p
    Set ParagraphsInRange = col
End Function
Public Sub ShowCleanupReport(reportTitle As String, results As Collection)
    If results.Count = 0 Then
        MsgBox "No items were changed.", vbInformation, reportTitle: Exit Sub
    End If
    Dim reportText As String: reportText = reportTitle & vbCrLf & vbCrLf
    Dim lineItem As Variant
    For Each lineItem In results
        reportText = reportText & "  " & lineItem & vbCrLf
    Next lineItem
    MsgBox reportText, vbInformation, reportTitle
End Sub
