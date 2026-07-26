Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Style Cleanup
'  Removes unused custom styles and optionally remaps common variant styles
'  (Normal (Web), Body Text 2, etc.) to the Normal style.  Built-in Word
'  styles are never deleted.  Always operates on the whole document.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    chkRemoveUnused.Value = True
    chkRemapVariants.Value = False
    fraScopeSelection.Enabled = False
    optScopeDocument.Value = True
    optScopeDocument.Caption = "Entire document (always)"
    optScopeSelection.Enabled = False
    chkRemoveUnused.Caption = "Remove unused styles"
    chkRemapVariants.Caption = "Remap style variants to canonical equivalents"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Function RemappableStyleNames() As Variant
    RemappableStyleNames = Array("Normal (Web)", "Body Text 2", "Body Text 3", "Body Text Indent", "Body Text Indent 2")
End Function

Private Function TryGetDocumentStyle(ByVal styleName As String, ByRef foundStyle As Style) As Boolean
    Set foundStyle = Nothing
    On Error Resume Next
    Set foundStyle = ActiveDocument.Styles(styleName)
    On Error GoTo 0
    TryGetDocumentStyle = Not foundStyle Is Nothing
End Function

Private Function ParagraphUsesStyle(ByVal paragraphItem As Paragraph, ByVal targetStyle As Style) As Boolean
    Dim appliedStyle As Style
    On Error GoTo NotMatched
    Set appliedStyle = paragraphItem.Range.Style
    ParagraphUsesStyle = (StrComp(appliedStyle.NameLocal, targetStyle.NameLocal, vbTextCompare) = 0)
    Exit Function
NotMatched:
    ParagraphUsesStyle = False
End Function

Private Function CountParagraphsUsingStyle(ByVal targetStyle As Style) As Long
    Dim firstStory As Range
    Dim storyPart As Range
    Dim paragraphItem As Paragraph
    For Each firstStory In ActiveDocument.StoryRanges
        Set storyPart = firstStory
        Do While Not storyPart Is Nothing
            For Each paragraphItem In storyPart.Paragraphs
                If ParagraphUsesStyle(paragraphItem, targetStyle) Then
                    CountParagraphsUsingStyle = CountParagraphsUsingStyle + 1
                End If
            Next paragraphItem
            Set storyPart = storyPart.NextStoryRange
        Loop
    Next firstStory
End Function

Private Function StyleInUse(ByVal targetStyle As Style) As Boolean
    If CountParagraphsUsingStyle(targetStyle) > 0 Then
        StyleInUse = True
        Exit Function
    End If

    Dim firstStory As Range
    Dim storyPart As Range
    Dim searchRange As Range
    For Each firstStory In ActiveDocument.StoryRanges
        Set storyPart = firstStory
        Do While Not storyPart Is Nothing
            Set searchRange = storyPart.Duplicate
            With searchRange.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "^?"
                .Style = targetStyle
                .Format = True
                .Forward = True
                .Wrap = wdFindStop
                .MatchWildcards = False
                If .Execute Then
                    StyleInUse = True
                    Exit Function
                End If
            End With
            Set storyPart = storyPart.NextStoryRange
        Loop
    Next firstStory
End Function

Private Function RemapStyle(ByVal fromName As String, ByVal toName As String) As Long
    Dim sourceStyle As Style
    Dim targetStyle As Style
    If Not TryGetDocumentStyle(fromName, sourceStyle) Then Exit Function
    If Not TryGetDocumentStyle(toName, targetStyle) Then Exit Function

    Dim firstStory As Range
    Dim storyPart As Range
    Dim paragraphItem As Paragraph
    For Each firstStory In ActiveDocument.StoryRanges
        Set storyPart = firstStory
        Do While Not storyPart Is Nothing
            For Each paragraphItem In storyPart.Paragraphs
                If ParagraphUsesStyle(paragraphItem, sourceStyle) Then
                    paragraphItem.Range.Style = targetStyle
                    RemapStyle = RemapStyle + 1
                End If
            Next paragraphItem
            Set storyPart = storyPart.NextStoryRange
        Loop
    Next firstStory
End Function
Private Sub chkRemoveUnused_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkRemapVariants_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub cmdRiskPlacementStyleRemoveUnused_Click(): ShowToolRiskChoiceExplanation "Remove unused styles", "Structure change", "Deletes custom styles that are not applied in the document." : End Sub
Private Sub cmdRiskPlacementStyleRemap_Click(): ShowToolRiskChoiceExplanation "Remap style variants", "Inspect first", "Moves content from lookalike styles into canonical styles." : End Sub
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
    If Not GuardBeforeCleanup("Style Cleanup") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim doUnused As Boolean: doUnused = chkRemoveUnused.Value
    Dim doRemap As Boolean: doRemap = chkRemapVariants.Value
    If Not (doUnused Or doRemap) Then MsgBox "Nothing selected.", vbInformation: Exit Sub
    Dim st As Style
    If previewOnly Then
        Dim previewUnused As Long
        Dim previewRemaps As Long
        If doUnused Then
            For Each st In ActiveDocument.Styles
                If Not st.BuiltIn Then
                    If Not StyleInUse(st) Then previewUnused = previewUnused + 1
                End If
            Next st
        End If
        If doRemap Then previewRemaps = HighlightStyleVariantsForRemap()
        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Unused custom styles", previewUnused, True, (Not doUnused)
        AddPreviewSummaryRow previewRows, "Variant style remaps", previewRemaps, False, (Not doRemap)
        ShowPreviewActionsSummary Me, "Style Cleanup", previewRows
        Exit Sub
    End If
    MarkCleanupStart "Style Cleanup"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Style Cleanup"
    On Error GoTo RunErr
    Dim cntRemap As Long: cntRemap = 0
    Dim cntUnused As Long: cntUnused = 0
    If doRemap Then
        Dim variants As Variant, vi As Long
        variants = RemappableStyleNames()
        For vi = LBound(variants) To UBound(variants)
            UpdateCleanupProgress "Remapping style variants", vi - LBound(variants) + 1, UBound(variants) - LBound(variants) + 1
            cntRemap = cntRemap + RemapStyle(CStr(variants(vi)), "Normal")
        Next vi
    End If
    If doUnused Then
        Dim i As Long
        For i = ActiveDocument.Styles.Count To 1 Step -1
            If i Mod 25 = 0 Or i = ActiveDocument.Styles.Count Then UpdateCleanupProgress "Checking unused styles", ActiveDocument.Styles.Count - i + 1, ActiveDocument.Styles.Count
            Set st = ActiveDocument.Styles(i)
            If Not st.BuiltIn Then
                If Not StyleInUse(st) Then
                    On Error Resume Next
                    st.Delete
                    If Err.Number = 0 Then cntUnused = cntUnused + 1
                    Err.Clear
                    On Error GoTo RunErr
                End If
            End If
        Next i
    End If
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

Private Function CountStyleVariantsForRemap() As Long
    Dim variants As Variant
    Dim vi As Long
    Dim variantStyle As Style
    variants = RemappableStyleNames()
    For vi = LBound(variants) To UBound(variants)
        If TryGetDocumentStyle(CStr(variants(vi)), variantStyle) Then
            CountStyleVariantsForRemap = CountStyleVariantsForRemap + CountParagraphsUsingStyle(variantStyle)
        End If
    Next vi
End Function

Private Function HighlightStyleVariantsForRemap() As Long
    Dim variants As Variant
    Dim vi As Long
    Dim variantStyle As Style
    Dim firstStory As Range
    Dim storyPart As Range
    Dim paragraphItem As Paragraph
    variants = RemappableStyleNames()

    For vi = LBound(variants) To UBound(variants)
        If TryGetDocumentStyle(CStr(variants(vi)), variantStyle) Then
            For Each firstStory In ActiveDocument.StoryRanges
                Set storyPart = firstStory
                Do While Not storyPart Is Nothing
                    For Each paragraphItem In storyPart.Paragraphs
                        If ParagraphUsesStyle(paragraphItem, variantStyle) Then
                            HighlightStyleVariantsForRemap = HighlightStyleVariantsForRemap + 1
                            ApplyPreviewMinimalMarker paragraphItem.Range, True
                        End If
                    Next paragraphItem
                    Set storyPart = storyPart.NextStoryRange
                Loop
            Next firstStory
        End If
    Next vi
End Function
