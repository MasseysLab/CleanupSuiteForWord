Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  List Normalizer
'  Normalizes list formatting: converts hyphen/asterisk lines to proper
'  Word bullet lists, standardizes existing bullet styles, converts
'  manually-typed numbered lists (1. 2. 3.) to Word auto-numbering,
'  and fixes list indentation to a standard hanging list indent.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optAll.GroupName = "ListMode"
    optBullets.GroupName = "ListMode"
    optNumbering.GroupName = "ListMode"
    optIndent.GroupName = "ListMode"
    optCustom.GroupName = "ListMode"
    optScopeDocument.GroupName = "ListScope"
    optScopeSelection.GroupName = "ListScope"
    optAll.Value = True: UpdateCustomVisibility: ClearCustomChecks
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    RefreshScopeSelectionState Me, True
    optAll.Caption = "All list fixes"
    optBullets.Caption = "Convert manual bullet lists"
    optNumbering.Caption = "Convert manual numbered lists"
    optIndent.Caption = "Fix list indentation"
    optCustom.Caption = "Custom"
    chkNormalizeBullets.Caption = "Convert manual hyphens / asterisks to List Bullet style"
    chkNormalizeNumbering.Caption = "Convert manual numbers to List Number style"
    chkFixIndent.Caption = "Fix list indentation"
    chkHyphenToBullets.Caption = "Convert hyphen-dash lines to bullets"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub UpdateCustomVisibility(): fraCustom.Visible = False: End Sub
Private Sub optAll_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub optBullets_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub optNumbering_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub optIndent_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub optCustom_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub ClearCustomChecks(): chkNormalizeBullets.Value = False: chkNormalizeNumbering.Value = False: chkFixIndent.Value = False: chkHyphenToBullets.Value = False: End Sub
Private Sub cmdSelectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    chkNormalizeBullets.Value = True
    chkNormalizeNumbering.Value = True
    chkFixIndent.Value = True
    chkHyphenToBullets.Value = True
    LayoutCleanupToolForm Me
End Sub
Private Sub cmdDeselectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    ClearCustomChecks
    LayoutCleanupToolForm Me
End Sub
Private Sub chkNormalizeBullets_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkNormalizeNumbering_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkFixIndent_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkHyphenToBullets_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub cmdRiskPlacementListAll_Click(): ShowToolRiskChoiceExplanation "All list fixes", "Structure change", "Converts manual list text and adjusts list indentation." : End Sub
Private Sub cmdRiskPlacementListBullets_Click(): ShowToolRiskChoiceExplanation "Bullet conversion", "Structure change", "Turns typed bullet-like lines into Word list formatting." : End Sub
Private Sub cmdRiskPlacementListNumbering_Click(): ShowToolRiskChoiceExplanation "Number conversion", "Structure change", "Turns typed numbered lines into Word numbering." : End Sub
Private Sub cmdRiskPlacementListIndent_Click(): ShowToolRiskChoiceExplanation "Fix list indentation", "Structure change", "Adjusts list paragraph indents." : End Sub
Private Sub cmdRiskPlacementListCustom_Click(): ShowToolRiskChoiceExplanation "Custom list cleanup", "Inspect first", "Lets you choose individual list structure repairs." : End Sub
Private Sub cmdRiskPlacementListNormalizeBullets_Click(): ShowToolRiskChoiceExplanation "Normalize bullets", "Structure change", "Turns manual bullets into the document's standard bullet list form." : End Sub
Private Sub cmdRiskPlacementListNormalizeNumbering_Click(): ShowToolRiskChoiceExplanation "Normalize numbering", "Structure change", "Turns manual numbering into Word-managed numbered lists." : End Sub
Private Sub cmdRiskPlacementListFixIndent_Click(): ShowToolRiskChoiceExplanation "Fix list indentation", "Structure change", "Straightens list indentation without changing list text." : End Sub
Private Sub cmdRiskPlacementListHyphenBullets_Click(): ShowToolRiskChoiceExplanation "Hyphen lines to bullets", "Structure change", "Promotes hyphen-led lines into bullet list items." : End Sub
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
    If Not GuardBeforeCleanup("List Normalizer") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doBullets As Boolean, doNumbering As Boolean, doIndent As Boolean, doHyphen As Boolean
    If optAll.Value Then doBullets = True: doNumbering = True: doIndent = True: doHyphen = True
    If optBullets.Value Then doBullets = True
    If optNumbering.Value Then doNumbering = True
    If optIndent.Value Then doIndent = True
    If optCustom.Value Then doBullets = chkNormalizeBullets.Value: doNumbering = chkNormalizeNumbering.Value: doIndent = chkFixIndent.Value: doHyphen = chkHyphenToBullets.Value
    If Not (doBullets Or doNumbering Or doIndent Or doHyphen) Then MsgBox "No list options selected.", vbInformation: Exit Sub
    Dim p As Paragraph
    Dim txt As String
    Dim changed As Long
    Dim previewHyphen As Long, previewBullets As Long, previewNumbering As Long, previewIndent As Long
    Dim actionHyphen As Boolean, actionBullets As Boolean, actionNumbering As Boolean, actionIndent As Boolean
    If previewOnly Then
        ' Only typed prefixes can be highlighted. Word-generated markers and
        ' ruler indents are reported in the summary without artificial marks.
        For Each p In targetRange.Paragraphs
            If ParagraphOverlapsRangeOutsideTable(p, targetRange) Then
                txt = Trim$(ParagraphBodyText(p))
                ResolveListParagraphActions p, txt, doBullets, doNumbering, doIndent, doHyphen, actionHyphen, actionBullets, actionNumbering, actionIndent
                If actionHyphen Or actionNumbering Then HighlightManualListPrefix p, actionNumbering
                If actionBullets Or actionIndent Then ApplyPreviewMinimalMarker p.Range, True
                If actionHyphen Or actionBullets Or actionNumbering Or actionIndent Then changed = changed + 1
                If actionHyphen Then previewHyphen = previewHyphen + 1
                If actionBullets Then previewBullets = previewBullets + 1
                If actionNumbering Then previewNumbering = previewNumbering + 1
                If actionIndent Then previewIndent = previewIndent + 1
            End If
        Next p

        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Hyphen list items", previewHyphen, False, (Not doHyphen)
        AddPreviewSummaryRow previewRows, "Bullet list items", previewBullets, False, (Not doBullets)
        AddPreviewSummaryRow previewRows, "Numbered list items", previewNumbering, False, (Not doNumbering)
        AddPreviewSummaryRow previewRows, "List indents", previewIndent, False, (Not doIndent)
        ShowPreviewActionsSummary Me, "List Normalizer", previewRows
        Exit Sub
    End If
    MarkCleanupStart "List Normalizer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - List Normalizer"
    On Error GoTo RunErr
    Dim cntHyphen As Long, cntBullets As Long, cntNumbering As Long, cntIndent As Long
    If doBullets   Then cntBullets   = NormalizeBullets(targetRange)
    If doHyphen    Then cntHyphen    = ConvertHyphenListsToBullets(targetRange)
    If doNumbering Then cntNumbering = NormalizeNumbering(targetRange)
    If doIndent    Then cntIndent    = FixListIndentation(targetRange)
    RemoveAllHighlighting targetRange
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
Private Sub HighlightManualListPrefix(ByVal p As Paragraph, ByVal isNumbered As Boolean)
    Dim visibleText As String
    Dim prefixLength As Long
    visibleText = ParagraphBodyText(p)
    If isNumbered Then
        prefixLength = InStr(visibleText, ".")
        If prefixLength > 0 Then
            Do While prefixLength < Len(visibleText)
                If Mid$(visibleText, prefixLength + 1, 1) <> " " And Mid$(visibleText, prefixLength + 1, 1) <> vbTab Then Exit Do
                prefixLength = prefixLength + 1
            Loop
        End If
    ElseIf Left$(visibleText, 2) = "- " Or Left$(visibleText, 2) = "* " Then
        prefixLength = 2
    End If
    If prefixLength > 0 Then ApplyPreviewHighlight ActiveDocument.Range(p.Range.Start, p.Range.Start + prefixLength)
End Sub
Private Function ConvertHyphenListsToBullets(scopeRange As Range) As Long
    Dim p As Paragraph, listText As String, cnt As Long
    Dim shadingColor As Long
    For Each p In scopeRange.Paragraphs
        If ParagraphOverlapsRangeOutsideTable(p, scopeRange) Then
            listText = Trim$(ParagraphBodyText(p))
            If Left$(listText, 2) = "- " Or Left$(listText, 2) = "* " Then
                shadingColor = p.Range.Shading.BackgroundPatternColor
                Set p = ReplaceAndRefreshListParagraph(p, Mid$(listText, 3))
                ApplyListBulletStyle p
                ApplyStandardListIndent p
                TightenListParagraphSpacing p
                RestoreListParagraphShading p, shadingColor
                cnt = cnt + 1
            End If
        End If
    Next p
    ConvertHyphenListsToBullets = cnt
End Function
Private Function NormalizeBullets(scopeRange As Range) As Long
    Dim p As Paragraph, cnt As Long
    For Each p In scopeRange.Paragraphs
        If ParagraphOverlapsRangeOutsideTable(p, scopeRange) Then
            If p.Range.ListFormat.ListType <> wdListNoNumbering Then
                If p.Range.ListFormat.ListType = wdListBullet Then
                    ApplyStandardListIndent p
                    TightenListParagraphSpacing p
                    cnt = cnt + 1
                End If
            End If
        End If
    Next p
    NormalizeBullets = cnt
End Function
Private Sub ApplyListBulletStyle(ByVal p As Paragraph)
    If p.Range.ListFormat.ListType <> wdListBullet Then p.Range.ListFormat.ApplyBulletDefault
End Sub
Private Function NormalizeNumbering(scopeRange As Range) As Long
    Dim p As Paragraph, listText As String, cnt As Long
    Dim inNumberRun As Boolean
    Dim shadingColor As Long
    For Each p In scopeRange.Paragraphs
        If ParagraphOverlapsRangeOutsideTable(p, scopeRange) Then
            listText = Trim$(ParagraphBodyText(p))
            If IsManualNumberedListText(listText) Then
                Dim pos As Long: pos = InStr(listText, ".")
                shadingColor = p.Range.Shading.BackgroundPatternColor
                If pos > 0 Then Set p = ReplaceAndRefreshListParagraph(p, LTrim$(Mid$(listText, pos + 1)))
                ApplyListNumberStyle p, inNumberRun
                ApplyStandardListIndent p
                TightenListParagraphSpacing p
                RestoreListParagraphShading p, shadingColor
                inNumberRun = True
                cnt = cnt + 1
            Else
                If Len(listText) > 0 Then inNumberRun = False
            End If
        End If
    Next p
    NormalizeNumbering = cnt
End Function
Private Sub ApplyListNumberStyle(ByVal p As Paragraph, ByVal continuePreviousList As Boolean)
    p.Range.ListFormat.ApplyListTemplateWithLevel _
        ListTemplate:=ListGalleries(wdNumberGallery).ListTemplates(1), _
        ContinuePreviousList:=continuePreviousList, _
        ApplyTo:=wdListApplyToWholeList, _
        DefaultListBehavior:=wdWord10ListBehavior
End Sub
Private Sub ResolveListParagraphActions(ByVal p As Paragraph, ByVal listText As String, ByVal doBullets As Boolean, ByVal doNumbering As Boolean, ByVal doIndent As Boolean, ByVal doHyphen As Boolean, ByRef actionHyphen As Boolean, ByRef actionBullets As Boolean, ByRef actionNumbering As Boolean, ByRef actionIndent As Boolean)
    actionHyphen = False
    actionBullets = False
    actionNumbering = False
    actionIndent = False

    If doHyphen And (Left$(listText, 2) = "- " Or Left$(listText, 2) = "* ") Then actionHyphen = True
    If doNumbering And IsManualNumberedListText(listText) Then actionNumbering = True

    If actionHyphen Or actionNumbering Then Exit Sub

    If doBullets And p.Range.ListFormat.ListType = wdListBullet Then actionBullets = True
    If doIndent And ListIndentNeedsFix(p) Then actionIndent = True
End Sub
Private Function IsManualNumberedListText(ByVal listText As String) As Boolean
    Dim pos As Long
    pos = InStr(listText, ".")
    If pos <= 1 Then Exit Function
    If Not Left$(listText, pos - 1) Like String$(Len(Left$(listText, pos - 1)), "#") Then Exit Function
    IsManualNumberedListText = (Len(Trim$(Mid$(listText, pos + 1))) > 0)
End Function
Private Function ReplaceAndRefreshListParagraph(ByVal p As Paragraph, ByVal replacementText As String) As Paragraph
    Dim paragraphStart As Long
    paragraphStart = p.Range.Start

    ReplaceParagraphBodyText p, replacementText
    Set ReplaceAndRefreshListParagraph = ActiveDocument.Range(paragraphStart, paragraphStart).Paragraphs(1)
End Function
Private Sub RestoreListParagraphShading(ByVal p As Paragraph, ByVal shadingColor As Long)
    p.Range.Shading.BackgroundPatternColor = shadingColor
End Sub
Private Sub ApplyStandardListIndent(ByVal p As Paragraph)
    With p.Range.ParagraphFormat
        .LeftIndent = InchesToPoints(0.5)
        .FirstLineIndent = InchesToPoints(-0.25)
    End With
End Sub
Private Function ListIndentNeedsFix(ByVal p As Paragraph) As Boolean
    If p.Range.ListFormat.ListType = wdListNoNumbering Then Exit Function
    ListIndentNeedsFix = (Abs(p.Range.ParagraphFormat.LeftIndent - InchesToPoints(0.5)) > 0.5 Or Abs(p.Range.ParagraphFormat.FirstLineIndent - InchesToPoints(-0.25)) > 0.5)
End Function
Private Sub TightenListParagraphSpacing(ByVal p As Paragraph)
    With p.Range.ParagraphFormat
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceSingle
    End With
End Sub
Private Function FixListIndentation(scopeRange As Range) As Long
    Dim p As Paragraph, cnt As Long
    For Each p In scopeRange.Paragraphs
        If ParagraphOverlapsRangeOutsideTable(p, scopeRange) Then
            If ListIndentNeedsFix(p) Then
                ApplyStandardListIndent p
                TightenListParagraphSpacing p
                cnt = cnt + 1
            End If
        End If
    Next p
    FixListIndentation = cnt
End Function
