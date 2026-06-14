Option Explicit
' Shared helper functions for cleanup forms
Public gPreviewActionPanel As Object
Private gPreviewActionPanelHasPosition As Boolean
Private gPreviewActionPanelLeft As Single
Private gPreviewActionPanelTop As Single
Private gCleanupToolExitReason As String
Private Const CLEANUP_TOOL_EXIT_APPLY As String = "apply"
Private Const CLEANUP_TOOL_EXIT_CLOSE As String = "close"

Public Sub RememberPreviewActionPanelPosition(panelLeft As Single, panelTop As Single)
    gPreviewActionPanelLeft = panelLeft
    gPreviewActionPanelTop = panelTop
    gPreviewActionPanelHasPosition = True
End Sub

Private Sub ApplyPreviewActionPanelPosition(ByVal panel As Object)
    On Error Resume Next
    If gPreviewActionPanelHasPosition Then
        panel.StartUpPosition = 0
        panel.Left = gPreviewActionPanelLeft
        panel.Top = gPreviewActionPanelTop
    End If
    On Error GoTo 0
End Sub

Public Sub ShowPreviewActions(ByVal sourceForm As Object, toolName As String, summaryText As String)
    On Error GoTo Fallback
    sourceForm.Hide
    Set gPreviewActionPanel = New frmPreviewActions
    gPreviewActionPanel.Configure sourceForm, toolName, summaryText
    ApplyPreviewActionPanelPosition gPreviewActionPanel
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
Public Function GetShowCompletionReviewAfterApplySetting() As Boolean
    ' Returns True by default so users see the completion summary unless they opt out.
    On Error Resume Next
    Dim prop As Object
    Set prop = ActiveDocument.CustomDocumentProperties("CleanupSuiteShowCompletionReviewAfterApply")
    If Err.Number <> 0 Then
        GetShowCompletionReviewAfterApplySetting = True
        Err.Clear
    Else
        GetShowCompletionReviewAfterApplySetting = (prop.Value = "True")
    End If
    On Error GoTo 0
End Function
Public Sub SetShowCompletionReviewAfterApplySetting(enabled As Boolean)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("CleanupSuiteShowCompletionReviewAfterApply").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteShowCompletionReviewAfterApply", _
        LinkToContent:=False, Type:=msoPropertyTypeString, _
        Value:=IIf(enabled, "True", "False")
    On Error GoTo 0
End Sub
Public Function GetReturnToMainAfterCloseSetting() As Boolean
    ' Returns True by default so closing a tool menu restores the launcher.
    On Error Resume Next
    Dim prop As Object
    Set prop = ActiveDocument.CustomDocumentProperties("CleanupSuiteReturnToMainAfterClose")
    If Err.Number <> 0 Then
        GetReturnToMainAfterCloseSetting = True
        Err.Clear
    Else
        GetReturnToMainAfterCloseSetting = (prop.Value = "True")
    End If
    On Error GoTo 0
End Function
Public Sub SetReturnToMainAfterCloseSetting(enabled As Boolean)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("CleanupSuiteReturnToMainAfterClose").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteReturnToMainAfterClose", _
        LinkToContent:=False, Type:=msoPropertyTypeString, _
        Value:=IIf(enabled, "True", "False")
    On Error GoTo 0
End Sub
Public Sub BeginCleanupToolSession()
    gCleanupToolExitReason = ""
End Sub
Public Sub MarkCleanupToolApplied()
    gCleanupToolExitReason = CLEANUP_TOOL_EXIT_APPLY
End Sub
Public Sub MarkCleanupToolClosedByUser()
    If Len(gCleanupToolExitReason) = 0 Then gCleanupToolExitReason = CLEANUP_TOOL_EXIT_CLOSE
End Sub
Public Function ShouldReturnToLauncherAfterToolSession() As Boolean
    Select Case gCleanupToolExitReason
        Case CLEANUP_TOOL_EXIT_APPLY
            ShouldReturnToLauncherAfterToolSession = GetReturnToMainAfterApplySetting()
        Case CLEANUP_TOOL_EXIT_CLOSE
            ShouldReturnToLauncherAfterToolSession = GetReturnToMainAfterCloseSetting()
        Case Else
            ShouldReturnToLauncherAfterToolSession = False
    End Select
    gCleanupToolExitReason = ""
End Function
Public Sub ReturnToMainAfterToolClose(ByVal closingForm As Object, ByRef Cancel As Integer, ByVal CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MarkCleanupToolClosedByUser
    End If
End Sub
Public Sub ResetCleanupSuiteDefaults()
    On Error Resume Next
    RemoveAllHighlighting ActiveDocument.Content
    SetAutoSaveSetting True
    SetReturnToMainAfterApplySetting True
    SetShowCompletionReviewAfterApplySetting True
    SetReturnToMainAfterCloseSetting True
    BeginCleanupToolSession
    gPreviewActionPanelHasPosition = False
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
Public Sub LayoutCleanupToolForm(ByVal toolForm As Object)
    On Error Resume Next
    Const FORM_W As Single = 420
    Const M As Single = 14
    Const GAP As Single = 6
    Const BH As Single = 24
    Dim contentW As Single
    contentW = FORM_W - (2 * M)

    toolForm.Width = FORM_W + 8
    toolForm.BackColor = RGB(248, 249, 251)
    ApplyMSFormTitleStrategy toolForm, True
    toolForm.Controls("chkPreviewOnly").Visible = False
    toolForm.Controls("fraScopeSelection").Visible = False

    Dim titleLabel As Object
    Dim introLabel As Object
    Set titleLabel = EnsureGuidedLabel(toolForm, "lblGuidedToolName")
    Set introLabel = EnsureGuidedLabel(toolForm, "lblGuidedIntro")

    With titleLabel
        .Caption = GuidedToolName(toolForm.Name)
        .Move M, 8, contentW, 18
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = True
        .ForeColor = RGB(32, 37, 45)
        .BackColor = RGB(248, 249, 251)
        .TextAlign = 2
        .WordWrap = False
        .Visible = True
        .ZOrder 0
    End With

    With introLabel
        .Caption = GuidedToolIntro(toolForm.Name)
        .Move M, 28, contentW, 30
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .Font.Bold = False
        .ForeColor = RGB(88, 101, 116)
        .BackColor = RGB(248, 249, 251)
        .TextAlign = 2
        .WordWrap = True
        .Visible = True
        .ZOrder 0
    End With

    Dim y As Single
    y = 66
    Dim pendingButton As String
    pendingButton = ""
    Dim pendingChoice As String
    pendingChoice = ""
    Dim ctl As Object
    For Each ctl In toolForm.Controls
        If ShouldSkipGuidedLayoutControl(ctl.Name) Then GoTo NextGuidedControl
        If GuidedControlMatchesToolTitle(toolForm, ctl) Then
            HideGuidedControl ctl
            GoTo NextGuidedControl
        End If
        If ShouldHideGuidedControl(toolForm, ctl.Name) Then
            FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
            FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
            HideGuidedControl ctl
            GoTo NextGuidedControl
        End If
        If Left$(ctl.Name, 3) = "fra" Then
            FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
            FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
            HideGuidedControl ctl
            GoTo NextGuidedControl
        End If

        Select Case ctl.Name
            Case "cmdPreview", "cmdRun", "cmdReset", "cmdHelp"
                ' action buttons are laid out after scope
            Case Else
                Select Case Left$(ctl.Name, 3)
                    Case "cmd"
                        FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
                        If Len(pendingButton) > 0 Then
                            StyleGuidedButton toolForm.Controls(pendingButton), False
                            StyleGuidedButton ctl, False
                            toolForm.Controls(pendingButton).Move M, y, (contentW - GAP) / 2, BH
                            ctl.Move M + ((contentW - GAP) / 2) + GAP, y, (contentW - GAP) / 2, BH
                            toolForm.Controls(pendingButton).Visible = True
                            ctl.Visible = True
                            y = y + BH + GAP
                            pendingButton = ""
                        Else
                            pendingButton = ctl.Name
                        End If
                    Case "lbl"
                        FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
                        FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
                        StyleGuidedBodyLabel ctl
                        ctl.Move M, y, contentW, GuidedLabelHeight(ctl)
                        ctl.Visible = True
                        y = y + ctl.Height + GAP
                    Case "opt", "chk"
                        FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
                        If Len(pendingChoice) > 0 Then
                            PositionGuidedChoicePair toolForm, pendingChoice, ctl, M, y, contentW, GAP
                        Else
                            pendingChoice = ctl.Name
                        End If
                    Case Else
                        FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
                        FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
                End Select
        End Select
NextGuidedControl:
    Next ctl
    FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
    FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
    y = y + 2

    LayoutGuidedScopeRow toolForm, M, y, contentW
    LayoutGuidedActionButtons toolForm, M, y, contentW
    toolForm.Height = y + 34
End Sub
Private Function EnsureGuidedLabel(ByVal toolForm As Object, ByVal controlName As String) As Object
    On Error Resume Next
    Set EnsureGuidedLabel = toolForm.Controls(controlName)
    On Error GoTo 0
    If EnsureGuidedLabel Is Nothing Then
        On Error Resume Next
        Set EnsureGuidedLabel = toolForm.Controls.Add("Forms.Label.1", controlName, True)
        On Error GoTo 0
    End If
End Function
Private Function ShouldSkipGuidedLayoutControl(ByVal controlName As String) As Boolean
    Select Case controlName
        Case "lblGuidedToolName", "lblGuidedIntro", "chkPreviewOnly", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdPreview", "cmdRun", "cmdReset", "cmdHelp"
            ShouldSkipGuidedLayoutControl = True
        Case Else
            ShouldSkipGuidedLayoutControl = False
    End Select
End Function
Private Sub HideGuidedControl(ByVal ctl As Object)
    On Error Resume Next
    ctl.Visible = False
    ctl.Move 0, 0, 0, 0
End Sub
Private Function GuidedControlMatchesToolTitle(ByVal toolForm As Object, ByVal ctl As Object) As Boolean
    On Error Resume Next
    If ctl.Name = "lblGuidedToolName" Then Exit Function
    If Left$(LCase$(ctl.Name), 3) <> "lbl" Then Exit Function
    GuidedControlMatchesToolTitle = (Trim$(ctl.Caption) = GuidedToolName(toolForm.Name))
End Function
Private Function ShouldHideGuidedControl(ByVal toolForm As Object, ByVal controlName As String) As Boolean
    On Error Resume Next
    ShouldHideGuidedControl = False

    If Not GuidedCustomModeIsActive(toolForm) Then
        Select Case controlName
            Case "chkNBSP", "chkZWSP", "chkZWNJ", "chkZWJ", "chkBOM", "chkSoftHyphen", "chkNBHyphen", _
                 "chkCurlyDouble", "chkCurlySingle", "chkEmDash", "chkEnDash", "chkEllipses", _
                 "chkDoubleSpaces", "chkTrimSpaces", "chkSpaceBeforePunct", "chkNormalizeAfterPunct", "chkExtraBlankLines", _
                 "chkNormalizeBullets", "chkNormalizeNumbering", "chkFixIndent", "chkHyphenToBullets", _
                 "chkRemoveEmpty", "chkCollapseBreaks", "chkNormalizeParaSpacing", _
                 "cmdSelectAll", "cmdDeselectAll"
                ShouldHideGuidedControl = True
        End Select
    End If

    If toolForm.Name = "frmBreakNormalizer" And Not toolForm.Controls("chkConvertSectionBreaks").Value Then
        Select Case controlName
            Case "optConvertNextPage", "optConvertContinuous", "optConvertEvenPage", "optConvertOddPage"
                ShouldHideGuidedControl = True
        End Select
    End If

    If toolForm.Name = "frmDuplicateDetector" And Not toolForm.Controls("optMatchFuzzy").Value Then
        Select Case controlName
            Case "optFuzzyLoose", "optFuzzyMedium", "optFuzzyStrict"
                ShouldHideGuidedControl = True
        End Select
    End If

    If toolForm.Name = "frmHeaderFooterStandardizer" Then
        If toolForm.Controls("optClearAll").Value Then
            Select Case controlName
                Case "chkFont", "chkSpacing", "chkAlignment", "optAlignLeft", "optAlignCenter", "optAlignRight", "chkBreakLinks"
                    ShouldHideGuidedControl = True
            End Select
        ElseIf Not toolForm.Controls("chkAlignment").Value Then
            Select Case controlName
                Case "optAlignLeft", "optAlignCenter", "optAlignRight"
                    ShouldHideGuidedControl = True
            End Select
        End If
    End If
End Function
Private Function GuidedCustomModeIsActive(ByVal toolForm As Object) As Boolean
    On Error Resume Next
    GuidedCustomModeIsActive = toolForm.Controls("optCustom").Value
End Function
Private Sub FlushGuidedPendingChoice(ByVal toolForm As Object, ByRef pendingChoice As String, ByVal M As Single, ByRef y As Single, ByVal contentW As Single, ByVal GAP As Single)
    On Error Resume Next
    If Len(pendingChoice) = 0 Then Exit Sub
    Dim choiceW As Single
    choiceW = (contentW - GAP) / 2
    Dim rowH As Single
    rowH = GuidedChoiceRowHeight(toolForm.Controls(pendingChoice))
    StyleGuidedOption toolForm.Controls(pendingChoice)
    toolForm.Controls(pendingChoice).Move M + 8, y, choiceW - 8, rowH
    toolForm.Controls(pendingChoice).Visible = True
    y = y + rowH + GAP
    pendingChoice = ""
End Sub
Private Sub PositionGuidedChoicePair(ByVal toolForm As Object, ByRef pendingChoice As String, ByVal choiceCtl As Object, ByVal M As Single, ByRef y As Single, ByVal contentW As Single, ByVal GAP As Single)
    On Error Resume Next
    Dim choiceW As Single
    choiceW = (contentW - GAP) / 2
    Dim rowH As Single
    rowH = GuidedChoiceRowHeight(toolForm.Controls(pendingChoice), choiceCtl)
    StyleGuidedOption toolForm.Controls(pendingChoice)
    StyleGuidedOption choiceCtl
    toolForm.Controls(pendingChoice).Move M + 8, y, choiceW - 8, rowH
    choiceCtl.Move M + choiceW + GAP + 8, y, choiceW - 8, rowH
    toolForm.Controls(pendingChoice).Visible = True
    choiceCtl.Visible = True
    y = y + rowH + GAP
    pendingChoice = ""
End Sub
Private Function GuidedChoiceRowHeight(ByVal firstControl As Object, Optional ByVal secondControl As Object = Nothing) As Single
    On Error Resume Next
    GuidedChoiceRowHeight = GuidedOptionHeight(firstControl)
    If Not secondControl Is Nothing Then
        If GuidedOptionHeight(secondControl) > GuidedChoiceRowHeight Then GuidedChoiceRowHeight = GuidedOptionHeight(secondControl)
    End If
End Function
Private Sub FlushGuidedPendingButton(ByVal toolForm As Object, ByRef pendingButton As String, ByVal M As Single, ByRef y As Single, ByVal contentW As Single, ByVal BH As Single, ByVal GAP As Single)
    On Error Resume Next
    If Len(pendingButton) = 0 Then Exit Sub
    StyleGuidedButton toolForm.Controls(pendingButton), False
    toolForm.Controls(pendingButton).Move M, y, contentW, BH
    toolForm.Controls(pendingButton).Visible = True
    y = y + BH + GAP
    pendingButton = ""
End Sub
Private Sub StyleGuidedBodyLabel(ByVal ctl As Object)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 8.5
    ctl.Font.Bold = False
    ctl.ForeColor = RGB(67, 80, 96)
    ctl.BackColor = RGB(255, 255, 255)
    ctl.BorderStyle = 1
    ctl.SpecialEffect = 0
    ctl.WordWrap = True
End Sub
Private Sub StyleGuidedOption(ByVal ctl As Object)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 8.5
    ctl.Font.Bold = False
    ctl.ForeColor = RGB(32, 37, 45)
    ctl.BackColor = RGB(248, 249, 251)
    ctl.WordWrap = True
    ctl.AutoSize = False
End Sub
Private Sub StyleGuidedButton(ByVal ctl As Object, Optional ByVal primary As Boolean = False)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 9
    ctl.Font.Bold = primary
    ctl.BackColor = RGB(255, 255, 255)
    If primary Then
        ctl.ForeColor = RGB(24, 90, 145)
    Else
        ctl.ForeColor = RGB(32, 37, 45)
    End If
    ctl.Enabled = True
End Sub
Private Function GuidedLabelHeight(ByVal ctl As Object) As Single
    On Error Resume Next
    If Len(ctl.Caption) > 72 Then
        GuidedLabelHeight = 42
    Else
        GuidedLabelHeight = 32
    End If
End Function
Private Function GuidedOptionHeight(ByVal ctl As Object) As Single
    On Error Resume Next
    If Len(ctl.Caption) > 72 Then
        GuidedOptionHeight = 44
    ElseIf Len(ctl.Caption) > 34 Then
        GuidedOptionHeight = 28
    Else
        GuidedOptionHeight = 18
    End If
End Function
Private Sub LayoutGuidedScopeRow(ByVal toolForm As Object, ByVal M As Single, ByRef y As Single, ByVal contentW As Single)
    On Error Resume Next
    Const GAP As Single = 6
    Dim halfW As Single
    halfW = (contentW - GAP) / 2
    StyleGuidedOption toolForm.Controls("optScopeDocument")
    StyleGuidedOption toolForm.Controls("optScopeSelection")
    toolForm.Controls("optScopeDocument").Move M, y, halfW, 18
    toolForm.Controls("optScopeSelection").Move M + halfW + GAP, y, halfW, 18
    toolForm.Controls("optScopeDocument").Visible = True
    If Len(Trim$(toolForm.Controls("optScopeSelection").Caption)) = 0 Then
        HideGuidedControl toolForm.Controls("optScopeSelection")
    Else
        toolForm.Controls("optScopeSelection").Visible = True
    End If
    y = y + 28
End Sub
Private Sub LayoutGuidedActionButtons(ByVal toolForm As Object, ByVal M As Single, ByRef y As Single, ByVal contentW As Single)
    On Error Resume Next
    Const GAP As Single = 6
    Const BH As Single = 24
    StyleGuidedButton toolForm.Controls("cmdPreview"), True
    toolForm.Controls("cmdPreview").Caption = "Preview"
    toolForm.Controls("cmdPreview").Move M, y, contentW, BH
    toolForm.Controls("cmdPreview").Visible = True
    y = y + BH + GAP

    Dim thirdW As Single
    thirdW = (contentW - (2 * GAP)) / 3
    StyleGuidedButton toolForm.Controls("cmdRun"), False
    StyleGuidedButton toolForm.Controls("cmdReset"), False
    StyleGuidedButton toolForm.Controls("cmdHelp"), False
    toolForm.Controls("cmdRun").Caption = "Apply"
    toolForm.Controls("cmdReset").Caption = "Reset"
    toolForm.Controls("cmdHelp").Caption = "Help"
    toolForm.Controls("cmdRun").Move M, y, thirdW, BH
    toolForm.Controls("cmdReset").Move M + thirdW + GAP, y, thirdW, BH
    toolForm.Controls("cmdHelp").Move M + (2 * (thirdW + GAP)), y, thirdW, BH
    toolForm.Controls("cmdRun").Visible = True
    toolForm.Controls("cmdReset").Visible = True
    toolForm.Controls("cmdHelp").Visible = True
    y = y + BH + 28
End Sub
Private Function GuidedToolName(ByVal formName As String) As String
    Select Case formName
        Case "frmPunctuationCleanup": GuidedToolName = "Punctuation Normalizer"
        Case "frmUnicodeCleanup": GuidedToolName = "Invisible Unicode Cleaner"
        Case "frmSpacingCleanup": GuidedToolName = "Spacing Fixer"
        Case "frmListCleanup": GuidedToolName = "List Normalizer"
        Case "frmParagraphCleanup": GuidedToolName = "Paragraph Structure Fixer"
        Case "frmDuplicateDetector": GuidedToolName = "Duplicate Paragraph Detector"
        Case "frmFontNormalizer": GuidedToolName = "Font Normalizer"
        Case "frmTableCleaner": GuidedToolName = "Table Cleaner"
        Case "frmBreakNormalizer": GuidedToolName = "Break Normalizer"
        Case "frmDocumentTrim": GuidedToolName = "Document Trim"
        Case "frmFormattingStripper": GuidedToolName = "Formatting Stripper"
        Case "frmHyperlinkRemover": GuidedToolName = "Hyperlink Remover"
        Case "frmSoftReturnConverter": GuidedToolName = "Soft Return Converter"
        Case "frmMetadataScrubber": GuidedToolName = "Metadata Scrubber"
        Case "frmStyleCleanup": GuidedToolName = "Style Cleanup"
        Case "frmFootnoteRemover": GuidedToolName = "Footnote / Endnote Remover"
        Case "frmHeaderFooterStandardizer": GuidedToolName = "Header / Footer Standardizer"
        Case "frmObjectRemover": GuidedToolName = "Object Remover"
        Case Else: GuidedToolName = formName
    End Select
End Function
Private Function GuidedToolIntro(ByVal formName As String) As String
    Select Case formName
        Case "frmPunctuationCleanup": GuidedToolIntro = "Choose which punctuation patterns to normalize."
        Case "frmUnicodeCleanup": GuidedToolIntro = "Choose which invisible or problem characters to clean."
        Case "frmSpacingCleanup": GuidedToolIntro = "Choose which spacing problems to fix."
        Case "frmListCleanup": GuidedToolIntro = "Choose how lists, bullets, numbering, and indents should be cleaned."
        Case "frmParagraphCleanup": GuidedToolIntro = "Choose which paragraph structure issues to clean."
        Case "frmDuplicateDetector": GuidedToolIntro = "Choose how repeated or similar paragraphs should be reviewed."
        Case "frmFontNormalizer": GuidedToolIntro = "Choose which direct font overrides to reset."
        Case "frmTableCleaner": GuidedToolIntro = "Choose which table structure and formatting issues to clean."
        Case "frmBreakNormalizer": GuidedToolIntro = "Choose how page and section breaks should be normalized."
        Case "frmDocumentTrim": GuidedToolIntro = "Remove extra empty paragraphs at the end of the document."
        Case "frmFormattingStripper": GuidedToolIntro = "Choose how direct formatting should be stripped while keeping content."
        Case "frmHyperlinkRemover": GuidedToolIntro = "Choose whether link styling should be removed with the hyperlinks."
        Case "frmSoftReturnConverter": GuidedToolIntro = "Choose which line-break conversion to apply."
        Case "frmMetadataScrubber": GuidedToolIntro = "Choose which document metadata and review artifacts to remove."
        Case "frmStyleCleanup": GuidedToolIntro = "Choose which style cleanup actions to run."
        Case "frmFootnoteRemover": GuidedToolIntro = "Choose which notes to remove and whether to keep their text."
        Case "frmHeaderFooterStandardizer": GuidedToolIntro = "Choose how headers and footers should be standardized or cleared."
        Case "frmObjectRemover": GuidedToolIntro = "Choose which embedded objects or hidden content to remove."
        Case Else: GuidedToolIntro = "Choose the cleanup options for this tool."
    End Select
End Function
Public Sub ApplyMSFormTitleStrategy(ByVal toolForm As Object, Optional bodyOwnsTitle As Boolean = False)
    ' Method 1: if the form body already identifies the workflow, blank the native UserForm caption.
    ' Method 2: otherwise keep the native caption and hide only legacy body-title labels.
    On Error Resume Next
    If bodyOwnsTitle Then toolForm.Caption = ""
    HideLegacyTitleControls toolForm
    On Error GoTo 0
End Sub
Private Sub HideLegacyTitleControls(ByVal toolForm As Object)
    On Error Resume Next
    Dim ctl As Object
    For Each ctl In toolForm.Controls
        If IsLegacyTitleControlName(ctl.Name) Or GuidedControlMatchesToolTitle(toolForm, ctl) Then
            ctl.Visible = False
            ctl.Move 0, 0, 0, 0
        End If
    Next ctl
End Sub
Private Function IsLegacyTitleControlName(controlName As String) As Boolean
    Dim nm As String
    nm = LCase$(controlName)
    IsLegacyTitleControlName = (nm = "lbltitle" Or Left$(nm, 8) = "lbltitle")
End Function
Public Sub LayoutCapitalizationCleanupForm(ByVal toolForm As Object)
    On Error Resume Next
    Const M As Single = 14
    Const GAP As Single = 6
    Const BH As Single = 24
    Const FORM_W As Single = 420
    Dim contentW As Single
    contentW = FORM_W - (2 * M)

    toolForm.Width = FORM_W + 8
    toolForm.BackColor = RGB(248, 249, 251)
    ApplyMSFormTitleStrategy toolForm, True
    toolForm.Controls("chkPreviewOnly").Visible = False
    toolForm.Controls("fraScopeSelection").Visible = False

    With toolForm.Controls("lblIntro")
        .Move M, 10, contentW, 18
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = False
        .ForeColor = RGB(88, 101, 116)
        .BackColor = RGB(248, 249, 251)
        .TextAlign = 2
        .WordWrap = False
        .Visible = True
    End With

    Dim y As Single
    y = 36
    Dim optionW As Single
    optionW = (contentW - GAP) / 2
    PositionGuidedChoice toolForm, "optAll", M, y, optionW
    PositionGuidedChoice toolForm, "optSentence", M + optionW + GAP, y, optionW
    y = y + 24
    PositionGuidedChoice toolForm, "optTitle", M, y, optionW
    PositionGuidedChoice toolForm, "optUpper", M + optionW + GAP, y, optionW
    y = y + 24
    PositionGuidedChoice toolForm, "optLower", M, y, optionW
    PositionGuidedChoice toolForm, "optCustom", M + optionW + GAP, y, optionW
    y = y + 28

    With toolForm.Controls("lblModeSummary")
        .Move M, y, contentW, 34
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .Font.Bold = False
        .ForeColor = RGB(67, 80, 96)
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = 1
        .SpecialEffect = 0
        .WordWrap = True
        .Visible = True
    End With
    y = y + 42

    toolForm.Controls("fraCustom").Visible = False
    If toolForm.Controls("optCustom").Value Then
        PositionGuidedCheck toolForm, "chkSentence", M + 8, y, contentW - 16
        PositionGuidedCheck toolForm, "chkTitle", M + 8, y + 18, contentW - 16
        PositionGuidedCheck toolForm, "chkUpper", M + 8, y + 36, contentW - 16
        PositionGuidedCheck toolForm, "chkLower", M + 8, y + 54, contentW - 16
        PositionGuidedCheck toolForm, "chkSmartSentences", M + 8, y + 72, contentW - 16
        y = y + 94
        With toolForm.Controls("cmdSelectAll")
            .Move M, y, (contentW - GAP) / 2, BH
            .Visible = True
            .ZOrder 0
        End With
        With toolForm.Controls("cmdDeselectAll")
            .Move M + ((contentW - GAP) / 2) + GAP, y, (contentW - GAP) / 2, BH
            .Visible = True
            .ZOrder 0
        End With
        y = y + BH + 12
    Else
        HideGuidedCustom toolForm
    End If

    With toolForm.Controls("optScopeDocument")
        .Move M, y, (contentW - GAP) / 2, 18
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .BackColor = RGB(248, 249, 251)
        .Visible = True
    End With
    With toolForm.Controls("optScopeSelection")
        .Move M + ((contentW - GAP) / 2) + GAP, y, (contentW - GAP) / 2, 18
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .BackColor = RGB(248, 249, 251)
        .Visible = True
    End With
    y = y + 28

    With toolForm.Controls("cmdPreview")
        .Move M, y, contentW, BH
        .Caption = "Preview"
        .Font.Bold = True
        .Visible = True
        .Enabled = True
    End With
    y = y + BH + GAP
    With toolForm.Controls("cmdRun")
        .Move M, y, (contentW - (2 * GAP)) / 3, BH
        .Caption = "Apply"
        .Visible = True
        .Enabled = True
    End With
    With toolForm.Controls("cmdReset")
        .Move M + ((contentW - (2 * GAP)) / 3) + GAP, y, (contentW - (2 * GAP)) / 3, BH
        .Caption = "Reset"
        .Visible = True
        .Enabled = True
    End With
    With toolForm.Controls("cmdHelp")
        .Move M + (2 * (((contentW - (2 * GAP)) / 3) + GAP)), y, (contentW - (2 * GAP)) / 3, BH
        .Caption = "Help"
        .Visible = True
        .Enabled = True
    End With
    y = y + BH + 28

    toolForm.Height = y + 34
End Sub
Private Sub PositionGuidedChoice(ByVal toolForm As Object, nm As String, L As Single, T As Single, W As Single)
    On Error Resume Next
    With toolForm.Controls(nm)
        .Move L, T, W, 18
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .Font.Bold = False
        .BackColor = RGB(248, 249, 251)
        .ForeColor = RGB(32, 37, 45)
        .Visible = True
    End With
End Sub
Private Sub PositionGuidedCheck(ByVal toolForm As Object, nm As String, L As Single, T As Single, W As Single)
    On Error Resume Next
    With toolForm.Controls(nm)
        .Move L, T, W, 16
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .BackColor = RGB(248, 249, 251)
        .Visible = True
        .ZOrder 0
    End With
End Sub
Private Sub HideGuidedCustom(ByVal toolForm As Object)
    On Error Resume Next
    toolForm.Controls("fraCustom").Visible = False
    toolForm.Controls("chkSentence").Visible = False
    toolForm.Controls("chkTitle").Visible = False
    toolForm.Controls("chkUpper").Visible = False
    toolForm.Controls("chkLower").Visible = False
    toolForm.Controls("chkSmartSentences").Visible = False
    toolForm.Controls("cmdSelectAll").Visible = False
    toolForm.Controls("cmdDeselectAll").Visible = False
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
    If Not GetShowCompletionReviewAfterApplySetting() Then Exit Sub
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
