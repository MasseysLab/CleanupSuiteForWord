Option Explicit
Private Sub UserForm_Initialize()
    ApplyMSFormTitleStrategy Me, True
    lblLauncherTitle.Caption = "Cleanup Suite"
    lblLauncherTitle.TextAlign = 1
    lblLauncherSubtitle.Caption = "Choose the kind of cleanup you need. Tools are grouped by what you are trying to fix."
    lblLauncherSubtitle.TextAlign = 1
    cmdFootnote.Caption = "Footnote / Endnote"
    RefreshLauncherIntentLabels
    RefreshReviewPrivacyLauncherSection
    chkReturnToMainAfterApply.Caption = "Global default: return to main menu after apply"
    chkReturnToMainAfterApply.Value = GetReturnToMainAfterApplySetting()
    chkReturnToMainAfterClose.Caption = "Global default: return to main menu after closing individual tool menus"
    chkReturnToMainAfterClose.Value = GetReturnToMainAfterCloseSetting()
    chkAutoSave.Caption = "Global default: auto-save before each tool"
    chkAutoSave.Value = GetAutoSaveSetting()
    chkUpdateChecks.Caption = "Check periodically for updates"
    chkUpdateChecks.Value = GetCleanupSuiteUpdateChecksEnabled()
    cmdCheckUpdates.Caption = "Check for Updates"
    lblAutoSaveWarning.Caption = "Dangerous! Should be selected"
    RefreshAutoSaveWarning
    RefreshLauncherFooterButtons
End Sub
Private Sub RefreshLauncherIntentLabels()
    On Error Resume Next
    lblCatText.Caption = "Clean Pasted Text"
    lblCatPara.Caption = "Fix Paragraphs and Lists"
    lblCatReview.Caption = "Review and Inspect"
    lblCatLayout.Caption = "Clean Document Objects"
    lblCatFormat.Caption = "Clean Formatting"
    lblCatShare.Caption = "Prepare for Sharing"
    lblCatAlpha.Caption = "Risky Alpha Tools"

    cmdDocTrim.Caption = "Trim End of Doc"
    cmdFormatStrip.Caption = "Formatting Cleaner"
    lblDescHyperlink.Caption = "Keep links but Hide styling" & vbCrLf & "Keep text but remove links"

    ApplyLauncherRiskLabel lblRiskUnicode, "Safe cleanup"
    ApplyLauncherRiskLabel lblRiskPunctuation, "Safe cleanup"
    ApplyLauncherRiskLabel lblRiskSpacing, "Safe cleanup"
    ApplyLauncherRiskLabel lblRiskCapitalization, "Text caution"
    ApplyLauncherRiskLabel lblRiskList, "Structure change"
    ApplyLauncherRiskLabel lblRiskParagraph, "Structure change"
    ApplyLauncherRiskLabel lblRiskSoftReturn, "Structure change"
    ApplyLauncherRiskLabel lblRiskDuplicate, "Inspect First"
    ApplyLauncherRiskLabel lblRiskTableClean, "Structure change"
    ApplyLauncherRiskLabel lblRiskHeaderFooter, "Structure change"
    ApplyLauncherRiskLabel lblRiskFootnote, "Removes content"
    ApplyLauncherRiskLabel lblRiskObjectRemover, "Removes content"
    ApplyLauncherRiskLabel lblRiskFontNorm, "Formatting change"
    ApplyLauncherRiskLabel lblRiskFormatStrip, "Formatting change"
    ApplyLauncherRiskLabel lblRiskMetadata, "Privacy cleanup"
    ApplyLauncherRiskLabel lblRiskFinalReview, "Finalizes review"
    ApplyLauncherRiskLabel lblRiskHyperlink, "Removes content"
    ApplyLauncherRiskLabel lblRiskStyleClean, "Alpha tool"
    ApplyLauncherRiskLabel lblRiskBreakNorm, "Alpha tool"
    ApplyLauncherRiskLabel lblRiskDocTrim, "Safe cleanup"

    LayoutLauncherIntentRows
End Sub
Private Sub ApplyLauncherRiskLabel(riskLabel As Object, ByVal labelText As String)
    On Error Resume Next
    riskLabel.Caption = LauncherRiskDisplayText(labelText)
    riskLabel.TextAlign = 2
    riskLabel.Font.Size = 7.5
    riskLabel.Font.Bold = True
    riskLabel.WordWrap = True
    riskLabel.TakeFocusOnClick = False
    riskLabel.BackStyle = 1
    riskLabel.SpecialEffect = 0
    riskLabel.BorderStyle = 1
    riskLabel.BackColor = LauncherRiskBackColor(labelText)
    riskLabel.ForeColor = LauncherRiskForeColor(labelText)
    riskLabel.BorderColor = riskLabel.ForeColor
    riskLabel.ControlTipText = LauncherRiskTip(labelText)
End Sub
Private Sub LayoutLauncherIntentRows()
    On Error Resume Next
    Const FORM_W As Single = 760
    Const MARGIN As Single = 14
    Const GAP As Single = 3
    Const COL_GAP As Single = 16
    Const COL_W As Single = (FORM_W - (2 * MARGIN) - COL_GAP) / 2
    Const COL1_X As Single = MARGIN
    Const COL2_X As Single = MARGIN + COL_W + COL_GAP
    Const TOP_Y As Single = 8
    Dim y As Single
    Dim yCol1 As Single
    Dim yCol2 As Single
    y = TOP_Y

    PositionLauncherControl lblLauncherTitle, COL1_X, y, COL_W, 18
    PositionLauncherControl lblLauncherSubtitle, COL1_X, y + 20, COL_W, 28
    lblLauncherTitle.TextAlign = 1
    lblLauncherSubtitle.TextAlign = 1

    yCol1 = y + 52
    yCol1 = LayoutLauncherCategory(lblCatText, COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(cmdHelpUnicode, cmdUnicode, lblDescUnicode, lblRiskUnicode, COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(cmdHelpPunct, cmdPunctuation, lblDescPunctuation, lblRiskPunctuation, COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(cmdHelpSpacing, cmdSpacing, lblDescSpacing, lblRiskSpacing, COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(cmdHelpCap, cmdCapitalization, lblDescCapitalization, lblRiskCapitalization, COL1_X, yCol1, COL_W)
    yCol1 = yCol1 + GAP
    yCol1 = LayoutLauncherCategory(lblCatPara, COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(cmdHelpList, cmdList, lblDescList, lblRiskList, COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(cmdHelpPara, cmdParagraph, lblDescParagraph, lblRiskParagraph, COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(cmdHelpSoftReturn, cmdSoftReturn, lblDescSoftReturn, lblRiskSoftReturn, COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(cmdHelpTrim, cmdDocTrim, lblDescDocTrim, lblRiskDocTrim, COL1_X, yCol1, COL_W)

    yCol2 = TOP_Y
    yCol2 = LayoutLauncherCategory(lblCatLayout, COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(cmdHelpTable, cmdTableClean, lblDescTableClean, lblRiskTableClean, COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(cmdHelpHeaderFooter, cmdHeaderFooter, lblDescHeaderFooter, lblRiskHeaderFooter, COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(cmdHelpFootnote, cmdFootnote, lblDescFootnote, lblRiskFootnote, COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(cmdHelpObject, cmdObjectRemover, lblDescObjectRemover, lblRiskObjectRemover, COL2_X, yCol2, COL_W)
    yCol2 = yCol2 + GAP
    yCol2 = LayoutLauncherCategory(lblCatFormat, COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(cmdHelpFont, cmdFontNorm, lblDescFontNorm, lblRiskFontNorm, COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(cmdHelpFormat, cmdFormatStrip, lblDescFormatStrip, lblRiskFormatStrip, COL2_X, yCol2, COL_W)
    yCol2 = yCol2 + GAP
    yCol2 = LayoutLauncherCategory(lblCatReview, COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(cmdHelpDuplicate, cmdDuplicate, lblDescDuplicate, lblRiskDuplicate, COL2_X, yCol2, COL_W)
    yCol2 = yCol2 + GAP
    yCol2 = LayoutLauncherCategory(lblCatShare, COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(cmdHelpMetadata, cmdMetadata, lblDescMetadata, lblRiskMetadata, COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(cmdHelpFinalReview, cmdFinalReview, lblDescFinalReview, lblRiskFinalReview, COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(cmdHelpHyperlink, cmdHyperlink, lblDescHyperlink, lblRiskHyperlink, COL2_X, yCol2, COL_W)
    yCol2 = yCol2 + GAP
    yCol2 = LayoutLauncherCategory(lblCatAlpha, COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(cmdHelpStyle, cmdStyleClean, lblDescStyleClean, lblRiskStyleClean, COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(cmdHelpBreak, cmdBreakNorm, lblDescBreakNorm, lblRiskBreakNorm, COL2_X, yCol2, COL_W)

    y = yCol1 + GAP
    PositionLauncherControl chkAutoSave, MARGIN + 2, y, 225, 18
    PositionLauncherControl lblAutoSaveWarning, MARGIN + 18, y + 16, COL_W - 18, 14
    RefreshAutoSaveWarning
    y = y + 32
    PositionLauncherControl chkReturnToMainAfterApply, MARGIN + 2, y, COL_W, 18
    y = y + 18
    PositionLauncherControl chkReturnToMainAfterClose, MARGIN + 2, y, COL_W, 18
    y = y + 20
    PositionLauncherControl chkUpdateChecks, MARGIN + 2, y, 214, 18
    PositionLauncherControl cmdCheckUpdates, MARGIN + 230, y - 3, 112, 24
    y = y + 26
    PositionLauncherControl cmdUserManual, MARGIN + 2, y, 96, 24
    PositionLauncherControl cmdResetAll, MARGIN + 104, y, 96, 24
    RefreshResetAllCallout
    Me.Width = FORM_W + 8
    Me.Height = MaxSingle(MaxSingle(lblResetAllCalloutBox.Top + lblResetAllCalloutBox.Height, lblResetAllCalloutText.Top + lblResetAllCalloutText.Height), yCol2) + 40
End Sub
Private Function LayoutLauncherCategory(catCtl As Object, ByVal L As Single, ByVal T As Single, ByVal W As Single) As Single
    On Error Resume Next
    PositionLauncherControl catCtl, L, T, W, 16
    LayoutLauncherCategory = T + 18
End Function
Private Function LayoutLauncherToolRow(helpCtl As Object, buttonCtl As Object, descCtl As Object, riskCtl As Object, ByVal L As Single, ByVal T As Single, ByVal W As Single) As Single
    On Error Resume Next
    Const ROW_H As Single = 30
    Const ROW_STEP As Single = 30
    Const BTN_W As Single = 108
    Const HELP_W As Single = 20
    Const RISK_W As Single = 56
    Const RISK_H As Single = 28
    Const INNER_GAP As Single = 6
    Dim descW As Single
    descW = W - BTN_W - HELP_W - RISK_W - (3 * INNER_GAP)

    PositionLauncherControl helpCtl, L, T + 5, HELP_W, 20
    PositionLauncherControl buttonCtl, L + HELP_W + INNER_GAP, T + 3, BTN_W, 24
    PositionLauncherControl riskCtl, L + HELP_W + BTN_W + (2 * INNER_GAP), T + 1, RISK_W, RISK_H
    PositionLauncherControl descCtl, L + HELP_W + BTN_W + RISK_W + (3 * INNER_GAP), T + 1, descW, ROW_H
    riskCtl.WordWrap = True
    LayoutLauncherToolRow = T + ROW_STEP
End Function
Private Sub PositionLauncherControl(ctl As Object, ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single)
    On Error Resume Next
    ctl.Left = L
    ctl.Top = T
    ctl.Width = W
    ctl.Height = H
    ctl.Visible = True
End Sub
Private Function MaxSingle(a As Single, b As Single) As Single
    If a >= b Then
        MaxSingle = a
    Else
        MaxSingle = b
    End If
End Function
Private Function LauncherRiskBackColor(labelText As String) As Long
    Select Case labelText
        Case "Safe cleanup": LauncherRiskBackColor = RGB(229, 245, 235)
        Case "Inspect First": LauncherRiskBackColor = RGB(228, 239, 251)
        Case "Text caution", "Formatting change": LauncherRiskBackColor = RGB(255, 246, 218)
        Case "Structure change", "Removes content": LauncherRiskBackColor = RGB(255, 235, 218)
        Case "Privacy cleanup", "Finalizes review": LauncherRiskBackColor = RGB(255, 226, 226)
        Case "Alpha tool": LauncherRiskBackColor = RGB(235, 236, 240)
        Case Else: LauncherRiskBackColor = RGB(240, 242, 245)
    End Select
End Function
Private Function LauncherRiskForeColor(labelText As String) As Long
    Select Case labelText
        Case "Safe cleanup": LauncherRiskForeColor = RGB(24, 105, 59)
        Case "Inspect First": LauncherRiskForeColor = RGB(31, 91, 145)
        Case "Text caution", "Formatting change": LauncherRiskForeColor = RGB(132, 89, 0)
        Case "Structure change", "Removes content": LauncherRiskForeColor = RGB(154, 73, 24)
        Case "Privacy cleanup", "Finalizes review": LauncherRiskForeColor = RGB(171, 40, 40)
        Case "Alpha tool": LauncherRiskForeColor = RGB(67, 72, 82)
        Case Else: LauncherRiskForeColor = RGB(84, 91, 104)
    End Select
End Function
Private Function LauncherRiskTip(labelText As String) As String
    Select Case labelText
        Case "Safe cleanup": LauncherRiskTip = "Common cleanup; preview is still recommended."
        Case "Inspect First": LauncherRiskTip = "Finds or reports issues before changing the document."
        Case "Text caution": LauncherRiskTip = "May change visible text; review the preview carefully."
        Case "Formatting change": LauncherRiskTip = "Changes visible formatting, not intended text content."
        Case "Structure change": LauncherRiskTip = "Changes document structure such as paragraphs, lists, tables, or sections."
        Case "Removes content": LauncherRiskTip = "Removes links, notes, objects, or other content."
        Case "Privacy cleanup": LauncherRiskTip = "Removes identity, metadata, or sharing-sensitive information."
        Case "Finalizes review": LauncherRiskTip = "Finalizes comments or tracked changes."
        Case "Alpha tool": LauncherRiskTip = "Specialist or not-yet-graduated tool; use on a copy."
        Case Else: LauncherRiskTip = "Review before applying."
    End Select
End Function
Private Sub ShowLauncherRiskExplanation(ByVal toolName As String, ByVal riskLabel As String, ByVal reasonText As String)
    Dim msg As String
    msg = toolName & vbCrLf & vbCrLf & _
          "Risk label: " & riskLabel & vbCrLf & _
          LauncherRiskTip(riskLabel) & vbCrLf & vbCrLf & _
          "Why this tool says that:" & vbCrLf & _
          reasonText & vbCrLf & vbCrLf & _
          "Before running it, use Preview when available and make sure the result matches what you intended."
    MsgBox msg, vbInformation, "Why This Risk?"
End Sub
Private Sub RefreshResetAllCallout()
    On Error Resume Next
    Dim calloutLeft As Single
    Dim calloutTop As Single
    Dim calloutWidth As Single
    Dim calloutHeight As Single
    Const FORM_W As Single = 760
    Const MARGIN As Single = 14
    Const COL_GAP As Single = 16
    Const COL_W As Single = (FORM_W - (2 * MARGIN) - COL_GAP) / 2

    lblResetAllArrow.Caption = ChrW$(&H2190)
    lblResetAllArrow.Font.Bold = True
    lblResetAllArrow.Font.Size = 12
    lblResetAllArrow.TextAlign = 2
    lblResetAllArrow.BackStyle = 0
    lblResetAllArrow.ForeColor = RGB(220, 0, 0)

    calloutLeft = cmdResetAll.Left + cmdResetAll.Width + 20
    calloutTop = cmdResetAll.Top - 2
    calloutWidth = (MARGIN + COL_W) - calloutLeft
    calloutHeight = 32

    lblResetAllArrow.Left = cmdResetAll.Left + cmdResetAll.Width + 2
    lblResetAllArrow.Top = cmdResetAll.Top + 7
    lblResetAllArrow.Width = 16
    lblResetAllArrow.Height = 16

    lblResetAllCalloutBox.Caption = ""
    lblResetAllCalloutBox.Left = calloutLeft
    lblResetAllCalloutBox.Top = calloutTop
    lblResetAllCalloutBox.Width = calloutWidth
    lblResetAllCalloutBox.Height = calloutHeight
    lblResetAllCalloutBox.BackStyle = 0
    lblResetAllCalloutBox.BorderStyle = 1
    lblResetAllCalloutBox.SpecialEffect = 0
    lblResetAllCalloutBox.ForeColor = RGB(220, 0, 0)
    lblResetAllCalloutBox.BorderColor = RGB(220, 0, 0)

    lblResetAllCalloutText.Caption = "If highlights remain after an error, select Reset All."
    lblResetAllCalloutText.Left = calloutLeft + 5
    lblResetAllCalloutText.Top = calloutTop + 4
    lblResetAllCalloutText.Width = calloutWidth - 10
    lblResetAllCalloutText.Height = calloutHeight - 8
    lblResetAllCalloutText.BackStyle = 0
    lblResetAllCalloutText.WordWrap = True
    lblResetAllCalloutText.Font.Size = 7.5
    lblResetAllCalloutText.ForeColor = RGB(32, 37, 45)
End Sub
Private Function LauncherRiskDisplayText(labelText As String) As String
    LauncherRiskDisplayText = Replace(labelText, " ", vbCrLf)
End Function
Private Sub RefreshReviewPrivacyLauncherSection()
    On Error Resume Next
    Const ROW_STEP As Single = 30
    cmdMetadata.Caption = "MetaDataSuite"
    lblDescMetadata.Caption = "Advanced metadata inspection, editing, and privacy management."
    cmdHelpFinalReview.Caption = "?"
    cmdFinalReview.Caption = "Final Review"
    lblDescFinalReview.Caption = "Remove review comments and accept tracked changes when markup is finalized."

    cmdHelpMetadata.Visible = True
    cmdMetadata.Visible = True
    lblDescMetadata.Visible = True
    lblRiskMetadata.Visible = True

    cmdHelpFinalReview.Font.Bold = True
    cmdHelpFinalReview.Left = cmdHelpMetadata.Left
    cmdFinalReview.Left = cmdMetadata.Left
    lblDescFinalReview.Left = lblDescMetadata.Left
    lblRiskFinalReview.Left = lblRiskMetadata.Left

    cmdFinalReview.Font.Bold = True
    lblDescFinalReview.Font.Bold = False
    lblDescFinalReview.Font.Size = lblDescMetadata.Font.Size
    lblDescFinalReview.ForeColor = lblDescMetadata.ForeColor
    lblDescFinalReview.BackColor = lblDescMetadata.BackColor
    lblRiskFinalReview.Font.Bold = lblRiskMetadata.Font.Bold
    cmdHelpFinalReview.Width = cmdHelpMetadata.Width
    cmdFinalReview.Width = cmdMetadata.Width
    lblDescFinalReview.Width = lblDescMetadata.Width
    lblRiskFinalReview.Width = lblRiskMetadata.Width

    cmdHelpFinalReview.BackColor = cmdHelpMetadata.BackColor
    cmdHelpFinalReview.ForeColor = cmdHelpMetadata.ForeColor

    cmdFinalReview.BackColor = cmdMetadata.BackColor
    cmdFinalReview.ForeColor = cmdMetadata.ForeColor

    cmdHelpFinalReview.Top = cmdHelpMetadata.Top + ROW_STEP
    cmdFinalReview.Top = cmdMetadata.Top + ROW_STEP
    lblDescFinalReview.Top = lblDescMetadata.Top + ROW_STEP
    lblRiskFinalReview.Top = lblRiskMetadata.Top + ROW_STEP

    cmdHelpHyperlink.Top = cmdHelpFinalReview.Top + ROW_STEP
    cmdHyperlink.Top = cmdFinalReview.Top + ROW_STEP
    lblDescHyperlink.Top = lblDescFinalReview.Top + ROW_STEP
    lblRiskHyperlink.Top = lblRiskFinalReview.Top + ROW_STEP
End Sub
Private Sub RefreshLauncherFooterButtons()
    On Error Resume Next
    cmdUserManual.Caption = "User Manual"
    cmdUserManual.Font.Bold = cmdResetAll.Font.Bold
    cmdUserManual.BackColor = cmdResetAll.BackColor
    cmdUserManual.ForeColor = cmdResetAll.ForeColor
    cmdUserManual.Left = 16
    cmdUserManual.Width = 96
    cmdResetAll.Left = 118
    cmdResetAll.Width = 96
    cmdResetAll.Top = chkReturnToMainAfterClose.Top + 20
    chkUpdateChecks.Top = cmdResetAll.Top
    chkUpdateChecks.Left = 16
    chkUpdateChecks.Width = 214
    cmdCheckUpdates.Left = 244
    cmdCheckUpdates.Top = chkUpdateChecks.Top - 3
    cmdCheckUpdates.Width = 112
    cmdCheckUpdates.Height = 24
    cmdResetAll.Top = chkUpdateChecks.Top + 26
    cmdUserManual.Top = cmdResetAll.Top
    RefreshResetAllCallout
End Sub
Private Sub chkReturnToMainAfterApply_Click()
    SetReturnToMainAfterApplySetting chkReturnToMainAfterApply.Value
End Sub
Private Sub chkReturnToMainAfterClose_Click()
    SetReturnToMainAfterCloseSetting chkReturnToMainAfterClose.Value
End Sub
Private Sub chkAutoSave_Click()
    If Not chkAutoSave.Value Then
        Dim msg As String
        msg = "Auto-save is your safety net." & vbCrLf & vbCrLf & _
              "Without it, if something goes wrong during cleanup, " & _
              "Word's recovery file may not have a clean pre-run version to fall back to." & vbCrLf & vbCrLf & _
              "Are you sure you want to turn auto-save off?"
        If MsgBox(msg, vbExclamation + vbYesNo, "Turn Off Auto-Save?") = vbNo Then
            chkAutoSave.Value = True
            RefreshAutoSaveWarning
            Exit Sub
        End If
    End If
    SetAutoSaveSetting chkAutoSave.Value
    RefreshAutoSaveWarning
End Sub
Private Sub chkUpdateChecks_Click()
    SetCleanupSuiteUpdateChecksEnabled chkUpdateChecks.Value
    If Not chkUpdateChecks.Value Then ExplainCleanupSuiteUpdateChecksDisabled
End Sub
Private Sub cmdCheckUpdates_Click()
    CheckForCleanupSuiteUpdates True
End Sub
Private Sub RefreshAutoSaveWarning()
    On Error Resume Next
    With lblAutoSaveWarning
        .Caption = "Dangerous! Should be selected"
        .Visible = (Not chkAutoSave.Value)
        .ForeColor = RGB(220, 0, 0)
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .BackStyle = 0
        .TextAlign = 1
    End With
    On Error GoTo 0
End Sub
Private Sub cmdResetAll_Click()
    ResetCleanupSuiteDefaults
    chkReturnToMainAfterApply.Value = GetReturnToMainAfterApplySetting()
    chkReturnToMainAfterClose.Value = GetReturnToMainAfterCloseSetting()
    chkAutoSave.Value = GetAutoSaveSetting()
    chkUpdateChecks.Value = GetCleanupSuiteUpdateChecksEnabled()
    RefreshAutoSaveWarning
    RefreshReviewPrivacyLauncherSection
    MsgBox "All Cleanup Suite tool settings have been reset to defaults.", vbInformation, "Cleanup Suite"
End Sub
Private Sub cmdUserManual_Click()
    OpenCleanupUserManualPdf
End Sub
Private Sub OpenCleanupTool(ByVal formName As String)
    Dim toolForm As Object
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    Dim originalErrSource As String

    On Error GoTo OpenErr
    Me.Hide
    BeginCleanupToolSession
    Set toolForm = CreateCleanupToolForm(formName)
    toolForm.Show
    Set toolForm = Nothing
    If ShouldReturnToLauncherAfterToolSession() Then Me.Show
    Exit Sub
OpenErr:
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    originalErrSource = Err.Source
    BeginCleanupToolSession
    MsgBox "Could not open that cleanup tool." & vbCrLf & vbCrLf & _
           "Error " & originalErrNumber & " from " & originalErrSource & ":" & vbCrLf & _
           originalErrDescription, vbExclamation, "Cleanup Suite"
    Me.Show
End Sub

Private Function CreateCleanupToolForm(ByVal formName As String) As Object
    Select Case formName
        Case "frmMetaDataSuite"
            Set CreateCleanupToolForm = New frmMetaDataSuite
        Case Else
            Set CreateCleanupToolForm = VBA.UserForms.Add(formName)
    End Select
End Function
Private Sub cmdUnicode_Click(): OpenCleanupTool "frmUnicodeCleanup": End Sub
Private Sub cmdPunctuation_Click(): OpenCleanupTool "frmPunctuationCleanup": End Sub
Private Sub cmdSpacing_Click(): OpenCleanupTool "frmSpacingCleanup": End Sub
Private Sub cmdCapitalization_Click(): OpenCleanupTool "frmCapitalizationCleanup": End Sub
Private Sub cmdList_Click(): OpenCleanupTool "frmListCleanup": End Sub
Private Sub cmdParagraph_Click(): OpenCleanupTool "frmParagraphCleanup": End Sub
Private Sub cmdDuplicate_Click(): OpenCleanupTool "frmDuplicateDetector": End Sub
Private Sub cmdFontNorm_Click(): OpenCleanupTool "frmFontNormalizer": End Sub
Private Sub cmdTableClean_Click(): OpenCleanupTool "frmTableCleaner": End Sub
Private Sub cmdBreakNorm_Click(): OpenCleanupTool "frmBreakNormalizer": End Sub
Private Sub cmdDocTrim_Click(): OpenCleanupTool "frmDocumentTrim": End Sub
Private Sub cmdFormatStrip_Click(): OpenCleanupTool "frmFormattingStripper": End Sub
Private Sub cmdHyperlink_Click(): OpenCleanupTool "frmHyperlinkRemover": End Sub
Private Sub cmdSoftReturn_Click(): OpenCleanupTool "frmSoftReturnConverter": End Sub
Private Sub cmdMetadata_Click(): OpenCleanupTool "frmMetaDataSuite": End Sub
Private Sub cmdFinalReview_Click(): OpenCleanupTool "frmFinalReview": End Sub
Private Sub cmdStyleClean_Click(): OpenCleanupTool "frmStyleCleanup": End Sub
Private Sub cmdFootnote_Click(): OpenCleanupTool "frmFootnoteRemover": End Sub
Private Sub cmdHeaderFooter_Click(): OpenCleanupTool "frmHeaderFooterStandardizer": End Sub
Private Sub cmdObjectRemover_Click(): OpenCleanupTool "frmObjectRemover": End Sub
Private Sub lblRiskUnicode_Click()
    ShowLauncherRiskExplanation "Invisible Unicode", "Safe cleanup", "It removes hidden Unicode characters that usually have no visible document meaning but can break search, sorting, or export."
End Sub
Private Sub lblRiskPunctuation_Click()
    ShowLauncherRiskExplanation "Punctuation", "Safe cleanup", "It normalizes punctuation marks and spacing that are commonly damaged by copy and paste."
End Sub
Private Sub lblRiskSpacing_Click()
    ShowLauncherRiskExplanation "Spacing", "Safe cleanup", "It fixes extra spaces and repeated blank lines that usually do not change the intended words."
End Sub
Private Sub lblRiskCapitalization_Click()
    ShowLauncherRiskExplanation "Capitalization", "Text caution", "It can change visible letters in the document, so names, acronyms, and headings should be reviewed."
End Sub
Private Sub lblRiskList_Click()
    ShowLauncherRiskExplanation "Lists", "Structure change", "It can turn typed bullets or numbers into real Word list formatting and adjust indentation."
End Sub
Private Sub lblRiskParagraph_Click()
    ShowLauncherRiskExplanation "Paragraphs", "Structure change", "It can change paragraph spacing, empty paragraphs, and manual indents."
End Sub
Private Sub lblRiskSoftReturn_Click()
    ShowLauncherRiskExplanation "Soft Returns", "Structure change", "It converts line breaks and paragraph marks, which changes how text is divided into paragraphs."
End Sub
Private Sub lblRiskDocTrim_Click()
    ShowLauncherRiskExplanation "Trim End of Doc", "Safe cleanup", "It removes trailing empty paragraphs at the end of the document."
End Sub
Private Sub lblRiskDuplicate_Click()
    ShowLauncherRiskExplanation "Duplicates", "Inspect First", "It finds repeated or near-duplicate paragraphs so you can decide what, if anything, should be removed."
End Sub
Private Sub lblRiskTableClean_Click()
    ShowLauncherRiskExplanation "Tables", "Structure change", "It can clean or convert table structure, padding, borders, and table formatting."
End Sub
Private Sub lblRiskHeaderFooter_Click()
    ShowLauncherRiskExplanation "Headers / Footers", "Structure change", "It can standardize or clear header and footer content across document sections."
End Sub
Private Sub lblRiskFootnote_Click()
    ShowLauncherRiskExplanation "Footnote / Endnote", "Removes content", "It can remove footnotes or endnotes, with an option to keep note text inline."
End Sub
Private Sub lblRiskObjectRemover_Click()
    ShowLauncherRiskExplanation "Objects", "Removes content", "It removes embedded objects such as images, text boxes, controls, hidden text, or tables."
End Sub
Private Sub lblRiskFontNorm_Click()
    ShowLauncherRiskExplanation "Fonts", "Formatting change", "It changes direct font overrides such as face, size, color, bold, or italic."
End Sub
Private Sub lblRiskFormatStrip_Click()
    ShowLauncherRiskExplanation "Formatting Cleaner", "Formatting change", "It strips direct character and paragraph formatting while trying to preserve styles."
End Sub
Private Sub lblRiskMetadata_Click()
    ShowLauncherRiskExplanation "MetaDataSuite", "Privacy cleanup", "It inspects and manages metadata that may contain identity or sharing-sensitive information."
End Sub
Private Sub lblRiskFinalReview_Click()
    ShowLauncherRiskExplanation "Final Review", "Finalizes review", "It can remove comments and accept tracked changes when markup is finalized."
End Sub
Private Sub lblRiskHyperlink_Click()
    ShowLauncherRiskExplanation "Hyperlinks", "Removes content", "It can hide hyperlink styling while keeping link targets, or remove link targets while keeping visible text."
End Sub
Private Sub lblRiskStyleClean_Click()
    ShowLauncherRiskExplanation "Styles", "Alpha tool", "It is a specialist cleanup tool that can change document styles and should be used on a copy."
End Sub
Private Sub lblRiskBreakNorm_Click()
    ShowLauncherRiskExplanation "Breaks", "Alpha tool", "It is a specialist cleanup tool that can change section and page breaks from assembled documents."
End Sub
Private Sub cmdHelpUnicode_Click()
    ShowCleanupToolHelp "Unicode"
End Sub
Private Sub cmdHelpPunct_Click()
    ShowCleanupToolHelp "Punctuation"
End Sub
Private Sub cmdHelpSpacing_Click()
    ShowCleanupToolHelp "Spacing"
End Sub
Private Sub cmdHelpCap_Click()
    ShowCleanupToolHelp "Capitalization"
End Sub
Private Sub cmdHelpList_Click()
    ShowCleanupToolHelp "List"
End Sub
Private Sub cmdHelpPara_Click()
    ShowCleanupToolHelp "Paragraph"
End Sub
Private Sub cmdHelpDuplicate_Click()
    ShowCleanupToolHelp "Duplicate"
End Sub
Private Sub cmdHelpFont_Click()
    ShowCleanupToolHelp "Font"
End Sub
Private Sub cmdHelpTable_Click()
    ShowCleanupToolHelp "Table"
End Sub
Private Sub cmdHelpBreak_Click()
    ShowCleanupToolHelp "Break"
End Sub
Private Sub cmdHelpTrim_Click()
    ShowCleanupToolHelp "DocumentTrim"
End Sub
Private Sub cmdHelpFormat_Click()
    ShowCleanupToolHelp "Formatting"
End Sub
Private Sub cmdHelpHyperlink_Click()
    ShowCleanupToolHelp "Hyperlink"
End Sub
Private Sub cmdHelpSoftReturn_Click()
    ShowCleanupToolHelp "SoftReturn"
End Sub
Private Sub cmdHelpMetadata_Click()
    ShowCleanupToolHelp "MetaDataSuite"
End Sub
Private Sub cmdHelpFinalReview_Click()
    ShowCleanupToolHelp "FinalReview"
End Sub
Private Sub cmdHelpStyle_Click()
    ShowCleanupToolHelp "Style"
End Sub
Private Sub cmdHelpFootnote_Click()
    ShowCleanupToolHelp "Footnote"
End Sub
Private Sub cmdHelpHeaderFooter_Click()
    ShowCleanupToolHelp "HeaderFooter"
End Sub
Private Sub cmdHelpObject_Click()
    ShowCleanupToolHelp "Object"
End Sub
