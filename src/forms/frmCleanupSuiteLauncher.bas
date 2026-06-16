Option Explicit
Private Sub UserForm_Initialize()
    ApplyMSFormTitleStrategy Me, True
    lblLauncherTitle.Caption = "Cleanup Suite"
    lblLauncherTitle.TextAlign = 2
    lblLauncherSubtitle.Caption = "Choose the kind of cleanup you need. Tools are grouped by what you are trying to fix."
    lblLauncherSubtitle.TextAlign = 2
    cmdFootnote.Caption = "Footnote / Endnote"
    chkShowCompletionReviewAfterApply.Caption = "Show completion review after apply"
    chkShowCompletionReviewAfterApply.Value = GetShowCompletionReviewAfterApplySetting()
    RefreshApplyReturnCaption
    chkReturnToMainAfterApply.Value = GetReturnToMainAfterApplySetting()
    chkReturnToMainAfterClose.Caption = "Return to main menu after closing individual tool menus"
    chkReturnToMainAfterClose.Value = GetReturnToMainAfterCloseSetting()
    chkAutoSave.Caption = "Auto-save before running each tool"
    chkAutoSave.Value = GetAutoSaveSetting()
End Sub
Private Sub chkReturnToMainAfterApply_Click()
    SetReturnToMainAfterApplySetting chkReturnToMainAfterApply.Value
End Sub
Private Sub chkShowCompletionReviewAfterApply_Click()
    SetShowCompletionReviewAfterApplySetting chkShowCompletionReviewAfterApply.Value
    RefreshApplyReturnCaption
End Sub
Private Sub RefreshApplyReturnCaption()
    If chkShowCompletionReviewAfterApply.Value Then
        chkReturnToMainAfterApply.Caption = "Return to main menu after completion review"
    Else
        chkReturnToMainAfterApply.Caption = "Return to main menu after apply"
    End If
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
            Exit Sub
        End If
    End If
    SetAutoSaveSetting chkAutoSave.Value
End Sub
Private Sub cmdResetAll_Click()
    ResetCleanupSuiteDefaults
    chkReturnToMainAfterApply.Value = GetReturnToMainAfterApplySetting()
    chkShowCompletionReviewAfterApply.Value = GetShowCompletionReviewAfterApplySetting()
    RefreshApplyReturnCaption
    chkReturnToMainAfterClose.Value = GetReturnToMainAfterCloseSetting()
    chkAutoSave.Value = GetAutoSaveSetting()
    MsgBox "All Cleanup Suite tool settings have been reset to defaults.", vbInformation, "Cleanup Suite"
End Sub
Private Sub cmdUserManual_Click()
    OpenCleanupUserManualPdf
End Sub
Private Sub OpenCleanupTool(ByVal formName As String)
    On Error GoTo OpenErr
    Me.Hide
    BeginCleanupToolSession
    Dim toolForm As Object
    Set toolForm = VBA.UserForms.Add(formName)
    toolForm.Show
    Set toolForm = Nothing
    If ShouldReturnToLauncherAfterToolSession() Then Me.Show
    Exit Sub
OpenErr:
    BeginCleanupToolSession
    MsgBox "Could not open that cleanup tool: " & Err.Description, vbExclamation, "Cleanup Suite"
    Me.Show
End Sub
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
Private Sub cmdMetadata_Click(): OpenCleanupTool "frmMetadataScrubber": End Sub
Private Sub cmdStyleClean_Click(): OpenCleanupTool "frmStyleCleanup": End Sub
Private Sub cmdFootnote_Click(): OpenCleanupTool "frmFootnoteRemover": End Sub
Private Sub cmdHeaderFooter_Click(): OpenCleanupTool "frmHeaderFooterStandardizer": End Sub
Private Sub cmdObjectRemover_Click(): OpenCleanupTool "frmObjectRemover": End Sub
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
    ShowCleanupToolHelp "Metadata"
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
