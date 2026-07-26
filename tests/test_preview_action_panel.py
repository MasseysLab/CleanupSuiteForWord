from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class PreviewActionPanelTests(unittest.TestCase):
    def test_preview_action_form_is_registered_in_manifest_and_installer(self):
        manifest = read("src/manifest.txt")
        installer = read("src/installer/installer.bas")

        self.assertIn("forms/frmPreviewActions.bas  frmPreviewActions_Code", manifest)
        self.assertIn('CreateOrReplaceForm vbProj, "frmPreviewActions", frmPreviewActions_Code', installer)
        self.assertIn('AttemptCreateControls vbProj, "frmPreviewActions"', installer)
        self.assertIn('"frmPreviewActions"', installer)
        self.assertIn('Case "frmPreviewActions": FriendlyFormCaption = "Preview Actions"', installer)

    def test_preview_panel_helpers_manage_panel_lifecycle_and_state(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        panel = read("src/forms/frmPreviewActions.bas")
        installer = read("src/installer/installer.bas")

        self.assertIn("Public gPreviewActionPanel As Object", helpers)
        self.assertIn("Private gPreviewActionPanelHasPosition As Boolean", helpers)
        self.assertIn("Private gPreviewActionPanelLeft As Single", helpers)
        self.assertIn("Private gPreviewActionPanelTop As Single", helpers)
        self.assertIn("Public Sub RememberPreviewActionPanelPosition", helpers)
        self.assertIn("Private Sub ApplyPreviewActionPanelPosition", helpers)
        self.assertIn("gPreviewActionPanelHasPosition = True", helpers)
        self.assertIn("panel.StartUpPosition = 0", helpers)
        self.assertIn("panel.Left = gPreviewActionPanelLeft", helpers)
        self.assertIn("panel.Top = gPreviewActionPanelTop", helpers)
        self.assertIn("Public Sub ShowPreviewActions", helpers)
        self.assertIn("Public Function NewPreviewSummaryRows() As Collection", helpers)
        self.assertIn("Public Sub AddPreviewSummaryRow(ByVal rows As Collection, ByVal itemText As String, ByVal countValue As Variant, Optional ByVal isUnhighlighted As Boolean = False, Optional ByVal isInactive As Boolean = False)", helpers)
        self.assertIn("Public Sub ShowPreviewActionsSummary(ByVal sourceForm As Object, ByVal toolName As String, ByVal rows As Collection)", helpers)
        self.assertIn("rows.Add PreviewSummaryRowText(itemText, countValue, isUnhighlighted, isInactive)", helpers)
        self.assertIn("PreviewSummaryRowText", helpers)
        self.assertIn('If isUnhighlighted Then displayText = displayText & " (unhighlighted)"', helpers)
        self.assertIn('displayText & vbTab & CStr(countValue) & vbTab & IIf(isInactive, "inactive", "")', helpers)
        self.assertIn('"inactive"', helpers)
        self.assertIn("isUnhighlighted", helpers)
        self.assertIn("isInactive", helpers)
        self.assertNotIn('"not previewed"', helpers)
        self.assertIn("Private gPreviewScopeHasRange As Boolean", helpers)
        self.assertIn("Private gPreviewScopeStart As Long", helpers)
        self.assertIn("Private gPreviewScopeEnd As Long", helpers)
        self.assertIn("Private gPreviewScopeDocumentKey As String", helpers)
        self.assertIn("Private Sub CapturePreviewScopeRange(ByVal sourceForm As Object)", helpers)
        self.assertIn("Public Sub ClearPreviewScopeRange()", helpers)
        self.assertIn("CapturePreviewScopeRange sourceForm", helpers)
        self.assertIn("ClearPreviewScopeRange", helpers)
        self.assertIn("If gPreviewScopeHasRange Then", helpers)
        self.assertIn("Set GetTargetRange = ActiveDocument.Range(gPreviewScopeStart, gPreviewScopeEnd)", helpers)
        self.assertLess(
            helpers.index("If gPreviewScopeHasRange Then"),
            helpers.index("If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then"),
        )
        self.assertIn("gPreviewActionPanel.Show vbModal", helpers)
        self.assertIn("gPreviewActionPanel.Show vbModeless", helpers)
        self.assertIn("ResumePreviewActionPanelModelessIfRequested", helpers)
        self.assertLess(
            helpers.index("ApplyPreviewActionPanelPosition gPreviewActionPanel"),
            helpers.index("gPreviewActionPanel.Show"),
        )
        self.assertIn("vbModeless", helpers)
        self.assertIn("Public Sub Configure", panel)
        self.assertIn("Public Sub ConfigureSummary(sourceForm As Object, toolName As String, rows As Collection)", panel)
        self.assertIn("Me.Caption = \"\"", panel)
        self.assertIn('CallByName mSourceForm, "RunAfterPreview", VbMethod', panel)
        self.assertNotIn('CallByName src, "PreviewFromPanel", VbMethod', panel)
        self.assertIn("RemovePreviewHighlighting mPreviewDocument", panel)
        self.assertIn("ClearPreviewScopeRange", panel)
        self.assertIn("ConsumeModelessResumeRequest", panel)
        self.assertIn("Private Sub cmdPreviousPage_Click()", panel)
        self.assertIn("Private Sub cmdNextPage_Click()", panel)
        self.assertIn("Private Sub cmdPreviousChange_Click()", panel)
        self.assertIn("Private Sub cmdNextChange_Click()", panel)
        self.assertIn("TurnPreviewPage(moveForward)", panel)
        self.assertIn("Application.Browser.Target = wdBrowsePage", helpers)
        self.assertIn("Application.Browser.Previous", helpers)
        self.assertIn("Application.Browser.Next", helpers)
        self.assertIn("Me.Hide", panel)
        self.assertIn("src.Show", panel)
        self.assertIn("ReturnToLauncherAfterDetachedToolSession", panel)
        self.assertNotIn("ShowCleanupSuiteLauncher", panel)
        self.assertIn("Private Sub LayoutPanel()", panel)
        self.assertIn("LayoutPanel", panel)
        self.assertIn("RenderSummaryTable rows", panel)
        self.assertIn('RenderSummaryCell "lblSummaryHeaderItem", "Item"', panel)
        self.assertIn('RenderSummaryCell "lblSummaryHeaderCount", "#"', panel)
        self.assertNotIn("RenderSummaryStatusCell", panel)
        self.assertIn('If UBound(rowParts) >= 2 Then isInactive = (CStr(rowParts(2)) = "inactive")', panel)
        self.assertIn('tableLeft = M + ((CONTENT_W - TABLE_W) / 2)', panel)
        self.assertIn('RenderSummaryCell "lblSummaryItem" & CStr(rowIndex), itemText, tableLeft, y, ITEM_W, ROW_H, False, isInactive', panel)
        self.assertIn('RenderSummaryCell "lblSummaryCount" & CStr(rowIndex), countText, tableLeft + ITEM_W, y, COUNT_W, ROW_H, False, isInactive', panel)
        self.assertIn("ctl.BorderStyle = fmBorderStyleSingle", panel)
        self.assertIn("If isInactive Then", panel)
        self.assertIn("ctl.ForeColor = RGB(145, 145, 145)", panel)
        self.assertIn("ctl.BackColor = RGB(245, 245, 245)", panel)
        self.assertIn('InStr(1, controlName, "Count", vbTextCompare) > 0', panel)
        self.assertIn("ctl.TextAlign = 2", panel)
        self.assertNotIn("lblSummaryStatus", panel)
        self.assertIn("SummaryTableHeight", panel)
        self.assertIn("Me.Height = MaxSingle(MIN_FORM_H, desiredInsideH)", panel)
        self.assertIn("If Me.InsideHeight < desiredInsideH Then", panel)
        self.assertIn("mResumeModelessAfterModal = True", panel)
        self.assertIn("RememberPreviewActionPanelPosition Me.Left, Me.Top", panel)
        self.assertIn("If Not mPreviewOn Then Exit Sub", panel)
        self.assertIn("Private Sub StyleActionBar()", panel)
        self.assertIn("Private Sub StyleLabel", panel)
        self.assertIn("Private Sub StyleButton", panel)
        self.assertIn("StyleLabel lblHint, 8, True", panel)
        self.assertIn("ctl.Font.Size = 8", panel)
        self.assertIn("Const FORM_W As Single = 283", panel)
        self.assertIn("Const MIN_FORM_H As Single = 90", panel)
        self.assertIn("Const CONTENT_W As Single = 257", panel)
        self.assertIn("Const BTN_W As Single = (CONTENT_W - (2 * GAP)) / 3", panel)
        self.assertIn("Private Function MeasuredPreviewButtonHeight() As Single", panel)
        self.assertIn("Private Sub FitPreviewButtonCaption", panel)
        self.assertIn("MeasuredPreviewButtonHeight = MaxSingle(32, maximumTextHeight + 8)", panel)
        self.assertIn("Const GAP As Single = 5", panel)
        self.assertIn('lblTitle.Caption = ""', panel)
        self.assertIn("lblTitle.Visible = False", panel)
        self.assertIn("lblTitle.Move 0, 0, 0, 0", panel)
        self.assertIn("lblHint.Caption = toolName", panel)
        self.assertNotIn("lblSummary.Caption = mSummaryText", panel)
        self.assertIn("lblHint.Move M, 2, CONTENT_W, 12", panel)
        self.assertIn("cmdPreview.Move M, ACTION_TOP, BTN_W, buttonHeight", panel)
        self.assertIn("cmdClear.Move M + BTN_W + GAP, ACTION_TOP, BTN_W, buttonHeight", panel)
        self.assertIn("cmdApply.Move M + (2 * (BTN_W + GAP)), ACTION_TOP, BTN_W, buttonHeight", panel)
        self.assertIn("navigationWidth = (CONTENT_W - (3 * GAP)) / 4", panel)
        self.assertIn("cmdPreviousChange.Move M, navigationTop, navigationWidth, buttonHeight", panel)
        self.assertIn("cmdPreviousPage.Move M + navigationWidth + GAP, navigationTop, navigationWidth, buttonHeight", panel)
        self.assertIn("cmdNextPage.Move M + (2 * (navigationWidth + GAP)), navigationTop, navigationWidth, buttonHeight", panel)
        self.assertIn("cmdNextChange.Move M + (3 * (navigationWidth + GAP)), navigationTop, navigationWidth, buttonHeight", panel)
        self.assertIn("If Not DocumentHasVisibleContent() Then", panel)
        self.assertIn("This document appears to be blank. There is nothing to apply.", panel)
        self.assertIn('cmdPreview.Caption = "Preview is ON"', panel)
        self.assertIn("cmdPreview.Enabled = True", panel)
        self.assertIn('cmdPreview.Caption = "Preview is OFF"', panel)
        self.assertIn("cmdPreview.Enabled = False", panel)
        self.assertIn('AddPreviewSummaryRow offRows, "Edit if needed. Choose Reconfigure to preview again.", ""', panel)
        self.assertIn("Set mSummaryRows = offRows", panel)
        self.assertLess(
            panel.index('cmdPreview.Caption = "Preview is OFF"'),
            panel.index('AddPreviewSummaryRow offRows, "Edit if needed. Choose Reconfigure to preview again.", ""'),
        )
        self.assertIn("mResumeModelessAfterModal = True", panel)
        self.assertIn('cmdClear.Caption = "Reconfigure"', panel)
        self.assertNotIn('cmdClear.Caption = "Settings"', panel)
        self.assertNotIn('cmdClear.Caption = "Reset Preview"', panel)
        self.assertNotIn("UserForm_Initialize", panel)
        self.assertNotIn("cmdBack", panel)
        self.assertNotIn("cmdClose", panel)

    def test_remove_all_highlighting_restores_preview_owned_markers_only(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        start = helpers.index("Public Sub RemoveAllHighlighting")
        end = helpers.index("Public Function DocumentHasVisibleContent")
        helper = helpers[start:end]

        self.assertIn("Set hlRange = scopeRange.Duplicate", helper)
        self.assertIn("RestorePreviewHighlighting", helper)
        self.assertIn("hlRange.Find.ClearHitHighlight", helper)
        self.assertNotIn("hlRange.HighlightColorIndex = wdNoHighlight", helper)
        self.assertNotIn(".Replacement.Highlight = False", helper)

        self.assertIn("Private gPreviewHighlightDocuments As Collection", helpers)
        self.assertIn("Public Sub ApplyPreviewHighlight", helpers)
        self.assertIn("RegisterPreviewHighlightRange targetRange", helpers)
        self.assertIn("Private Sub RestorePreviewHighlighting", helpers)

    def test_preview_highlight_registry_preserves_uniform_and_mixed_user_colors(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("If originalColor <> wdUndefined Then", helpers)
        self.assertIn("For characterIndex = targetRange.Start + 1 To targetRange.End - 1", helpers)
        self.assertIn("If characterColor <> runColor Then", helpers)
        self.assertIn("For recordIndex = gPreviewHighlightDocuments.Count To 1 Step -1", helpers)
        self.assertIn("restoreRange.HighlightColorIndex = expectedColor", helpers)
        self.assertIn("If allLiveRecordsRestored Then", helpers)

    def test_preview_change_navigation_uses_page_deduplicated_anchors(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        panel = read("src/forms/frmPreviewActions.bas")
        punctuation = read("src/forms/frmPunctuationCleanup.bas")

        self.assertIn("Private gPreviewChangePages As Collection", helpers)
        self.assertIn("BeginPreviewChangeCollection", helpers)
        self.assertIn("RegisterPreviewChangeAnchor targetRange", helpers)
        self.assertIn("RegisterPreviewChangeAnchor paragraphRange", helpers)
        self.assertIn('pageKey = PreviewShadingDocumentIdentity(anchorRange.Document) & "|" & CStr(pageNumber)', helpers)
        self.assertIn("Public Function NavigatePreviewChangePage", helpers)
        self.assertIn("anchorPage > currentPage And anchorPage < candidatePage", helpers)
        self.assertIn("anchorPage < currentPage And anchorPage > candidatePage", helpers)
        self.assertIn("Public Function TurnPreviewPage", helpers)
        self.assertIn('Application.CommandBars.ExecuteMso "GoToNextPage"', helpers)
        self.assertIn('Application.CommandBars.ExecuteMso "GoToPreviousPage"', helpers)
        self.assertIn("For pageStep = 1 To pageDistance", helpers)
        self.assertIn("If Not TurnPreviewPage(moveForward) Then Exit Function", helpers)
        self.assertIn("Application.Browser.Next", helpers)
        self.assertIn("Application.Browser.Previous", helpers)
        self.assertIn("If candidateEnd < previewDocument.Content.End Then candidateEnd = candidateEnd + 1", helpers)
        self.assertIn("Set targetRange = previewDocument.Range(candidateStart, candidateEnd)", helpers)
        self.assertIn("targetRange.Select", helpers)
        self.assertIn("Application.ActiveWindow.ScrollIntoView targetRange, True", helpers)
        self.assertIn("Application.ScreenRefresh", helpers)
        self.assertIn("DoEvents", helpers)
        self.assertLess(helpers.index("Next pageStep"), helpers.index("targetRange.Select"))
        self.assertLess(helpers.index("targetRange.Select"), helpers.index("ScrollIntoView targetRange, True"))
        self.assertIn("Private Sub NavigatePreviewChange", panel)
        self.assertIn("Private Function DetachPreviewPanelForNavigationIfNeeded() As Boolean", panel)
        self.assertIn("mResumeModelessAfterModal = True", panel)
        self.assertIn("RememberPreviewActionPanelPosition Me.Left, Me.Top", panel)
        self.assertIn("Me.Hide", panel)
        self.assertIn("mNavigationDetached = True", panel)
        self.assertIn("Private Sub ResumePreviewPanelAfterNavigation", panel)
        self.assertIn("Me.Show vbModeless", panel)
        self.assertIn("Private Sub FocusPreviewDocumentWindow()", panel)
        self.assertIn("mPreviewWindow.Activate", panel)
        self.assertGreaterEqual(panel.count("FocusPreviewDocumentWindow"), 3)
        self.assertIn("If Not TurnPreviewPage(moveForward) Then GoTo NavigationErr", panel)
        self.assertIn("UpdateNavigationButtonStates", panel)
        self.assertIn("readingViewActive = (ActiveWindow.View.Type = wdReadingView)", helpers)
        self.assertIn("If readingViewActive Then", helpers)
        self.assertIn("Private Function RepaintReadingViewAfterPreviousPage", helpers)
        self.assertIn("mouse_event MOUSEEVENTF_LEFTDOWN", helpers)
        self.assertIn("mouse_event MOUSEEVENTF_LEFTUP", helpers)
        self.assertIn('VBA.SendKeys "{LEFT}", True', helpers)
        self.assertIn("If cursorCaptured Then SetCursorPos originalCursor.X, originalCursor.Y", helpers)
        self.assertIn("originalStart = Selection.Start", helpers)
        self.assertIn("If Selection.Start >= originalStart Then", helpers)
        self.assertIn('Application.CommandBars.ExecuteMso "GoToPreviousPage"', helpers)
        self.assertIn("CountPreviewFindMatches(targetRange, ChrW(&H2014), True)", punctuation)
        self.assertIn("highlightRange.Find.HitHighlight", helpers)

    def test_preview_navigation_controls_are_equal_sized_and_semantically_colored(self):
        panel = read("src/forms/frmPreviewActions.bas")

        self.assertIn("navigationWidth = (CONTENT_W - (3 * GAP)) / 4", panel)
        for control in ("cmdPreviousChange", "cmdPreviousPage", "cmdNextPage", "cmdNextChange"):
            self.assertIn(f"{control}.Move", panel)
        self.assertIn("buttonControl.BackColor = RGB(255, 243, 205)", panel)
        self.assertIn("buttonControl.BackColor = RGB(221, 235, 247)", panel)
        self.assertIn("buttonControl.BackColor = RGB(235, 235, 235)", panel)

    def test_preview_panel_exposes_only_decisive_actions(self):
        installer = read("src/installer/installer.bas")

        self.assertIn(
            'ControlsForForm = Array("lblTitle", "lblSummary", "lblHint", "lblButtonMeasure", _',
            installer,
        )
        self.assertIn('"cmdApply", "cmdPreview", "cmdClear", "cmdPreviousChange", "cmdPreviousPage", "cmdNextPage", "cmdNextChange")', installer)
        preview_controls_start = installer.index('Case "frmPreviewActions"')
        preview_controls_end = installer.index('Case "frmPunctuationCleanup"')
        preview_controls = installer[preview_controls_start:preview_controls_end]
        self.assertNotIn('"cmdBack"', installer)
        self.assertNotIn('"cmdClose"', preview_controls)

    def test_launcher_hides_while_a_subtool_is_open(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")

        self.assertIn("Private Sub OpenCleanupTool(ByVal formName As String)", launcher)
        self.assertIn("Me.Hide", launcher)
        self.assertIn("Set toolForm = CreateCleanupToolForm(formName)", launcher)
        self.assertIn("Set CreateCleanupToolForm = VBA.UserForms.Add(formName)", launcher)
        self.assertIn("toolForm.Show", launcher)
        self.assertIn("If ShouldReturnToLauncherAfterToolSession() Then Me.Show", launcher)
        self.assertNotIn("toolForm.Hide", launcher)
        self.assertNotIn("Private Sub cmdUnicode_Click(): Me.Hide: frmUnicodeCleanup.Show: End Sub", launcher)
        self.assertNotIn("Private Sub cmdCapitalization_Click(): Me.Hide: frmCapitalizationCleanup.Show: End Sub", launcher)
        self.assertIn('Private Sub cmdUnicode_Click(): OpenCleanupTool "frmUnicodeCleanup": End Sub', launcher)
        self.assertIn('Private Sub cmdCapitalization_Click(): OpenCleanupTool "frmCapitalizationCleanup": End Sub', launcher)

    def test_preview_capable_forms_can_apply_changes_from_the_panel(self):
        form_paths = [
            path
            for path in (ROOT / "src" / "forms").glob("frm*.bas")
            if "chkPreviewOnly" in path.read_text(encoding="utf-8")
        ]

        self.assertGreater(len(form_paths), 10)
        for path in form_paths:
            with self.subTest(form=path.name):
                source = path.read_text(encoding="utf-8")
                self.assertIn("Public Sub RunAfterPreview()", source)
                self.assertIn("chkPreviewOnly.Value = False", source)
                self.assertIn("cmdRun_Click", source)
                self.assertIn("ShowPreviewActionsSummary Me", source)
                self.assertIn("MarkCleanupToolApplied", source)

    def test_apply_and_preview_buttons_use_the_expected_labels(self):
        installer = read("src/installer/installer.bas")
        panel = read("src/forms/frmPreviewActions.bas")

        self.assertIn('If CStr(cn) = "cmdPreview" Then', installer)
        self.assertIn("Const FORM_W As Single = 283", installer)
        self.assertIn("Const FORM_H As Single = 126", installer)
        self.assertIn("Const CONTENT_W As Single = 257", installer)
        self.assertIn("Const BTN_W As Single = (CONTENT_W - (2 * GAP)) / 3", installer)
        self.assertIn("Const BTN_H As Single = 25", installer)
        self.assertIn("Const GAP As Single = 5", installer)
        self.assertIn('Case "frmPreviewActions.lblTitle": ControlCaptionText = ""', installer)
        self.assertIn('Case "frmPreviewActions.lblHint": ControlCaptionText = "Tool Name"', installer)
        self.assertIn('comp.Properties("Caption") = ""', installer)
        self.assertIn('designer.Caption = ""', installer)
        self.assertIn('Case "cmdClear": cap = "Reconfigure"', installer)
        self.assertIn('Case "cmdPreviousChange": cap = "Change" & vbCrLf', installer)
        self.assertIn('Case "cmdPreviousPage": cap = "Page" & vbCrLf', installer)
        self.assertIn('Case "cmdNextPage": cap = "Page" & vbCrLf', installer)
        self.assertIn('Case "cmdNextChange": cap = "Change" & vbCrLf', installer)
        self.assertIn('Case "cmdPreview": cap = "Preview"', installer)
        self.assertIn('Case "cmdReset": cap = "Reset"', installer)
        self.assertIn('Case "cmdResetAll": cap = "Reset All"', installer)
        self.assertIn('cmdApply.Caption = "Apply"', panel)
        self.assertIn('Case "cmdPreview": ButtonTipText = "Preview with the current settings"', installer)
        self.assertIn('Case "cmdReset": ButtonTipText = "Reset this tool to its defaults"', installer)
        self.assertIn('Case "cmdResetAll": ButtonTipText = "Reset all tools and global settings to defaults"', installer)
        self.assertIn("ctl.Enabled = True", installer)
        self.assertIn('ControlTypeByName = "Forms.CommandButton.1"', installer)
        self.assertNotIn('"chkPreviewOnly", "cmdRun"', installer)
        self.assertNotIn('"cmdDeselectAll", "cmdRun"', installer)

    def test_preview_panel_summary_height_resizes_with_content(self):
        panel = read("src/forms/frmPreviewActions.bas")

        self.assertIn("Private Function SummaryTableHeight", panel)
        self.assertIn("Private mSummaryTop As Single", panel)
        self.assertIn("Private Const SUMMARY_BOTTOM_GAP As Single = 6", panel)
        self.assertIn("Dim desiredInsideH As Single", panel)
        self.assertIn("desiredInsideH = mSummaryTop + SummaryTableHeight(mSummaryRows) + SUMMARY_BOTTOM_GAP", panel)
        self.assertIn("If Me.InsideHeight < desiredInsideH Then", panel)
        self.assertIn("Me.Height = Me.Height + (desiredInsideH - Me.InsideHeight)", panel)
        self.assertIn("Me.Height = Me.Height + (desiredInsideH - Me.InsideHeight) + 2", panel)
        self.assertIn("y = mSummaryTop", panel)
        self.assertIn("rowCount = rows.Count + 1", panel)
        self.assertIn("SummaryTableHeight = rowCount * ROW_H", panel)
        self.assertIn("If SummaryTableHeight < 22 Then SummaryTableHeight = 22", panel)
        self.assertIn("Private Function MaxSingle", panel)

    def test_preview_toggle_event_declares_objects_before_executable_code(self):
        panel = read("src/forms/frmPreviewActions.bas")
        proc = panel[
            panel.index("Private Sub cmdPreview_Click()"):
            panel.index("Private Sub cmdClear_Click()")
        ]

        self.assertLess(proc.index("Dim errorDescription As String"), proc.index("On Error GoTo PreviewErr"))

    def test_blank_document_apply_is_short_circuited_before_any_tool_runs(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        panel = read("src/forms/frmPreviewActions.bas")

        self.assertIn("Public Function DocumentHasVisibleContent", helpers)
        self.assertIn('textValue = Replace(textValue, vbCr, "")', helpers)
        self.assertIn("ActiveDocument.InlineShapes.Count", helpers)
        self.assertIn("ActiveDocument.Shapes.Count", helpers)
        self.assertIn("ActiveDocument.Tables.Count", helpers)
        self.assertIn("If Not DocumentHasVisibleContent() Then", panel)
        self.assertLess(
            panel.index("If Not DocumentHasVisibleContent() Then"),
            panel.index('CallByName mSourceForm, "RunAfterPreview", VbMethod'),
        )

    def test_preview_capable_forms_have_the_preview_button_hook(self):
        form_paths = [
            path
            for path in (ROOT / "src" / "forms").glob("frm*.bas")
            if "chkPreviewOnly" in path.read_text(encoding="utf-8")
        ]

        self.assertGreater(len(form_paths), 10)
        for path in form_paths:
            with self.subTest(form=path.name):
                source = path.read_text(encoding="utf-8")
                self.assertIn("Private Sub cmdPreview_Click()", source)
                self.assertIn("Public Sub PreviewFromPanel()", source)
                self.assertIn("Private Sub cmdReset_Click()", source)
                self.assertIn("UserForm_Initialize", source)
                self.assertIn("LayoutCleanupToolForm Me", source)
                self.assertIn("chkPreviewOnly.Value = True", source)
                self.assertIn("BeginPreviewActionIndicator Me", source)
                self.assertIn("EndPreviewActionIndicator", source)
                self.assertIn("cmdRun_Click", source)
                preview_from_panel = source[
                    source.index("Public Sub PreviewFromPanel()"):
                    source.index("Private Sub cmdReset_Click()")
                ]
                self.assertLess(
                    preview_from_panel.index("chkPreviewOnly.Value = True"),
                    preview_from_panel.index("BeginPreviewActionIndicator Me"),
                )
                self.assertLess(
                    preview_from_panel.index("BeginPreviewActionIndicator Me"),
                    preview_from_panel.index("cmdRun_Click"),
                )
                self.assertLess(
                    preview_from_panel.index("cmdRun_Click"),
                    preview_from_panel.index("EndPreviewActionIndicator"),
                )
                self.assertLess(
                    preview_from_panel.index("EndPreviewActionIndicator"),
                    preview_from_panel.rindex("chkPreviewOnly.Value = False"),
                )

    def test_return_to_main_after_apply_setting_is_wired_through_launcher_and_helpers(self):
        launcher = read("src/forms/frmCleanupSuiteLauncher.bas")
        helpers = read("src/modules/modCleanupHelpers.bas")
        panel = read("src/forms/frmPreviewActions.bas")
        installer = read("src/installer/installer.bas")

        self.assertIn("chkReturnToMainAfterApply.Value = GetReturnToMainAfterApplySetting()", launcher)
        self.assertIn("SetReturnToMainAfterApplySetting chkReturnToMainAfterApply.Value", launcher)
        self.assertIn("Public Function GetReturnToMainAfterApplySetting() As Boolean", helpers)
        self.assertIn("Public Sub SetReturnToMainAfterApplySetting(enabled As Boolean)", helpers)
        self.assertIn("Public Sub ResetCleanupSuiteDefaults()", helpers)
        self.assertIn("SetAutoSaveSetting True", helpers)
        self.assertIn("SetReturnToMainAfterApplySetting True", helpers)
        self.assertIn('GetSetting(CLEANUP_SETTINGS_APP, CLEANUP_SETTINGS_SECTION, settingName, defaultText)', helpers)
        self.assertIn('SaveSetting CLEANUP_SETTINGS_APP, CLEANUP_SETTINGS_SECTION, settingName', helpers)
        self.assertNotIn('ActiveDocument.CustomDocumentProperties("CleanupSuiteAutoSave")', helpers)
        self.assertNotIn("GetShowCompletionReviewAfterApplySetting", helpers)
        self.assertNotIn("SetShowCompletionReviewAfterApplySetting", helpers)
        self.assertNotIn("CleanupSuiteShowCompletionReviewAfterApply", helpers)
        self.assertNotIn("ShowCleanupReport", helpers)
        self.assertNotIn("chkShowCompletionReviewAfterApply", launcher)
        self.assertNotIn("chkShowCompletionReviewAfterApply", installer)
        self.assertIn('"Global default: return to main menu after apply"', launcher)
        self.assertIn("If ShouldReturnToLauncherAfterToolSession() Then Me.Show", launcher)
        self.assertNotIn("If GetReturnToMainAfterApplySetting() Then ShowCleanupSuiteLauncher", panel)
        self.assertIn('"chkReturnToMainAfterApply", "chkReturnToMainAfterClose", "chkAutoSave", "lblAutoSaveWarning", "chkUpdateChecks", "cmdCheckUpdates", "cmdResetAll"', installer)
        self.assertIn(
            'Case "frmCleanupSuiteLauncher.chkReturnToMainAfterApply": ControlCaptionText = "Global default: return to main menu after apply"',
            installer,
        )
        self.assertIn("Private Sub cmdResetAll_Click()", launcher)
        self.assertIn("ResetCleanupSuiteDefaults", launcher)

    def test_reconfigure_handles_return_to_main_after_preview_panel_goes_modeless(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        panel = read("src/forms/frmPreviewActions.bas")

        self.assertIn("Public Sub ReturnToLauncherAfterDetachedToolSession()", helpers)
        self.assertIn("If ShouldReturnToLauncherAfterToolSession() Then ShowCleanupSuiteLauncher", helpers)
        self.assertIn("src.Show", panel)
        self.assertLess(
            panel.index("src.Show"),
            panel.index("ReturnToLauncherAfterDetachedToolSession"),
        )

    def test_installer_report_writing_is_unique_and_nonfatal(self):
        installer = read("src/installer/installer.bas")

        self.assertIn("WriteInstallReport reportText", installer)
        self.assertIn("Private Sub WriteInstallReport(reportText As String)", installer)
        self.assertIn('CleanupSuite_Install_Report_" & Format$(Now, "yyyymmdd_hhnnss")', installer)
        self.assertIn("ReportErr:", installer)
        self.assertIn("Installer finished, but the report could not be written or opened.", installer)
        self.assertIn("The Cleanup Suite components were already installed before the report step.", installer)
        self.assertNotIn('reportPath = Environ$("TEMP") & "\\CleanupSuite_Install_Report.txt"', installer)


if __name__ == "__main__":
    unittest.main()
