from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]


def read(path: str) -> str:
    return (ROOT / path).read_text(encoding="utf-8")


class PreviewReadingViewAndProgressTests(unittest.TestCase):
    def test_preview_session_captures_and_restores_window_state(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("Private gPreviewViewDocument As Document", helpers)
        self.assertIn("Private gPreviewViewWindow As Window", helpers)
        self.assertIn("Private gPreviewViewSelection As Range", helpers)
        self.assertIn("Private gPreviewViewType As WdViewType", helpers)
        self.assertIn("Private gPreviewViewReadingLayout As Boolean", helpers)
        self.assertIn("Private gPreviewViewShowHighlight As Boolean", helpers)
        self.assertIn("Private gPreviewViewZoom As Long", helpers)
        self.assertIn("Private gPreviewViewVerticalScroll As Long", helpers)
        self.assertIn("Private gPreviewViewHasVerticalScroll As Boolean", helpers)
        self.assertIn("Set gPreviewViewSelection = previewSelection.Range.Duplicate", helpers)
        self.assertIn("gPreviewViewWindow.View.ReadingLayout = True", helpers)
        self.assertIn("gPreviewViewShowHighlight = gPreviewViewWindow.View.ShowHighlight", helpers)
        self.assertIn("gPreviewViewWindow.View.ShowHighlight = True", helpers)
        self.assertIn(".View.Type = gPreviewViewType", helpers)
        self.assertIn(".View.ShowHighlight = gPreviewViewShowHighlight", helpers)
        self.assertIn(".View.Zoom.Percentage = gPreviewViewZoom", helpers)
        self.assertIn("If gPreviewViewHasVerticalScroll Then gPreviewViewWindow.VerticalPercentScrolled = gPreviewViewVerticalScroll", helpers)
        begin = helpers[helpers.index("Public Function BeginPreviewViewSession()"):helpers.index("Public Sub RestorePreviewViewSession()")]
        self.assertLess(
            begin.index("gPreviewViewWindow.View.ShowHighlight = True"),
            begin.index("gPreviewViewWindow.View.ReadingLayout = True"),
        )

    def test_preview_restoration_is_idempotent_and_window_specific(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        start = helpers.index("Public Sub RestorePreviewViewSession()")
        end = helpers.index("Private Sub ClearPreviewViewSessionState()", start)
        restore = helpers[start:end]

        self.assertIn("If Not gPreviewViewSessionActive Then", restore)
        self.assertIn("gPreviewViewSessionActive = False", restore)
        self.assertIn("If liveDocument Is gPreviewViewDocument Then", restore)
        self.assertIn("If liveWindow Is gPreviewViewWindow Then", restore)
        self.assertIn("If liveWindow.Document Is gPreviewViewDocument Then previewWindowIsLive = True", restore)
        self.assertIn("If ActiveWindow Is gPreviewViewWindow Then", restore)
        self.assertIn("gPreviewViewSelection.Select", restore)
        self.assertLess(
            restore.index("gPreviewViewSessionActive = False"),
            restore.index("If previewWindowIsLive Then"),
        )

    def test_panel_cleans_up_every_preview_exit_path(self):
        panel = read("src/forms/frmPreviewActions.bas")

        self.assertIn("BeginPreviewViewSession", panel)
        self.assertIn("Private mPreviewDocument As Document", panel)
        self.assertIn("Private mPreviewWindow As Window", panel)
        self.assertIn("Private Function PreviewTargetIsActive() As Boolean", panel)
        self.assertIn("Private Sub RestoreEditablePreviewDocument()", panel)
        cleanup_start = panel.index("Private Sub RestoreEditablePreviewDocument()")
        cleanup_end = panel.index("Private Sub cmdApply_Click()", cleanup_start)
        cleanup = panel[cleanup_start:cleanup_end]
        self.assertIn("RemovePreviewHighlighting mPreviewDocument", cleanup)
        self.assertLess(
            cleanup.index("RestoreCleanupSuiteTransientState"),
            cleanup.index("RemovePreviewHighlighting mPreviewDocument"),
        )
        self.assertIn("Private Sub UserForm_Terminate()", panel)

        for procedure, next_procedure in [
            ("Private Sub cmdApply_Click()", "Private Sub cmdPreview_Click()"),
            ("Private Sub cmdPreview_Click()", "Private Sub cmdClear_Click()"),
            ("Private Sub cmdClear_Click()", "Private Function PreviewTargetIsActive()"),
        ]:
            block = panel[panel.index(procedure): panel.index(next_procedure)]
            self.assertIn("RestoreEditablePreviewDocument", block)

        query_close = panel[panel.index("Private Sub UserForm_QueryClose"):panel.index("Private Sub UserForm_Terminate")]
        terminate = panel[panel.index("Private Sub UserForm_Terminate"):]
        self.assertIn("RestoreEditablePreviewDocument", query_close)
        self.assertIn("RestoreEditablePreviewDocument", terminate)

    def test_preview_panel_is_modal_while_on_and_modeless_only_after_preview_off(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("gPreviewActionPanel.Show vbModal", helpers)
        self.assertIn("gPreviewActionPanel.Show vbModeless", helpers)
        self.assertLess(
            helpers.index("gPreviewActionPanel.Show vbModal"),
            helpers.index("gPreviewActionPanel.Show vbModeless"),
        )

    def test_modal_preview_navigation_uses_word_page_browser_and_restores_target(self):
        panel = read("src/forms/frmPreviewActions.bas")

        self.assertIn("Private mOriginalBrowseTarget As Long", panel)
        self.assertIn("mOriginalBrowseTarget = Application.Browser.Target", panel)
        self.assertIn("Application.Browser.Target = wdBrowsePage", panel)
        self.assertIn("Application.Browser.Previous", panel)
        self.assertIn("Application.Browser.Next", panel)
        self.assertIn("If mBrowseTargetCaptured Then Application.Browser.Target = mOriginalBrowseTarget", panel)
        self.assertIn('cmdPreviousPage.Caption = ChrW$(&H25C0) & "  Previous page"', panel)
        self.assertIn('cmdNextPage.Caption = "Next page  " & ChrW$(&H25B6)', panel)

    def test_preview_shading_is_restored_from_document_coordinates(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("Private gPreviewShadingDocuments As Collection", helpers)
        self.assertIn("Private gPreviewShadingStarts As Collection", helpers)
        self.assertIn("Private gPreviewShadingEnds As Collection", helpers)
        self.assertIn("Private gPreviewShadingKeys As Object", helpers)
        self.assertIn('Set gPreviewShadingKeys = CreateObject("Scripting.Dictionary")', helpers)
        self.assertIn("For Each paragraphItem In targetRange.Paragraphs", helpers)
        self.assertIn("ApplyPreviewShadingToParagraph paragraphRange, shadeColor", helpers)
        self.assertIn("Private Sub ApplyPreviewShadingToParagraph(ByVal paragraphRange As Range, ByVal shadeColor As Long)", helpers)
        self.assertIn("PreviewShadingDocumentIdentity(paragraphRange.Document)", helpers)
        self.assertIn("If gPreviewShadingKeys.Exists(shadingKey) Then", helpers)
        self.assertIn("gPreviewShadingDocuments.Add paragraphRange.Document", helpers)
        self.assertIn("gPreviewShadingStarts.Add CLng(paragraphRange.Start)", helpers)
        self.assertIn("gPreviewShadingEnds.Add CLng(paragraphRange.End)", helpers)
        self.assertIn("gPreviewShadingBackgrounds.Add CLng(paragraphRange.ParagraphFormat.Shading.BackgroundPatternColor)", helpers)
        self.assertIn("Set restoreRange = previewDocument.Range(rangeStart, rangeEnd)", helpers)
        self.assertIn("restoreRange.ParagraphFormat.Shading.Texture = expectedTexture", helpers)
        self.assertIn("restoreRange.ParagraphFormat.Shading.ForegroundPatternColor = expectedForeground", helpers)
        self.assertIn("restoreRange.ParagraphFormat.Shading.BackgroundPatternColor = expectedBackground", helpers)
        self.assertNotIn("gPreviewShadingRanges", helpers)

    def test_transient_cleanup_retries_verified_preview_shading_restoration(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        cleanup_start = helpers.index("Public Sub RestoreCleanupSuiteTransientState()")
        cleanup = helpers[cleanup_start:helpers.index("Public Sub RemoveAllHighlighting", cleanup_start)]
        restore_start = helpers.index("Private Sub RestorePreviewShading()")
        restore = helpers[restore_start:helpers.index("Public Function DocumentHasVisibleContent", restore_start)]

        self.assertIn("RestorePreviewShading", cleanup)
        self.assertIn("allLiveRecordsRestored = False", restore)
        self.assertIn("If allLiveRecordsRestored Then", restore)
        self.assertIn("ParagraphFormat.Shading.BackgroundPatternColor) <> expectedBackground", restore)

    def test_runtime_smoke_covers_reconfigure_and_unload_format_cleanup(self):
        runner = read("src/modules/modAlphaSmokeRunner.bas")

        self.assertIn('"Reconfigure left preview paragraph shading behind."', runner)
        self.assertIn('"Reconfigure left preview highlighting behind."', runner)
        self.assertIn("previewDocument.Paragraphs(3).Range.Words(1).Bold = True", runner)
        self.assertIn('"Panel unload left mixed-format preview paragraph shading behind."', runner)
        self.assertIn('"Panel unload left preview highlighting behind."', runner)

    def test_panel_does_not_apply_to_a_switched_document_or_window(self):
        panel = read("src/forms/frmPreviewActions.bas")
        apply_start = panel.index("Private Sub cmdApply_Click()")
        apply_end = panel.index("Private Sub cmdPreview_Click()", apply_start)
        apply_block = panel[apply_start:apply_end]

        self.assertIn("If Not PreviewTargetIsActive() Then", apply_block)
        self.assertIn("If Not ActiveDocument Is mPreviewDocument Then Exit Function", panel)
        self.assertIn("If Not ActiveWindow Is mPreviewWindow Then Exit Function", panel)
        self.assertLess(
            apply_block.index("If Not PreviewTargetIsActive() Then"),
            apply_block.index('CallByName mSourceForm, "RunAfterPreview", VbMethod'),
        )

    def test_reset_close_and_word_exit_restore_transient_state(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        launcher = read("src/modules/modCleanupLauncher.bas")

        self.assertIn("RestoreCleanupSuiteTransientState", helpers[helpers.index("Public Sub ResetCleanupSuiteDefaults()"):])
        self.assertIn("RestoreCleanupSuiteTransientState", helpers[helpers.index("Public Sub MarkCleanupToolClosedByUser()"):])
        self.assertIn("Public Sub AutoClose()", launcher)
        self.assertIn("Public Sub AutoExit()", launcher)
        self.assertEqual(launcher.count("RestoreCleanupSuiteTransientState"), 2)

    def test_shared_progress_service_is_balanced_and_behavior_neutral(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("Public Sub BeginCleanupProgress", helpers)
        self.assertIn("Public Sub UpdateCleanupProgress", helpers)
        self.assertIn("Public Sub EndCleanupProgress", helpers)
        self.assertIn("PublishCleanupProgress CleanupProgressText", helpers)
        self.assertIn('If gCleanupProgressActive Then Application.StatusBar = ""', helpers)
        self.assertNotIn("CStr(Application.StatusBar)", helpers)
        self.assertIn('BeginCleanupProgress toolName, "Applying changes"', helpers)
        self.assertIn("RestoreCleanupSuiteTransientState", helpers[helpers.index("Public Sub MarkCleanupEnd()"):])
        self.assertNotIn("DoEvents", helpers[helpers.index("Public Sub BeginCleanupProgress"): helpers.index("Public Sub RemoveAllHighlighting")])

    def test_progress_falls_back_to_the_existing_preview_panel(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        panel = read("src/forms/frmPreviewActions.bas")

        self.assertIn('CallByName gPreviewActionPanel, "ShowProgress", VbMethod, statusText', helpers)
        self.assertIn('CallByName gPreviewActionPanel, "ClearProgress", VbMethod', helpers)
        self.assertIn("Public Sub ShowProgress(ByVal progressText As String)", panel)
        self.assertIn("Public Sub ClearProgress()", panel)
        self.assertIn("lblHint.Caption = progressText", panel)
        self.assertIn('cmdApply.Caption = "Applying..."', panel)
        self.assertIn("cmdApply.BackColor = RGB(221, 235, 247)", panel)
        self.assertIn("cmdApply.ForeColor = RGB(31, 78, 121)", panel)
        self.assertIn('cmdApply.Caption = "Apply"', panel)
        self.assertIn("cmdApply.BackColor = mApplyButtonBackColor", panel)
        self.assertIn("cmdApply.ForeColor = mApplyButtonForeColor", panel)
        self.assertIn("cmdPreview.Enabled = False", panel)
        self.assertIn("cmdClear.Enabled = False", panel)
        self.assertIn("cmdApply.Enabled = False", panel)
        self.assertIn("cmdPreviousPage.Enabled = False", panel)
        self.assertIn("cmdNextPage.Enabled = False", panel)
        self.assertIn("Me.Repaint", panel)
        apply_block = panel[panel.index("Private Sub cmdApply_Click()"): panel.index("Private Sub cmdPreview_Click()")]
        self.assertIn('ShowProgress mToolName & ": Applying changes"', apply_block)
        self.assertNotIn("Me.Hide", apply_block[apply_block.index('ShowProgress mToolName & ": Applying changes"'):])

    def test_preview_indicator_uses_the_existing_button_and_restores_its_state(self):
        helpers = read("src/modules/modCleanupHelpers.bas")

        self.assertIn("Public Sub BeginPreviewActionIndicator(ByVal sourceForm As Object)", helpers)
        self.assertIn("Public Sub EndPreviewActionIndicator()", helpers)
        self.assertIn('previewButton.Caption = "Previewing..."', helpers)
        self.assertIn("previewButton.BackColor = RGB(221, 235, 247)", helpers)
        self.assertIn("previewButton.ForeColor = RGB(31, 78, 121)", helpers)
        self.assertIn("previewButton.BackColor = gPreviewActionIndicatorBackColor", helpers)
        self.assertIn("previewButton.ForeColor = gPreviewActionIndicatorForeColor", helpers)
        self.assertIn("EndPreviewActionIndicator", helpers[helpers.index("Public Sub RestoreCleanupSuiteTransientState()"):])

    def test_preview_on_caption_remains_bold(self):
        panel = read("src/forms/frmPreviewActions.bas")

        self.assertIn('cmdPreview.Caption = "Preview is ON"', panel)
        self.assertIn("StyleButton cmdPreview, True", panel)

    def test_slowest_structural_tools_report_meaningful_phases(self):
        expected = {
            "src/forms/frmParagraphCleanup.bas": [
                "Removing extra empty paragraphs",
                "Converting soft returns",
                "Normalizing paragraph spacing",
                "Fixing manual indents",
            ],
            "src/forms/frmBreakNormalizer.bas": [
                "Collapsing section breaks",
                "Collapsing page breaks",
                "Converting section breaks",
            ],
            "src/forms/frmStyleCleanup.bas": [
                "Remapping style variants",
                "Checking unused styles",
            ],
            "src/forms/frmObjectRemover.bas": [
                "Removing pictures",
                "Removing hidden text",
            ],
            "src/forms/frmTableCleaner.bas": ["Cleaning tables"],
        }

        for path, phases in expected.items():
            source = read(path)
            with self.subTest(path=path):
                self.assertIn("MarkCleanupStart", source)
                self.assertIn("MarkCleanupEnd", source)
                for phase in phases:
                    self.assertIn(f'UpdateCleanupProgress "{phase}"', source)

    def test_prototype_does_not_protect_the_document(self):
        helpers = read("src/modules/modCleanupHelpers.bas")
        panel = read("src/forms/frmPreviewActions.bas")
        preview_code = helpers[helpers.index("Public Function BeginPreviewViewSession"): helpers.index("Public Sub RemoveAllHighlighting")]

        self.assertNotIn(".Protect", preview_code)
        self.assertNotIn(".Unprotect", preview_code)
        self.assertNotIn(".Protect", panel)
        self.assertNotIn(".Unprotect", panel)

    def test_runtime_smoke_covers_required_lifecycle_cases(self):
        runner = read("src/modules/modAlphaSmokeRunner.bas")
        script = read("scripts/run_preview_lifecycle_smoke.ps1")

        self.assertIn("Public Function RunPreviewLifecycleSmokeChecks() As String", runner)
        for snippet in [
            'panel.Controls("cmdPreview").Value = True',
            'panel.Controls("cmdNextPage").Value = True',
            'panel.Controls("cmdPreviousPage").Value = True',
            'panel.Controls("cmdApply").Value = True',
            'panel.Controls("cmdClear").Value = True',
            "SmokeCloseForm panel",
            "previewWindow.View.ReadingLayout = True",
            "Set otherDocument = Documents.Add",
            "previewDocument.Close SaveChanges:=wdDoNotSaveChanges",
            "BeginCleanupProgress",
            "EndCleanupProgress",
        ]:
            with self.subTest(snippet=snippet):
                self.assertIn(snippet, runner)
        self.assertIn('"Next page did not advance the Reading View preview."', runner)
        self.assertIn('"Preview navigation did not restore Word\'s browse target."', runner)
        self.assertIn('modAlphaSmokeRunner.RunPreviewLifecycleSmokeChecks', script)
        self.assertIn('Get-Process WINWORD -ErrorAction SilentlyContinue | Stop-Process -Force', script)
        self.assertIn('$suiteDoc.Close([ref]$false)', script)

    def test_word_automation_cannot_leave_reading_view_as_the_last_view(self):
        scripts = [
            read("scripts/run_preview_lifecycle_smoke.ps1"),
            read("scripts/run_word_alpha_smoke.ps1"),
            read("scripts/sync_docm_code_only.ps1"),
        ]

        for script in scripts:
            with self.subTest(script=script[:80]):
                self.assertIn(".ActiveWindow.View.ReadingLayout = $false", script)
                self.assertIn(".ActiveWindow.View.Type = 3", script)


if __name__ == "__main__":
    unittest.main()
