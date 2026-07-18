Option Explicit

' Output format examples:
' PASS|Launcher|frmCleanupSuiteLauncher opened and exposed its main controls.
' FAIL|Preview Flow|Preview flow did not create frmPreviewActions.
' WARN|Visual review|Confirm the live Word forms still look visually clean.

Public Function RunAlphaSmokeChecklist(Optional ByVal repoRoot As String = "") As String
    Dim results As Collection
    Set results = New Collection

    If Len(Trim$(repoRoot)) = 0 Then repoRoot = SmokeRepoRoot()
    SmokeTrace "RunAlphaSmokeChecklist start"

    results.Add SmokeLauncherCheck()
    SmokeTrace "Launcher check complete"
    results.Add SmokeDocumentTrimCheck()
    SmokeTrace "Document Trim check complete"
    results.Add SmokeUnicodeCheck()
    SmokeTrace "Unicode check complete"
    results.Add SmokeCapitalizationCheck()
    SmokeTrace "Capitalization check complete"
    results.Add SmokeParagraphHelpersCheck()
    SmokeTrace "Paragraph helper check complete"
    results.Add SmokeHeaderFooterCheck()
    SmokeTrace "Header/Footer check complete"
    results.Add SmokeFormattingCleanerCheck()
    SmokeTrace "Formatting Cleaner check complete"
    results.Add SmokeStyleCleanupApplyCheck(repoRoot)
    SmokeTrace "Style Cleanup apply check complete"
    results.Add SmokeRiskControlsCheck()
    SmokeTrace "Risk controls check complete"
    results.Add SmokePreviewFlowCheck(repoRoot)
    SmokeTrace "Preview flow check complete"
    results.Add SmokePreviewSummaryTableCheck(repoRoot)
    SmokeTrace "Preview summary table check complete"
    results.Add SmokeResultLine("WARN", "Visual review", "Confirm the live Word forms still look visually clean and do not leave large wasted vertical space.")
    SmokeTrace "RunAlphaSmokeChecklist end"

    RunAlphaSmokeChecklist = JoinSmokeResults(results)
End Function

Public Function RunCapitalizationAlphaSmokeCheck() As String
    RunCapitalizationAlphaSmokeCheck = CStr(SmokeCapitalizationCheck())
End Function

Public Function RunPreviewLifecycleSmokeChecks() As String
    Dim originalDocument As Document
    Dim previewDocument As Document
    Dim otherDocument As Document
    Dim previewWindow As Window
    Dim panel As Object
    Dim rows As Collection
    Dim contentText As String
    Dim originalViewType As WdViewType
    Dim originalShowHighlight As Boolean
    Dim originalZoom As Long
    Dim originalScroll As Long
    Dim originalSelectionStart As Long
    Dim originalSelectionEnd As Long
    Dim otherViewType As WdViewType
    Dim scrollAvailable As Boolean
    Dim originalShade As Long
    Dim originalBrowseTarget As WdBrowseTarget
    Dim navigationStart As Long
    Dim navigationNextStart As Long
    Dim i As Long

    On Error GoTo SmokeFail
    Set originalDocument = ActiveDocument
    Set previewDocument = Documents.Add
    For i = 1 To 180
        contentText = contentText & "Preview lifecycle smoke paragraph " & CStr(i) & "." & vbCr
    Next i
    previewDocument.Content.Text = contentText
    previewDocument.Activate
    Set previewWindow = previewDocument.ActiveWindow
    previewWindow.View.ReadingLayout = False
    previewWindow.View.Type = wdPrintView
    previewWindow.View.Zoom.Percentage = 115
    previewWindow.Selection.SetRange 5, 12
    On Error Resume Next
    Err.Clear
    previewWindow.VerticalPercentScrolled = 35
    scrollAvailable = (Err.Number = 0)
    Err.Clear
    On Error GoTo SmokeFail
    originalViewType = previewWindow.View.Type
    previewWindow.View.ShowHighlight = False
    originalShowHighlight = previewWindow.View.ShowHighlight
    originalZoom = previewWindow.View.Zoom.Percentage
    If scrollAvailable Then originalScroll = previewWindow.VerticalPercentScrolled
    originalSelectionStart = previewWindow.Selection.Start
    originalSelectionEnd = previewWindow.Selection.End

    SmokeRequireValue BeginPreviewViewSession(), "Preview ON did not start a view session."
    SmokeRequireValue PreviewViewSessionIsActive(), "Preview ON did not mark the view session active."
    SmokeRequireValue previewWindow.View.ReadingLayout, "Preview ON did not enter Reading View."
    SmokeRequireValue previewWindow.View.ShowHighlight, "Preview ON did not make highlight formatting visible."
    previewWindow.Selection.SetRange 30, 30
    RestorePreviewViewSession
    SmokeRequireValue Not PreviewViewSessionIsActive(), "Preview OFF left the view session active."
    SmokeRequireValue Not previewWindow.View.ReadingLayout, "Preview OFF did not leave Reading View."
    SmokeRequireValue previewWindow.View.Type = originalViewType, "Preview OFF did not restore the original view type."
    SmokeRequireValue previewWindow.View.ShowHighlight = originalShowHighlight, "Preview OFF did not restore highlight visibility."
    SmokeRequireValue previewWindow.View.Zoom.Percentage = originalZoom, "Preview OFF did not restore zoom."
    If scrollAvailable Then SmokeRequireValue previewWindow.VerticalPercentScrolled = originalScroll, "Preview OFF did not restore scroll position."
    SmokeRequireValue previewWindow.Selection.Start = originalSelectionStart And previewWindow.Selection.End = originalSelectionEnd, "Preview OFF did not restore the selection."
    RestorePreviewViewSession

    previewWindow.View.ReadingLayout = True
    SmokeRequireValue BeginPreviewViewSession(), "An existing Reading View did not start a preview session."
    RestorePreviewViewSession
    SmokeRequireValue previewWindow.View.ReadingLayout, "An existing Reading View was not preserved."
    previewWindow.View.ReadingLayout = False

    Set rows = NewPreviewSummaryRows()
    AddPreviewSummaryRow rows, "Smoke row", 1
    Set panel = VBA.UserForms.Add("frmPreviewActions")
    originalBrowseTarget = Application.Browser.Target
    panel.ConfigureSummary Nothing, "Preview OFF Smoke", rows
    SmokeRequireValue previewWindow.View.ReadingLayout, "Panel Preview ON did not enter Reading View."
    SmokeRequireValue panel.Controls("cmdPreviousPage").Visible, "Preview did not expose Previous page navigation."
    SmokeRequireValue panel.Controls("cmdNextPage").Visible, "Preview did not expose Next page navigation."
    navigationStart = previewWindow.Selection.Start
    panel.Controls("cmdNextPage").Value = True
    navigationNextStart = previewWindow.Selection.Start
    SmokeRequireValue navigationNextStart > navigationStart, "Next page did not advance the Reading View preview."
    panel.Controls("cmdPreviousPage").Value = True
    SmokeRequireValue previewWindow.Selection.Start < navigationNextStart, "Previous page did not move back in the Reading View preview."
    panel.Controls("cmdPreview").Value = True
    SmokeRequireValue Not PreviewViewSessionIsActive(), "Panel Preview OFF left the session active."
    SmokeRequireValue Not previewWindow.View.ReadingLayout, "Panel Preview OFF did not restore the view."
    SmokeRequireValue panel.Controls("cmdPreview").Caption = "Preview is OFF", "Panel Preview OFF did not update its state label."
    SmokeRequireValue Not panel.Controls("cmdPreview").Enabled, "Panel Preview OFF still exposed the unstable direct re-preview shortcut."
    SmokeRequireValue Not panel.Controls("cmdPreviousPage").Visible And Not panel.Controls("cmdNextPage").Visible, "Panel Preview OFF left page navigation visible."
    SmokeRequireValue Application.Browser.Target = originalBrowseTarget, "Preview navigation did not restore Word's browse target."
    SmokeCloseForm panel

    Set rows = NewPreviewSummaryRows()
    AddPreviewSummaryRow rows, "Smoke row", 1
    Set panel = VBA.UserForms.Add("frmPreviewActions")
    panel.ConfigureSummary Nothing, "Apply Smoke", rows
    panel.Controls("cmdApply").Value = True
    SmokeRequireValue Not PreviewViewSessionIsActive(), "Apply left the preview session active."
    SmokeRequireValue Not previewWindow.View.ReadingLayout, "Apply did not restore the view."
    SmokeCloseForm panel

    Set rows = NewPreviewSummaryRows()
    AddPreviewSummaryRow rows, "Smoke row", 1
    originalShade = previewDocument.Paragraphs(1).Range.Shading.BackgroundPatternColor
    ApplyPreviewShading previewDocument.Paragraphs(1).Range
    previewDocument.Paragraphs(2).Range.HighlightColorIndex = wdBrightGreen
    Set panel = VBA.UserForms.Add("frmPreviewActions")
    panel.ConfigureSummary Nothing, "Reconfigure Smoke", rows
    panel.Controls("cmdClear").Value = True
    SmokeRequireValue Not PreviewViewSessionIsActive(), "Reconfigure left the preview session active."
    SmokeRequireValue Not previewWindow.View.ReadingLayout, "Reconfigure did not restore the view."
    SmokeRequireValue previewDocument.Paragraphs(1).Range.Shading.BackgroundPatternColor = originalShade, "Reconfigure left preview paragraph shading behind."
    SmokeRequireValue previewDocument.Paragraphs(2).Range.HighlightColorIndex = wdNoHighlight, "Reconfigure left preview highlighting behind."
    SmokeCloseForm panel

    Set rows = NewPreviewSummaryRows()
    AddPreviewSummaryRow rows, "Smoke row", 1
    previewDocument.Paragraphs(3).Range.Words(1).Bold = True
    originalShade = previewDocument.Paragraphs(3).Range.ParagraphFormat.Shading.BackgroundPatternColor
    ApplyPreviewShading previewDocument.Paragraphs(3).Range
    previewDocument.Paragraphs(4).Range.HighlightColorIndex = wdBrightGreen
    Set panel = VBA.UserForms.Add("frmPreviewActions")
    panel.ConfigureSummary Nothing, "Close Smoke", rows
    SmokeCloseForm panel
    SmokeRequireValue Not PreviewViewSessionIsActive(), "Panel unload left the preview session active."
    SmokeRequireValue Not previewWindow.View.ReadingLayout, "Panel unload did not restore the view."
    SmokeRequireValue previewDocument.Paragraphs(3).Range.ParagraphFormat.Shading.BackgroundPatternColor = originalShade, "Panel unload left mixed-format preview paragraph shading behind."
    SmokeRequireValue previewDocument.Paragraphs(4).Range.HighlightColorIndex = wdNoHighlight, "Panel unload left preview highlighting behind."

    previewDocument.Activate
    SmokeRequireValue BeginPreviewViewSession(), "Switched-window smoke did not start a preview session."
    Set otherDocument = Documents.Add
    otherDocument.Content.Text = "Other document"
    otherDocument.Activate
    otherViewType = otherDocument.ActiveWindow.View.Type
    RestorePreviewViewSession
    SmokeRequireValue (ActiveDocument Is otherDocument), "Restoration stole focus from the newly active document."
    SmokeRequireValue otherDocument.ActiveWindow.View.Type = otherViewType, "Restoration changed the newly active window."
    SmokeRequireValue Not previewWindow.View.ReadingLayout, "Restoration did not clean the original preview window."

    BeginCleanupProgress "Progress Smoke", "Starting"
    SmokeRequireValue CleanupProgressIsActive(), "Progress begin did not activate the shared service."
    UpdateCleanupProgress "Checking", 2, 4
    EndCleanupProgress
    SmokeRequireValue Not CleanupProgressIsActive(), "Progress end left the shared service active."

    otherDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set otherDocument = Nothing
    previewDocument.Activate
    SmokeRequireValue BeginPreviewViewSession(), "Closed-document smoke did not start a preview session."
    previewDocument.Close SaveChanges:=wdDoNotSaveChanges
    Set previewDocument = Nothing
    RestorePreviewViewSession
    SmokeRequireValue Not PreviewViewSessionIsActive(), "Closed-document cleanup left the session active."

    RunPreviewLifecycleSmokeChecks = "PASS|Preview Lifecycle|Reading View, modal page navigation, modeless Preview OFF, preview formatting, switched windows, closed documents, and progress cleanup passed."
    GoTo SmokeCleanup
SmokeFail:
    RunPreviewLifecycleSmokeChecks = "FAIL|Preview Lifecycle|" & CStr(Err.Number) & " - " & Err.Description
SmokeCleanup:
    On Error Resume Next
    RestoreCleanupSuiteTransientState
    SmokeCloseForm panel
    If Not otherDocument Is Nothing Then otherDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Not previewDocument Is Nothing Then previewDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Not originalDocument Is Nothing Then originalDocument.Activate
    On Error GoTo 0
End Function

Private Function SmokeStyleCleanupApplyCheck(ByVal repoRoot As String) As Variant
    Dim originalDocument As Document
    Dim smokeDocument As Document
    Dim toolForm As Object
    Dim paragraphItem As Paragraph
    Dim appliedStyle As Style
    Dim normalStyle As Style
    Dim variantStyle As Style
    Dim fixturePath As String
    Dim smokePath As String
    Dim oldReturnSetting As Boolean
    Dim originalVariantCount As Long
    Dim remainingVariantCount As Long
    Dim normalCount As Long
    On Error GoTo Fail

    oldReturnSetting = GetReturnToMainAfterApplySetting()
    fixturePath = repoRoot & "\Testing Files\TTF-19_StyleCleanup.docx"
    smokePath = Environ$("TEMP") & "\CleanupSuite_style_apply_smoke.docx"
    If Len(Dir$(fixturePath)) = 0 Then Err.Raise vbObjectError + 1410, "SmokeStyleCleanupApplyCheck", "Style Cleanup fixture is missing."
    On Error Resume Next
    Kill smokePath
    On Error GoTo Fail
    FileCopy fixturePath, smokePath

    Set originalDocument = ActiveDocument
    Set smokeDocument = Documents.Open(smokePath, ReadOnly:=False, AddToRecentFiles:=False)
    smokeDocument.Activate
    SetReturnToMainAfterApplySetting False
    Set variantStyle = smokeDocument.Styles("Body Text 2")
    Set normalStyle = smokeDocument.Styles("Normal")

    For Each paragraphItem In smokeDocument.Paragraphs
        If InStr(1, paragraphItem.Range.Text, "Style variant remapping", vbTextCompare) > 0 Then
            Set appliedStyle = paragraphItem.Range.Style
            If StrComp(appliedStyle.NameLocal, variantStyle.NameLocal, vbTextCompare) = 0 Then
                originalVariantCount = originalVariantCount + 1
            End If
        End If
    Next paragraphItem
    SmokeRequireValue originalVariantCount = 1, "Style Cleanup fixture did not start with exactly one Body Text 2 paragraph."

    Set toolForm = VBA.UserForms.Add("frmStyleCleanup")
    toolForm.Controls("chkRemoveUnused").Value = False
    toolForm.Controls("chkRemapVariants").Value = True
    toolForm.RunAfterPreview
    For Each paragraphItem In smokeDocument.Paragraphs
        If InStr(1, paragraphItem.Range.Text, "Style variant remapping", vbTextCompare) > 0 Then
            Set appliedStyle = paragraphItem.Range.Style
            If StrComp(appliedStyle.NameLocal, normalStyle.NameLocal, vbTextCompare) = 0 Then
                normalCount = normalCount + 1
            ElseIf StrComp(appliedStyle.NameLocal, variantStyle.NameLocal, vbTextCompare) = 0 Then
                remainingVariantCount = remainingVariantCount + 1
            End If
        End If
    Next paragraphItem
    SmokeRequireValue remainingVariantCount = 0, "Style Cleanup left the fixture paragraph in Body Text 2."
    SmokeRequireValue normalCount = 2, "Style Cleanup did not leave both paired fixture paragraphs in Normal."

    SmokeStyleCleanupApplyCheck = SmokeResultLine("PASS", "Style Cleanup", "Variant style remapping completed on a temporary fixture without Find/ReplaceAll or a Word crash.")
    GoTo CleanUp

Fail:
    SmokeStyleCleanupApplyCheck = SmokeResultLine("FAIL", "Style Cleanup", Err.Description)
CleanUp:
    On Error Resume Next
    SetReturnToMainAfterApplySetting oldReturnSetting
    If Not toolForm Is Nothing Then Unload toolForm
    If Not smokeDocument Is Nothing Then smokeDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Not originalDocument Is Nothing Then originalDocument.Activate
    Kill smokePath
    On Error GoTo 0
End Function

Private Function SmokeFormattingCleanerCheck() As Variant
    Dim toolForm As Object
    On Error GoTo Fail
    Set toolForm = VBA.UserForms.Add("frmFormattingStripper")
    SmokeRequireControl toolForm, "chkResetChar"
    SmokeRequireControl toolForm, "chkResetPara"
    SmokeFormattingCleanerCheck = SmokeResultLine("PASS", "Formatting Cleaner", "Formatting Cleaner opened with supported Word formatting APIs.")
    GoTo CleanUp
Fail:
    SmokeFormattingCleanerCheck = SmokeResultLine("FAIL", "Formatting Cleaner", Err.Description)
CleanUp:
    SmokeCloseForm toolForm
End Function

Private Function SmokeLauncherCheck() As Variant
    Dim launcher As Object
    On Error GoTo Fail

    Set launcher = VBA.UserForms.Add("frmCleanupSuiteLauncher")
    SmokeRequireControl launcher, "cmdDocTrim"
    SmokeRequireControl launcher, "cmdUnicode"
    SmokeRequireControl launcher, "chkReturnToMainAfterApply"
    SmokeRequireControl launcher, "chkReturnToMainAfterClose"
    SmokeRequireControl launcher, "chkAutoSave"
    SmokeRequireControl launcher, "lblAutoSaveWarning"
    SmokeRequireControl launcher, "chkUpdateChecks"
    SmokeRequireControl launcher, "cmdCheckUpdates"
    SmokeRequireValue Trim$(launcher.Controls("lblLauncherTitle").Caption) = "Cleanup Suite", "Launcher title did not initialize as expected."
    SmokeRequireValue launcher.Controls("chkAutoSave").Top < launcher.Controls("chkReturnToMainAfterApply").Top, "Auto-save should be the first global default."
    SmokeRequireValue launcher.Controls("lblAutoSaveWarning").Top > launcher.Controls("chkAutoSave").Top, "Auto-save warning should sit below its setting."
    SmokeRequireValue launcher.Controls("lblAutoSaveWarning").Top < launcher.Controls("chkReturnToMainAfterApply").Top, "Auto-save warning should not crowd the next setting."
    SmokeRequireValue launcher.Controls("chkUpdateChecks").Top > launcher.Controls("chkReturnToMainAfterClose").Top, "Update preference should follow the return settings."
    On Error Resume Next
    Dim obsoleteControl As Object
    Set obsoleteControl = launcher.Controls("chkShowCompletionReviewAfterApply")
    On Error GoTo Fail
    SmokeRequireValue obsoleteControl Is Nothing, "Obsolete completion-review setting is still present."

    SmokeLauncherCheck = SmokeResultLine("PASS", "Launcher", "frmCleanupSuiteLauncher opened and exposed its main controls.")
    GoTo CleanUp

Fail:
    SmokeLauncherCheck = SmokeResultLine("FAIL", "Launcher", Err.Description)
CleanUp:
    SmokeCloseForm launcher
End Function

Private Function SmokeDocumentTrimCheck() As Variant
    Dim toolForm As Object
    On Error GoTo Fail

    Set toolForm = VBA.UserForms.Add("frmDocumentTrim")
    SmokeRequireBlankCaption toolForm, "Document Trim"
    SmokeRequireControl toolForm, "cmdPreview"
    SmokeRequireControl toolForm, "cmdRun"
    SmokeRequireControl toolForm, "cmdReset"
    SmokeRequireValue toolForm.Controls("cmdRun").Visible = False, "Guided forms should not expose in-tool Apply."
    SmokeRequireValue toolForm.Controls("cmdReset").Visible = True, "Reset should remain available as a small utility."
    SmokeRequireValue toolForm.Controls("optScopeDocument").Visible = False, "Full-document-only tools should not show scope choices."
    SmokeRequireValue toolForm.Controls("optScopeSelection").Visible = False, "Full-document-only tools should not show selected-scope choices."

    SmokeDocumentTrimCheck = SmokeResultLine("PASS", "Document Trim", "The simplest guided form opened with Preview-only visible action flow.")
    GoTo CleanUp

Fail:
    SmokeDocumentTrimCheck = SmokeResultLine("FAIL", "Document Trim", Err.Description)
CleanUp:
    SmokeCloseForm toolForm
End Function

Private Function SmokeUnicodeCheck() As Variant
    Dim toolForm As Object
    On Error GoTo Fail

    Set toolForm = VBA.UserForms.Add("frmUnicodeCleanup")
    SmokeRequireBlankCaption toolForm, "Invisible Unicode Cleaner"
    SmokeRequireControl toolForm, "optCustom"
    SmokeRequireControl toolForm, "chkNBSP"
    SmokeRequireControl toolForm, "chkZWSP"
    SmokeRequireValue Not SmokeGuidedDividerVisible(toolForm), "Custom divider should be hidden before Custom mode is selected."

    toolForm.Controls("optCustom").Value = True
    LayoutCleanupToolForm toolForm
    SmokeRequireValue SmokeGuidedDividerVisible(toolForm), "Custom divider did not appear after selecting Custom mode."

    toolForm.Controls("optAll").Value = True
    LayoutCleanupToolForm toolForm
    SmokeRequireValue Not SmokeGuidedDividerVisible(toolForm), "Custom divider did not disappear after leaving Custom mode."

    SmokeUnicodeCheck = SmokeResultLine("PASS", "Invisible Unicode Cleaner", "Custom-mode divider and shared guided layout toggled correctly.")
    GoTo CleanUp

Fail:
    SmokeUnicodeCheck = SmokeResultLine("FAIL", "Invisible Unicode Cleaner", Err.Description)
CleanUp:
    SmokeCloseForm toolForm
End Function

Private Function SmokeCapitalizationCheck() As Variant
    Dim toolForm As Object
    Dim numberedHeading As Variant
    Dim repairedHeading As String
    On Error GoTo Fail

    Set toolForm = VBA.UserForms.Add("frmCapitalizationCleanup")
    SmokeRequireBlankCaption toolForm, "Capitalization Fixer"
    SmokeRequireControl toolForm, "optAll"
    SmokeRequireControl toolForm, "optCustom"
    SmokeRequireControl toolForm, "chkSmartSentences"
    SmokeRequireControl toolForm, "chkHeadingParentheses"
    SmokeRequireControl toolForm, "cmdEditExceptions"
    SmokeRequireValue Trim$(toolForm.Controls("optAll").Caption) = "Recommended repair", "Capitalization Fixer did not keep the expected Recommended repair label."
    SmokeRequireValue Trim$(toolForm.Controls("optCustom").Caption) = "Custom repair", "Capitalization Fixer did not keep the expected Custom repair label."
    SmokeRequireValue Trim$(toolForm.Controls("chkHeadingParentheses").Caption) = "Also title-case words inside parentheses", "Capitalization Fixer did not keep the subordinate parenthetical-heading label."
    SmokeRequireValue Not CBool(toolForm.Controls("chkHeadingParentheses").Value), "Capitalization parenthetical heading repair must default to off."
    SmokeRequireValue Trim$(toolForm.Controls("cmdEditExceptions").Caption) = "Edit custom exceptions", "Capitalization Fixer did not keep the custom-exceptions action."

    For Each numberedHeading In Array( _
        "1. The Problem of Evil (Logical + Evidential)", _
        "2. The Problem of Divine Hiddenness", _
        "3. Naturalism Explains Everything Without God", _
        "4. The Incoherence of God (Omnipotence, Omniscience, etc.)", _
        "5. The Success of Evolutionary Biology", _
        "6. The Universe Looks Indifferent, Not Designed for Humans", _
        "7. The Problem of Miracles (Lack of Reliable Evidence)", _
        "8. The Problem of Multiple Religions", _
        "9. The Argument from Nonbelief Across History", _
        "10. The Burden of Proof / Lack of Empirical Evidence")
        repairedHeading = CStr(toolForm.SmokeSmartRepairText(CStr(numberedHeading)))
        SmokeRequireValue repairedHeading = CStr(numberedHeading), _
            "Capitalization Fixer changed a correctly capitalized numbered heading from [" & CStr(numberedHeading) & "] to [" & repairedHeading & "]"
    Next numberedHeading

    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText(ChrW$(8226) & " the origin of species")) = ChrW$(8226) & " The origin of species", _
        "Capitalization Fixer did not repair a lowercase sentence start after a bullet."

    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText("Visit JW.Org, Example.COM, and Anything.BLAH.")) = _
        "Visit JW.org, Example.com, and Anything.blah.", _
        "Capitalization Fixer did not keep dot-attached website/domain segments lowercase."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText("Initials A.B., P.O.Box, and path Example.COM/MyPath stay readable.")) = _
        "Initials A.B., P.O.Box, and path Example.com/MyPath stay readable.", _
        "Capitalization Fixer changed initials or a case-sensitive path while repairing a domain."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairHeadingText("Resources at JW.Org and Example.COM")) = _
        "Resources at JW.org and Example.com", _
        "Capitalization Fixer changed a domain incorrectly inside a heading-style paragraph."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairHeadingText("Divine Hiddenness (weakened)")) = _
        "Divine Hiddenness (weakened)", _
        "Capitalization Fixer changed parenthetical wording while the subordinate option was off."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairHeadingParenthesesText("Consciousness (the Hard Problem)")) = _
        "Consciousness (The Hard Problem)", _
        "Capitalization Fixer did not title-case an opted-in parenthetical heading segment."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairHeadingParenthesesText("Heading (outer words (inner words))")) = _
        "Heading (Outer Words (Inner Words))", _
        "Capitalization Fixer did not safely title-case nested balanced parentheses."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairHeadingParenthesesText("Heading (first note) and (second note)")) = _
        "Heading (First Note) and (Second Note)", _
        "Capitalization Fixer did not safely title-case multiple balanced parenthetical groups."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairHeadingParenthesesText("Kalam Cosmological Argument (a specific version of #1)")) = _
        "Kalam Cosmological Argument (A Specific Version of #1)", _
        "Capitalization Fixer capitalized a minor word before a trailing numeric reference."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairHeadingParenthesesText("The Incoherence of God (Omnipotence, Omniscience, etc.)")) = _
        "The Incoherence of God (Omnipotence, Omniscience, etc.)", _
        "Capitalization Fixer capitalized lowercase etc. in a heading segment."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairHeadingParenthesesText("Teleological / Design in Nature (Biology, DNA)")) = _
        "Teleological / Design in Nature (Biology, DNA)", _
        "Capitalization Fixer changed the protected DNA acronym inside a heading segment."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairHeadingText("God's Attributes and Satan's Rulership")) = _
        "God's Attributes and Satan's Rulership", _
        "Capitalization Fixer capitalized a straight-apostrophe possessive suffix."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairHeadingText("God" & ChrW$(8217) & "s Attributes and JW.org" & ChrW$(8217) & "s Framework")) = _
        "God" & ChrW$(8217) & "s Attributes and JW.org" & ChrW$(8217) & "s Framework", _
        "Capitalization Fixer capitalized a curly-apostrophe possessive suffix."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText("""quoted sentence."" next sentence.")) = _
        """Quoted sentence."" Next sentence.", _
        "Capitalization Fixer did not recognize straight double quotes at sentence boundaries."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText(ChrW$(&H201C) & "quoted sentence." & ChrW$(&H201D) & " next sentence.")) = _
        ChrW$(&H201C) & "Quoted sentence." & ChrW$(&H201D) & " Next sentence.", _
        "Capitalization Fixer did not recognize Word smart double quotes at sentence boundaries."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText("He said,""quoted words.""")) = _
        "He said,""Quoted words.""", _
        "Capitalization Fixer did not recognize a direct quotation after a comma and straight double quote."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText("He said,'quoted words.'")) = _
        "He said,'Quoted words.'", _
        "Capitalization Fixer did not recognize a direct quotation after a comma and straight single quote."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText("He said," & ChrW$(&H201C) & "quoted words." & ChrW$(&H201D))) = _
        "He said," & ChrW$(&H201C) & "Quoted words." & ChrW$(&H201D), _
        "Capitalization Fixer did not recognize a direct quotation after a comma and smart double quote."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText("He said," & ChrW$(&H2018) & "quoted words." & ChrW$(&H2019))) = _
        "He said," & ChrW$(&H2018) & "Quoted words." & ChrW$(&H2019), _
        "Capitalization Fixer did not recognize a direct quotation after a comma and smart single quote."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText("DNA carries genetic information.")) = _
        "DNA carries genetic information.", _
        "Capitalization Fixer changed the protected DNA acronym in ordinary text."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText("DNA, RNA, mRNA, ATP, MRI, PDF, and XLSX remain protected.")) = _
        "DNA, RNA, mRNA, ATP, MRI, PDF, and XLSX remain protected.", _
        "Capitalization Fixer changed an expanded built-in acronym."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText("OpenAI uses GitLab, PowerShell, and macOS.")) = _
        "OpenAI uses GitLab, PowerShell, and macOS.", _
        "Capitalization Fixer changed an expanded protected mixed-case term."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText("Meet at Main Ave. near the park.")) = _
        "Meet at Main Ave. near the park.", _
        "Capitalization Fixer treated an expanded abbreviation as a sentence ending."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairHeadingText("Heading (unfinished words")) = _
        "Heading (unfinished words", _
        "Capitalization Fixer changed unmatched parenthetical wording."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairWithoutHeadingsText("MOST RESILIENT ARGUMENTS FOR GOD")) = _
        "MOST RESILIENT ARGUMENTS FOR GOD", _
        "Capitalization Fixer partially normalized an all-caps heading while heading repair was off."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairWithoutHeadingsText("This sentence has EXCESSIVELY loud text.")) = _
        "This sentence has Excessively loud text.", _
        "Capitalization Fixer stopped ordinary non-heading all-caps repair when heading repair was off."
    SmokeRequireValue CStr(toolForm.SmokeSmartRepairText("THE FINAL SHORTLIST: The Most Resilient Arguments Overall")) = _
        "The Final Shortlist: The Most Resilient Arguments Overall", _
        "Capitalization Fixer did not recognize a heading with a mixed all-caps and title-case pattern."
    SmokeRequireValue CStr(toolForm.SmokeFriendlyCapitalizationExceptions( _
        " MetaDataSuite ; CleanUpSuite ")) = "MetaDataSuite; CleanUpSuite", _
        "Capitalization custom exceptions did not keep the friendly semicolon list."

    toolForm.Controls("optAll").Value = False
    toolForm.Controls("optCustom").Value = True
    LayoutCleanupToolForm toolForm
    SmokeRequireValue toolForm.Controls("chkSentence").Top = toolForm.Controls("chkTitle").Top, _
        "Capitalization Custom row 1 is not aligned as two columns."
    SmokeRequireValue toolForm.Controls("chkUpper").Top = toolForm.Controls("chkLower").Top, _
        "Capitalization Custom row 2 is not aligned as two columns."
    SmokeRequireValue toolForm.Controls("chkSmartSentences").Top = toolForm.Controls("chkHeadingParentheses").Top, _
        "Capitalization Custom row 3 is not aligned as two columns."
    SmokeRequireValue toolForm.Controls("chkTitle").Left > toolForm.Controls("chkSentence").Left And _
                      toolForm.Controls("chkLower").Left > toolForm.Controls("chkUpper").Left And _
                      toolForm.Controls("chkHeadingParentheses").Left > toolForm.Controls("chkSmartSentences").Left, _
        "Capitalization Custom choices are not arranged in two columns."

    SmokeCapitalizationCheck = SmokeResultLine("PASS", "Capitalization Fixer", "Two-column Custom layout, mixed-case headings, expanded protection packs, possessives, quote boundaries, lowercase domains, and safe heading bypass passed.")
    GoTo CleanUp

Fail:
    SmokeCapitalizationCheck = SmokeResultLine("FAIL", "Capitalization Fixer", Err.Description)
CleanUp:
    SmokeCloseForm toolForm
End Function

Private Function SmokeHeaderFooterCheck() As Variant
    Dim toolForm As Object
    On Error GoTo Fail

    Set toolForm = VBA.UserForms.Add("frmHeaderFooterStandardizer")
    SmokeRequireBlankCaption toolForm, "Header / Footer Standardizer"
    SmokeRequireControl toolForm, "optStandardize"
    SmokeRequireControl toolForm, "optClearAll"
    SmokeRequireControl toolForm, "chkBreakLinks"
    SmokeRequireControl toolForm, "chkAlignment"
    SmokeRequireValue Trim$(toolForm.Controls("chkBreakLinks").Caption) = "Unlink sections (set each independently)", "Header / Footer Standardizer lost the unlink-sections caption."

    SmokeHeaderFooterCheck = SmokeResultLine("PASS", "Header / Footer Standardizer", "Special-case layout form opened with its alignment and unlink controls intact.")
    GoTo CleanUp

Fail:
    SmokeHeaderFooterCheck = SmokeResultLine("FAIL", "Header / Footer Standardizer", Err.Description)
CleanUp:
    SmokeCloseForm toolForm
End Function

Private Function SmokeRiskControlsCheck() As Variant
    Dim objectForm As Object
    Dim headerForm As Object
    On Error GoTo Fail

    Set objectForm = VBA.UserForms.Add("frmObjectRemover")
    SmokeRequireBlankCaption objectForm, "Object Remover"
    SmokeRequireControl objectForm, "cmdRiskPlacementObjectTables"
    SmokeRequireControl objectForm, "cmdRiskPlacementObjectPictures"
    SmokeRequireValue objectForm.Controls("cmdRun").Visible = False, "Object Remover should not expose in-tool Apply."
    SmokeRequireValue objectForm.Controls("cmdRiskPlacementObjectTables").Visible = True, "Object Remover table risk control should be visible."

    Set headerForm = VBA.UserForms.Add("frmHeaderFooterStandardizer")
    SmokeRequireBlankCaption headerForm, "Header / Footer Standardizer"
    SmokeRequireControl headerForm, "cmdRiskPlacementHeaderFooterClear"
    SmokeRequireControl headerForm, "cmdRiskPlacementHeaderFooterBreakLinks"
    SmokeRequireValue headerForm.Controls("cmdRun").Visible = False, "Header / Footer Standardizer should not expose in-tool Apply."
    SmokeRequireValue headerForm.Controls("cmdRiskPlacementHeaderFooterClear").Visible = True, "Header / Footer clear risk control should be visible."

    SmokeRiskControlsCheck = SmokeResultLine("PASS", "Risk Controls", "Shared risk dispatcher shows high-impact risk controls without exposing in-tool Apply.")
    GoTo CleanUp

Fail:
    SmokeRiskControlsCheck = SmokeResultLine("FAIL", "Risk Controls", Err.Description)
CleanUp:
    SmokeCloseForm objectForm
    SmokeCloseForm headerForm
End Function

Private Function SmokePreviewFlowCheck(ByVal repoRoot As String) As String
    Dim toolForm As Object
    Dim panel As Object
    Dim flowRows As Collection
    On Error GoTo Fail

    Set toolForm = VBA.UserForms.Add("frmDocumentTrim")
    Set flowRows = NewPreviewSummaryRows()
    AddPreviewSummaryRow flowRows, "Trailing empty paragraphs", 3
    Set panel = VBA.UserForms.Add("frmPreviewActions")
    panel.ConfigureSummary toolForm, "Document Trim", flowRows

    SmokeRequireValue Trim$(panel.Caption) = "", "Preview panel title bar should be blank."
    SmokeRequireValue Trim$(panel.Controls("cmdApply").Caption) = "Apply", "Preview panel Apply button label changed."
    SmokeRequireValue Trim$(panel.Controls("cmdClear").Caption) = "Reconfigure", "Preview panel Reconfigure button label changed."
    SmokeRequireValue Left$(Trim$(panel.Controls("cmdPreview").Caption), 10) = "Preview is", "Preview panel toggle label was not initialized."
    SmokeRequireValue panel.Controls("lblSummaryItem0").Caption = "Trailing empty paragraphs", "Preview panel did not render the configured summary row."

    SmokePreviewFlowCheck = SmokeResultLine("PASS", "Preview Flow", "Document Trim preview panel configured with Apply, Reconfigure, Preview, and summary controls.")
    GoTo CleanUp

Fail:
    SmokePreviewFlowCheck = SmokeResultLine("FAIL", "Preview Flow", Err.Description)
CleanUp:
    On Error Resume Next
    If Not panel Is Nothing Then
        panel.Hide
        Unload panel
    End If
    Set gPreviewActionPanel = Nothing
    SmokeCloseForm toolForm
    On Error GoTo 0
End Function

Private Function SmokePreviewSummaryTableCheck(ByVal repoRoot As String) As Variant
    Dim toolForm As Object
    Dim panel As Object
    Dim flowRows As Collection
    Dim wideRows As Collection
    On Error GoTo Fail

    Set toolForm = VBA.UserForms.Add("frmDocumentTrim")
    Set flowRows = NewPreviewSummaryRows()
    AddPreviewSummaryRow flowRows, "Trailing empty paragraphs", 3
    Set panel = VBA.UserForms.Add("frmPreviewActions")
    panel.ConfigureSummary toolForm, "Document Trim", flowRows
    SmokeRequireValue panel.Controls("lblSummaryHeaderItem").Caption = "Item", "Preview summary Item header missing."
    SmokeRequireValue panel.Controls("lblSummaryHeaderCount").Caption = "#", "Preview summary # header missing."
    SmokeRequireValue panel.Controls("lblSummaryItem0").Caption = "Trailing empty paragraphs", "Preview summary first row item was not rendered."
    SmokeRequireSummaryRowFits panel, "lblSummaryItem0"

    panel.Hide
    Unload panel
    Set panel = Nothing
    Set gPreviewActionPanel = Nothing

    Set wideRows = NewPreviewSummaryRows()
    AddPreviewSummaryRow wideRows, "Pictures", 0, True
    AddPreviewSummaryRow wideRows, "Text boxes", 0, True
    AddPreviewSummaryRow wideRows, "Frames", 0, True
    AddPreviewSummaryRow wideRows, "Horizontal lines", 0, True
    AddPreviewSummaryRow wideRows, "Form controls", 0, True
    AddPreviewSummaryRow wideRows, "Hidden text", 0, True
    AddPreviewSummaryRow wideRows, "Tables", 0, True
    Set panel = VBA.UserForms.Add("frmPreviewActions")
    panel.ConfigureSummary toolForm, "Object Remover", wideRows
    SmokeRequireValue panel.Controls("lblSummaryItem6").Caption = "Tables (unhighlighted)", "Preview summary unhighlighted suffix was not rendered."
    SmokeRequireSummaryRowFits panel, "lblSummaryItem6"

    SmokePreviewSummaryTableCheck = SmokeResultLine("PASS", "Preview Summary Table", "Preview panel rendered all fixed Item/# rows inside the visible form.")
    GoTo CleanUp

Fail:
    SmokePreviewSummaryTableCheck = SmokeResultLine("FAIL", "Preview Summary Table", Err.Description)
CleanUp:
    On Error Resume Next
    RemoveAllHighlighting ActiveDocument.Content
    If Not panel Is Nothing Then
        panel.Hide
        Unload panel
    End If
    Set gPreviewActionPanel = Nothing
    SmokeCloseForm toolForm
    On Error GoTo 0
End Function

Private Function SmokeParagraphHelpersCheck() As Variant
    Dim originalDocument As Document
    Dim testDocument As Document
    Dim normalParagraph As Paragraph
    Dim cellParagraph As Paragraph
    Dim scopeRange As Range
    Dim testTable As Table

    On Error GoTo Fail
    Set originalDocument = ActiveDocument
    Set testDocument = Documents.Add
    testDocument.Content.Text = "First body" & vbCr & "Second body" & vbCr
    Set normalParagraph = testDocument.Paragraphs(1)

    SmokeRequireValue ParagraphBodyText(normalParagraph) = "First body", "Paragraph body text included its paragraph marker."
    Set scopeRange = normalParagraph.Range.Duplicate
    SmokeRequireValue ParagraphContainedInRange(normalParagraph, scopeRange), "A paragraph was not contained in its own range."
    Set scopeRange = testDocument.Range(normalParagraph.Range.Start + 1, normalParagraph.Range.End - 1)
    SmokeRequireValue ParagraphOverlapsRange(normalParagraph, scopeRange), "A partial paragraph scope did not overlap its paragraph."
    SmokeRequireValue Not ParagraphContainedInRange(normalParagraph, scopeRange), "A partial paragraph scope was treated as containing the whole paragraph."

    ReplaceParagraphBodyText normalParagraph, "Changed body"
    SmokeRequireValue ParagraphBodyText(normalParagraph) = "Changed body", "Paragraph body replacement did not replace the visible text."
    SmokeRequireValue Right$(normalParagraph.Range.Text, 1) = vbCr, "Paragraph body replacement removed the paragraph marker."

    Set testTable = testDocument.Tables.Add(testDocument.Range(testDocument.Content.End - 1, testDocument.Content.End - 1), 1, 1)
    testTable.Cell(1, 1).Range.Text = "Cell body"
    Set cellParagraph = testTable.Cell(1, 1).Range.Paragraphs(1)
    SmokeRequireValue ParagraphBodyText(cellParagraph) = "Cell body", "Table-cell body text included its cell markers."
    ReplaceParagraphBodyText cellParagraph, "Changed cell"
    SmokeRequireValue ParagraphBodyText(cellParagraph) = "Changed cell", "Table-cell body replacement did not replace the visible text."
    SmokeRequireValue Right$(cellParagraph.Range.Text, 2) = vbCr & Chr$(7), "Table-cell body replacement removed a cell marker."
    Set scopeRange = cellParagraph.Range.Duplicate
    SmokeRequireValue Not ParagraphOverlapsRangeOutsideTable(cellParagraph, scopeRange), "The outside-table scope helper accepted a table paragraph."

    SmokeParagraphHelpersCheck = SmokeResultLine("PASS", "Paragraph Helpers", "Shared paragraph body and scope helpers preserved paragraph and table-cell markers.")
    GoTo CleanUp

Fail:
    SmokeParagraphHelpersCheck = SmokeResultLine("FAIL", "Paragraph Helpers", Err.Description)
CleanUp:
    On Error Resume Next
    If Not testDocument Is Nothing Then testDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Not originalDocument Is Nothing Then originalDocument.Activate
    On Error GoTo 0
End Function

Private Function SmokeRepoRoot() As Variant
    Dim hostPath As String
    hostPath = ThisDocument.Path
    If Right$(hostPath, Len("\Practice - Try CleanupSuite Here")) = "\Practice - Try CleanupSuite Here" Then
        SmokeRepoRoot = Left$(hostPath, Len(hostPath) - Len("\Practice - Try CleanupSuite Here"))
    Else
        SmokeRepoRoot = hostPath
    End If
End Function

Private Function JoinSmokeResults(ByVal results As Collection) As String
    Dim lineItem As Variant
    For Each lineItem In results
        If Len(JoinSmokeResults) > 0 Then JoinSmokeResults = JoinSmokeResults & vbCrLf
        JoinSmokeResults = JoinSmokeResults & CStr(lineItem)
    Next lineItem
End Function

Private Function SmokeResultLine(ByVal statusText As String, ByVal checkName As String, ByVal detailText As String) As String
    SmokeResultLine = statusText & "|" & checkName & "|" & detailText
End Function

Private Sub SmokeRequireControl(ByVal formObject As Object, ByVal controlName As String)
    Dim ctl As Object
    On Error Resume Next
    Set ctl = formObject.Controls(controlName)
    On Error GoTo 0
    If ctl Is Nothing Then
        Err.Raise vbObjectError + 1400, "SmokeRequireControl", "Missing control " & controlName & " on " & formObject.Name & "."
    End If
End Sub

Private Sub SmokeRequireBlankCaption(ByVal formObject As Object, ByVal friendlyName As String)
    SmokeRequireValue Trim$(formObject.Caption) = "", friendlyName & " should blank the native UserForm caption to avoid duplicate titles."
End Sub

Private Sub SmokeRequireValue(ByVal conditionHolds As Boolean, ByVal messageText As String)
    If Not conditionHolds Then
        Err.Raise vbObjectError + 1401, "SmokeRequireValue", messageText
    End If
End Sub

Private Sub SmokeRequireSummaryRowFits(ByVal panel As Object, ByVal controlName As String)
    Dim ctl As Object
    Set ctl = panel.Controls(controlName)
    SmokeRequireValue ctl.Top + ctl.Height <= panel.InsideHeight - 2, controlName & " extends below the visible preview panel area."
End Sub

Private Function SmokeGuidedDividerVisible(ByVal formObject As Object) As Boolean
    On Error Resume Next
    SmokeGuidedDividerVisible = formObject.Controls("lblGuidedDivider").Visible
    On Error GoTo 0
End Function

Private Sub SmokeCloseForm(ByRef formObject As Object)
    On Error Resume Next
    If Not formObject Is Nothing Then
        formObject.Hide
        Unload formObject
        Set formObject = Nothing
    End If
    On Error GoTo 0
End Sub

Private Sub SmokeTrace(ByVal messageText As String)
    On Error Resume Next
    Dim fileNum As Integer
    fileNum = FreeFile
    Open Environ$("TEMP") & "\CleanupSuite_alpha_smoke_trace.txt" For Append As #fileNum
    Print #fileNum, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & messageText
    Close #fileNum
    On Error GoTo 0
End Sub
