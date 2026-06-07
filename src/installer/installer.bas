' ===========================================================================
'  Professional Cleanup Suite for Microsoft Word
'  A self-installing collection of document cleanup tools.
'
'  MIT License
'
'  Copyright (c) 2026 Christopher Travis Massey
'
'  Permission is hereby granted, free of charge, to any person obtaining a copy
'  of this software and associated documentation files (the "Software"), to deal
'  in the Software without restriction, including without limitation the rights
'  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'  copies of the Software, and to permit persons to whom the Software is
'  furnished to do so, subject to the following conditions:
'
'  The above copyright notice and this permission notice shall be included in all
'  copies or substantial portions of the Software.
'
'  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'  SOFTWARE.
' ===========================================================================

' ---------------------------------------------------------------------------
'  DEVELOPERS: To add a new cleanup tool to this suite, see CONTRIBUTING.md
'  in the repository.  In brief, a new tool must be wired into EIGHT places:
'    1. a Private Function frmXxx_Code() builder (the form code, as a string)
'    2. CreateOrReplaceForm ... in InstallCleanupSuite
'    3. AttemptCreateControls ... in InstallCleanupSuite
'    4. a Case in ControlsForForm (control names, in top-to-bottom order)
'    5. a cmdXxx_Click() open-handler in frmCleanupSuiteLauncher_Code
'    6. a cmdHelpXxx_Click() help-handler in frmCleanupSuiteLauncher_Code
'    7. the tool + help button names in the launcher ControlsForForm array
'    8. frmXxx name added to the Array(...) in IsSuiteComponent
'
'  CRITICAL VBA TRAP: never put a block construct (With/For/Do/Select, or a
'  second If) inside a single-line If.  If...Then...: End If  and
'  If...Then With...: End With  do NOT compile.  Use full multi-line blocks.
'  Always set OptionButton .GroupName in UserForm_Initialize -- controls are
'  created flat, so frames do not group them.  CONTRIBUTING.md has the full
'  safety chain, control-naming rules, builder-string rules, and cautions.
' ---------------------------------------------------------------------------
' Paste this entire module into a new standard module and run InstallCleanupSuite
Option Explicit

' ---------------------------
' Option 2 Installer (corrected, robust)
' ---------------------------
Public Sub InstallCleanupSuite()
    If Not EnsurePrereqs() Then Exit Sub

    Dim vbProj As VBIDE.VBProject
    Set vbProj = ThisDocument.VBProject

    Dim installLog As Collection
    Set installLog = New Collection

    ' Create/replace modules and forms with robust error handling
    On Error GoTo InstallerErr

    ' 1) Create standard modules (code-only)
    CreateOrReplaceModule vbProj, "modCleanupLauncher", modCleanupLauncher_Code, installLog
    CreateOrReplaceModule vbProj, "modCleanupHelpers", modCleanupHelpers_Code, installLog

    ' 2) Create full code UserForms (code injection)
    CreateOrReplaceForm vbProj, "frmPunctuationCleanup", frmPunctuationCleanup_Code, installLog
    CreateOrReplaceForm vbProj, "frmUnicodeCleanup", frmUnicodeCleanup_Code, installLog
    CreateOrReplaceForm vbProj, "frmSpacingCleanup", frmSpacingCleanup_Code, installLog
    CreateOrReplaceForm vbProj, "frmCapitalizationCleanup", frmCapitalizationCleanup_Code, installLog
    CreateOrReplaceForm vbProj, "frmListCleanup", frmListCleanup_Code, installLog
    CreateOrReplaceForm vbProj, "frmParagraphCleanup", frmParagraphCleanup_Code, installLog
    CreateOrReplaceForm vbProj, "frmDuplicateDetector", frmDuplicateDetector_Code, installLog
    CreateOrReplaceForm vbProj, "frmFontNormalizer", frmFontNormalizer_Code, installLog
    CreateOrReplaceForm vbProj, "frmTableCleaner", frmTableCleaner_Code, installLog
    CreateOrReplaceForm vbProj, "frmBreakNormalizer", frmBreakNormalizer_Code, installLog
    CreateOrReplaceForm vbProj, "frmDocumentTrim", frmDocumentTrim_Code, installLog
    CreateOrReplaceForm vbProj, "frmFormattingStripper", frmFormattingStripper_Code, installLog
    CreateOrReplaceForm vbProj, "frmHyperlinkRemover", frmHyperlinkRemover_Code, installLog
    CreateOrReplaceForm vbProj, "frmSoftReturnConverter", frmSoftReturnConverter_Code, installLog
    CreateOrReplaceForm vbProj, "frmMetadataScrubber", frmMetadataScrubber_Code, installLog
    CreateOrReplaceForm vbProj, "frmStyleCleanup", frmStyleCleanup_Code, installLog
    CreateOrReplaceForm vbProj, "frmFootnoteRemover", frmFootnoteRemover_Code, installLog
    CreateOrReplaceForm vbProj, "frmHeaderFooterStandardizer", frmHeaderFooterStandardizer_Code, installLog
    CreateOrReplaceForm vbProj, "frmObjectRemover", frmObjectRemover_Code, installLog
    CreateOrReplaceForm vbProj, "frmCleanupSuiteLauncher", frmCleanupSuiteLauncher_Code, installLog

    ' 3) Attempt to create controls programmatically (best-effort)
    Dim uiFailures As Collection
    Set uiFailures = New Collection

    AttemptCreateControls vbProj, "frmCleanupSuiteLauncher", uiFailures
    AttemptCreateControls vbProj, "frmPunctuationCleanup", uiFailures
    AttemptCreateControls vbProj, "frmUnicodeCleanup", uiFailures
    AttemptCreateControls vbProj, "frmSpacingCleanup", uiFailures
    AttemptCreateControls vbProj, "frmCapitalizationCleanup", uiFailures
    AttemptCreateControls vbProj, "frmListCleanup", uiFailures
    AttemptCreateControls vbProj, "frmParagraphCleanup", uiFailures
    AttemptCreateControls vbProj, "frmDuplicateDetector", uiFailures
    AttemptCreateControls vbProj, "frmFontNormalizer", uiFailures
    AttemptCreateControls vbProj, "frmTableCleaner", uiFailures
    AttemptCreateControls vbProj, "frmBreakNormalizer", uiFailures
    AttemptCreateControls vbProj, "frmDocumentTrim", uiFailures
    AttemptCreateControls vbProj, "frmFormattingStripper", uiFailures
    AttemptCreateControls vbProj, "frmHyperlinkRemover", uiFailures
    AttemptCreateControls vbProj, "frmSoftReturnConverter", uiFailures
    AttemptCreateControls vbProj, "frmMetadataScrubber", uiFailures
    AttemptCreateControls vbProj, "frmStyleCleanup", uiFailures
    AttemptCreateControls vbProj, "frmFootnoteRemover", uiFailures
    AttemptCreateControls vbProj, "frmHeaderFooterStandardizer", uiFailures
    AttemptCreateControls vbProj, "frmObjectRemover", uiFailures

    ' 4) Post-install report
    Dim reportText As String
    reportText = "Cleanup Suite Installer Report" & vbCrLf & "Generated: " & Now & vbCrLf & vbCrLf

    Dim logEntry As Variant, failureMsg As Variant
    reportText = reportText & "Code injection results:" & vbCrLf
    For Each logEntry In installLog
        reportText = reportText & "- " & logEntry & vbCrLf
    Next logEntry

    reportText = reportText & vbCrLf & "UI automation results:" & vbCrLf
    If uiFailures.Count = 0 Then
        reportText = reportText & "- All UI controls created programmatically (best-effort)." & vbCrLf
    Else
        reportText = reportText & "- Some forms could not be fully created programmatically. See manual checklist below." & vbCrLf
        For Each failureMsg In uiFailures
            reportText = reportText & "  * " & failureMsg & vbCrLf
        Next failureMsg
    End If

    reportText = reportText & vbCrLf & "Manual checklist (if any UI items failed to create):" & vbCrLf
    reportText = reportText & ManualChecklistText()

    ' Write report to temp file and open it
    Dim reportPath As String
    reportPath = Environ$("TEMP") & "\CleanupSuite_Install_Report.txt"
    Dim ff As Integer: ff = FreeFile
    Open reportPath For Output As #ff
    Print #ff, reportText
    Close #ff
    Shell "notepad.exe " & Chr(34) & reportPath & Chr(34), vbNormalFocus

    MsgBox "Installer finished. The report has been opened in Notepad." & vbCrLf & _
           "Follow the manual checklist if needed, then compile the project (VBA editor -> Debug -> Compile).", vbInformation, "Install Complete"

    Exit Sub

InstallerErr:
    MsgBox "Installer encountered an error: " & Err.Number & " - " & Err.Description, vbCritical, "Installer Error"
End Sub

Sub UninstallCleanupSuite()
    ' --------------------------------------------------------------
    '  DEVELOPER TOOL -- removes every Cleanup Suite component (all forms
    '  and the two injected modules) from this project, but leaves THIS
    '  host module untouched.  Run this BEFORE re-running InstallCleanupSuite
    '  over an existing install.
    '
    '  Why a separate tool: the VBA editor defers component removal until the
    '  running macro ends, so removing and re-adding the same form in one run
    '  spawns duplicates (frmX1) or fails.  This Sub only removes; the deletions
    '  commit when it returns.  Then run InstallCleanupSuite as a SEPARATE run.
    ' --------------------------------------------------------------
    If Not EnsurePrereqs() Then Exit Sub

    Dim vbProj As VBIDE.VBProject
    Set vbProj = ThisDocument.VBProject

    Dim removed As String, failed As String
    Dim removedCount As Long, failedCount As Long
    Dim i As Long
    Dim comp As VBIDE.VBComponent
    Dim nm As String

    For i = vbProj.VBComponents.Count To 1 Step -1
        Set comp = vbProj.VBComponents(i)
        If IsSuiteComponent(comp.Name) Then
            nm = comp.Name
            On Error Resume Next
            Err.Clear
            vbProj.VBComponents.Remove comp
            If Err.Number = 0 Then
                removed = removed & "  " & nm & vbCrLf
                removedCount = removedCount + 1
            Else
                failed = failed & "  " & nm & "  (" & Err.Description & ")" & vbCrLf
                failedCount = failedCount + 1
            End If
            On Error GoTo 0
        End If
    Next i

    Dim msg As String
    msg = "Cleanup Suite -- Uninstaller" & vbCrLf & vbCrLf
    If removedCount = 0 And failedCount = 0 Then
        msg = msg & "No Cleanup Suite components were found. Nothing to remove."
    Else
        msg = msg & "Removed " & removedCount & " component(s):" & vbCrLf & removed
        If failedCount > 0 Then
            msg = msg & vbCrLf & "Could NOT remove " & failedCount & " component(s):" & vbCrLf & failed
            msg = msg & vbCrLf & "These are usually open in the VBA editor. Close each one's"
            msg = msg & vbCrLf & "form/code window (click its X) and run the uninstaller again."
        End If
        msg = msg & vbCrLf & vbCrLf & "NEXT STEP: to reinstall, run InstallCleanupSuite as a SEPARATE"
        msg = msg & vbCrLf & "run (open it and press F5, or pick it from the macro list)."
        msg = msg & vbCrLf & "It will not re-create cleanly if called in the same run as removal."
    End If
    MsgBox msg, vbInformation, "Cleanup Suite Uninstaller"
End Sub

Sub ReinstallCleanupSuite()
    ' --------------------------------------------------------------
    '  OPTIONAL convenience: removes all components now, then auto-runs
    '  InstallCleanupSuite about 2 seconds later, in a SEPARATE run (via
    '  Application.OnTime) so removal and re-creation never collide.
    '
    '  Requires this host code to live in a STANDARD module so OnTime can
    '  resolve "InstallCleanupSuite" by name.  If the auto-step does not
    '  fire, just run InstallCleanupSuite yourself -- the removal is done.
    ' --------------------------------------------------------------
    If Not EnsurePrereqs() Then Exit Sub

    Dim vbProj As VBIDE.VBProject
    Set vbProj = ThisDocument.VBProject

    Dim i As Long, removedCount As Long
    Dim comp As VBIDE.VBComponent
    For i = vbProj.VBComponents.Count To 1 Step -1
        Set comp = vbProj.VBComponents(i)
        If IsSuiteComponent(comp.Name) Then
            On Error Resume Next
            Err.Clear
            vbProj.VBComponents.Remove comp
            If Err.Number = 0 Then removedCount = removedCount + 1
            On Error GoTo 0
        End If
    Next i

    Dim scheduled As Boolean: scheduled = True
    On Error Resume Next
    Err.Clear
    Application.OnTime Now + TimeValue("00:00:02"), "InstallCleanupSuite"
    If Err.Number <> 0 Then scheduled = False
    On Error GoTo 0

    If Not scheduled Then
        MsgBox "Removed " & removedCount & " old component(s), but could not auto-schedule" & vbCrLf & _
               "the reinstall. Please run InstallCleanupSuite manually now.", vbExclamation, "Reinstall"
    End If
    ' On success this Sub ends silently: the deferred removals commit, then
    ' the scheduled InstallCleanupSuite runs and shows its own report.
End Sub

Private Function IsSuiteComponent(compName As String) As Boolean
    ' True if compName matches a Cleanup Suite component, ignoring any trailing
    ' digits the editor may have appended to a duplicate (e.g. frmSpacingCleanup1).
    Dim baseName As String
    baseName = compName
    Do While Len(baseName) > 0 And IsNumeric(Right(baseName, 1))
        baseName = Left(baseName, Len(baseName) - 1)
    Loop

    Dim names As Variant
    names = Array( _
        "frmPunctuationCleanup", "frmUnicodeCleanup", "frmSpacingCleanup", _
        "frmCapitalizationCleanup", "frmListCleanup", "frmParagraphCleanup", _
        "frmDuplicateDetector", "frmFontNormalizer", "frmTableCleaner", _
        "frmBreakNormalizer", "frmDocumentTrim", "frmFormattingStripper", _
        "frmHyperlinkRemover", "frmSoftReturnConverter", "frmMetadataScrubber", _
        "frmStyleCleanup", "frmFootnoteRemover", "frmHeaderFooterStandardizer", _
        "frmObjectRemover", "frmCleanupSuiteLauncher", "modCleanupLauncher", _
        "modCleanupHelpers")

    Dim i As Long
    For i = LBound(names) To UBound(names)
        If baseName = names(i) Then IsSuiteComponent = True: Exit Function
    Next i
    IsSuiteComponent = False
End Function

' ---------------------------
' Prereq check
' ---------------------------
Private Function EnsurePrereqs() As Boolean
    EnsurePrereqs = False
    On Error Resume Next
    Dim testProj As VBIDE.VBProject
    Set testProj = ThisDocument.VBProject
    If Err.Number <> 0 Then
        MsgBox "This installer requires temporary access to the VBA project object model and the VBIDE Extensibility reference." & vbCrLf & _
               "Please enable: File → Options → Trust Center → Trust Center Settings → Macro Settings → check 'Trust access to the VBA project object model'." & vbCrLf & _
               "Then in the VBA editor: Tools → References → check 'Microsoft Visual Basic for Applications Extensibility 5.3'." & vbCrLf & vbCrLf & _
               "After enabling, re-run this installer.", vbExclamation, "Prerequisites Required"
        Exit Function
    End If
    EnsurePrereqs = True
End Function

' ---------------------------
' Create or replace a standard module with code
' ---------------------------
Private Sub CreateOrReplaceModule(vbProj As VBIDE.VBProject, moduleName As String, sourceCode As String, installLog As Collection)
    Dim comp As VBIDE.VBComponent
    On Error Resume Next
    Set comp = vbProj.VBComponents(moduleName)
    If Not comp Is Nothing Then vbProj.VBComponents.Remove comp
    On Error GoTo 0

    Set comp = vbProj.VBComponents.Add(vbext_ct_StdModule)
    comp.Name = moduleName
    comp.CodeModule.AddFromString sourceCode
    installLog.Add moduleName & " module created/updated."
End Sub

' ---------------------------
' Create or replace a UserForm and inject code
' ---------------------------
Private Sub CreateOrReplaceForm(vbProj As VBIDE.VBProject, formName As String, sourceCode As String, installLog As Collection)
    Dim comp As VBIDE.VBComponent
    On Error Resume Next
    Set comp = vbProj.VBComponents(formName)
    If Not comp Is Nothing Then vbProj.VBComponents.Remove comp
    On Error GoTo 0

    Set comp = vbProj.VBComponents.Add(vbext_ct_MSForm)
    comp.Name = formName
    comp.Properties("Caption") = formName
    comp.CodeModule.AddFromString sourceCode
    installLog.Add formName & " created with code injected."
End Sub

' ---------------------------
' Attempt to create controls programmatically (best-effort)
' Adds common controls used by our forms. If Designer is not available or blocked, records failure.
' ---------------------------
Private Sub AttemptCreateControls(vbProj As VBIDE.VBProject, formName As String, failures As Collection)
    On Error GoTo FailHandler
    Dim comp As VBIDE.VBComponent
    Set comp = vbProj.VBComponents(formName)
    If comp Is Nothing Then
        failures.Add formName & ": form not found"
        Exit Sub
    End If

    Dim designer As Object
    Set designer = comp.Designer
    If designer Is Nothing Then
        failures.Add formName & ": Designer not available (likely blocked by security settings)"
        Exit Sub
    End If

    ' Best-effort: add a few controls only if they don't exist.
    ' We add minimal visible controls so the form is usable; layout adjustments are recommended in the designer.
    Dim ctrlNames As Variant
    ctrlNames = ControlsForForm(formName)

    Dim ctrlName As Variant
    For Each ctrlName In ctrlNames
        On Error Resume Next
        If designer.Controls(ctrlName) Is Nothing Then
            ' Choose MSForms ProgID by control name prefix
            Dim ctrlProgId As String
            ctrlProgId = ControlTypeByName(CStr(ctrlName))
            designer.Controls.Add ctrlProgId, CStr(ctrlName), True
            ' Make visible and set a default caption for command buttons
            On Error Resume Next
            designer.Controls(ctrlName).Visible = True
            If Left$(CStr(ctrlName), 3) = "cmd" Then
                designer.Controls(ctrlName).Caption = Replace(CStr(ctrlName), "cmd", "")
            End If
        End If
        On Error GoTo 0
    Next ctrlName

    LayoutFormControls comp

    Exit Sub

FailHandler:
    failures.Add formName & ": failed to create controls (" & Err.Number & " - " & Err.Description & ")"
    Err.Clear
End Sub

Private Sub LayoutFormControls(comp As VBIDE.VBComponent)
    ' Best-effort automatic layout so installed forms open usable rather than
    ' with every control stacked at the top-left.  Controls are placed in a
    ' single vertical column in their declared order: option buttons and
    ' check boxes indented, command buttons paired two per row, labels
    ' word-wrapped, and the form is sized to fit.  Any per-control failure is
    ' ignored so one odd control cannot abort the whole pass.
    On Error Resume Next
    Dim designer As Object
    Set designer = comp.Designer
    If designer Is Nothing Then Exit Sub
    Const FORM_W As Single = 372
    Const MARGIN As Single = 12
    Const GAP As Single = 4
    Const BTN_H As Single = 22
    Dim contentW As Single: contentW = FORM_W - 2 * MARGIN
    Dim halfW As Single: halfW = (contentW - 8) / 2
    Dim y As Single: y = 6
    Dim pend As String: pend = ""
    Dim ctrlNames As Variant: ctrlNames = ControlsForForm(comp.Name)
    Dim cn As Variant, pfx As String
    Dim h As Single, lft As Single, w As Single
    For Each cn In ctrlNames
        pfx = Left$(CStr(cn), 3)
        If pfx = "cmd" Then
            SetButtonCaption designer, CStr(cn)
            If Len(pend) > 0 Then
                PositionControl designer, pend, MARGIN, y, halfW, BTN_H
                PositionControl designer, CStr(cn), MARGIN + halfW + 8, y, halfW, BTN_H
                y = y + BTN_H + GAP
                pend = ""
            Else
                pend = CStr(cn)
            End If
        Else
            If Len(pend) > 0 Then
                PositionControl designer, pend, MARGIN, y, contentW, BTN_H
                y = y + BTN_H + GAP
                pend = ""
            End If
            Select Case pfx
                Case "lbl"
                    h = 28: lft = MARGIN: w = contentW
                    designer.Controls(CStr(cn)).WordWrap = True
                    designer.Controls(CStr(cn)).AutoSize = False
                Case "opt", "chk"
                    h = 15: lft = MARGIN + 10: w = contentW - 10
                Case "fra"
                    h = 20: lft = MARGIN: w = contentW
                Case Else
                    h = 18: lft = MARGIN: w = contentW
            End Select
            PositionControl designer, CStr(cn), lft, y, w, h
            y = y + h + GAP
        End If
    Next cn
    If Len(pend) > 0 Then
        PositionControl designer, pend, MARGIN, y, contentW, BTN_H
        y = y + BTN_H + GAP
    End If
    comp.Properties("Width") = FORM_W + 8
    comp.Properties("Height") = y + 40
    designer.Width = FORM_W + 8
    designer.Height = y + 40
End Sub

Private Sub PositionControl(designer As Object, nm As String, L As Single, T As Single, W As Single, H As Single)
    On Error Resume Next
    With designer.Controls(nm)
        .Left = L
        .Top = T
        .Width = W
        .Height = H
        .Visible = True
    End With
End Sub

Private Sub SetButtonCaption(designer As Object, nm As String)
    On Error Resume Next
    Dim cap As String
    Select Case nm
        Case "cmdRun": cap = "Run"
        Case "cmdSelectAll": cap = "Select All"
        Case "cmdDeselectAll": cap = "Deselect All"
        Case "cmdHelp": cap = "Help"
        Case Else: cap = Replace(nm, "cmd", "")
    End Select
    designer.Controls(nm).Caption = cap
End Sub


' ---------------------------
' Helper: return control names expected for each form
' ---------------------------
Private Function ControlsForForm(formName As String) As Variant
    Select Case formName
        Case "frmCleanupSuiteLauncher"
            ControlsForForm = Array("cmdUnicode", "cmdPunctuation", "cmdSpacing", "cmdCapitalization", "cmdList", "cmdParagraph", _
                                    "cmdDuplicate", "cmdFontNorm", "cmdTableClean", "cmdBreakNorm", "cmdDocTrim", "cmdFormatStrip", "cmdHyperlink", "cmdSoftReturn", "cmdMetadata", "cmdStyleClean", "cmdFootnote", "cmdHeaderFooter", "cmdObjectRemover", "chkAutoSave", _
                                    "cmdHelpUnicode", "cmdHelpPunct", "cmdHelpSpacing", "cmdHelpCap", "cmdHelpList", "cmdHelpPara", _
                                    "cmdHelpDuplicate", "cmdHelpFont", "cmdHelpTable", "cmdHelpBreak", "cmdHelpTrim", "cmdHelpFormat", "cmdHelpHyperlink", "cmdHelpSoftReturn", "cmdHelpMetadata", "cmdHelpStyle", "cmdHelpFootnote", "cmdHelpHeaderFooter", "cmdHelpObject")
        Case "frmPunctuationCleanup"
            ControlsForForm = Array("optAll", "optQuotes", "optDashes", "optEllipses", "optCustom", "fraCustom", _
                                    "chkCurlyDouble", "chkCurlySingle", "chkEmDash", "chkEnDash", "chkEllipses", _
                                    "chkPreviewOnly", "cmdSelectAll", "cmdDeselectAll", "cmdRun", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdHelp")
        Case "frmUnicodeCleanup"
            ControlsForForm = Array("optAll", "optNBSP", "optZeroWidth", "optCustom", "fraCustom", _
                                    "chkNBSP", "chkZWSP", "chkZWNJ", "chkZWJ", "chkBOM", "chkSoftHyphen", "chkNBHyphen", _
                                    "chkPreviewOnly", "cmdSelectAll", "cmdDeselectAll", "cmdRun", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdHelp")
        Case "frmSpacingCleanup"
            ControlsForForm = Array("optAll", "optDoubleSpaces", "optTrim", "optCustom", "fraCustom", _
                                    "chkDoubleSpaces", "chkTrimSpaces", "chkSpaceBeforePunct", "chkNormalizeAfterPunct", "chkExtraBlankLines", _
                                    "chkPreviewOnly", "cmdSelectAll", "cmdDeselectAll", "cmdRun", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdHelp")
        Case "frmCapitalizationCleanup"
            ControlsForForm = Array("optAll", "optSentence", "optTitle", "optUpper", "optLower", "optCustom", "fraCustom", _
                                    "chkSentence", "chkTitle", "chkUpper", "chkLower", "chkSmartSentences", _
                                    "chkPreviewOnly", "cmdSelectAll", "cmdDeselectAll", "cmdRun", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdHelp")
        Case "frmListCleanup"
            ControlsForForm = Array("optAll", "optBullets", "optNumbering", "optIndent", "optCustom", "fraCustom", _
                                    "chkNormalizeBullets", "chkNormalizeNumbering", "chkFixIndent", "chkHyphenToBullets", _
                                    "chkPreviewOnly", "cmdSelectAll", "cmdDeselectAll", "cmdRun", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdHelp")
        Case "frmParagraphCleanup"
            ControlsForForm = Array("optAll", "optRemoveEmpty", "optNormalizeSpacing", "optCustom", "fraCustom", _
                                    "chkRemoveEmpty", "chkCollapseBreaks", "chkNormalizeParaSpacing", "chkFixIndent", _
                                    "chkPreviewOnly", "cmdSelectAll", "cmdDeselectAll", "cmdRun", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdHelp")
        Case "frmDuplicateDetector"
            ControlsForForm = Array("optHighlightOnly", "optRemoveDupes", "fraMatching", _
                                    "optMatchExact", "optMatchNormalized", "optMatchFuzzy", _
                                    "fraThreshold", "optFuzzyLoose", "optFuzzyMedium", "optFuzzyStrict", "lblFuzzyWarning", _
                                    "chkPreviewOnly", "cmdRun", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdHelp")
        Case "frmFontNormalizer"
            ControlsForForm = Array("chkFontFace", "chkFontSize", "chkBold", "chkItalic", "chkFontColor", _
                                    "chkPreviewOnly", "cmdRun", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdHelp")
        Case "frmTableCleaner"
            ControlsForForm = Array("chkRemoveEmptyRows", "chkRemoveEmptyCols", "chkNormalizePadding", _
                                    "chkStripDirectFormat", "chkNormalizeBorders", "chkRemoveBorders", "chkConvertToText", _
                                    "chkPreviewOnly", "cmdRun", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdHelp")
        Case "frmBreakNormalizer"
            ControlsForForm = Array("chkCollapseSectionBreaks", "chkCollapsePageBreaks", _
                                    "chkConvertSectionBreaks", "fraConvertTo", _
                                    "optConvertNextPage", "optConvertContinuous", "optConvertEvenPage", "optConvertOddPage", _
                                    "chkPreviewOnly", "cmdRun", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdHelp")
        Case "frmDocumentTrim"
            ControlsForForm = Array("chkPreviewOnly", "cmdRun", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdHelp")
        Case "frmFormattingStripper"
            ControlsForForm = Array("chkResetChar", "chkResetPara", "fraEmphasis", _
                                    "optEmphQuick", "optEmphThorough", "optEmphStrip", "lblSpeedWarning", _
                                    "chkPreserveHighlight", "chkPreserveDropCaps", "chkPreviewOnly", "cmdRun", "cmdHelp", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmHyperlinkRemover"
            ControlsForForm = Array("chkRemoveFormat", "chkPreviewOnly", "cmdRun", "cmdHelp", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmSoftReturnConverter"
            ControlsForForm = Array("optSoftToPara", "optParaToSoft", "chkPreviewOnly", "cmdRun", "cmdHelp", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmMetadataScrubber"
            ControlsForForm = Array("chkProperties", "chkPersonalInfo", "chkComments", "chkRevisions", _
                                    "chkPreviewOnly", "cmdRun", "cmdHelp", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmStyleCleanup"
            ControlsForForm = Array("chkRemoveUnused", "chkRemapVariants", "chkPreviewOnly", "cmdRun", "cmdHelp", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmFootnoteRemover"
            ControlsForForm = Array("chkFootnotes", "chkEndnotes", "chkKeepTextInline", _
                                    "chkPreviewOnly", "cmdRun", "cmdHelp", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmHeaderFooterStandardizer"
            ControlsForForm = Array("optStandardize", "optClearAll", "fraOptions", _
                                    "chkHeaders", "chkFooters", "chkFont", "chkSpacing", "chkAlignment", _
                                    "fraAlign", "optAlignLeft", "optAlignCenter", "optAlignRight", "chkBreakLinks", _
                                    "chkPreviewOnly", "cmdRun", "cmdHelp", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmObjectRemover"
            ControlsForForm = Array("chkPictures", "chkTextBoxes", "chkFrames", "chkHorizontalLines", _
                                    "chkHtmlControls", "chkHiddenText", "chkTables", _
                                    "chkPreviewOnly", "cmdRun", "cmdHelp", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case Else
            ControlsForForm = Array()
    End Select
End Function

' ---------------------------
' Helper: choose MSForms control ProgID by name prefix
' ---------------------------
Private Function ControlTypeByName(ctrlName As String) As String
    If Left$(ctrlName, 3) = "opt" Then
        ControlTypeByName = "Forms.OptionButton.1"
    ElseIf Left$(ctrlName, 3) = "chk" Then
        ControlTypeByName = "Forms.CheckBox.1"
    ElseIf Left$(ctrlName, 3) = "cmd" Then
        ControlTypeByName = "Forms.CommandButton.1"
    ElseIf Left$(ctrlName, 3) = "fra" Then
        ControlTypeByName = "Forms.Frame.1"
    ElseIf Left$(ctrlName, 3) = "lbl" Then
        ControlTypeByName = "Forms.Label.1"
    Else
        ControlTypeByName = "Forms.TextBox.1"
    End If
End Function