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
               "Please enable: File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> check 'Trust access to the VBA project object model'." & vbCrLf & _
               "Then in the VBA editor: Tools -> References -> check 'Microsoft Visual Basic for Applications Extensibility 5.3'." & vbCrLf & vbCrLf & _
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

    ' Best-effort: add controls only if they do not exist, then apply a
    ' consistent generated layout so the installed forms are usable without
    ' hand-adjusting them in the designer.
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
            On Error Resume Next
            designer.Controls(ctrlName).Visible = True
        End If
        ApplyControlStyle designer, formName, CStr(ctrlName)
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
    ' with every control stacked at the top-left.  The launcher gets a compact
    ' two-column tool grid; individual tool forms use a readable vertical flow.
    ' Any per-control failure is ignored so one odd control cannot abort the pass.
    On Error Resume Next
    Dim designer As Object
    Set designer = comp.Designer
    If designer Is Nothing Then Exit Sub

    ApplyFormStyle comp, designer
    If comp.Name = "frmCleanupSuiteLauncher" Then
        LayoutLauncherControls comp, designer
        Exit Sub
    End If

    Const FORM_W As Single = 372
    Const MARGIN As Single = 12
    Const GAP As Single = 5
    Const BTN_H As Single = 24
    Dim contentW As Single: contentW = FORM_W - 2 * MARGIN
    Dim halfW As Single: halfW = (contentW - 8) / 2
    Dim y As Single: y = 10
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
                    h = 34: lft = MARGIN: w = contentW
                    designer.Controls(CStr(cn)).WordWrap = True
                    designer.Controls(CStr(cn)).AutoSize = False
                Case "opt", "chk"
                    h = 18: lft = MARGIN + 8: w = contentW - 8
                Case "fra"
                    h = 22: lft = MARGIN: w = contentW
                Case Else
                    h = 20: lft = MARGIN: w = contentW
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

Private Sub LayoutLauncherControls(comp As VBIDE.VBComponent, designer As Object)
    On Error Resume Next
    Const FORM_W As Single = 840
    Const MARGIN As Single = 14
    Const GAP As Single = 6
    Const COL_GAP As Single = 14
    Dim contentW As Single: contentW = FORM_W - (2 * MARGIN)
    Const COL_W As Single = (FORM_W - (2 * MARGIN) - COL_GAP) / 2
    Const LEFT_X As Single = MARGIN
    Const RIGHT_X As Single = MARGIN + COL_W + COL_GAP
    Dim y As Single: y = 12
    Dim yLeft As Single
    Dim yRight As Single

    PositionControl designer, "lblLauncherTitle", MARGIN, y, contentW, 20
    y = y + 22
    PositionControl designer, "lblLauncherSubtitle", MARGIN, y, contentW, 28
    y = y + 34

    yLeft = y
    yLeft = LayoutLauncherCategory(designer, "lblCatText", LEFT_X, yLeft, COL_W)
    yLeft = LayoutLauncherToolRow(designer, "cmdHelpUnicode", "cmdUnicode", "lblDescUnicode", LEFT_X, yLeft, COL_W)
    yLeft = LayoutLauncherToolRow(designer, "cmdHelpPunct", "cmdPunctuation", "lblDescPunctuation", LEFT_X, yLeft, COL_W)
    yLeft = LayoutLauncherToolRow(designer, "cmdHelpSpacing", "cmdSpacing", "lblDescSpacing", LEFT_X, yLeft, COL_W)
    yLeft = LayoutLauncherToolRow(designer, "cmdHelpCap", "cmdCapitalization", "lblDescCapitalization", LEFT_X, yLeft, COL_W)
    yLeft = yLeft + GAP

    yLeft = LayoutLauncherCategory(designer, "lblCatPara", LEFT_X, yLeft, COL_W)
    yLeft = LayoutLauncherToolRow(designer, "cmdHelpList", "cmdList", "lblDescList", LEFT_X, yLeft, COL_W)
    yLeft = LayoutLauncherToolRow(designer, "cmdHelpPara", "cmdParagraph", "lblDescParagraph", LEFT_X, yLeft, COL_W)
    yLeft = LayoutLauncherToolRow(designer, "cmdHelpSoftReturn", "cmdSoftReturn", "lblDescSoftReturn", LEFT_X, yLeft, COL_W)
    yLeft = LayoutLauncherToolRow(designer, "cmdHelpDuplicate", "cmdDuplicate", "lblDescDuplicate", LEFT_X, yLeft, COL_W)
    yLeft = yLeft + GAP

    yRight = y
    yRight = LayoutLauncherCategory(designer, "lblCatLayout", RIGHT_X, yRight, COL_W)
    yRight = LayoutLauncherToolRow(designer, "cmdHelpTable", "cmdTableClean", "lblDescTableClean", RIGHT_X, yRight, COL_W)
    yRight = LayoutLauncherToolRow(designer, "cmdHelpBreak", "cmdBreakNorm", "lblDescBreakNorm", RIGHT_X, yRight, COL_W)
    yRight = LayoutLauncherToolRow(designer, "cmdHelpHeaderFooter", "cmdHeaderFooter", "lblDescHeaderFooter", RIGHT_X, yRight, COL_W)
    yRight = LayoutLauncherToolRow(designer, "cmdHelpTrim", "cmdDocTrim", "lblDescDocTrim", RIGHT_X, yRight, COL_W)
    yRight = yRight + GAP

    yRight = LayoutLauncherCategory(designer, "lblCatFormat", RIGHT_X, yRight, COL_W)
    yRight = LayoutLauncherToolRow(designer, "cmdHelpFont", "cmdFontNorm", "lblDescFontNorm", RIGHT_X, yRight, COL_W)
    yRight = LayoutLauncherToolRow(designer, "cmdHelpFormat", "cmdFormatStrip", "lblDescFormatStrip", RIGHT_X, yRight, COL_W)
    yRight = LayoutLauncherToolRow(designer, "cmdHelpStyle", "cmdStyleClean", "lblDescStyleClean", RIGHT_X, yRight, COL_W)
    yRight = LayoutLauncherToolRow(designer, "cmdHelpHyperlink", "cmdHyperlink", "lblDescHyperlink", RIGHT_X, yRight, COL_W)
    yRight = yRight + GAP

    yRight = LayoutLauncherCategory(designer, "lblCatReview", RIGHT_X, yRight, COL_W)
    yRight = LayoutLauncherToolRow(designer, "cmdHelpMetadata", "cmdMetadata", "lblDescMetadata", RIGHT_X, yRight, COL_W)
    yRight = LayoutLauncherToolRow(designer, "cmdHelpFootnote", "cmdFootnote", "lblDescFootnote", RIGHT_X, yRight, COL_W)
    yRight = LayoutLauncherToolRow(designer, "cmdHelpObject", "cmdObjectRemover", "lblDescObjectRemover", RIGHT_X, yRight, COL_W)
    y = MaxSingle(yLeft, yRight) + 8

    PositionControl designer, "chkAutoSave", MARGIN + 2, y, FORM_W - (2 * MARGIN), 18

    comp.Properties("Width") = FORM_W + 8
    comp.Properties("Height") = y + 58
    designer.Width = FORM_W + 8
    designer.Height = y + 58
End Sub

Private Function LayoutLauncherCategory(designer As Object, labelName As String, L As Single, T As Single, W As Single) As Single
    On Error Resume Next
    PositionControl designer, labelName, L, T, W, 18
    LayoutLauncherCategory = T + 20
End Function

Private Function LayoutLauncherToolRow(designer As Object, helpName As String, buttonName As String, descName As String, L As Single, T As Single, W As Single) As Single
    On Error Resume Next
    Const ROW_H As Single = 32
    Const BTN_W As Single = 132
    Const HELP_W As Single = 24
    Const INNER_GAP As Single = 8
    Dim descW As Single
    descW = W - BTN_W - HELP_W - (2 * INNER_GAP)

    PositionControl designer, helpName, L, T + 4, HELP_W, 24
    PositionControl designer, buttonName, L + HELP_W + INNER_GAP, T, BTN_W, ROW_H
    PositionControl designer, descName, L + HELP_W + BTN_W + (2 * INNER_GAP), T + 1, descW, ROW_H
    LayoutLauncherToolRow = T + ROW_H + 4
End Function

Private Function MaxSingle(a As Single, b As Single) As Single
    If a >= b Then
        MaxSingle = a
    Else
        MaxSingle = b
    End If
End Function

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

Private Sub ApplyFormStyle(comp As VBIDE.VBComponent, designer As Object)
    On Error Resume Next
    comp.Properties("Caption") = FriendlyFormCaption(comp.Name)
    designer.Caption = FriendlyFormCaption(comp.Name)
    designer.BackColor = RGB(248, 249, 251)
    designer.SpecialEffect = 0
End Sub

Private Sub ApplyControlStyle(designer As Object, formName As String, nm As String)
    On Error Resume Next
    Dim ctl As Object
    Set ctl = designer.Controls(nm)
    If ctl Is Nothing Then Exit Sub

    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 9
    ctl.ForeColor = RGB(32, 37, 45)

    Select Case Left$(nm, 3)
        Case "cmd"
            SetButtonCaption designer, nm
            ctl.Font.Bold = False
            ctl.BackColor = RGB(255, 255, 255)
            ctl.ControlTipText = ButtonTipText(nm)
            If Left$(nm, 7) = "cmdHelp" And nm <> "cmdHelp" Then
                ctl.Caption = "?"
                ctl.ControlTipText = "Show help for this tool"
                ctl.Font.Bold = True
            End If
            If nm = "cmdRun" Then
                ctl.Font.Bold = True
                ctl.ControlTipText = "Run the selected cleanup options"
            End If
        Case "fra"
            ctl.Caption = FrameCaption(nm)
            ctl.Font.Bold = True
            ctl.BackColor = RGB(248, 249, 251)
            ctl.ForeColor = RGB(64, 70, 80)
        Case "lbl"
            ctl.BackColor = RGB(248, 249, 251)
            ctl.ForeColor = RGB(84, 91, 104)
            If Len(ControlCaptionText(formName, nm)) > 0 Then ctl.Caption = ControlCaptionText(formName, nm)
            If formName = "frmCleanupSuiteLauncher" Then
                ctl.WordWrap = True
                ctl.AutoSize = False
                If nm = "lblLauncherTitle" Then
                    ctl.Font.Size = 14
                    ctl.Font.Bold = True
                    ctl.ForeColor = RGB(28, 43, 58)
                ElseIf nm = "lblLauncherSubtitle" Then
                    ctl.Font.Size = 8.5
                    ctl.ForeColor = RGB(88, 101, 116)
                ElseIf Left$(nm, 6) = "lblCat" Then
                    ctl.Font.Size = 9
                    ctl.Font.Bold = True
                    ctl.ForeColor = RGB(38, 91, 137)
                ElseIf Left$(nm, 7) = "lblDesc" Then
                    ctl.Font.Size = 8
                    ctl.ForeColor = RGB(93, 105, 119)
                End If
            End If
        Case "chk", "opt"
            ctl.BackColor = RGB(248, 249, 251)
            If Len(ControlCaptionText(formName, nm)) > 0 Then ctl.Caption = ControlCaptionText(formName, nm)
    End Select
End Sub

Private Function ControlCaptionText(formName As String, nm As String) As String
    Select Case formName & "." & nm
        Case "frmBreakNormalizer.chkCollapsePageBreaks": ControlCaptionText = "Collapse consecutive page breaks"
        Case "frmBreakNormalizer.chkCollapseSectionBreaks": ControlCaptionText = "Collapse consecutive section breaks"
        Case "frmBreakNormalizer.chkConvertSectionBreaks": ControlCaptionText = "Convert section breaks to page breaks"
        Case "frmBreakNormalizer.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmBreakNormalizer.optConvertContinuous": ControlCaptionText = "Continuous"
        Case "frmBreakNormalizer.optConvertEvenPage": ControlCaptionText = "Even Page"
        Case "frmBreakNormalizer.optConvertNextPage": ControlCaptionText = "Next Page"
        Case "frmBreakNormalizer.optConvertOddPage": ControlCaptionText = "Odd Page"
        Case "frmBreakNormalizer.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmBreakNormalizer.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmCapitalizationCleanup.chkLower": ControlCaptionText = "Fix all-lowercase paragraphs"
        Case "frmCapitalizationCleanup.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmCapitalizationCleanup.chkSentence": ControlCaptionText = "Capitalize sentence starts"
        Case "frmCapitalizationCleanup.chkSmartSentences": ControlCaptionText = "Smart sentence detection (skip abbreviations)"
        Case "frmCapitalizationCleanup.chkTitle": ControlCaptionText = "Apply title case to headings"
        Case "frmCapitalizationCleanup.chkUpper": ControlCaptionText = "Convert ALL-CAPS words (5+ letters)"
        Case "frmCapitalizationCleanup.optAll": ControlCaptionText = "All capitalization fixes"
        Case "frmCapitalizationCleanup.optCustom": ControlCaptionText = "Custom (choose below)"
        Case "frmCapitalizationCleanup.optLower": ControlCaptionText = "lowercase -> Sentence case"
        Case "frmCapitalizationCleanup.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmCapitalizationCleanup.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmCapitalizationCleanup.optSentence": ControlCaptionText = "Sentence case"
        Case "frmCapitalizationCleanup.optTitle": ControlCaptionText = "Title case"
        Case "frmCapitalizationCleanup.optUpper": ControlCaptionText = "ALL CAPS -> Title Case"
        Case "frmCleanupSuiteLauncher.chkAutoSave": ControlCaptionText = "Auto-save before running each tool"
        Case "frmCleanupSuiteLauncher.lblLauncherTitle": ControlCaptionText = "Cleanup Suite"
        Case "frmCleanupSuiteLauncher.lblLauncherSubtitle": ControlCaptionText = "Choose the kind of cleanup you need. Tools are grouped by what you are trying to fix, not by how the VBA code is built."
        Case "frmCleanupSuiteLauncher.lblCatText": ControlCaptionText = "Text and Characters"
        Case "frmCleanupSuiteLauncher.lblCatPara": ControlCaptionText = "Paragraphs, Breaks, and Lists"
        Case "frmCleanupSuiteLauncher.lblCatLayout": ControlCaptionText = "Layout and Document Structure"
        Case "frmCleanupSuiteLauncher.lblCatFormat": ControlCaptionText = "Formatting, Links, and Styles"
        Case "frmCleanupSuiteLauncher.lblCatReview": ControlCaptionText = "Review, Privacy, and Removals"
        Case "frmCleanupSuiteLauncher.lblDescUnicode": ControlCaptionText = "Remove invisible Unicode characters that break search, sorting, and export."
        Case "frmCleanupSuiteLauncher.lblDescPunctuation": ControlCaptionText = "Normalize smart quotes, dashes, ellipses, and plain-text punctuation."
        Case "frmCleanupSuiteLauncher.lblDescSpacing": ControlCaptionText = "Fix extra spaces, punctuation spacing, and repeated blank lines."
        Case "frmCleanupSuiteLauncher.lblDescCapitalization": ControlCaptionText = "Apply sentence, title, upper, lower, or smart capitalization cleanup."
        Case "frmCleanupSuiteLauncher.lblDescList": ControlCaptionText = "Normalize manual bullets, numbering, and list indentation."
        Case "frmCleanupSuiteLauncher.lblDescParagraph": ControlCaptionText = "Clean empty paragraphs, paragraph spacing, and manual indents."
        Case "frmCleanupSuiteLauncher.lblDescSoftReturn": ControlCaptionText = "Convert Shift+Enter line breaks to paragraph marks, or reverse that conversion."
        Case "frmCleanupSuiteLauncher.lblDescDuplicate": ControlCaptionText = "Find repeated or near-duplicate paragraphs before deciding what to remove."
        Case "frmCleanupSuiteLauncher.lblDescTableClean": ControlCaptionText = "Clean table structure, padding, borders, direct formatting, or convert simple tables."
        Case "frmCleanupSuiteLauncher.lblDescBreakNorm": ControlCaptionText = "Collapse or convert section and page breaks from assembled documents."
        Case "frmCleanupSuiteLauncher.lblDescHeaderFooter": ControlCaptionText = "Standardize or clear headers and footers across document sections."
        Case "frmCleanupSuiteLauncher.lblDescDocTrim": ControlCaptionText = "Remove trailing empty paragraphs at the end of the document."
        Case "frmCleanupSuiteLauncher.lblDescFontNorm": ControlCaptionText = "Reset direct font face, size, color, bold, and italic overrides."
        Case "frmCleanupSuiteLauncher.lblDescFormatStrip": ControlCaptionText = "Strip direct character and paragraph formatting while preserving styles."
        Case "frmCleanupSuiteLauncher.lblDescStyleClean": ControlCaptionText = "Remove unused custom styles or remap common style variants."
        Case "frmCleanupSuiteLauncher.lblDescHyperlink": ControlCaptionText = "Remove links while keeping visible text, with optional hyperlink styling cleanup."
        Case "frmCleanupSuiteLauncher.lblDescMetadata": ControlCaptionText = "Scrub document properties, personal info, comments, and tracked changes."
        Case "frmCleanupSuiteLauncher.lblDescFootnote": ControlCaptionText = "Remove footnotes or endnotes, optionally keeping note text inline."
        Case "frmCleanupSuiteLauncher.lblDescObjectRemover": ControlCaptionText = "Remove embedded clutter such as images, text boxes, controls, hidden text, or tables."
        Case "frmDocumentTrim.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmDocumentTrim.optScopeDocument": ControlCaptionText = "Entire document (always)"
        Case "frmDuplicateDetector.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmDuplicateDetector.lblFuzzyWarning": ControlCaptionText = "Fuzzy match: groups paragraphs sharing the chosen percentage of words.  Slower on large documents."
        Case "frmDuplicateDetector.optFuzzyLoose": ControlCaptionText = "Loose  (50% word overlap)"
        Case "frmDuplicateDetector.optFuzzyMedium": ControlCaptionText = "Medium  (70% word overlap)"
        Case "frmDuplicateDetector.optFuzzyStrict": ControlCaptionText = "Strict  (90% word overlap)"
        Case "frmDuplicateDetector.optHighlightOnly": ControlCaptionText = "Highlight duplicates (preview only)"
        Case "frmDuplicateDetector.optMatchExact": ControlCaptionText = "Exact match"
        Case "frmDuplicateDetector.optMatchFuzzy": ControlCaptionText = "Fuzzy match (word overlap)"
        Case "frmDuplicateDetector.optMatchNormalized": ControlCaptionText = "Normalized match (ignore punctuation and spaces)"
        Case "frmDuplicateDetector.optRemoveDupes": ControlCaptionText = "Remove duplicate paragraphs"
        Case "frmDuplicateDetector.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmDuplicateDetector.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmFontNormalizer.chkBold": ControlCaptionText = "Normalize direct bold overrides"
        Case "frmFontNormalizer.chkFontColor": ControlCaptionText = "Remove direct font color overrides"
        Case "frmFontNormalizer.chkFontFace": ControlCaptionText = "Normalize font face to document default"
        Case "frmFontNormalizer.chkFontSize": ControlCaptionText = "Normalize font size to style-defined size"
        Case "frmFontNormalizer.chkItalic": ControlCaptionText = "Normalize direct italic overrides"
        Case "frmFontNormalizer.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmFontNormalizer.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmFontNormalizer.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmFootnoteRemover.chkEndnotes": ControlCaptionText = "Remove endnotes"
        Case "frmFootnoteRemover.chkFootnotes": ControlCaptionText = "Remove footnotes"
        Case "frmFootnoteRemover.chkKeepTextInline": ControlCaptionText = "Keep note text inline [in brackets] instead of deleting it"
        Case "frmFootnoteRemover.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmFootnoteRemover.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmFootnoteRemover.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmFormattingStripper.chkPreserveDropCaps": ControlCaptionText = "Preserve drop caps"
        Case "frmFormattingStripper.chkPreserveHighlight": ControlCaptionText = "Preserve highlighting"
        Case "frmFormattingStripper.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmFormattingStripper.chkResetChar": ControlCaptionText = "Strip direct character formatting"
        Case "frmFormattingStripper.chkResetPara": ControlCaptionText = "Strip direct paragraph formatting"
        Case "frmFormattingStripper.lblSpeedWarning": ControlCaptionText = "Strip mode removes all direct formatting, including bold and italic emphasis."
        Case "frmFormattingStripper.optEmphQuick": ControlCaptionText = "Quick  (preserve paragraph-level emphasis)"
        Case "frmFormattingStripper.optEmphStrip": ControlCaptionText = "Strip  (remove all direct formatting)"
        Case "frmFormattingStripper.optEmphThorough": ControlCaptionText = "Thorough  (preserve word-level emphasis)"
        Case "frmFormattingStripper.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmFormattingStripper.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmHeaderFooterStandardizer.chkAlignment": ControlCaptionText = "Set alignment"
        Case "frmHeaderFooterStandardizer.chkBreakLinks": ControlCaptionText = "Unlink sections (set each independently)"
        Case "frmHeaderFooterStandardizer.chkFont": ControlCaptionText = "Reset font to the document default"
        Case "frmHeaderFooterStandardizer.chkFooters": ControlCaptionText = "Include footers"
        Case "frmHeaderFooterStandardizer.chkHeaders": ControlCaptionText = "Include headers"
        Case "frmHeaderFooterStandardizer.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmHeaderFooterStandardizer.chkSpacing": ControlCaptionText = "Remove paragraph spacing"
        Case "frmHeaderFooterStandardizer.optAlignCenter": ControlCaptionText = "Center"
        Case "frmHeaderFooterStandardizer.optAlignLeft": ControlCaptionText = "Left"
        Case "frmHeaderFooterStandardizer.optAlignRight": ControlCaptionText = "Right"
        Case "frmHeaderFooterStandardizer.optClearAll": ControlCaptionText = "Clear all header/footer content"
        Case "frmHeaderFooterStandardizer.optScopeDocument": ControlCaptionText = "Entire document (always)"
        Case "frmHeaderFooterStandardizer.optStandardize": ControlCaptionText = "Standardize formatting"
        Case "frmHyperlinkRemover.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmHyperlinkRemover.chkRemoveFormat": ControlCaptionText = "Remove hyperlink character style (blue underline)"
        Case "frmHyperlinkRemover.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmHyperlinkRemover.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmListCleanup.chkFixIndent": ControlCaptionText = "Fix list indentation"
        Case "frmListCleanup.chkHyphenToBullets": ControlCaptionText = "Convert hyphen-dash lines to bullets"
        Case "frmListCleanup.chkNormalizeBullets": ControlCaptionText = "Convert manual hyphens / asterisks to List Bullet style"
        Case "frmListCleanup.chkNormalizeNumbering": ControlCaptionText = "Convert manual numbers to List Number style"
        Case "frmListCleanup.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmListCleanup.optAll": ControlCaptionText = "All list fixes"
        Case "frmListCleanup.optBullets": ControlCaptionText = "Convert manual bullet lists"
        Case "frmListCleanup.optCustom": ControlCaptionText = "Custom (choose below)"
        Case "frmListCleanup.optIndent": ControlCaptionText = "Fix list indentation"
        Case "frmListCleanup.optNumbering": ControlCaptionText = "Convert manual numbered lists"
        Case "frmListCleanup.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmListCleanup.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmMetadataScrubber.chkComments": ControlCaptionText = "Remove all comments"
        Case "frmMetadataScrubber.chkPersonalInfo": ControlCaptionText = "Clear personal information (Last Saved By, revision authors)"
        Case "frmMetadataScrubber.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmMetadataScrubber.chkProperties": ControlCaptionText = "Clear document properties (Title, Author, Company, etc.)"
        Case "frmMetadataScrubber.chkRevisions": ControlCaptionText = "Accept tracked changes and clear revision marks"
        Case "frmMetadataScrubber.optScopeDocument": ControlCaptionText = "Entire document (always)"
        Case "frmObjectRemover.chkFrames": ControlCaptionText = "Frames (remove frame, keep the text)"
        Case "frmObjectRemover.chkHiddenText": ControlCaptionText = "Hidden text"
        Case "frmObjectRemover.chkHorizontalLines": ControlCaptionText = "Horizontal lines"
        Case "frmObjectRemover.chkHtmlControls": ControlCaptionText = "HTML / ActiveX form controls"
        Case "frmObjectRemover.chkPictures": ControlCaptionText = "Pictures (inline and floating images)"
        Case "frmObjectRemover.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmObjectRemover.chkTables": ControlCaptionText = "Tables -- DELETES tables and all their contents"
        Case "frmObjectRemover.chkTextBoxes": ControlCaptionText = "Text boxes (and their contents)"
        Case "frmObjectRemover.optScopeDocument": ControlCaptionText = "Entire document (always)"
        Case "frmParagraphCleanup.chkCollapseBreaks": ControlCaptionText = "Convert soft returns (Shift+Enter) to paragraph marks"
        Case "frmParagraphCleanup.chkFixIndent": ControlCaptionText = "Remove manual tab / space indentation at paragraph start"
        Case "frmParagraphCleanup.chkNormalizeParaSpacing": ControlCaptionText = "Normalize Space Before / Space After settings"
        Case "frmParagraphCleanup.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmParagraphCleanup.chkRemoveEmpty": ControlCaptionText = "Collapse multiple consecutive empty paragraphs"
        Case "frmParagraphCleanup.optAll": ControlCaptionText = "All paragraph fixes"
        Case "frmParagraphCleanup.optCustom": ControlCaptionText = "Custom (choose below)"
        Case "frmParagraphCleanup.optNormalizeSpacing": ControlCaptionText = "Normalize paragraph spacing"
        Case "frmParagraphCleanup.optRemoveEmpty": ControlCaptionText = "Remove empty paragraphs"
        Case "frmParagraphCleanup.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmParagraphCleanup.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmPunctuationCleanup.chkCurlyDouble": ControlCaptionText = "Curly double quotes  -> straight double quotes"
        Case "frmPunctuationCleanup.chkCurlySingle": ControlCaptionText = "Curly single quotes / apostrophes  -> straight single quotes"
        Case "frmPunctuationCleanup.chkEllipses": ControlCaptionText = "Ellipsis character  ->  three dots  ..."
        Case "frmPunctuationCleanup.chkEmDash": ControlCaptionText = "Em dash  ->  hyphen-minus  -"
        Case "frmPunctuationCleanup.chkEnDash": ControlCaptionText = "En dash  ->  hyphen-minus  -"
        Case "frmPunctuationCleanup.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmPunctuationCleanup.optAll": ControlCaptionText = "All punctuation types"
        Case "frmPunctuationCleanup.optCustom": ControlCaptionText = "Custom (choose below)"
        Case "frmPunctuationCleanup.optDashes": ControlCaptionText = "Dashes only"
        Case "frmPunctuationCleanup.optEllipses": ControlCaptionText = "Ellipses only"
        Case "frmPunctuationCleanup.optQuotes": ControlCaptionText = "Quotes only"
        Case "frmPunctuationCleanup.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmPunctuationCleanup.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmSoftReturnConverter.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmSoftReturnConverter.optParaToSoft": ControlCaptionText = "Convert paragraph marks to soft returns"
        Case "frmSoftReturnConverter.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmSoftReturnConverter.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmSoftReturnConverter.optSoftToPara": ControlCaptionText = "Convert soft returns (Shift+Enter) to paragraph marks"
        Case "frmSpacingCleanup.chkDoubleSpaces": ControlCaptionText = "Collapse double (and multiple) spaces"
        Case "frmSpacingCleanup.chkExtraBlankLines": ControlCaptionText = "Collapse extra blank lines between paragraphs"
        Case "frmSpacingCleanup.chkNormalizeAfterPunct": ControlCaptionText = "Normalize spacing after punctuation"
        Case "frmSpacingCleanup.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmSpacingCleanup.chkSpaceBeforePunct": ControlCaptionText = "Remove space before punctuation  ( , . : ; ! ? )"
        Case "frmSpacingCleanup.chkTrimSpaces": ControlCaptionText = "Remove leading and trailing spaces"
        Case "frmSpacingCleanup.optAll": ControlCaptionText = "All spacing fixes"
        Case "frmSpacingCleanup.optCustom": ControlCaptionText = "Custom (choose below)"
        Case "frmSpacingCleanup.optDoubleSpaces": ControlCaptionText = "Double spaces only"
        Case "frmSpacingCleanup.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmSpacingCleanup.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmSpacingCleanup.optTrim": ControlCaptionText = "Trim leading / trailing spaces"
        Case "frmStyleCleanup.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmStyleCleanup.chkRemapVariants": ControlCaptionText = "Remap style variants to canonical equivalents"
        Case "frmStyleCleanup.chkRemoveUnused": ControlCaptionText = "Remove unused styles"
        Case "frmStyleCleanup.optScopeDocument": ControlCaptionText = "Entire document (always)"
        Case "frmTableCleaner.chkConvertToText": ControlCaptionText = "Convert single-column tables to plain text"
        Case "frmTableCleaner.chkNormalizeBorders": ControlCaptionText = "Normalize border styles"
        Case "frmTableCleaner.chkNormalizePadding": ControlCaptionText = "Normalize cell padding"
        Case "frmTableCleaner.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmTableCleaner.chkRemoveBorders": ControlCaptionText = "Remove all borders"
        Case "frmTableCleaner.chkRemoveEmptyCols": ControlCaptionText = "Remove empty columns"
        Case "frmTableCleaner.chkRemoveEmptyRows": ControlCaptionText = "Remove empty rows"
        Case "frmTableCleaner.chkStripDirectFormat": ControlCaptionText = "Strip direct formatting from cells"
        Case "frmTableCleaner.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmTableCleaner.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmUnicodeCleanup.chkBOM": ControlCaptionText = "Byte order marks (U+FEFF)"
        Case "frmUnicodeCleanup.chkNBHyphen": ControlCaptionText = "Non-breaking hyphens (U+2011)"
        Case "frmUnicodeCleanup.chkNBSP": ControlCaptionText = "Non-breaking spaces (U+00A0)"
        Case "frmUnicodeCleanup.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmUnicodeCleanup.chkSoftHyphen": ControlCaptionText = "Soft hyphens (U+00AD)"
        Case "frmUnicodeCleanup.chkZWJ": ControlCaptionText = "Zero-width joiners (U+200D)"
        Case "frmUnicodeCleanup.chkZWNJ": ControlCaptionText = "Zero-width non-joiners (U+200C)"
        Case "frmUnicodeCleanup.chkZWSP": ControlCaptionText = "Zero-width spaces (U+200B)"
        Case "frmUnicodeCleanup.optAll": ControlCaptionText = "All invisible / problem characters"
        Case "frmUnicodeCleanup.optCustom": ControlCaptionText = "Custom (choose below)"
        Case "frmUnicodeCleanup.optNBSP": ControlCaptionText = "Non-breaking spaces only"
        Case "frmUnicodeCleanup.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmUnicodeCleanup.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmUnicodeCleanup.optZeroWidth": ControlCaptionText = "Zero-width characters only"
        Case Else: ControlCaptionText = ""
    End Select
End Function
Private Sub SetButtonCaption(designer As Object, nm As String)
    On Error Resume Next
    Dim cap As String
    Select Case nm
        Case "cmdRun": cap = "Run"
        Case "cmdSelectAll": cap = "Select All"
        Case "cmdDeselectAll": cap = "Deselect All"
        Case "cmdHelp": cap = "Help"
        Case "cmdUnicode": cap = "Invisible Unicode"
        Case "cmdPunctuation": cap = "Punctuation"
        Case "cmdSpacing": cap = "Spacing"
        Case "cmdCapitalization": cap = "Capitalization"
        Case "cmdList": cap = "Lists"
        Case "cmdParagraph": cap = "Paragraphs"
        Case "cmdDuplicate": cap = "Duplicates"
        Case "cmdFontNorm": cap = "Fonts"
        Case "cmdTableClean": cap = "Tables"
        Case "cmdBreakNorm": cap = "Breaks"
        Case "cmdDocTrim": cap = "Document Trim"
        Case "cmdFormatStrip": cap = "Strip Formatting"
        Case "cmdHyperlink": cap = "Hyperlinks"
        Case "cmdSoftReturn": cap = "Soft Returns"
        Case "cmdMetadata": cap = "Metadata"
        Case "cmdStyleClean": cap = "Styles"
        Case "cmdFootnote": cap = "Footnotes"
        Case "cmdHeaderFooter": cap = "Headers / Footers"
        Case "cmdObjectRemover": cap = "Objects"
        Case Else: cap = Replace(nm, "cmd", "")
    End Select
    designer.Controls(nm).Caption = cap
End Sub

Private Function FriendlyFormCaption(formName As String) As String
    Select Case formName
        Case "frmCleanupSuiteLauncher": FriendlyFormCaption = "Cleanup Suite"
        Case "frmPunctuationCleanup": FriendlyFormCaption = "Punctuation Normalizer"
        Case "frmUnicodeCleanup": FriendlyFormCaption = "Invisible Unicode Cleaner"
        Case "frmSpacingCleanup": FriendlyFormCaption = "Spacing Fixer"
        Case "frmCapitalizationCleanup": FriendlyFormCaption = "Capitalization Fixer"
        Case "frmListCleanup": FriendlyFormCaption = "List Normalizer"
        Case "frmParagraphCleanup": FriendlyFormCaption = "Paragraph Structure Fixer"
        Case "frmDuplicateDetector": FriendlyFormCaption = "Duplicate Paragraph Detector"
        Case "frmFontNormalizer": FriendlyFormCaption = "Font Normalizer"
        Case "frmTableCleaner": FriendlyFormCaption = "Table Cleaner"
        Case "frmBreakNormalizer": FriendlyFormCaption = "Break Normalizer"
        Case "frmDocumentTrim": FriendlyFormCaption = "Document Trim"
        Case "frmFormattingStripper": FriendlyFormCaption = "Formatting Stripper"
        Case "frmHyperlinkRemover": FriendlyFormCaption = "Hyperlink Remover"
        Case "frmSoftReturnConverter": FriendlyFormCaption = "Soft Return Converter"
        Case "frmMetadataScrubber": FriendlyFormCaption = "Metadata Scrubber"
        Case "frmStyleCleanup": FriendlyFormCaption = "Style Cleanup"
        Case "frmFootnoteRemover": FriendlyFormCaption = "Footnote / Endnote Remover"
        Case "frmHeaderFooterStandardizer": FriendlyFormCaption = "Header / Footer Standardizer"
        Case "frmObjectRemover": FriendlyFormCaption = "Object Remover"
        Case Else: FriendlyFormCaption = formName
    End Select
End Function

Private Function FrameCaption(nm As String) As String
    Select Case nm
        Case "fraCustom": FrameCaption = "Custom options"
        Case "fraScopeSelection": FrameCaption = "Scope"
        Case "fraMatching": FrameCaption = "Matching"
        Case "fraThreshold": FrameCaption = "Fuzzy sensitivity"
        Case "fraConvertTo": FrameCaption = "Convert section breaks to"
        Case "fraEmphasis": FrameCaption = "Emphasis handling"
        Case "fraOptions": FrameCaption = "Options"
        Case "fraAlign": FrameCaption = "Alignment"
        Case Else: FrameCaption = ""
    End Select
End Function

Private Function ButtonTipText(nm As String) As String
    Select Case nm
        Case "cmdRun": ButtonTipText = "Run this cleanup tool"
        Case "cmdSelectAll": ButtonTipText = "Select all custom options"
        Case "cmdDeselectAll": ButtonTipText = "Clear all custom options"
        Case "cmdUnicode": ButtonTipText = "Remove invisible Unicode characters"
        Case "cmdPunctuation": ButtonTipText = "Normalize smart punctuation"
        Case "cmdSpacing": ButtonTipText = "Fix common spacing issues"
        Case "cmdCapitalization": ButtonTipText = "Apply consistent capitalization"
        Case "cmdList": ButtonTipText = "Normalize bullets, numbering, and list indents"
        Case "cmdParagraph": ButtonTipText = "Clean paragraph spacing and structure"
        Case "cmdDuplicate": ButtonTipText = "Find repeated or similar paragraphs"
        Case "cmdFontNorm": ButtonTipText = "Reset direct font overrides"
        Case "cmdTableClean": ButtonTipText = "Clean table structure and formatting"
        Case "cmdBreakNorm": ButtonTipText = "Normalize section and page breaks"
        Case "cmdDocTrim": ButtonTipText = "Trim document start and end whitespace"
        Case "cmdFormatStrip": ButtonTipText = "Strip direct formatting while preserving content"
        Case "cmdHyperlink": ButtonTipText = "Remove hyperlinks"
        Case "cmdSoftReturn": ButtonTipText = "Convert soft returns and paragraph breaks"
        Case "cmdMetadata": ButtonTipText = "Scrub document metadata and review artifacts"
        Case "cmdStyleClean": ButtonTipText = "Clean unused and duplicate styles"
        Case "cmdFootnote": ButtonTipText = "Remove or inline footnotes and endnotes"
        Case "cmdHeaderFooter": ButtonTipText = "Standardize headers and footers"
        Case "cmdObjectRemover": ButtonTipText = "Remove pictures, text boxes, and other objects"
        Case Else: ButtonTipText = ""
    End Select
End Function


' ---------------------------
' Helper: return control names expected for each form
' ---------------------------
Private Function ControlsForForm(formName As String) As Variant
    Select Case formName
        Case "frmCleanupSuiteLauncher"
            ControlsForForm = Array("lblLauncherTitle", "lblLauncherSubtitle", _
                                    "lblCatText", "cmdHelpUnicode", "cmdUnicode", "lblDescUnicode", "cmdHelpPunct", "cmdPunctuation", "lblDescPunctuation", "cmdHelpSpacing", "cmdSpacing", "lblDescSpacing", "cmdHelpCap", "cmdCapitalization", "lblDescCapitalization", _
                                    "lblCatPara", "cmdHelpList", "cmdList", "lblDescList", "cmdHelpPara", "cmdParagraph", "lblDescParagraph", "cmdHelpSoftReturn", "cmdSoftReturn", "lblDescSoftReturn", "cmdHelpDuplicate", "cmdDuplicate", "lblDescDuplicate", _
                                    "lblCatLayout", "cmdHelpTable", "cmdTableClean", "lblDescTableClean", "cmdHelpBreak", "cmdBreakNorm", "lblDescBreakNorm", "cmdHelpHeaderFooter", "cmdHeaderFooter", "lblDescHeaderFooter", "cmdHelpTrim", "cmdDocTrim", "lblDescDocTrim", _
                                    "lblCatFormat", "cmdHelpFont", "cmdFontNorm", "lblDescFontNorm", "cmdHelpFormat", "cmdFormatStrip", "lblDescFormatStrip", "cmdHelpStyle", "cmdStyleClean", "lblDescStyleClean", "cmdHelpHyperlink", "cmdHyperlink", "lblDescHyperlink", _
                                    "lblCatReview", "cmdHelpMetadata", "cmdMetadata", "lblDescMetadata", "cmdHelpFootnote", "cmdFootnote", "lblDescFootnote", "cmdHelpObject", "cmdObjectRemover", "lblDescObjectRemover", "chkAutoSave")
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
