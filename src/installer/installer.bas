' ===========================================================================
'  Professional Cleanup Suite for Microsoft Word
'  A self-installing collection of document cleanup tools.
'
'  MIT License
'
'  Copyright (c) 2026 MasseysLab
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
'  Generic tool menus use a two-column guided layout for opt/chk choices,
'  including Custom choices revealed after optCustom is selected.
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
    CreateOrReplaceModule vbProj, "modHybridJson", modHybridJson_Code, installLog
    CreateOrReplaceModule vbProj, "modHybridBridge", modHybridBridge_Code, installLog
    CreateOrReplaceModule vbProj, "modMetaDataSuite", modMetaDataSuite_Code, installLog
    CreateOrReplaceModule vbProj, "modUpdateChecker", modUpdateChecker_Code, installLog

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
    CreateOrReplaceForm vbProj, "frmMetaDataSuite", frmMetaDataSuite_Code, installLog
    CreateOrReplaceForm vbProj, "frmSafeMetadataEditor", frmSafeMetadataEditor_Code, installLog
    CreateOrReplaceForm vbProj, "frmFinalReview", frmFinalReview_Code, installLog
    CreateOrReplaceForm vbProj, "frmStyleCleanup", frmStyleCleanup_Code, installLog
    CreateOrReplaceForm vbProj, "frmFootnoteRemover", frmFootnoteRemover_Code, installLog
    CreateOrReplaceForm vbProj, "frmHeaderFooterStandardizer", frmHeaderFooterStandardizer_Code, installLog
    CreateOrReplaceForm vbProj, "frmObjectRemover", frmObjectRemover_Code, installLog
    CreateOrReplaceForm vbProj, "frmPreviewActions", frmPreviewActions_Code, installLog
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
    AttemptCreateControls vbProj, "frmMetaDataSuite", uiFailures
    AttemptCreateControls vbProj, "frmSafeMetadataEditor", uiFailures
    AttemptCreateControls vbProj, "frmFinalReview", uiFailures
    AttemptCreateControls vbProj, "frmStyleCleanup", uiFailures
    AttemptCreateControls vbProj, "frmFootnoteRemover", uiFailures
    AttemptCreateControls vbProj, "frmHeaderFooterStandardizer", uiFailures
    AttemptCreateControls vbProj, "frmObjectRemover", uiFailures
    AttemptCreateControls vbProj, "frmPreviewActions", uiFailures

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

    WriteInstallReport reportText

    Exit Sub

InstallerErr:
    MsgBox "Installer encountered an error: " & Err.Number & " - " & Err.Description, vbCritical, "Installer Error"
End Sub

Private Sub WriteInstallReport(reportText As String)
    On Error GoTo ReportErr

    Dim tempFolder As String
    tempFolder = Environ$("TEMP")
    If Len(tempFolder) = 0 Then tempFolder = Environ$("TMP")
    If Len(tempFolder) = 0 Then tempFolder = CurDir$
    If Right$(tempFolder, 1) = "\" Then tempFolder = Left$(tempFolder, Len(tempFolder) - 1)

    Dim reportPath As String
    reportPath = tempFolder & "\CleanupSuite_Install_Report_" & Format$(Now, "yyyymmdd_hhnnss") & ".txt"

    Dim ff As Integer
    ff = FreeFile
    Open reportPath For Output As #ff
    Print #ff, reportText
    Close #ff

    On Error Resume Next
    Shell "notepad.exe " & Chr(34) & reportPath & Chr(34), vbNormalFocus
    On Error GoTo 0

    MsgBox "Installer finished. The report has been saved to:" & vbCrLf & _
           reportPath & vbCrLf & vbCrLf & _
           "Follow the manual checklist if needed, then compile the project (VBA editor -> Debug -> Compile).", vbInformation, "Install Complete"
    Exit Sub

ReportErr:
    Dim reportErrNumber As Long
    Dim reportErrDescription As String
    reportErrNumber = Err.Number
    reportErrDescription = Err.Description

    On Error Resume Next
    If ff <> 0 Then Close #ff
    On Error GoTo 0

    MsgBox "Installer finished, but the report could not be written or opened." & vbCrLf & _
           "Report error " & reportErrNumber & ": " & reportErrDescription & vbCrLf & vbCrLf & _
           "The Cleanup Suite components were already installed before the report step.", vbExclamation, "Install Complete"
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
        "frmHyperlinkRemover", "frmSoftReturnConverter", "frmMetaDataSuite", "frmSafeMetadataEditor", "frmFinalReview", _
        "frmStyleCleanup", "frmFootnoteRemover", "frmHeaderFooterStandardizer", _
        "frmObjectRemover", "frmPreviewActions", "frmCleanupSuiteLauncher", "modCleanupLauncher", _
        "modCleanupHelpers", "modHybridJson", "modHybridBridge", _
        "modMetaDataSuite", "modUpdateChecker")

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

    RemoveToolFormHelpButton formName, designer

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
    If comp.Name = "frmPreviewActions" Then
        LayoutPreviewActionsControls comp, designer
        Exit Sub
    End If
    If comp.Name = "frmCapitalizationCleanup" Then
        LayoutCapitalizationControls comp, designer
        Exit Sub
    End If

    LayoutGenericToolControls comp, designer
End Sub

Private Sub LayoutGenericToolControls(comp As VBIDE.VBComponent, designer As Object)
    On Error Resume Next
    Const FORM_W As Single = 420
    Const MARGIN As Single = 14
    Const GAP As Single = 6
    Const BTN_H As Single = 24
    Dim contentW As Single: contentW = FORM_W - (2 * MARGIN)
    Dim halfW As Single: halfW = (contentW - GAP) / 2
    Dim y As Single: y = 36
    Dim pend As String: pend = ""
    Dim pendingChoice As String: pendingChoice = ""
    Dim ctrlNames As Variant: ctrlNames = ControlsForForm(comp.Name)
    Dim cn As Variant, pfx As String
    Dim h As Single, lft As Single, w As Single

    PositionControl designer, "lblTitle", 0, 0, 0, 0
    designer.Controls("lblTitle").Caption = ""
    designer.Controls("lblTitle").Visible = False
    HideGenericTitleLabels designer
    PositionControl designer, "cmdReset", MARGIN + contentW - 50, 8, 50, 18
    designer.Controls("cmdReset").Caption = "Reset"
    designer.Controls("cmdReset").Font.Size = 7.5
    PositionControl designer, "cmdRun", 0, 0, 0, 0
    designer.Controls("cmdRun").Visible = False

    For Each cn In ctrlNames
        If InStr(1, CStr(cn), "RiskPlacement", vbTextCompare) > 0 Then
            designer.Controls(CStr(cn)).Visible = False
            PositionControl designer, CStr(cn), 0, 0, 0, 0
            GoTo NextControl
        End If
        If CStr(cn) = "chkPreviewOnly" Or CStr(cn) = "fraScopeSelection" Then
            FlushGenericPendingChoice designer, pendingChoice, MARGIN, y, contentW, GAP
            FlushGenericPendingButton designer, pend, MARGIN, y, contentW, BTN_H, GAP
            designer.Controls(CStr(cn)).Visible = False
            PositionControl designer, CStr(cn), 0, 0, 0, 0
            GoTo NextControl
        End If
        If CStr(cn) = "cmdPreview" Then
            GoTo NextControl
        End If
        If CStr(cn) = "cmdRun" Or CStr(cn) = "cmdReset" Then
            GoTo NextControl
        End If
        If Left$(CStr(cn), 3) = "fra" Then
            FlushGenericPendingChoice designer, pendingChoice, MARGIN, y, contentW, GAP
            FlushGenericPendingButton designer, pend, MARGIN, y, contentW, BTN_H, GAP
            PositionControl designer, CStr(cn), 0, 0, 0, 0
            designer.Controls(CStr(cn)).Visible = False
            GoTo NextControl
        End If
        pfx = Left$(CStr(cn), 3)
        If pfx = "cmd" Then
            FlushGenericPendingChoice designer, pendingChoice, MARGIN, y, contentW, GAP
            SetButtonCaption designer, CStr(cn)
            If Len(pend) > 0 Then
                PositionControl designer, pend, MARGIN, y, halfW, BTN_H
                PositionControl designer, CStr(cn), MARGIN + halfW + GAP, y, halfW, BTN_H
                y = y + BTN_H + GAP
                pend = ""
            Else
                pend = CStr(cn)
            End If
        Else
            Select Case pfx
                Case "lbl"
                    FlushGenericPendingChoice designer, pendingChoice, MARGIN, y, contentW, GAP
                    FlushGenericPendingButton designer, pend, MARGIN, y, contentW, BTN_H, GAP
                    h = 34: lft = MARGIN: w = contentW
                    designer.Controls(CStr(cn)).WordWrap = True
                    designer.Controls(CStr(cn)).AutoSize = False
                    PositionControl designer, CStr(cn), lft, y, w, h
                    y = y + h + GAP
                Case "opt", "chk"
                    FlushGenericPendingButton designer, pend, MARGIN, y, contentW, BTN_H, GAP
                    If Len(pendingChoice) > 0 Then
                        PositionGenericChoicePair designer, pendingChoice, CStr(cn), MARGIN, y, contentW, GAP
                    Else
                        pendingChoice = CStr(cn)
                    End If
                Case Else
                    FlushGenericPendingChoice designer, pendingChoice, MARGIN, y, contentW, GAP
                    FlushGenericPendingButton designer, pend, MARGIN, y, contentW, BTN_H, GAP
                    h = 20: lft = MARGIN: w = contentW
                    PositionControl designer, CStr(cn), lft, y, w, h
                    y = y + h + GAP
            End Select
        End If
NextControl:
    Next cn
    FlushGenericPendingChoice designer, pendingChoice, MARGIN, y, contentW, GAP
    FlushGenericPendingButton designer, pend, MARGIN, y, contentW, BTN_H, GAP
    y = y + 2
    LayoutGenericPreviewAndScopeRow designer, y, MARGIN, contentW, BTN_H
    PositionControl designer, "cmdRun", 0, 0, 0, 0
    designer.Controls("cmdRun").Visible = False
    y = y + 20
    comp.Properties("Width") = FORM_W + 8
    comp.Properties("Height") = y + 40
    designer.Width = FORM_W + 8
    designer.Height = y + 40
End Sub

Private Sub LayoutGenericPreviewAndScopeRow(designer As Object, ByRef y As Single, ByVal MARGIN As Single, ByVal contentW As Single, ByVal BTN_H As Single)
    On Error Resume Next
    If DesignerScopeSelectionAvailable(designer) Then
        Dim scopeW As Single
        Dim previewW As Single
        Dim previewLeft As Single
        Dim scopeLeft As Single
        scopeW = 154
        previewW = contentW - scopeW - 10
        scopeLeft = MARGIN
        previewLeft = MARGIN + scopeW + 10
        designer.Controls("optScopeSelection").Caption = "Selected text only"
        designer.Controls("optScopeDocument").Caption = "Entire document"
        PositionControl designer, "optScopeSelection", scopeLeft, y, scopeW, 18
        PositionControl designer, "optScopeDocument", scopeLeft, y + 20, scopeW, 18
        PositionControl designer, "cmdPreview", previewLeft, y, previewW, 42
        y = y + 48
    Else
        PositionControl designer, "optScopeDocument", 0, 0, 0, 0
        PositionControl designer, "optScopeSelection", 0, 0, 0, 0
        designer.Controls("optScopeDocument").Visible = False
        designer.Controls("optScopeSelection").Visible = False
        PositionControl designer, "cmdPreview", MARGIN, y, contentW, BTN_H + 6
        y = y + BTN_H + 12
    End If
End Sub

Private Sub FlushGenericPendingButton(designer As Object, ByRef pendingButton As String, ByVal MARGIN As Single, ByRef y As Single, ByVal contentW As Single, ByVal BTN_H As Single, ByVal GAP As Single)
    On Error Resume Next
    If Len(pendingButton) = 0 Then Exit Sub
    PositionControl designer, pendingButton, MARGIN, y, contentW, BTN_H
    y = y + BTN_H + GAP
    pendingButton = ""
End Sub

Private Sub FlushGenericPendingChoice(designer As Object, ByRef pendingChoice As String, ByVal MARGIN As Single, ByRef y As Single, ByVal contentW As Single, ByVal GAP As Single)
    On Error Resume Next
    If Len(pendingChoice) = 0 Then Exit Sub
    Dim singleChoiceW As Single
    singleChoiceW = GenericSingleChoiceWidth(designer, pendingChoice, contentW)
    Dim rowH As Single
    rowH = GenericChoiceHeight(designer, pendingChoice)
    PositionControl designer, pendingChoice, MARGIN + ((contentW - singleChoiceW) / 2), y, singleChoiceW, rowH
    designer.Controls(pendingChoice).WordWrap = True
    designer.Controls(pendingChoice).AutoSize = False
    y = y + rowH + GAP
    pendingChoice = ""
End Sub

Private Sub PositionGenericChoicePair(designer As Object, ByRef pendingChoice As String, ByVal controlName As String, ByVal MARGIN As Single, ByRef y As Single, ByVal contentW As Single, ByVal GAP As Single)
    On Error Resume Next
    Dim choiceW As Single
    choiceW = (contentW - GAP) / 2
    Dim rowH As Single
    rowH = GenericChoiceRowHeight(designer, pendingChoice, controlName)
    PositionControl designer, pendingChoice, MARGIN + 8, y, choiceW - 8, rowH
    PositionControl designer, controlName, MARGIN + choiceW + GAP + 8, y, choiceW - 8, rowH
    designer.Controls(pendingChoice).WordWrap = True
    designer.Controls(pendingChoice).AutoSize = False
    designer.Controls(controlName).WordWrap = True
    designer.Controls(controlName).AutoSize = False
    y = y + rowH + GAP
    pendingChoice = ""
End Sub

Private Function GenericChoiceRowHeight(designer As Object, ByVal firstName As String, ByVal secondName As String) As Single
    On Error Resume Next
    GenericChoiceRowHeight = GenericChoiceHeight(designer, firstName)
    If GenericChoiceHeight(designer, secondName) > GenericChoiceRowHeight Then GenericChoiceRowHeight = GenericChoiceHeight(designer, secondName)
End Function

Private Function GenericChoiceHeight(designer As Object, ByVal controlName As String) As Single
    On Error Resume Next
    Dim captionLength As Long
    captionLength = Len(CStr(designer.Controls(controlName).Caption))
    If captionLength > 72 Then
        GenericChoiceHeight = 44
    ElseIf captionLength > 34 Then
        GenericChoiceHeight = 28
    Else
        GenericChoiceHeight = 18
    End If
End Function
Private Function GenericSingleChoiceWidth(designer As Object, ByVal controlName As String, ByVal contentW As Single) As Single
    On Error Resume Next
    Dim pairW As Single
    pairW = ((contentW - 6) / 2) - 8
    GenericSingleChoiceWidth = pairW + 12
    If GenericChoiceHeight(designer, controlName) > 18 Then GenericSingleChoiceWidth = pairW + 28
    If Len(Trim$(CStr(designer.Controls(controlName).Caption))) > 42 Then
        GenericSingleChoiceWidth = contentW - 24
    ElseIf Len(Trim$(CStr(designer.Controls(controlName).Caption))) > 30 Then
        GenericSingleChoiceWidth = pairW + 72
    End If
    If GenericSingleChoiceWidth < pairW Then GenericSingleChoiceWidth = pairW
    If GenericSingleChoiceWidth > contentW Then GenericSingleChoiceWidth = contentW
End Function
Private Function DesignerScopeSelectionAvailable(designer As Object) As Boolean
    On Error Resume Next
    DesignerScopeSelectionAvailable = _
        InStr(1, CStr(designer.Controls("optScopeDocument").Caption), "(always)", vbTextCompare) = 0 _
        And Len(Trim$(CStr(designer.Controls("optScopeSelection").Caption))) > 0
End Function
Private Sub HideGenericTitleLabels(designer As Object)
    On Error Resume Next
    Dim ctl As Object
    For Each ctl In designer.Controls
        If Left$(LCase$(ctl.Name), 3) <> "lbl" Then GoTo NextGenericTitleControl
        If ctl.Name = "lblIntro" Or ctl.Name = "lblModeSummary" Or ctl.Name = "lblSpeedWarning" Or ctl.Name = "lblFuzzyWarning" Then GoTo NextGenericTitleControl
        If Len(Trim$(CStr(ctl.Caption))) = 0 Then GoTo NextGenericTitleControl
        If ctl.Top <= 24 And ctl.Height <= 24 Then
            PositionControl designer, ctl.Name, 0, 0, 0, 0
            ctl.Visible = False
        End If
NextGenericTitleControl:
    Next ctl
End Sub

Private Sub LayoutCapitalizationControls(comp As VBIDE.VBComponent, designer As Object)
    On Error Resume Next
    Const FORM_W As Single = 420
    Const MARGIN As Single = 14
    Const GAP As Single = 6
    Const BTN_H As Single = 24
    Dim contentW As Single: contentW = FORM_W - (2 * MARGIN)
    Dim optionW As Single: optionW = (contentW - GAP) / 2
    Dim y As Single

    PositionControl designer, "lblIntro", MARGIN, 10, contentW, 18
    designer.Controls("lblIntro").TextAlign = 2

    y = 36
    PositionControl designer, "optAll", MARGIN, y, optionW, 18
    PositionControl designer, "optCustom", MARGIN + optionW + GAP, y, optionW, 18
    y = y + 28

    PositionControl designer, "lblModeSummary", MARGIN, y, contentW, 34
    designer.Controls("lblModeSummary").WordWrap = True
    y = y + 42

    PositionControl designer, "chkSentence", MARGIN + 8, y, contentW - 16, 18
    PositionControl designer, "chkTitle", MARGIN + 8, y + 22, contentW - 16, 18
    PositionControl designer, "chkUpper", MARGIN + 8, y + 44, contentW - 16, 18
    PositionControl designer, "chkLower", MARGIN + 8, y + 66, contentW - 16, 18
    PositionControl designer, "chkHeadingParentheses", MARGIN + 28, y + 88, contentW - 40, 18
    PositionControl designer, "chkSmartSentences", MARGIN + 8, y + 110, contentW - 16, 18
    PositionControl designer, "cmdEditExceptions", MARGIN + 8, y + 132, contentW - 16, BTN_H
    PositionControl designer, "cmdSelectAll", MARGIN, y + 162, (contentW - GAP) / 2, BTN_H
    PositionControl designer, "cmdDeselectAll", MARGIN + ((contentW - GAP) / 2) + GAP, y + 162, (contentW - GAP) / 2, BTN_H
    designer.Controls("cmdEditExceptions").Visible = False
    designer.Controls("cmdSelectAll").Visible = False
    designer.Controls("cmdDeselectAll").Visible = False

    y = y + 182
    designer.Controls("fraScopeSelection").Visible = False
    designer.Controls("chkPreviewOnly").Visible = False
    PositionControl designer, "cmdReset", MARGIN + contentW - 50, 8, 50, 18
    designer.Controls("cmdReset").Caption = "Reset"
    designer.Controls("cmdReset").Font.Size = 7.5
    LayoutGenericPreviewAndScopeRow designer, y, MARGIN, contentW, BTN_H
    PositionControl designer, "cmdRun", 0, 0, 0, 0
    designer.Controls("cmdRun").Visible = False
    y = y + 20

    comp.Properties("Width") = FORM_W + 8
    comp.Properties("Height") = y + 34
    designer.Width = FORM_W + 8
    designer.Height = y + 34
End Sub

Private Sub LayoutPreviewActionsControls(comp As VBIDE.VBComponent, designer As Object)
    On Error Resume Next
    Const FORM_W As Single = 283
    Const FORM_H As Single = 126
    Const CONTENT_W As Single = 257
    Const MARGIN As Single = 8
    Const BTN_W As Single = (CONTENT_W - (2 * GAP)) / 3
    Const BTN_H As Single = 25
    Const GAP As Single = 5
    Const ACTION_TOP As Single = 17
    Const NAV_TOP As Single = ACTION_TOP + BTN_H + GAP
    Const SUMMARY_TOP As Single = NAV_TOP + BTN_H + GAP
    Const NAV_W As Single = (CONTENT_W - (3 * GAP)) / 4

    ' Keep the design-time MSForms canvas caption blank; Configure sets the native title bar.
    comp.Properties("Caption") = ""
    designer.Caption = ""
    PositionControl designer, "lblTitle", 0, 0, 0, 0
    PositionControl designer, "lblHint", MARGIN, 2, CONTENT_W, 12
    PositionControl designer, "cmdPreview", MARGIN, ACTION_TOP, BTN_W, BTN_H
    PositionControl designer, "cmdClear", MARGIN + BTN_W + GAP, ACTION_TOP, BTN_W, BTN_H
    PositionControl designer, "cmdApply", MARGIN + (2 * (BTN_W + GAP)), ACTION_TOP, BTN_W, BTN_H
    PositionControl designer, "cmdPreviousChange", MARGIN, NAV_TOP, NAV_W, BTN_H
    PositionControl designer, "cmdPreviousPage", MARGIN + NAV_W + GAP, NAV_TOP, NAV_W, BTN_H
    PositionControl designer, "cmdNextPage", MARGIN + (2 * (NAV_W + GAP)), NAV_TOP, NAV_W, BTN_H
    PositionControl designer, "cmdNextChange", MARGIN + (3 * (NAV_W + GAP)), NAV_TOP, NAV_W, BTN_H
    PositionControl designer, "lblButtonMeasure", 0, 0, 0, 0
    PositionControl designer, "lblSummary", MARGIN, SUMMARY_TOP, CONTENT_W, 26
    designer.Controls("lblTitle").Caption = ""
    designer.Controls("lblTitle").Visible = False
    designer.Controls("lblButtonMeasure").Caption = ""
    designer.Controls("lblButtonMeasure").Visible = False
    designer.Controls("lblSummary").WordWrap = True

    comp.Properties("Width") = FORM_W + 8
    comp.Properties("Height") = FORM_H
    designer.Width = FORM_W + 8
    designer.Height = FORM_H
End Sub

Private Sub LayoutLauncherControls(comp As VBIDE.VBComponent, designer As Object)
    On Error Resume Next
    Const FORM_W As Single = 760
    Const MARGIN As Single = 14
    Const GAP As Single = 3
    Const COL_GAP As Single = 16
    Const COL_W As Single = (FORM_W - (2 * MARGIN) - COL_GAP) / 2
    Const COL1_X As Single = MARGIN
    Const COL2_X As Single = MARGIN + COL_W + COL_GAP
    Const TOP_Y As Single = 8
    Dim y As Single: y = TOP_Y
    Dim yCol1 As Single
    Dim yCol2 As Single
    Dim calloutLeft As Single
    Dim calloutTop As Single
    Dim calloutWidth As Single
    Dim calloutHeight As Single

    comp.Properties("Caption") = ""
    designer.Caption = ""

    PositionControl designer, "lblLauncherTitle", COL1_X, y, COL_W, 18
    PositionControl designer, "lblLauncherSubtitle", COL1_X, y + 20, COL_W, 28
    designer.Controls("lblLauncherTitle").TextAlign = 1
    designer.Controls("lblLauncherSubtitle").TextAlign = 1

    yCol1 = y + 52
    yCol1 = LayoutLauncherCategory(designer, "lblCatText", COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(designer, "cmdHelpUnicode", "cmdUnicode", "lblDescUnicode", "lblRiskUnicode", COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(designer, "cmdHelpPunct", "cmdPunctuation", "lblDescPunctuation", "lblRiskPunctuation", COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(designer, "cmdHelpSpacing", "cmdSpacing", "lblDescSpacing", "lblRiskSpacing", COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(designer, "cmdHelpCap", "cmdCapitalization", "lblDescCapitalization", "lblRiskCapitalization", COL1_X, yCol1, COL_W)
    yCol1 = yCol1 + GAP
    yCol1 = LayoutLauncherCategory(designer, "lblCatPara", COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(designer, "cmdHelpList", "cmdList", "lblDescList", "lblRiskList", COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(designer, "cmdHelpPara", "cmdParagraph", "lblDescParagraph", "lblRiskParagraph", COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(designer, "cmdHelpSoftReturn", "cmdSoftReturn", "lblDescSoftReturn", "lblRiskSoftReturn", COL1_X, yCol1, COL_W)
    yCol1 = LayoutLauncherToolRow(designer, "cmdHelpTrim", "cmdDocTrim", "lblDescDocTrim", "lblRiskDocTrim", COL1_X, yCol1, COL_W)

    yCol2 = TOP_Y
    yCol2 = LayoutLauncherCategory(designer, "lblCatLayout", COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(designer, "cmdHelpTable", "cmdTableClean", "lblDescTableClean", "lblRiskTableClean", COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(designer, "cmdHelpHeaderFooter", "cmdHeaderFooter", "lblDescHeaderFooter", "lblRiskHeaderFooter", COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(designer, "cmdHelpFootnote", "cmdFootnote", "lblDescFootnote", "lblRiskFootnote", COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(designer, "cmdHelpObject", "cmdObjectRemover", "lblDescObjectRemover", "lblRiskObjectRemover", COL2_X, yCol2, COL_W)
    yCol2 = yCol2 + GAP
    yCol2 = LayoutLauncherCategory(designer, "lblCatFormat", COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(designer, "cmdHelpFont", "cmdFontNorm", "lblDescFontNorm", "lblRiskFontNorm", COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(designer, "cmdHelpFormat", "cmdFormatStrip", "lblDescFormatStrip", "lblRiskFormatStrip", COL2_X, yCol2, COL_W)
    yCol2 = yCol2 + GAP
    yCol2 = LayoutLauncherCategory(designer, "lblCatReview", COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(designer, "cmdHelpDuplicate", "cmdDuplicate", "lblDescDuplicate", "lblRiskDuplicate", COL2_X, yCol2, COL_W)
    yCol2 = yCol2 + GAP
    yCol2 = LayoutLauncherCategory(designer, "lblCatShare", COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(designer, "cmdHelpMetadata", "cmdMetadata", "lblDescMetadata", "lblRiskMetadata", COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(designer, "cmdHelpFinalReview", "cmdFinalReview", "lblDescFinalReview", "lblRiskFinalReview", COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(designer, "cmdHelpHyperlink", "cmdHyperlink", "lblDescHyperlink", "lblRiskHyperlink", COL2_X, yCol2, COL_W)
    yCol2 = yCol2 + GAP
    yCol2 = LayoutLauncherCategory(designer, "lblCatAlpha", COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(designer, "cmdHelpStyle", "cmdStyleClean", "lblDescStyleClean", "lblRiskStyleClean", COL2_X, yCol2, COL_W)
    yCol2 = LayoutLauncherToolRow(designer, "cmdHelpBreak", "cmdBreakNorm", "lblDescBreakNorm", "lblRiskBreakNorm", COL2_X, yCol2, COL_W)
    y = yCol1 + GAP

    PositionControl designer, "chkAutoSave", MARGIN + 2, y, 225, 18
    PositionControl designer, "lblAutoSaveWarning", MARGIN + 18, y + 16, COL_W - 18, 14
    y = y + 32
    PositionControl designer, "chkReturnToMainAfterApply", MARGIN + 2, y, COL_W, 18
    y = y + 18
    PositionControl designer, "chkReturnToMainAfterClose", MARGIN + 2, y, COL_W, 18
    y = y + 20
    PositionControl designer, "chkUpdateChecks", MARGIN + 2, y, 214, 18
    PositionControl designer, "cmdCheckUpdates", MARGIN + 230, y - 3, 112, 24
    y = y + 26
    PositionControl designer, "cmdUserManual", MARGIN + 2, y, 96, 24
    PositionControl designer, "cmdResetAll", MARGIN + 104, y, 96, 24
    calloutLeft = MARGIN + 220
    calloutTop = y - 2
    calloutWidth = (MARGIN + COL_W) - calloutLeft
    calloutHeight = 32
    PositionControl designer, "lblResetAllArrow", MARGIN + 202, y + 7, 16, 16
    PositionControl designer, "lblResetAllCalloutBox", calloutLeft, calloutTop, calloutWidth, calloutHeight
    PositionControl designer, "lblResetAllCalloutText", calloutLeft + 5, calloutTop + 4, calloutWidth - 10, calloutHeight - 8

    comp.Properties("Width") = FORM_W + 8
    comp.Properties("Height") = MaxSingle(calloutTop + calloutHeight, yCol2) + 40
    designer.Width = FORM_W + 8
    designer.Height = MaxSingle(calloutTop + calloutHeight, yCol2) + 40
End Sub

Private Function LayoutLauncherCategory(designer As Object, labelName As String, L As Single, T As Single, W As Single) As Single
    On Error Resume Next
    PositionControl designer, labelName, L, T, W, 16
    LayoutLauncherCategory = T + 18
End Function

Private Function LayoutLauncherToolRow(designer As Object, helpName As String, buttonName As String, descName As String, riskName As String, L As Single, T As Single, W As Single) As Single
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

    PositionControl designer, helpName, L, T + 5, HELP_W, 20
    PositionControl designer, buttonName, L + HELP_W + INNER_GAP, T + 3, BTN_W, 24
    PositionControl designer, riskName, L + HELP_W + BTN_W + (2 * INNER_GAP), T + 1, RISK_W, RISK_H
    PositionControl designer, descName, L + HELP_W + BTN_W + RISK_W + (3 * INNER_GAP), T + 1, descW, ROW_H
    StyleLauncherRiskLabel designer, riskName
    LayoutLauncherToolRow = T + ROW_STEP
End Function

Private Sub StyleLauncherRiskLabel(designer As Object, riskName As String)
    On Error Resume Next
    Dim labelText As String
    With designer.Controls(riskName)
        labelText = Replace(CStr(.Caption), vbCrLf, " ")
        .Caption = LauncherRiskDisplayText(labelText)
        .TextAlign = 2
        .Font.Size = 7.5
        .Font.Bold = True
        .WordWrap = True
        .TakeFocusOnClick = False
        .BackStyle = 1
        .SpecialEffect = 0
        .BorderStyle = 1
        .BackColor = LauncherRiskBackColor(labelText)
        .ForeColor = LauncherRiskForeColor(labelText)
        .BorderColor = .ForeColor
        .ControlTipText = LauncherRiskTip(labelText)
    End With
End Sub

Private Function LauncherRiskDisplayText(labelText As String) As String
    LauncherRiskDisplayText = Replace(labelText, " ", vbCrLf)
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
    If FormBodyOwnsTitle(comp.Name) Then
        comp.Properties("Caption") = ""
        designer.Caption = ""
    End If
    designer.BackColor = RGB(248, 249, 251)
    designer.SpecialEffect = 0
End Sub
Private Function FormBodyOwnsTitle(formName As String) As Boolean
    Select Case formName
        Case "frmCleanupSuiteLauncher", "frmPreviewActions"
            FormBodyOwnsTitle = True
        Case Else
            FormBodyOwnsTitle = False
    End Select
End Function

Private Sub RemoveToolFormHelpButton(formName As String, designer As Object)
    On Error Resume Next
    If formName = "frmCleanupSuiteLauncher" Or formName = "frmPreviewActions" Then Exit Sub
    If Not FormBodyOwnsTitle(formName) Then Exit Sub
    designer.Controls.Remove "cmdHelp"
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
            ctl.Enabled = True
            ctl.Font.Bold = False
            ctl.BackColor = RGB(255, 255, 255)
            ctl.ControlTipText = ButtonTipText(nm)
            If formName = "frmPreviewActions" Then
                ctl.Font.Size = 10
                ctl.Font.Bold = True
                ctl.ForeColor = RGB(0, 0, 0)
            End If
            If Left$(nm, 7) = "cmdHelp" And nm <> "cmdHelp" Then
                ctl.Caption = "?"
                ctl.ControlTipText = "Show help for this tool"
                ctl.Font.Bold = True
            End If
            If nm = "cmdRun" Then
                ctl.Font.Bold = True
                ctl.ControlTipText = "Apply the selected cleanup options"
            End If
            If nm = "cmdPreview" And formName <> "frmPreviewActions" Then
                ctl.ControlTipText = "Preview with the current settings"
                ctl.Font.Bold = True
                ctl.ForeColor = RGB(24, 90, 145)
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
            If formName = "frmPreviewActions" Then
                ctl.WordWrap = False
                ctl.AutoSize = False
                ctl.ForeColor = RGB(0, 0, 0)
                ctl.Font.Size = 11
                ctl.Font.Bold = False
                If nm = "lblHint" Then
                    ctl.Font.Bold = True
                ElseIf nm = "lblSummary" Then
                    ctl.Font.Size = 10
                    ctl.Font.Bold = True
                End If
            End If
            If formName = "frmCleanupSuiteLauncher" Then
                ctl.WordWrap = True
                ctl.AutoSize = False
                If nm = "lblLauncherTitle" Then
                    ctl.TextAlign = 1
                    ctl.Font.Size = 14
                    ctl.Font.Bold = True
                    ctl.ForeColor = RGB(28, 43, 58)
                ElseIf nm = "lblLauncherSubtitle" Then
                    ctl.TextAlign = 1
                    ctl.Font.Size = 8.5
                    ctl.ForeColor = RGB(88, 101, 116)
                ElseIf Left$(nm, 6) = "lblCat" Then
                    ctl.Font.Size = 9
                    ctl.Font.Bold = True
                    ctl.ForeColor = RGB(38, 91, 137)
                ElseIf Left$(nm, 7) = "lblDesc" Then
                    ctl.Font.Size = 8
                    ctl.ForeColor = RGB(93, 105, 119)
                ElseIf Left$(nm, 7) = "lblRisk" Then
                    StyleLauncherRiskLabel designer, nm
                ElseIf nm = "lblResetAllArrow" Then
                    ctl.Caption = ChrW$(&H2190)
                    ctl.TextAlign = 2
                    ctl.Font.Size = 12
                    ctl.Font.Bold = True
                    ctl.BackStyle = 0
                    ctl.ForeColor = RGB(220, 0, 0)
                ElseIf nm = "lblResetAllCalloutBox" Then
                    ctl.Caption = ""
                    ctl.BackStyle = 0
                    ctl.BorderStyle = 1
                    ctl.SpecialEffect = 0
                    ctl.ForeColor = RGB(220, 0, 0)
                    ctl.BorderColor = RGB(220, 0, 0)
                ElseIf nm = "lblResetAllCalloutText" Then
                    ctl.Font.Size = 7.5
                    ctl.BackStyle = 0
                    ctl.ForeColor = RGB(32, 37, 45)
                ElseIf nm = "lblAutoSaveWarning" Then
                    ctl.Font.Size = 8
                    ctl.Font.Bold = True
                    ctl.BackStyle = 0
                    ctl.ForeColor = RGB(220, 0, 0)
                End If
            End If
            If formName = "frmCapitalizationCleanup" Then
                ctl.WordWrap = True
                ctl.AutoSize = False
                If nm = "lblIntro" Then
                    ctl.TextAlign = 2
                    ctl.Font.Size = 9
                    ctl.ForeColor = RGB(88, 101, 116)
                ElseIf nm = "lblModeSummary" Then
                    ctl.Font.Size = 8.5
                    ctl.ForeColor = RGB(67, 80, 96)
                    ctl.BackColor = RGB(255, 255, 255)
                    ctl.BorderStyle = 1
                    ctl.SpecialEffect = 0
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
        Case "frmCapitalizationCleanup.chkLower": ControlCaptionText = "Repair likely headings in title case"
        Case "frmCapitalizationCleanup.chkHeadingParentheses": ControlCaptionText = "Also title-case words inside parentheses"
        Case "frmCapitalizationCleanup.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmCapitalizationCleanup.chkSentence": ControlCaptionText = "Recognize common abbreviations"
        Case "frmCapitalizationCleanup.chkSmartSentences": ControlCaptionText = "Recognize quotes, bullets, and sentence boundaries"
        Case "frmCapitalizationCleanup.chkTitle": ControlCaptionText = "Protect common acronyms"
        Case "frmCapitalizationCleanup.chkUpper": ControlCaptionText = "Protect names and brands"
        Case "frmCapitalizationCleanup.cmdEditExceptions": ControlCaptionText = "Edit custom exceptions"
        Case "frmCapitalizationCleanup.lblIntro": ControlCaptionText = "Use the recommended repair, or customize its protections."
        Case "frmCapitalizationCleanup.lblModeSummary": ControlCaptionText = "Repairs sentence starts, obvious all-caps damage, headings, acronyms, names, and brands."
        Case "frmCapitalizationCleanup.optAll": ControlCaptionText = "Recommended repair"
        Case "frmCapitalizationCleanup.optCustom": ControlCaptionText = "Custom repair"
        Case "frmCapitalizationCleanup.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmCapitalizationCleanup.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmCleanupSuiteLauncher.chkAutoSave": ControlCaptionText = "Global default: auto-save before each tool"
        Case "frmCleanupSuiteLauncher.chkReturnToMainAfterApply": ControlCaptionText = "Global default: return to main menu after apply"
        Case "frmCleanupSuiteLauncher.chkReturnToMainAfterClose": ControlCaptionText = "Global default: return to main menu after closing individual tool menus"
        Case "frmCleanupSuiteLauncher.chkUpdateChecks": ControlCaptionText = "Check periodically for updates"
        Case "frmCleanupSuiteLauncher.lblAutoSaveWarning": ControlCaptionText = "It is dangerous to turn off auto save!"
        Case "frmMetaDataSuite.chkShowSuiteMetadata": ControlCaptionText = "Show CleanupSuite metadata"
        Case "frmCleanupSuiteLauncher.lblLauncherTitle": ControlCaptionText = "Cleanup Suite"
        Case "frmCleanupSuiteLauncher.lblLauncherSubtitle": ControlCaptionText = "Choose the kind of cleanup you need. Tools are grouped by what you are trying to fix."
        Case "frmCleanupSuiteLauncher.lblResetAllCalloutText": ControlCaptionText = "If highlights remain after an error, select Reset All."
        Case "frmCleanupSuiteLauncher.lblCatText": ControlCaptionText = "Clean Pasted Text"
        Case "frmCleanupSuiteLauncher.lblCatPara": ControlCaptionText = "Fix Paragraphs and Lists"
        Case "frmCleanupSuiteLauncher.lblCatReview": ControlCaptionText = "Review and Inspect"
        Case "frmCleanupSuiteLauncher.lblCatLayout": ControlCaptionText = "Clean Document Objects"
        Case "frmCleanupSuiteLauncher.lblCatFormat": ControlCaptionText = "Clean Formatting"
        Case "frmCleanupSuiteLauncher.lblCatShare": ControlCaptionText = "Prepare for Sharing"
        Case "frmCleanupSuiteLauncher.lblCatAlpha": ControlCaptionText = "Risky Alpha Tools"
        Case "frmCleanupSuiteLauncher.lblDescUnicode": ControlCaptionText = "Remove invisible Unicode characters that break search, sorting, and export."
        Case "frmCleanupSuiteLauncher.lblDescPunctuation": ControlCaptionText = "Normalize smart quotes, dashes, ellipses, and plain-text punctuation."
        Case "frmCleanupSuiteLauncher.lblDescSpacing": ControlCaptionText = "Fix extra spaces, punctuation spacing, and repeated blank lines."
        Case "frmCleanupSuiteLauncher.lblDescCapitalization": ControlCaptionText = "Apply sentence, title, upper, lower, or smart capitalization cleanup."
        Case "frmCleanupSuiteLauncher.lblDescList": ControlCaptionText = "Normalize manual bullets, numbering, and list indentation."
        Case "frmCleanupSuiteLauncher.lblDescParagraph": ControlCaptionText = "Clean empty paragraphs, paragraph spacing, and manual indents."
        Case "frmCleanupSuiteLauncher.lblDescSoftReturn": ControlCaptionText = "Convert Shift+Enter line breaks to paragraph marks, or the reverse of that."
        Case "frmCleanupSuiteLauncher.lblDescDuplicate": ControlCaptionText = "Find repeated or near-duplicate paragraphs before removing them."
        Case "frmCleanupSuiteLauncher.lblDescTableClean": ControlCaptionText = "Clean table structure, padding, borders, direct formatting, or convert simple tables."
        Case "frmCleanupSuiteLauncher.lblDescBreakNorm": ControlCaptionText = "Collapse or convert section and page breaks from assembled documents."
        Case "frmCleanupSuiteLauncher.lblDescHeaderFooter": ControlCaptionText = "Standardize or clear headers and footers across document sections."
        Case "frmCleanupSuiteLauncher.lblDescDocTrim": ControlCaptionText = "Remove trailing empty paragraphs at the end of the document."
        Case "frmCleanupSuiteLauncher.lblDescFontNorm": ControlCaptionText = "Reset direct font face, size, color, bold, and italic overrides."
        Case "frmCleanupSuiteLauncher.lblDescFormatStrip": ControlCaptionText = "Strip direct character and paragraph formatting while preserving styles."
        Case "frmCleanupSuiteLauncher.lblDescStyleClean": ControlCaptionText = "Remove unused custom styles or remap common style variants."
        Case "frmCleanupSuiteLauncher.lblDescHyperlink": ControlCaptionText = "Keep links but Hide styling" & vbCrLf & "Keep text but remove links"
        Case "frmCleanupSuiteLauncher.lblDescMetadata": ControlCaptionText = "Advanced metadata inspection, editing, and privacy management."
        Case "frmCleanupSuiteLauncher.lblDescFinalReview": ControlCaptionText = "Remove review comments and accept tracked changes when markup is finalized."
        Case "frmCleanupSuiteLauncher.lblDescFootnote": ControlCaptionText = "Remove footnotes or endnotes, optionally keeping note text inline."
        Case "frmCleanupSuiteLauncher.lblDescObjectRemover": ControlCaptionText = "Remove embedded clutter such as images, text boxes, controls, hidden text, or tables."
        Case "frmCleanupSuiteLauncher.lblRiskUnicode": ControlCaptionText = "Safe cleanup"
        Case "frmCleanupSuiteLauncher.lblRiskPunctuation": ControlCaptionText = "Safe cleanup"
        Case "frmCleanupSuiteLauncher.lblRiskSpacing": ControlCaptionText = "Safe cleanup"
        Case "frmCleanupSuiteLauncher.lblRiskCapitalization": ControlCaptionText = "Text caution"
        Case "frmCleanupSuiteLauncher.lblRiskList": ControlCaptionText = "Structure change"
        Case "frmCleanupSuiteLauncher.lblRiskParagraph": ControlCaptionText = "Structure change"
        Case "frmCleanupSuiteLauncher.lblRiskSoftReturn": ControlCaptionText = "Structure change"
        Case "frmCleanupSuiteLauncher.lblRiskDuplicate": ControlCaptionText = "Inspect First"
        Case "frmCleanupSuiteLauncher.lblRiskTableClean": ControlCaptionText = "Structure change"
        Case "frmCleanupSuiteLauncher.lblRiskHeaderFooter": ControlCaptionText = "Structure change"
        Case "frmCleanupSuiteLauncher.lblRiskFootnote": ControlCaptionText = "Removes content"
        Case "frmCleanupSuiteLauncher.lblRiskObjectRemover": ControlCaptionText = "Removes content"
        Case "frmCleanupSuiteLauncher.lblRiskFontNorm": ControlCaptionText = "Formatting change"
        Case "frmCleanupSuiteLauncher.lblRiskFormatStrip": ControlCaptionText = "Formatting change"
        Case "frmCleanupSuiteLauncher.lblRiskMetadata": ControlCaptionText = "Privacy cleanup"
        Case "frmCleanupSuiteLauncher.lblRiskFinalReview": ControlCaptionText = "Finalizes review"
        Case "frmCleanupSuiteLauncher.lblRiskHyperlink": ControlCaptionText = "Removes content"
        Case "frmCleanupSuiteLauncher.lblRiskStyleClean": ControlCaptionText = "Alpha tool"
        Case "frmCleanupSuiteLauncher.lblRiskBreakNorm": ControlCaptionText = "Alpha tool"
        Case "frmCleanupSuiteLauncher.lblRiskDocTrim": ControlCaptionText = "Safe cleanup"
        Case "frmDocumentTrim.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmDocumentTrim.optScopeDocument": ControlCaptionText = "Entire document (always)"
        Case "frmDuplicateDetector.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmDuplicateDetector.lblFuzzyWarning": ControlCaptionText = "Fuzzy match: groups paragraphs sharing the chosen percentage of words.  Slower on large documents."
        Case "frmDuplicateDetector.optFuzzyLoose": ControlCaptionText = "Loose  (50% word overlap)"
        Case "frmDuplicateDetector.optFuzzyMedium": ControlCaptionText = "Medium  (70% word overlap)"
        Case "frmDuplicateDetector.optFuzzyStrict": ControlCaptionText = "Strict  (90% word overlap)"
        Case "frmDuplicateDetector.optMatchExact": ControlCaptionText = "Exact match"
        Case "frmDuplicateDetector.optMatchFuzzy": ControlCaptionText = "Fuzzy match (word overlap)"
        Case "frmDuplicateDetector.optMatchNormalized": ControlCaptionText = "Normalized match (ignore punctuation and spaces)"
        Case "frmDuplicateDetector.chkIncludeEmptyParagraphs": ControlCaptionText = "Include empty paragraphs"
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
        Case "frmHyperlinkRemover.optHideLinks": ControlCaptionText = "Hide hyperlink styling, keep links"
        Case "frmHyperlinkRemover.optRemoveLinks": ControlCaptionText = "Remove hyperlinks, keep visible text"
        Case "frmHyperlinkRemover.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmHyperlinkRemover.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmListCleanup.chkFixIndent": ControlCaptionText = "Fix list indentation"
        Case "frmListCleanup.chkHyphenToBullets": ControlCaptionText = "Convert hyphen-dash lines to bullets"
        Case "frmListCleanup.chkNormalizeBullets": ControlCaptionText = "Convert manual hyphens / asterisks to List Bullet style"
        Case "frmListCleanup.chkNormalizeNumbering": ControlCaptionText = "Convert manual numbers to List Number style"
        Case "frmListCleanup.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmListCleanup.optAll": ControlCaptionText = "All list fixes"
        Case "frmListCleanup.optBullets": ControlCaptionText = "Convert manual bullet lists"
        Case "frmListCleanup.optCustom": ControlCaptionText = "Custom"
        Case "frmListCleanup.optIndent": ControlCaptionText = "Fix list indentation"
        Case "frmListCleanup.optNumbering": ControlCaptionText = "Convert manual numbered lists"
        Case "frmListCleanup.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmListCleanup.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmFinalReview.chkComments": ControlCaptionText = "Remove all comments"
        Case "frmFinalReview.cmdAuthorCleanupMinimal": ControlCaptionText = "Author Cleanup - Minimal"
        Case "frmFinalReview.cmdAuthorCleanupNormal": ControlCaptionText = "Author Cleanup - Normal"
        Case "frmFinalReview.cmdAuthorCleanupBroad": ControlCaptionText = "Author Cleanup - Broad"
        Case "frmFinalReview.chkPreviewOnly": ControlCaptionText = "Preview only (highlight, do not change)"
        Case "frmFinalReview.chkRevisions": ControlCaptionText = "Accept tracked changes and clear revision marks"
        Case "frmFinalReview.optScopeDocument": ControlCaptionText = "Entire document (always)"
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
        Case "frmParagraphCleanup.optCustom": ControlCaptionText = "Custom"
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
        Case "frmPunctuationCleanup.optCustom": ControlCaptionText = "Custom"
        Case "frmPunctuationCleanup.optDashes": ControlCaptionText = "Dashes only"
        Case "frmPunctuationCleanup.optEllipses": ControlCaptionText = "Ellipses only"
        Case "frmPunctuationCleanup.optQuotes": ControlCaptionText = "Quotes only"
        Case "frmPunctuationCleanup.optScopeDocument": ControlCaptionText = "Entire document"
        Case "frmPunctuationCleanup.optScopeSelection": ControlCaptionText = "Selected text only"
        Case "frmPreviewActions.lblTitle": ControlCaptionText = ""
        Case "frmPreviewActions.lblSummary": ControlCaptionText = ""
        Case "frmPreviewActions.lblHint": ControlCaptionText = "Tool Name"
        Case "frmPreviewActions.lblButtonMeasure": ControlCaptionText = ""
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
        Case "frmSpacingCleanup.optCustom": ControlCaptionText = "Custom"
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
        Case "frmUnicodeCleanup.optCustom": ControlCaptionText = "Custom"
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
        Case "cmdRun": cap = "Apply"
        Case "cmdPreview": cap = "Preview"
        Case "cmdApply": cap = "Apply"
        Case "cmdClear": cap = "Reconfigure"
        Case "cmdPreviousChange": cap = "Change" & vbCrLf & ChrW$(&H25C0) & " Previous"
        Case "cmdPreviousPage": cap = "Page" & vbCrLf & ChrW$(&H25C0) & " Previous"
        Case "cmdNextPage": cap = "Page" & vbCrLf & "Next " & ChrW$(&H25B6)
        Case "cmdNextChange": cap = "Change" & vbCrLf & "Next " & ChrW$(&H25B6)
        Case "cmdReset": cap = "Reset"
        Case "cmdResetAll": cap = "Reset All"
        Case "cmdUserManual": cap = "User Manual"
        Case "cmdCheckUpdates": cap = "Check for Updates"
        Case "cmdSelectAll": cap = "Select All"
        Case "cmdDeselectAll": cap = "Deselect All"
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
        Case "cmdDocTrim": cap = "Trim End of Doc"
        Case "cmdFormatStrip": cap = "Formatting Cleaner"
        Case "cmdHyperlink": cap = "Hyperlinks"
        Case "cmdSoftReturn": cap = "Soft Returns"
        Case "cmdMetadata": cap = "MetaDataSuite"
        Case "cmdAuthorCleanupMinimal": cap = "Author Cleanup - Minimal"
        Case "cmdAuthorCleanupNormal": cap = "Author Cleanup - Normal"
        Case "cmdAuthorCleanupBroad": cap = "Author Cleanup - Broad"
        Case "cmdClearSharingProperties": cap = "Clear Sharing Properties"
        Case "cmdFinalReview": cap = "Final Review"
        Case "cmdStyleClean": cap = "Styles"
        Case "cmdFootnote": cap = "Footnote / Endnote"
        Case "cmdHeaderFooter": cap = "Header / Footer"
        Case "cmdObjectRemover": cap = "Objects"
        Case Else: cap = Replace(nm, "cmd", "")
    End Select
    designer.Controls(nm).Caption = cap
End Sub

Private Function FriendlyFormCaption(formName As String) As String
    Select Case formName
        Case "frmCleanupSuiteLauncher": FriendlyFormCaption = "Cleanup Suite"
        Case "frmPreviewActions": FriendlyFormCaption = "Preview Actions"
        Case "frmPunctuationCleanup": FriendlyFormCaption = "Punctuation Normalizer"
        Case "frmUnicodeCleanup": FriendlyFormCaption = "Invisible Unicode Cleaner"
        Case "frmSpacingCleanup": FriendlyFormCaption = "Spacing Fixer"
        Case "frmCapitalizationCleanup": FriendlyFormCaption = "Capitalization Fixer"
        Case "frmListCleanup": FriendlyFormCaption = "List Normalizer"
        Case "frmParagraphCleanup": FriendlyFormCaption = "Paragraph Structure Fixer"
        Case "frmDuplicateDetector": FriendlyFormCaption = "Duplicate Paragraph Remover"
        Case "frmFontNormalizer": FriendlyFormCaption = "Font Normalizer"
        Case "frmTableCleaner": FriendlyFormCaption = "Table Cleaner"
        Case "frmBreakNormalizer": FriendlyFormCaption = "Break Normalizer"
        Case "frmDocumentTrim": FriendlyFormCaption = "Document Trim"
        Case "frmFormattingStripper": FriendlyFormCaption = "Formatting Stripper"
        Case "frmHyperlinkRemover": FriendlyFormCaption = "Hyperlink Cleaner"
        Case "frmSoftReturnConverter": FriendlyFormCaption = "Soft Return Converter"
        Case "frmMetaDataSuite": FriendlyFormCaption = "MetaDataSuite"
        Case "frmSafeMetadataEditor": FriendlyFormCaption = "Safe Metadata Editor"
        Case "frmFinalReview": FriendlyFormCaption = "Final Review"
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
        Case "cmdRun": ButtonTipText = "Apply this cleanup tool"
        Case "cmdPreview": ButtonTipText = "Preview with the current settings"
        Case "cmdApply": ButtonTipText = "Apply this preview using the same settings"
        Case "cmdClear": ButtonTipText = "Return to this tool without resetting its controls"
        Case "cmdPreviousChange": ButtonTipText = "Skip to the previous page that contains a preview change"
        Case "cmdPreviousPage": ButtonTipText = "Turn to the previous preview page"
        Case "cmdNextPage": ButtonTipText = "Turn to the next preview page"
        Case "cmdNextChange": ButtonTipText = "Skip to the next page that contains a preview change"
        Case "cmdReset": ButtonTipText = "Reset this tool to its defaults"
        Case "cmdResetAll": ButtonTipText = "Reset all tools and global settings to defaults"
        Case "cmdUserManual": ButtonTipText = "Open the PDF User Manual"
        Case "cmdCheckUpdates": ButtonTipText = "Check for a newer CleanupSuite release"
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
        Case "cmdHyperlink": ButtonTipText = "Clean or hide hyperlinks"
        Case "cmdSoftReturn": ButtonTipText = "Convert soft returns and paragraph breaks"
        Case "cmdMetadata": ButtonTipText = "Open the advanced metadata inspection and management dashboard"
        Case "cmdAuthorCleanupMinimal": ButtonTipText = "Clear the most common author-identifying metadata fields"
        Case "cmdAuthorCleanupNormal": ButtonTipText = "Clear a broader set of author-identifying metadata fields"
        Case "cmdAuthorCleanupBroad": ButtonTipText = "Apply the broadest author cleanup; treat saved results as effectively final"
        Case "cmdClearSharingProperties": ButtonTipText = "Clear document info fields before sharing"
        Case "cmdFinalReview": ButtonTipText = "Remove comments and finalize tracked changes"
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
                                    "lblCatText", "cmdHelpUnicode", "cmdUnicode", "lblDescUnicode", "lblRiskUnicode", "cmdHelpPunct", "cmdPunctuation", "lblDescPunctuation", "lblRiskPunctuation", "cmdHelpSpacing", "cmdSpacing", "lblDescSpacing", "lblRiskSpacing", "cmdHelpCap", "cmdCapitalization", "lblDescCapitalization", "lblRiskCapitalization", _
                                    "lblCatPara", "cmdHelpList", "cmdList", "lblDescList", "lblRiskList", "cmdHelpPara", "cmdParagraph", "lblDescParagraph", "lblRiskParagraph", "cmdHelpSoftReturn", "cmdSoftReturn", "lblDescSoftReturn", "lblRiskSoftReturn", "cmdHelpTrim", "cmdDocTrim", "lblDescDocTrim", "lblRiskDocTrim", _
                                    "lblCatReview", "cmdHelpDuplicate", "cmdDuplicate", "lblDescDuplicate", "lblRiskDuplicate", _
                                    "lblCatLayout", "cmdHelpTable", "cmdTableClean", "lblDescTableClean", "lblRiskTableClean", "cmdHelpHeaderFooter", "cmdHeaderFooter", "lblDescHeaderFooter", "lblRiskHeaderFooter", "cmdHelpFootnote", "cmdFootnote", "lblDescFootnote", "lblRiskFootnote", "cmdHelpObject", "cmdObjectRemover", "lblDescObjectRemover", "lblRiskObjectRemover", _
                                    "lblCatFormat", "cmdHelpFont", "cmdFontNorm", "lblDescFontNorm", "lblRiskFontNorm", "cmdHelpFormat", "cmdFormatStrip", "lblDescFormatStrip", "lblRiskFormatStrip", _
                                    "lblCatShare", "cmdHelpMetadata", "cmdMetadata", "lblDescMetadata", "lblRiskMetadata", "cmdHelpFinalReview", "cmdFinalReview", "lblDescFinalReview", "lblRiskFinalReview", "cmdHelpHyperlink", "cmdHyperlink", "lblDescHyperlink", "lblRiskHyperlink", _
                                    "lblCatAlpha", "cmdHelpStyle", "cmdStyleClean", "lblDescStyleClean", "lblRiskStyleClean", "cmdHelpBreak", "cmdBreakNorm", "lblDescBreakNorm", "lblRiskBreakNorm", _
                                    "chkReturnToMainAfterApply", "chkReturnToMainAfterClose", "chkAutoSave", "lblAutoSaveWarning", "chkUpdateChecks", "cmdCheckUpdates", "cmdResetAll", "cmdUserManual", "lblResetAllArrow", "lblResetAllCalloutBox", "lblResetAllCalloutText")
        Case "frmPreviewActions"
            ControlsForForm = Array("lblTitle", "lblSummary", "lblHint", "lblButtonMeasure", _
                                    "cmdApply", "cmdPreview", "cmdClear", "cmdPreviousChange", "cmdPreviousPage", "cmdNextPage", "cmdNextChange")
        Case "frmPunctuationCleanup"
            ControlsForForm = Array("optAll", "optQuotes", "optDashes", "optEllipses", "optCustom", _
                                    "cmdRiskPlacementPunctuationAll", "cmdRiskPlacementPunctuationQuotes", "cmdRiskPlacementPunctuationDashes", "cmdRiskPlacementPunctuationEllipses", "cmdRiskPlacementPunctuationCustom", "fraCustom", _
                                    "chkCurlyDouble", "chkCurlySingle", "chkEmDash", "chkEnDash", "chkEllipses", _
                                    "chkPreviewOnly", "cmdSelectAll", "cmdDeselectAll", "cmdPreview", "cmdRun", "cmdReset", "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmUnicodeCleanup"
            ControlsForForm = Array("optAll", "optNBSP", "optZeroWidth", "optCustom", _
                                    "cmdRiskPlacementUnicodeAll", "cmdRiskPlacementUnicodeNBSP", "cmdRiskPlacementUnicodeZeroWidth", "cmdRiskPlacementUnicodeCustom", "fraCustom", _
                                    "chkNBSP", "chkZWSP", "chkZWNJ", "chkZWJ", "chkBOM", "chkSoftHyphen", "chkNBHyphen", _
                                    "chkPreviewOnly", "cmdSelectAll", "cmdDeselectAll", "cmdPreview", "cmdRun", "cmdReset", "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmSpacingCleanup"
            ControlsForForm = Array("optAll", "optDoubleSpaces", "optTrim", "optCustom", _
                                    "cmdRiskPlacementSpacingAll", "cmdRiskPlacementSpacingDouble", "cmdRiskPlacementSpacingTrim", "cmdRiskPlacementSpacingCustom", "fraCustom", _
                                    "chkDoubleSpaces", "chkTrimSpaces", "chkSpaceBeforePunct", "chkNormalizeAfterPunct", "chkExtraBlankLines", _
                                    "chkPreviewOnly", "cmdSelectAll", "cmdDeselectAll", "cmdPreview", "cmdRun", "cmdReset", "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmCapitalizationCleanup"
            ControlsForForm = Array("lblIntro", "optAll", "optCustom", _
                                    "cmdRiskPlacementCapitalizationRecommended", "cmdRiskPlacementCapitalizationCustom", "lblModeSummary", _
                                    "chkSentence", "chkTitle", "chkUpper", "chkLower", "chkHeadingParentheses", "chkSmartSentences", _
                                    "cmdEditExceptions", "chkPreviewOnly", "cmdSelectAll", "cmdDeselectAll", "cmdPreview", "cmdRun", "cmdReset", "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmListCleanup"
            ControlsForForm = Array("optAll", "optBullets", "optNumbering", "optIndent", "optCustom", _
                                    "cmdRiskPlacementListAll", "cmdRiskPlacementListBullets", "cmdRiskPlacementListNumbering", "cmdRiskPlacementListIndent", "cmdRiskPlacementListCustom", "fraCustom", _
                                    "cmdRiskPlacementListNormalizeBullets", "cmdRiskPlacementListNormalizeNumbering", "cmdRiskPlacementListFixIndent", "cmdRiskPlacementListHyphenBullets", _
                                    "chkNormalizeBullets", "chkNormalizeNumbering", "chkFixIndent", "chkHyphenToBullets", _
                                    "chkPreviewOnly", "cmdSelectAll", "cmdDeselectAll", "cmdPreview", "cmdRun", "cmdReset", "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmParagraphCleanup"
            ControlsForForm = Array("optAll", "optRemoveEmpty", "optNormalizeSpacing", "optCustom", _
                                    "cmdRiskPlacementParagraphAll", "cmdRiskPlacementParagraphRemoveEmpty", "cmdRiskPlacementParagraphSpacing", "cmdRiskPlacementParagraphCustom", "fraCustom", _
                                    "cmdRiskPlacementParagraphChkRemoveEmpty", "cmdRiskPlacementParagraphChkBreaks", "cmdRiskPlacementParagraphChkSpacing", "cmdRiskPlacementParagraphChkIndent", _
                                    "chkRemoveEmpty", "chkCollapseBreaks", "chkNormalizeParaSpacing", "chkFixIndent", _
                                    "chkPreviewOnly", "cmdSelectAll", "cmdDeselectAll", "cmdPreview", "cmdRun", "cmdReset", "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmDuplicateDetector"
            ControlsForForm = Array("optMatchExact", "optMatchNormalized", "optMatchFuzzy", _
                                    "cmdRiskPlacementDuplicateExact", "cmdRiskPlacementDuplicateNormalized", "cmdRiskPlacementDuplicateFuzzy", _
                                    "chkIncludeEmptyParagraphs", "cmdRiskPlacementDuplicateEmpty", _
                                    "fraThreshold", "optFuzzyLoose", "optFuzzyMedium", "optFuzzyStrict", "lblFuzzyWarning", _
                                    "cmdRiskPlacementDuplicateFuzzyLoose", "cmdRiskPlacementDuplicateFuzzyMedium", "cmdRiskPlacementDuplicateFuzzyStrict", _
                                    "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmFontNormalizer"
            ControlsForForm = Array("chkFontFace", "chkFontSize", "chkBold", "chkItalic", "chkFontColor", _
                                    "cmdRiskPlacementFontFace", "cmdRiskPlacementFontSize", "cmdRiskPlacementFontBold", "cmdRiskPlacementFontItalic", "cmdRiskPlacementFontColor", _
                                    "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmTableCleaner"
            ControlsForForm = Array("chkRemoveEmptyRows", "chkRemoveEmptyCols", "chkNormalizePadding", _
                                    "chkStripDirectFormat", "chkNormalizeBorders", "chkRemoveBorders", "chkConvertToText", _
                                    "cmdRiskPlacementTableRows", "cmdRiskPlacementTableCols", "cmdRiskPlacementTablePadding", "cmdRiskPlacementTableStripFormat", "cmdRiskPlacementTableBorders", "cmdRiskPlacementTableRemoveBorders", "cmdRiskPlacementTableConvertToText", _
                                    "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmBreakNormalizer"
            ControlsForForm = Array("chkCollapseSectionBreaks", "chkCollapsePageBreaks", _
                                    "chkConvertSectionBreaks", "fraConvertTo", _
                                    "optConvertNextPage", "optConvertContinuous", "optConvertEvenPage", "optConvertOddPage", _
                                    "cmdRiskPlacementBreakSections", "cmdRiskPlacementBreakPages", "cmdRiskPlacementBreakConvert", "cmdRiskPlacementBreakNextPage", "cmdRiskPlacementBreakContinuous", "cmdRiskPlacementBreakEvenPage", "cmdRiskPlacementBreakOddPage", _
                                    "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmDocumentTrim"
            ControlsForForm = Array("lblRiskPlacementDocumentTrimSummary", "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmFormattingStripper"
            ControlsForForm = Array("chkResetChar", "chkResetPara", "fraEmphasis", _
                                    "optEmphQuick", "optEmphThorough", "optEmphStrip", "lblSpeedWarning", _
                                    "cmdRiskPlacementFormattingChar", "cmdRiskPlacementFormattingPara", "cmdRiskPlacementFormattingQuick", "cmdRiskPlacementFormattingThorough", "cmdRiskPlacementFormattingStrip", _
                                    "chkPreserveHighlight", "chkPreserveDropCaps", "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", _
                                    "cmdRiskPlacementFormattingHighlight", "cmdRiskPlacementFormattingDropCaps", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmHyperlinkRemover"
            ControlsForForm = Array("optHideLinks", "optRemoveLinks", "cmdRiskPlacementHyperlinkHide", "cmdRiskPlacementHyperlinkRemove", _
                                    "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmSoftReturnConverter"
            ControlsForForm = Array("optSoftToPara", "optParaToSoft", "cmdRiskPlacementSoftReturnToPara", "cmdRiskPlacementSoftReturnToSoft", "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmMetaDataSuite"
            ControlsForForm = Array("lblTitle", "lblTargetTitle", "lblTargetPath", "lblOverallIndicator", "lblOverallStatus", "lblHint", _
                                    "cmdRefreshCurrent", "cmdAuditAnother", "cmdOpenWordReport", "cmdSafeEditCurrent", "cmdGroupEditCurrent", "cmdClearSharingProperties", "cmdClose", "cmdExpandAll", "chkShowSuiteMetadata", _
                                    "lblCoreHeader", "chkCoreDetails", "lblCoreIndicator", "lblCoreSummary", "txtCoreDetails", _
                                    "lblAppHeader", "chkAppDetails", "lblAppIndicator", "lblAppSummary", "txtAppDetails", _
                                    "lblCustomHeader", "chkCustomDetails", "lblCustomIndicator", "lblCustomSummary", "txtCustomDetails", _
                                    "lblPackageHeader", "chkPackageDetails", "lblPackageIndicator", "lblPackageSummary", "txtPackageDetails", _
                                    "lblMetadataHeader", "chkMetadataDetails", "lblMetadataIndicator", "lblMetadataSummary", "txtMetadataDetails", "cmdMetadataInvestigate", _
                                    "lblCryptoHeader", "chkCryptoDetails", "lblCryptoIndicator", "lblCryptoSummary", "txtCryptoDetails", "cmdCryptoInvestigate")
        Case "frmSafeMetadataEditor"
            ControlsForForm = Array("lblEditorTitle", "lblDocumentPath", "lblSummaryHeader", _
                                    "lblTitle", "txtTitle", "lblSubject", "txtSubject", "lblAuthor", "txtAuthor", _
                                    "lblKeywords", "txtKeywords", "lblComments", "txtComments", "lblCategory", "txtCategory", _
                                    "lblCompany", "txtCompany", "lblManager", "txtManager", _
                                    "lblStatusHeader", "lblLastAuthor", "txtLastAuthor", "lblContentStatus", "txtContentStatus", _
                                    "lblContentType", "txtContentType", "lblLanguage", "txtLanguage", _
                                    "lblCustomHeader", "lblCustomHint", "lblCustomNameHeader", "lblCustomTypeHeader", "lblCustomValueHeader", _
                                    "lstCustomProperties", "lblCustomName", "txtCustomName", "lblCustomType", "cboCustomType", _
                                    "lblCustomValue", "txtCustomValue", "cmdCustomAddUpdate", "cmdCustomDelete", _
                                    "lblAuditOnlyNote", "cmdSave", "cmdCancel")
        Case "frmFinalReview"
            ControlsForForm = Array("chkComments", "chkRevisions", "cmdAuthorCleanupMinimal", "cmdAuthorCleanupNormal", "cmdAuthorCleanupBroad", _
                                    "lblRiskPlacementFinalReviewSummary", _
                                    "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmStyleCleanup"
            ControlsForForm = Array("chkRemoveUnused", "chkRemapVariants", _
                                    "cmdRiskPlacementStyleRemoveUnused", "cmdRiskPlacementStyleRemap", _
                                    "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmFootnoteRemover"
            ControlsForForm = Array("chkFootnotes", "chkEndnotes", "chkKeepTextInline", _
                                    "cmdRiskPlacementFootnoteFootnotes", "cmdRiskPlacementFootnoteEndnotes", "cmdRiskPlacementFootnoteKeepText", _
                                    "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmHeaderFooterStandardizer"
            ControlsForForm = Array("optStandardize", "optClearAll", "fraOptions", _
                                    "cmdRiskPlacementHeaderFooterStandardize", "cmdRiskPlacementHeaderFooterClear", _
                                    "chkHeaders", "chkFooters", "chkFont", "chkSpacing", "chkBreakLinks", "chkAlignment", _
                                    "cmdRiskPlacementHeaderFooterHeaders", "cmdRiskPlacementHeaderFooterFooters", "cmdRiskPlacementHeaderFooterFont", "cmdRiskPlacementHeaderFooterSpacing", "cmdRiskPlacementHeaderFooterBreakLinks", "cmdRiskPlacementHeaderFooterAlignment", _
                                    "fraAlign", "optAlignLeft", "optAlignRight", "optAlignCenter", _
                                    "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case "frmObjectRemover"
            ControlsForForm = Array("chkPictures", "chkTextBoxes", "chkFrames", "chkHorizontalLines", _
                                    "chkHtmlControls", "chkHiddenText", "chkTables", _
                                    "cmdRiskPlacementObjectPictures", "cmdRiskPlacementObjectTextBoxes", "cmdRiskPlacementObjectFrames", "cmdRiskPlacementObjectLines", "cmdRiskPlacementObjectControls", "cmdRiskPlacementObjectHiddenText", "cmdRiskPlacementObjectTables", _
                                    "chkPreviewOnly", "cmdPreview", "cmdRun", "cmdReset", _
                                    "fraScopeSelection", "optScopeDocument", "optScopeSelection")
        Case Else
            ControlsForForm = Array()
    End Select
End Function

' ---------------------------
' Helper: choose MSForms control ProgID by name prefix
' ---------------------------
Private Function ControlTypeByName(ctrlName As String) As String
    If Left$(ctrlName, 7) = "lblRisk" Then
        ControlTypeByName = "Forms.CommandButton.1"
    ElseIf Left$(ctrlName, 3) = "opt" Then
        ControlTypeByName = "Forms.OptionButton.1"
    ElseIf Left$(ctrlName, 3) = "chk" Then
        ControlTypeByName = "Forms.CheckBox.1"
    ElseIf Left$(ctrlName, 3) = "cmd" Then
        ControlTypeByName = "Forms.CommandButton.1"
    ElseIf Left$(ctrlName, 3) = "fra" Then
        ControlTypeByName = "Forms.Frame.1"
    ElseIf Left$(ctrlName, 3) = "lst" Then
        ControlTypeByName = "Forms.ListBox.1"
    ElseIf Left$(ctrlName, 3) = "cbo" Then
        ControlTypeByName = "Forms.ComboBox.1"
    ElseIf Left$(ctrlName, 3) = "lbl" Then
        ControlTypeByName = "Forms.Label.1"
    Else
        ControlTypeByName = "Forms.TextBox.1"
    End If
End Function
