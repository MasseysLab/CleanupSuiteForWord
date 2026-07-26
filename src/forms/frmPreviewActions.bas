Option Explicit

Private mSourceForm As Object
Private mToolName As String
Private mPreviewOn As Boolean
Private mSummaryRows As Collection
Private mPreviewDocument As Document
Private mPreviewWindow As Window
Private mApplyButtonBackColor As Long
Private mApplyButtonForeColor As Long
Private mSummaryTop As Single
Private mResumeModelessAfterModal As Boolean
Private mOriginalBrowseTarget As Long
Private mBrowseTargetCaptured As Boolean
Private mEditablePreviewRestored As Boolean
Private mNavigationDetached As Boolean

Private Const FORM_W As Single = 283
Private Const MIN_FORM_H As Single = 90
Private Const CONTENT_W As Single = 257
Private Const M As Single = 8
Private Const GAP As Single = 5
Private Const BTN_W As Single = (CONTENT_W - (2 * GAP)) / 3
Private Const SUMMARY_BOTTOM_GAP As Single = 6
Private Const ACTION_TOP As Single = 17

Public Sub Configure(sourceForm As Object, toolName As String, summaryText As String)
    Dim rows As Collection
    Set rows = NewPreviewSummaryRows()
    AddPreviewSummaryRow rows, summaryText, ""
    ConfigureSummary sourceForm, toolName, rows
End Sub

Public Sub ConfigureSummary(sourceForm As Object, toolName As String, rows As Collection)
    Set mSourceForm = sourceForm
    mToolName = toolName
    Set mSummaryRows = rows
    mPreviewOn = True
    mEditablePreviewRestored = False
    mResumeModelessAfterModal = False
    mNavigationDetached = False
    Set mPreviewDocument = ActiveDocument
    Set mPreviewWindow = ActiveWindow
    On Error Resume Next
    mOriginalBrowseTarget = Application.Browser.Target
    mBrowseTargetCaptured = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    BeginPreviewViewSession
    Me.Caption = ""
    ApplyMSFormTitleStrategy Me, True
    lblTitle.Caption = ""
    lblHint.Caption = toolName
    cmdApply.Caption = "Apply"
    cmdPreview.Caption = "Preview is ON"
    cmdPreview.Enabled = True
    cmdPreview.Visible = True
    cmdClear.Caption = "Reconfigure"
    cmdPreviousChange.Caption = "Change" & vbCrLf & ChrW$(&H25C0) & " Previous"
    cmdPreviousPage.Caption = "Page" & vbCrLf & ChrW$(&H25C0) & " Previous"
    cmdNextPage.Caption = "Page" & vbCrLf & "Next " & ChrW$(&H25B6)
    cmdNextChange.Caption = "Change" & vbCrLf & "Next " & ChrW$(&H25B6)
    cmdPreviousChange.Enabled = True
    cmdPreviousPage.Enabled = True
    cmdNextPage.Enabled = True
    cmdNextChange.Enabled = True
    cmdPreviousChange.Visible = True
    cmdPreviousPage.Visible = True
    cmdNextPage.Visible = True
    cmdNextChange.Visible = True
    LayoutPanel
    UpdateNavigationButtonStates
    mApplyButtonBackColor = cmdApply.BackColor
    mApplyButtonForeColor = cmdApply.ForeColor
End Sub

Private Sub LayoutPanel()
    Dim rows As Collection
    Dim desiredInsideH As Single
    Dim buttonHeight As Single
    Dim navigationTop As Single
    Dim navigationWidth As Single
    Set rows = mSummaryRows
    Me.Caption = ""
    Me.Width = FORM_W
    buttonHeight = MeasuredPreviewButtonHeight()
    navigationTop = ACTION_TOP + buttonHeight + GAP
    navigationWidth = (CONTENT_W - (3 * GAP)) / 4
    If mPreviewOn Then
        mSummaryTop = navigationTop + buttonHeight + GAP
    Else
        mSummaryTop = ACTION_TOP + buttonHeight + GAP
    End If
    desiredInsideH = mSummaryTop + SummaryTableHeight(mSummaryRows) + SUMMARY_BOTTOM_GAP
    Me.Height = MaxSingle(MIN_FORM_H, desiredInsideH)
    If Me.InsideHeight < desiredInsideH Then
        Me.Height = Me.Height + (desiredInsideH - Me.InsideHeight)
    End If
    If Me.InsideHeight < desiredInsideH Then
        Me.Height = Me.Height + (desiredInsideH - Me.InsideHeight) + 2
    End If
    lblTitle.Caption = ""
    lblTitle.Visible = False
    StyleActionBar
    lblTitle.Move 0, 0, 0, 0
    lblHint.Move M, 2, CONTENT_W, 12
    cmdPreview.Move M, ACTION_TOP, BTN_W, buttonHeight
    cmdClear.Move M + BTN_W + GAP, ACTION_TOP, BTN_W, buttonHeight
    cmdApply.Move M + (2 * (BTN_W + GAP)), ACTION_TOP, BTN_W, buttonHeight
    cmdPreviousChange.Move M, navigationTop, navigationWidth, buttonHeight
    cmdPreviousPage.Move M + navigationWidth + GAP, navigationTop, navigationWidth, buttonHeight
    cmdNextPage.Move M + (2 * (navigationWidth + GAP)), navigationTop, navigationWidth, buttonHeight
    cmdNextChange.Move M + (3 * (navigationWidth + GAP)), navigationTop, navigationWidth, buttonHeight
    cmdPreviousChange.Visible = mPreviewOn
    cmdPreviousPage.Visible = mPreviewOn
    cmdNextPage.Visible = mPreviewOn
    cmdNextChange.Visible = mPreviewOn
    FitPreviewButtonCaption cmdPreview
    FitPreviewButtonCaption cmdClear
    FitPreviewButtonCaption cmdApply
    FitPreviewButtonCaption cmdPreviousChange
    FitPreviewButtonCaption cmdPreviousPage
    FitPreviewButtonCaption cmdNextPage
    FitPreviewButtonCaption cmdNextChange
    lblSummary.Move M, mSummaryTop, 0, 0
    lblSummary.Visible = False
    lblTitle.Visible = False
    lblTitle.WordWrap = False
    lblHint.WordWrap = False
    lblHint.TextAlign = 2
    RenderSummaryTable rows
End Sub

Private Function SummaryTableHeight(ByVal rows As Collection) As Single
    Const ROW_H As Single = 11
    Dim rowCount As Long
    If rows Is Nothing Then
        rowCount = 1
    Else
        rowCount = rows.Count + 1
    End If
    SummaryTableHeight = rowCount * ROW_H
    If SummaryTableHeight < 22 Then SummaryTableHeight = 22
End Function

Private Sub RenderSummaryTable(ByVal rows As Collection)
    Const ITEM_W As Single = 215
    Const COUNT_W As Single = 31
    Const ROW_H As Single = 11
    Const TABLE_W As Single = ITEM_W + COUNT_W

    Dim tableLeft As Single
    Dim y As Single
    Dim rowText As Variant
    Dim rowParts As Variant
    Dim rowIndex As Long
    Dim itemText As String
    Dim countText As String
    Dim isInactive As Boolean

    HideSummaryTableControls
    tableLeft = M + ((CONTENT_W - TABLE_W) / 2)
    y = mSummaryTop
    RenderSummaryCell "lblSummaryHeaderItem", "Item", tableLeft, y, ITEM_W, ROW_H, True
    RenderSummaryCell "lblSummaryHeaderCount", "#", tableLeft + ITEM_W, y, COUNT_W, ROW_H, True

    y = y + ROW_H
    rowIndex = 0
    If rows Is Nothing Then Exit Sub
    For Each rowText In rows
        rowParts = Split(CStr(rowText), vbTab)
        itemText = ""
        countText = ""
        isInactive = False
        If UBound(rowParts) >= 0 Then itemText = CStr(rowParts(0))
        If UBound(rowParts) >= 1 Then countText = CStr(rowParts(1))
        If UBound(rowParts) >= 2 Then isInactive = (CStr(rowParts(2)) = "inactive")
        RenderSummaryCell "lblSummaryItem" & CStr(rowIndex), itemText, tableLeft, y, ITEM_W, ROW_H, False, isInactive
        RenderSummaryCell "lblSummaryCount" & CStr(rowIndex), countText, tableLeft + ITEM_W, y, COUNT_W, ROW_H, False, isInactive
        y = y + ROW_H
        rowIndex = rowIndex + 1
    Next rowText
End Sub

Private Sub HideSummaryTableControls()
    On Error Resume Next
    Dim ctl As Object
    For Each ctl In Me.Controls
        If Left$(ctl.Name, 10) = "lblSummary" And ctl.Name <> "lblSummary" Then
            ctl.Visible = False
        End If
    Next ctl
    On Error GoTo 0
End Sub

Private Sub RenderSummaryCell(ByVal controlName As String, ByVal captionText As String, ByVal L As Single, ByVal T As Single, ByVal W As Single, ByVal H As Single, ByVal isHeader As Boolean, Optional ByVal isInactive As Boolean = False)
    Dim ctl As Object
    Set ctl = EnsureSummaryLabel(controlName)
    ctl.Caption = captionText
    ctl.Move L, T, W, H
    ctl.Visible = True
    ctl.BorderStyle = fmBorderStyleSingle
    ctl.SpecialEffect = fmSpecialEffectFlat
    ctl.BackStyle = fmBackStyleOpaque
    ctl.BackColor = RGB(255, 255, 255)
    ctl.ForeColor = RGB(0, 0, 0)
    If isInactive Then
        ctl.ForeColor = RGB(145, 145, 145)
        ctl.BackColor = RGB(245, 245, 245)
    End If
    If InStr(1, controlName, "Count", vbTextCompare) > 0 Then
        ctl.TextAlign = 2
    Else
        ctl.TextAlign = 1
    End If
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 7.5
    ctl.Font.Bold = isHeader
    ctl.WordWrap = False
End Sub

Private Function EnsureSummaryLabel(ByVal controlName As String) As Object
    On Error Resume Next
    Set EnsureSummaryLabel = Me.Controls(controlName)
    If EnsureSummaryLabel Is Nothing Then
        Set EnsureSummaryLabel = Me.Controls.Add("Forms.Label.1", controlName, True)
    End If
    On Error GoTo 0
End Function

Private Function MaxSingle(a As Single, b As Single) As Single
    If a > b Then
        MaxSingle = a
    Else
        MaxSingle = b
    End If
End Function

Private Sub StyleActionBar()
    On Error Resume Next
    Me.BackColor = RGB(248, 249, 251)
    Me.SpecialEffect = 0
    StyleLabel lblTitle, 11, False, RGB(0, 0, 0)
    StyleLabel lblHint, 8, True, RGB(0, 0, 0)
    StyleLabel lblSummary, 7.5, False, RGB(0, 0, 0)
    StyleButton cmdApply, True
    StyleButton cmdPreview, True
    StyleButton cmdClear, True
    StyleButton cmdPreviousChange, True
    StyleButton cmdPreviousPage, True
    StyleButton cmdNextPage, True
    StyleButton cmdNextChange, True
    lblButtonMeasure.Visible = False
End Sub

Private Sub StyleLabel(ctl As Object, fontSize As Single, isBold As Boolean, fontColor As Long)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = fontSize
    ctl.Font.Bold = isBold
    ctl.ForeColor = fontColor
    ctl.BackColor = RGB(248, 249, 251)
    ctl.AutoSize = False
End Sub

Private Sub StyleButton(ctl As Object, isBold As Boolean)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 8
    ctl.Font.Bold = isBold
    ctl.BackColor = RGB(255, 255, 255)
    ctl.ForeColor = RGB(0, 0, 0)
End Sub

Private Function MeasuredPreviewButtonHeight() As Single
    Dim captionValue As Variant
    Dim maximumTextHeight As Single
    Dim captions As Variant

    On Error Resume Next
    captions = Array(cmdPreview.Caption, cmdClear.Caption, cmdApply.Caption, _
                     cmdPreviousChange.Caption, cmdPreviousPage.Caption, _
                     cmdNextPage.Caption, cmdNextChange.Caption)
    With lblButtonMeasure
        .Visible = False
        .AutoSize = True
        .WordWrap = False
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        For Each captionValue In captions
            .Caption = CStr(captionValue)
            If .Height > maximumTextHeight Then maximumTextHeight = .Height
        Next captionValue
        .Caption = ""
        .AutoSize = False
    End With
    MeasuredPreviewButtonHeight = MaxSingle(32, maximumTextHeight + 8)
    On Error GoTo 0
End Function

Private Sub FitPreviewButtonCaption(ByVal buttonControl As Object)
    Dim measuredWidth As Single
    Dim fittedSize As Single

    On Error Resume Next
    fittedSize = 8
    With lblButtonMeasure
        .Visible = False
        .AutoSize = True
        .WordWrap = False
        .Font.Name = "Segoe UI"
        .Font.Bold = buttonControl.Font.Bold
        Do
            .Font.Size = fittedSize
            .Caption = CStr(buttonControl.Caption)
            measuredWidth = .Width
            If measuredWidth <= buttonControl.Width - 12 Then Exit Do
            fittedSize = fittedSize - 0.25
        Loop While fittedSize > 7
        .Caption = ""
        .AutoSize = False
    End With
    buttonControl.Font.Size = fittedSize
    On Error GoTo 0
End Sub

Public Function ConsumeModelessResumeRequest() As Boolean
    ConsumeModelessResumeRequest = mResumeModelessAfterModal
    mResumeModelessAfterModal = False
End Function

Public Sub ShowProgress(ByVal progressText As String)
    On Error Resume Next
    lblHint.Caption = progressText
    cmdApply.Caption = "Applying..."
    cmdApply.BackColor = RGB(221, 235, 247)
    cmdApply.ForeColor = RGB(31, 78, 121)
    cmdPreview.Enabled = False
    cmdClear.Enabled = False
    cmdApply.Enabled = False
    StyleNavigationButton cmdPreviousChange, True, False
    StyleNavigationButton cmdPreviousPage, False, False
    StyleNavigationButton cmdNextPage, False, False
    StyleNavigationButton cmdNextChange, True, False
    Me.Repaint
    On Error GoTo 0
End Sub

Public Sub ClearProgress()
    On Error Resume Next
    lblHint.Caption = mToolName
    cmdApply.Caption = "Apply"
    cmdApply.BackColor = mApplyButtonBackColor
    cmdApply.ForeColor = mApplyButtonForeColor
    cmdPreview.Enabled = True
    cmdClear.Enabled = True
    cmdApply.Enabled = True
    UpdateNavigationButtonStates
    Me.Repaint
    On Error GoTo 0
End Sub

Private Sub RestoreEditablePreviewDocument()
    If mEditablePreviewRestored Then Exit Sub
    On Error Resume Next
    RestoreCleanupSuiteTransientState
    RemovePreviewHighlighting mPreviewDocument
    RestorePreviewBrowseTarget
    mEditablePreviewRestored = True
    On Error GoTo 0
End Sub

Private Sub RestorePreviewBrowseTarget()
    On Error Resume Next
    If mBrowseTargetCaptured Then Application.Browser.Target = mOriginalBrowseTarget
    mBrowseTargetCaptured = False
    On Error GoTo 0
End Sub

Private Sub cmdPreviousChange_Click()
    NavigatePreviewChange False
End Sub

Private Sub cmdPreviousPage_Click()
    NavigatePreviewPage False
End Sub

Private Sub cmdNextPage_Click()
    NavigatePreviewPage True
End Sub

Private Sub cmdNextChange_Click()
    NavigatePreviewChange True
End Sub

Private Sub NavigatePreviewPage(ByVal moveForward As Boolean)
    Dim resumePanelHere As Boolean

    On Error GoTo NavigationErr
    If Not mPreviewOn Then Exit Sub
    If Not PreviewTargetIsActive() Then Exit Sub
    resumePanelHere = DetachPreviewPanelForNavigationIfNeeded()
    FocusPreviewDocumentWindow
    If Not TurnPreviewPage(moveForward) Then GoTo NavigationErr
    lblHint.Caption = mToolName
    UpdateNavigationButtonStates
    Me.Repaint
    GoTo NavigationExit
NavigationErr:
    lblHint.Caption = "Could not turn the preview page"
    Me.Repaint
NavigationExit:
    ResumePreviewPanelAfterNavigation resumePanelHere
End Sub

Private Sub NavigatePreviewChange(ByVal moveForward As Boolean)
    Dim resumePanelHere As Boolean

    On Error GoTo NavigationErr
    If Not mPreviewOn Then Exit Sub
    If Not PreviewTargetIsActive() Then Exit Sub
    resumePanelHere = DetachPreviewPanelForNavigationIfNeeded()
    FocusPreviewDocumentWindow
    If Not NavigatePreviewChangePage(mPreviewDocument, moveForward) Then GoTo NavigationExit
    lblHint.Caption = mToolName
    UpdateNavigationButtonStates
    Me.Repaint
    GoTo NavigationExit
NavigationErr:
    lblHint.Caption = "Could not turn to the preview change"
    Me.Repaint
NavigationExit:
    ResumePreviewPanelAfterNavigation resumePanelHere
End Sub

Private Function DetachPreviewPanelForNavigationIfNeeded() As Boolean
    RememberPreviewActionPanelPosition Me.Left, Me.Top
    If mNavigationDetached Then
        ' The panel is already modeless. Hide it for this page turn so Word's
        ' Reading View window owns the native left/right page key.
        Me.Hide
        DoEvents
        DetachPreviewPanelForNavigationIfNeeded = True
        Exit Function
    End If

    ' A modal VBA form can prevent Reading View from repainting a backward
    ' page turn even though Word moves Selection to the correct page. Hand the
    ' first navigation click back to the document window, then resume this
    ' same form modelessly. Reading View remains the editing safety boundary.
    mNavigationDetached = True
    mResumeModelessAfterModal = True
    Me.Hide
    DoEvents
End Function

Private Sub ResumePreviewPanelAfterNavigation(ByVal resumePanelHere As Boolean)
    If Not resumePanelHere Then Exit Sub
    On Error Resume Next
    Me.Show vbModeless
    On Error GoTo 0
End Sub

Private Sub FocusPreviewDocumentWindow()
    ' Modeless form clicks return keyboard focus to the form. Word can update
    ' Selection without repainting a backward Reading View page while that is
    ' true, so reactivate the captured Word window before each turn.
    On Error Resume Next
    mPreviewWindow.Activate
    DoEvents
    On Error GoTo 0
End Sub

Private Sub UpdateNavigationButtonStates()
    Dim targetIsActive As Boolean

    On Error Resume Next
    targetIsActive = (mPreviewOn And PreviewTargetIsActive())
    If Not targetIsActive Then
        StyleNavigationButton cmdPreviousChange, True, False
        StyleNavigationButton cmdPreviousPage, False, False
        StyleNavigationButton cmdNextPage, False, False
        StyleNavigationButton cmdNextChange, True, False
        Exit Sub
    End If

    StyleNavigationButton cmdPreviousChange, True, PreviewChangeNavigationAvailable(mPreviewDocument, False)
    StyleNavigationButton cmdPreviousPage, False, PreviewPageNavigationAvailable(mPreviewDocument, False)
    StyleNavigationButton cmdNextPage, False, PreviewPageNavigationAvailable(mPreviewDocument, True)
    StyleNavigationButton cmdNextChange, True, PreviewChangeNavigationAvailable(mPreviewDocument, True)
    On Error GoTo 0
End Sub

Private Sub StyleNavigationButton(ByVal buttonControl As Object, ByVal isChangeButton As Boolean, ByVal enabledState As Boolean)
    On Error Resume Next
    buttonControl.Enabled = enabledState
    buttonControl.WordWrap = True
    buttonControl.Font.Size = 7.5
    buttonControl.Font.Bold = True
    If Not enabledState Then
        buttonControl.BackColor = RGB(235, 235, 235)
        buttonControl.ForeColor = RGB(128, 128, 128)
    ElseIf isChangeButton Then
        buttonControl.BackColor = RGB(255, 243, 205)
        buttonControl.ForeColor = RGB(112, 78, 0)
    Else
        buttonControl.BackColor = RGB(221, 235, 247)
        buttonControl.ForeColor = RGB(31, 78, 121)
    End If
    On Error GoTo 0
End Sub

Private Sub cmdApply_Click()
    Dim errorDescription As String

    On Error GoTo ApplyErr
    If Not PreviewTargetIsActive() Then
        MsgBox "The preview document is no longer active. Return to that document and try again.", vbInformation, "Preview Actions"
        Exit Sub
    End If
    RestoreEditablePreviewDocument
    If Not DocumentHasVisibleContent() Then
        RemoveAllHighlighting ActiveDocument.Content
        MsgBox "This document appears to be blank. There is nothing to apply.", vbInformation, "Preview Actions"
        Me.Hide
        If Not mSourceForm Is Nothing Then mSourceForm.Hide
        MarkCleanupToolApplied
        Set gPreviewActionPanel = Nothing
        Unload Me
        Exit Sub
    End If
    ShowProgress mToolName & ": Applying changes"
    If Not mSourceForm Is Nothing Then
        CallByName mSourceForm, "RunAfterPreview", VbMethod
    End If
    ClearPreviewScopeRange
    Set gPreviewActionPanel = Nothing
    Unload Me
    Exit Sub
ApplyErr:
    errorDescription = Err.Description
    RestoreEditablePreviewDocument
    MsgBox "Could not apply the preview: " & errorDescription, vbExclamation, "Preview Actions"
End Sub

Private Sub cmdPreview_Click()
    Dim errorDescription As String
    Dim offRows As Collection

    On Error GoTo PreviewErr
    If Not mPreviewOn Then Exit Sub

    Set offRows = NewPreviewSummaryRows()
    RestoreEditablePreviewDocument
    mPreviewOn = False
    cmdPreview.Caption = "Preview is OFF"
    ' Re-previewing directly from this modeless form can crash Word on long
    ' documents. Reconfigure returns to the tool form and starts a fresh,
    ' stable preview call stack.
    cmdPreview.Enabled = False
    cmdPreview.Visible = True
    AddPreviewSummaryRow offRows, "Edit if needed. Choose Reconfigure to preview again.", ""
    Set mSummaryRows = offRows
    LayoutPanel
    mResumeModelessAfterModal = True
    RememberPreviewActionPanelPosition Me.Left, Me.Top
    Me.Hide
    Exit Sub
PreviewErr:
    errorDescription = Err.Description
    RestoreEditablePreviewDocument
    MsgBox "Could not refresh the preview: " & errorDescription, vbExclamation, "Preview Actions"
End Sub

Private Sub cmdClear_Click()
    Dim src As Object
    Dim errorDescription As String

    On Error GoTo ClearErr
    If Not PreviewTargetIsActive() Then
        RestoreEditablePreviewDocument
        ClearPreviewScopeRange
        Set gPreviewActionPanel = Nothing
        Me.Hide
        Unload Me
        Exit Sub
    End If
    RestoreEditablePreviewDocument
    ClearPreviewScopeRange
    Set src = mSourceForm
    Set gPreviewActionPanel = Nothing
    Me.Hide
    Unload Me
    If Not src Is Nothing Then
        LayoutCleanupToolFormForReconfigure src
        src.Show
        ReturnToLauncherAfterDetachedToolSession
    End If
    Exit Sub
ClearErr:
    errorDescription = Err.Description
    RestoreEditablePreviewDocument
    ClearPreviewScopeRange
    MsgBox "Could not reconfigure the preview: " & errorDescription, vbExclamation, "Preview Actions"
End Sub

Private Function PreviewTargetIsActive() As Boolean
    On Error GoTo NotActive
    If mPreviewDocument Is Nothing Then Exit Function
    If mPreviewWindow Is Nothing Then Exit Function
    If Not ActiveDocument Is mPreviewDocument Then Exit Function
    If Not ActiveWindow Is mPreviewWindow Then Exit Function
    PreviewTargetIsActive = True
NotActive:
End Function

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        If mPreviewOn Then
            cmdClear_Click
        Else
            Set gPreviewActionPanel = Nothing
            Me.Hide
        End If
    Else
        RestoreEditablePreviewDocument
        ClearPreviewScopeRange
        Set gPreviewActionPanel = Nothing
    End If
End Sub

Private Sub UserForm_Terminate()
    RestoreEditablePreviewDocument
    ClearPreviewScopeRange
    RestorePreviewBrowseTarget
    Set gPreviewActionPanel = Nothing
End Sub
