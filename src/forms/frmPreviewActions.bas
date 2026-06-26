Option Explicit

Private mSourceForm As Object
Private mToolName As String
Private mPreviewOn As Boolean
Private mSummaryText As String

Private Const FORM_W As Single = 283
Private Const MIN_FORM_H As Single = 90
Private Const CONTENT_W As Single = 257
Private Const M As Single = 8
Private Const BH As Single = 18
Private Const GAP As Single = 5
Private Const BTN_W As Single = (CONTENT_W - (2 * GAP)) / 3
Private Const SUMMARY_LINE_H As Single = 11
Private Const MAX_SUMMARY_H As Single = 88

Public Sub Configure(sourceForm As Object, toolName As String, summaryText As String)
    Set mSourceForm = sourceForm
    mToolName = toolName
    mSummaryText = summaryText
    mPreviewOn = True
    Me.Caption = ""
    ApplyMSFormTitleStrategy Me, True
    lblTitle.Caption = ""
    lblHint.Caption = toolName
    lblSummary.Caption = mSummaryText
    cmdApply.Caption = "Apply"
    cmdPreview.Caption = "Preview is ON"
    cmdPreview.Enabled = True
    cmdPreview.Visible = True
    cmdClear.Caption = "Reconfigure"
    LayoutPanel
End Sub

Private Sub LayoutPanel()
    Dim summaryH As Single
    summaryH = RequiredSummaryHeight(mSummaryText, CONTENT_W)
    Me.Caption = ""
    Me.Width = FORM_W
    Me.Height = MaxSingle(MIN_FORM_H, 74 + summaryH)
    lblTitle.Caption = ""
    lblTitle.Visible = False
    StyleActionBar
    lblTitle.Move 0, 0, 0, 0
    lblHint.Move M, 2, CONTENT_W, 12
    cmdPreview.Move M, 17, BTN_W, BH
    cmdClear.Move M + BTN_W + GAP, 17, BTN_W, BH
    cmdApply.Move M + (2 * (BTN_W + GAP)), 17, BTN_W, BH
    lblSummary.Move M, 40, CONTENT_W, summaryH
    lblTitle.Visible = False
    lblTitle.WordWrap = False
    lblSummary.WordWrap = True
    lblHint.WordWrap = False
    lblHint.TextAlign = 2
End Sub

Private Function RequiredSummaryHeight(textValue As String, contentWidth As Single) As Single
    RequiredSummaryHeight = MinSingle(MAX_SUMMARY_H, CountSummaryRows(textValue, contentWidth) * SUMMARY_LINE_H)
    If RequiredSummaryHeight < 16 Then RequiredSummaryHeight = 16
End Function

Private Function CountSummaryRows(textValue As String, contentWidth As Single) As Long
    Dim charsPerLine As Long
    charsPerLine = CLng(contentWidth / 4)
    If charsPerLine < 20 Then charsPerLine = 20

    Dim rows As Long
    Dim parts As Variant
    Dim part As Variant
    parts = Split(textValue, vbCrLf)
    For Each part In parts
        rows = rows + 1 + (Len(part) \ charsPerLine)
    Next part
    If rows < 1 Then rows = 1
    CountSummaryRows = rows
End Function

Private Function MinSingle(a As Single, b As Single) As Single
    If a < b Then
        MinSingle = a
    Else
        MinSingle = b
    End If
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
    ctl.Font.Size = 7.5
    ctl.Font.Bold = isBold
    ctl.BackColor = RGB(255, 255, 255)
    ctl.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub ShowModelessAtCurrentPosition()
    Dim panelLeft As Single
    Dim panelTop As Single
    panelLeft = Me.Left
    panelTop = Me.Top
    Me.Hide
    Me.Show vbModeless
    Me.Left = panelLeft
    Me.Top = panelTop
End Sub

Private Sub cmdApply_Click()
    On Error GoTo ApplyErr
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
    Me.Hide
    If Not mSourceForm Is Nothing Then
        CallByName mSourceForm, "RunAfterPreview", VbMethod
    End If
    Set gPreviewActionPanel = Nothing
    Unload Me
    Exit Sub
ApplyErr:
    MsgBox "Could not apply the preview: " & Err.Description, vbExclamation, "Preview Actions"
End Sub

Private Sub cmdPreview_Click()
    On Error GoTo PreviewErr
    If mPreviewOn Then
        RemoveAllHighlighting ActiveDocument.Content
        mPreviewOn = False
        cmdPreview.Caption = "Preview is OFF"
        cmdPreview.Enabled = True
        cmdPreview.Visible = True
        mSummaryText = "You may now edit your document if you wish."
        lblSummary.Caption = mSummaryText
        LayoutPanel
        ShowModelessAtCurrentPosition
    Else
        Dim src As Object
        Set src = mSourceForm
        RememberPreviewActionPanelPosition Me.Left, Me.Top
        Set gPreviewActionPanel = Nothing
        Me.Hide
        Unload Me
        If Not src Is Nothing Then CallByName src, "PreviewFromPanel", VbMethod
    End If
    Exit Sub
PreviewErr:
    MsgBox "Could not refresh the preview: " & Err.Description, vbExclamation, "Preview Actions"
End Sub

Private Sub cmdClear_Click()
    RemoveAllHighlighting ActiveDocument.Content
    Me.Hide
    If Not mSourceForm Is Nothing Then mSourceForm.Show
    Set gPreviewActionPanel = Nothing
    Unload Me
End Sub

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
        Set gPreviewActionPanel = Nothing
    End If
End Sub
