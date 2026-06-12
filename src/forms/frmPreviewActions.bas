Option Explicit

Private mSourceForm As Object
Private mToolName As String
Private mPreviewOn As Boolean
Private mSummaryText As String

Public Sub Configure(sourceForm As Object, toolName As String, summaryText As String)
    Set mSourceForm = sourceForm
    mToolName = toolName
    mSummaryText = summaryText
    mPreviewOn = True
    Me.Caption = "Preview Actions"
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
    Const FORM_W As Single = 530
    Const FORM_H As Single = 160
    Const CONTENT_W As Single = 250
    Const M As Single = 8
    Const BH As Single = 17
    Const GAP As Single = 5
    Const BTN_W As Single = 80
    Me.Width = FORM_W
    Me.Height = FORM_H
    lblTitle.Caption = ""
    lblTitle.Visible = False
    StyleActionBar
    lblTitle.Move 0, 0, 0, 0
    lblHint.Move M, 2, CONTENT_W, 12
    cmdPreview.Move M, 17, BTN_W, BH
    cmdClear.Move M + BTN_W + GAP, 17, BTN_W, BH
    cmdApply.Move M + (2 * (BTN_W + GAP)), 17, BTN_W, BH
    lblSummary.Move M, 39, CONTENT_W, 18
    lblTitle.Visible = False
    lblTitle.WordWrap = False
    lblSummary.WordWrap = True
    lblHint.WordWrap = False
    lblHint.TextAlign = 2
End Sub

Private Sub StyleActionBar()
    On Error Resume Next
    Me.BackColor = RGB(248, 249, 251)
    Me.SpecialEffect = 0
    StyleLabel lblTitle, 11, False, RGB(0, 0, 0)
    StyleLabel lblHint, 9, True, RGB(0, 0, 0)
    StyleLabel lblSummary, 8.5, False, RGB(0, 0, 0)
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
    ctl.Font.Size = 8.5
    ctl.Font.Bold = isBold
    ctl.BackColor = RGB(255, 255, 255)
    ctl.ForeColor = RGB(0, 0, 0)
End Sub

Private Sub cmdApply_Click()
    On Error GoTo ApplyErr
    If Not DocumentHasVisibleContent() Then
        RemoveAllHighlighting ActiveDocument.Content
        MsgBox "This document appears to be blank. There is nothing to apply.", vbInformation, "Preview Actions"
        Me.Hide
        If Not mSourceForm Is Nothing Then mSourceForm.Hide
        Set gPreviewActionPanel = Nothing
        Unload Me
        If GetReturnToMainAfterApplySetting() Then ShowCleanupSuiteLauncher
        Exit Sub
    End If
    Me.Hide
    If Not mSourceForm Is Nothing Then
        CallByName mSourceForm, "RunAfterPreview", VbMethod
    End If
    Set gPreviewActionPanel = Nothing
    Unload Me
    If GetReturnToMainAfterApplySetting() Then ShowCleanupSuiteLauncher
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
        lblSummary.Caption = mSummaryText
        LayoutPanel
        Me.Hide
        Me.Show vbModeless
    Else
        Dim src As Object
        Set src = mSourceForm
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
