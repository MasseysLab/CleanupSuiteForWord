Option Explicit

Private mSourceForm As Object
Private mToolName As String

Public Sub Configure(sourceForm As Object, toolName As String, summaryText As String)
    Set mSourceForm = sourceForm
    mToolName = toolName
    lblTitle.Caption = toolName & " Preview"
    lblSummary.Caption = summaryText
    lblHint.Caption = "Apply reruns the same settings without Preview. Clear Preview removes the yellow marks."
    cmdApply.Caption = "Apply"
    cmdClear.Caption = "Reset Preview"
    cmdBack.Caption = "Back"
    cmdClose.Caption = "Close"
End Sub

Private Sub cmdApply_Click()
    On Error GoTo ApplyErr
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

Private Sub cmdClear_Click()
    RemoveAllHighlighting ActiveDocument.Content
    CloseSourceForm
    Set gPreviewActionPanel = Nothing
    Unload Me
End Sub

Private Sub cmdBack_Click()
    Me.Hide
    If Not mSourceForm Is Nothing Then mSourceForm.Show
    Set gPreviewActionPanel = Nothing
    Unload Me
End Sub

Private Sub cmdClose_Click()
    CloseSourceForm
    Set gPreviewActionPanel = Nothing
    Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set gPreviewActionPanel = Nothing
End Sub

Private Sub CloseSourceForm()
    On Error Resume Next
    If Not mSourceForm Is Nothing Then Unload mSourceForm
    Set mSourceForm = Nothing
    On Error GoTo 0
End Sub
