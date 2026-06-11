Option Explicit

Private mSourceForm As Object
Private mToolName As String

Public Sub Configure(sourceForm As Object, toolName As String, summaryText As String)
    Set mSourceForm = sourceForm
    mToolName = toolName
    lblTitle.Caption = toolName & " Preview"
    lblSummary.Caption = summaryText
    lblHint.Caption = "Apply reruns the same settings without Preview. Reset Preview clears yellow marks and returns here."
    cmdApply.Caption = "Apply"
    cmdClear.Caption = "Reset Preview"
End Sub

Private Sub cmdApply_Click()
    On Error GoTo ApplyErr
    Me.Hide
    If Not mSourceForm Is Nothing Then
        CallByName mSourceForm, "RunAfterPreview", VbMethod
    End If
    Set gPreviewActionPanel = Nothing
    Unload Me
    ShowCleanupSuiteLauncher
    Exit Sub
ApplyErr:
    MsgBox "Could not apply the preview: " & Err.Description, vbExclamation, "Preview Actions"
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
        cmdClear_Click
    Else
        Set gPreviewActionPanel = Nothing
    End If
End Sub
