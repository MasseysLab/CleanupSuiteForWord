Option Explicit

Private Sub UserForm_Initialize()
    InitializeSafeMetadataEditor Me
End Sub

Private Sub lstCustomProperties_Click()
    LoadSelectedSafeCustomProperty Me
End Sub

Private Sub cmdCustomAddUpdate_Click()
    AddOrUpdateSafeCustomProperty Me
End Sub

Private Sub cmdCustomDelete_Click()
    DeleteSelectedSafeCustomProperty Me
End Sub

Private Sub cmdSave_Click()
    SaveSafeMetadataEditor Me
End Sub

Private Sub cmdCancel_Click()
    CancelSafeMetadataEditor Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        NoteSafeMetadataEditorCancelled
    End If
End Sub
