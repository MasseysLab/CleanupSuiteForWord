Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Final Review
'  Finalizes review markup before a document is shared by
'  removing comments and/or accepting tracked changes. Always
'  operates on the whole document.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    chkComments.Value = False
    chkRevisions.Value = False
    fraScopeSelection.Enabled = False
    optScopeDocument.Value = True
    optScopeDocument.Caption = "Entire document (always)"
    optScopeSelection.Enabled = False
    chkComments.Caption = "Remove all comments"
    chkRevisions.Caption = "Accept tracked changes and clear revision marks"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    cmdAuthorCleanupMinimal.Caption = "Author Cleanup - Minimal"
    cmdAuthorCleanupNormal.Caption = "Author Cleanup - Normal"
    cmdAuthorCleanupBroad.Caption = "Author Cleanup - Broad"
    LayoutFinalReviewForm
End Sub
Private Sub LayoutFinalReviewForm()
    LayoutCleanupToolForm Me
End Sub
Private Sub chkComments_Click(): LayoutFinalReviewForm: End Sub
Private Sub chkRevisions_Click(): LayoutFinalReviewForm: End Sub
Private Sub cmdPreview_Click()
    PreviewFromPanel
End Sub
Public Sub PreviewFromPanel()
    chkPreviewOnly.Value = True
    BeginPreviewActionIndicator Me
    cmdRun_Click
    EndPreviewActionIndicator
    chkPreviewOnly.Value = False
End Sub
Private Sub cmdReset_Click()
    UserForm_Initialize
    chkPreviewOnly.Value = False
End Sub
Public Sub RunAfterPreview()
    chkPreviewOnly.Value = False
    cmdRun_Click
End Sub
Private Sub cmdAuthorCleanupMinimal_Click()
    If Not ConfirmAuthorCleanup("minimal", "Undoability: partially undoable or not reliably undoable.") Then Exit Sub
    RunFinalReviewAuthorCleanup "minimal"
    MsgBox "Author cleanup completed." & vbCrLf & vbCrLf & "Undoability: partially undoable or not reliably undoable.", vbInformation, "Final Review"
End Sub
Private Sub cmdAuthorCleanupNormal_Click()
    If Not ConfirmAuthorCleanup("normal", "Undoability: partially undoable or not reliably undoable.") Then Exit Sub
    RunFinalReviewAuthorCleanup "normal"
    MsgBox "Author cleanup completed." & vbCrLf & vbCrLf & "Undoability: partially undoable or not reliably undoable.", vbInformation, "Final Review"
End Sub
Private Sub cmdAuthorCleanupBroad_Click()
    If Not ConfirmAuthorCleanup("broad", "Undoability: effectively final; treat as permanent once saved.") Then Exit Sub
    RunFinalReviewAuthorCleanup "broad"
    MsgBox "Author cleanup completed." & vbCrLf & vbCrLf & "Undoability: effectively final; treat as permanent once saved.", vbInformation, "Final Review"
End Sub
Private Sub cmdRun_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected. Please remove protection before running cleanup.", vbExclamation, "Document Protected": Exit Sub
    End If
    If Not GuardBeforeCleanup("Final Review") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim doComments As Boolean: doComments = chkComments.Value
    Dim doRevisions As Boolean: doRevisions = chkRevisions.Value
    If Not (doComments Or doRevisions) Then MsgBox "Nothing selected to finalize.", vbInformation: Exit Sub
    If previewOnly Then
        Dim previewComments As Long
        Dim previewRevisions As Long
        If doComments Then
            Dim previewComment As Comment
            For Each previewComment In ActiveDocument.Comments
                previewComments = previewComments + 1
                ApplyPreviewMinimalMarker previewComment.Scope, True
            Next previewComment
        End If
        If doRevisions Then
            Dim previewRevision As Revision
            For Each previewRevision In ActiveDocument.Revisions
                previewRevisions = previewRevisions + 1
                ApplyPreviewMinimalMarker previewRevision.Range, True
            Next previewRevision
        End If
        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Comments", previewComments, False, (Not doComments)
        AddPreviewSummaryRow previewRows, "Tracked changes", previewRevisions, False, (Not doRevisions)
        ShowPreviewActionsSummary Me, "Final Review", previewRows
        Exit Sub
    End If
    MarkCleanupStart "Final Review"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Final Review"
    On Error GoTo RunErr
    If doComments Then
        Dim commentCount As Long: commentCount = ActiveDocument.Comments.Count
        If commentCount > 0 Then
            ActiveDocument.DeleteAllComments
        End If
    End If
    If doRevisions Then
        Dim revisionCount As Long: revisionCount = ActiveDocument.Revisions.Count
        If revisionCount > 0 Then
            ActiveDocument.Revisions.AcceptAll
        End If
    End If
    undoRec.EndCustomRecord
    MarkCleanupEnd
    MarkCleanupToolApplied
    Unload Me
    Exit Sub
RunErr:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    On Error Resume Next: undoRec.EndCustomRecord: On Error GoTo 0
    MarkCleanupEnd
    Dim errMsg As String
    errMsg = "An unexpected error occurred: " & originalErrNumber & " - " & originalErrDescription
    errMsg = errMsg & vbCrLf & vbCrLf & _
             "Some changes may have been made to the document." & vbCrLf & _
             "Use Word's Undo command (Ctrl+Z) after closing this message if you want to roll them back."
    MsgBox errMsg, vbCritical, "Cleanup Error"
End Sub

Private Function ConfirmAuthorCleanup(ByVal removalLevel As String, ByVal undoabilityWarning As String) As Boolean
    Dim previewText As String
    Dim promptText As String
    previewText = PreviewFinalReviewAuthorCleanup(removalLevel)
    promptText = "Final Review author cleanup preview:" & vbCrLf & vbCrLf & _
                 previewText & vbCrLf & vbCrLf & _
                 undoabilityWarning & vbCrLf & vbCrLf & _
                 "Continue?"

    If MsgBox(promptText, vbQuestion + vbYesNo + vbDefaultButton2, "Final Review Author Cleanup") <> vbYes Then
        Exit Function
    End If
    ConfirmAuthorCleanup = True
End Function
