Option Explicit
Public Const SUITE_VERSION As String = "0.8.0"
Public gRibbon As Object   ' IRibbonUI -- stored for optional Invalidate calls
Sub ShowCleanupSuiteLauncher()
    frmCleanupSuiteLauncher.Show
End Sub
Public Sub RibbonOnLoad(ribbon As Object)
    Set gRibbon = ribbon
End Sub
Public Sub AutoOpen()
    ' Fires whenever any document is opened while this template is active.
    ' Detects an interrupted cleanup operation and warns the user.
    On Error Resume Next
    Dim cp As Object
    Set cp = ActiveDocument.CustomDocumentProperties("CleanupSuiteInProgress")
    If Not cp Is Nothing Then
        Dim toolName As String: toolName = cp.Value
        Dim msg As String
        msg = "A Cleanup Suite operation (" & toolName & ") on this document was interrupted unexpectedly." & vbCrLf & vbCrLf
        msg = msg & "The document may be in a partially modified state." & vbCrLf
        msg = msg & "If you saved a backup before running cleanup, compare it with this version." & vbCrLf & vbCrLf
        msg = msg & "Press Ctrl+Z to undo any unwanted changes."
        MsgBox msg, vbExclamation, "Cleanup Suite - Interrupted Operation"
        cp.Delete
    End If
    On Error GoTo 0
End Sub
