Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then ReturnToMainAfterToolClose
End Sub
' --------------------------------------------------------------
'  Metadata Scrubber
'  Strips identifying information (author, company, comment and revision
'  author names, and optionally comments and tracked changes) before a
'  document is shared.  A one-click version of Word Document Inspector.
'  Always operates on the whole document.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    chkProperties.Value = True
    chkPersonalInfo.Value = True
    chkComments.Value = False
    chkRevisions.Value = False
    fraScopeSelection.Enabled = False
    optScopeDocument.Value = True
    optScopeDocument.Caption = "Entire document (always)"
    optScopeSelection.Enabled = False
    chkProperties.Caption = "Clear document properties (Title, Author, Company, etc.)"
    chkPersonalInfo.Caption = "Clear personal information (Last Saved By, revision authors)"
    chkComments.Caption = "Remove all comments"
    chkRevisions.Caption = "Accept tracked changes and clear revision marks"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Metadata Scrubber" & vbCrLf & vbCrLf
    h = h & "Strips identifying information from the document before you share it " & _
            "externally -- a one-click version of Word built-in Document Inspector." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Document properties  --  author, company, manager, title, and other" & vbCrLf
    h = h & "                           built-in properties." & vbCrLf
    h = h & "  Personal information --  author names attached to comments and tracked" & vbCrLf
    h = h & "                           changes (anonymises them)." & vbCrLf
    h = h & "  Delete all comments  --  removes every comment from the document." & vbCrLf
    h = h & "  Accept all tracked changes  --  finalises the document and removes the" & vbCrLf
    h = h & "                                  revision history." & vbCrLf
    h = h & vbCrLf
    h = h & "NOTE: This tool always works on the whole document.  Metadata removal " & _
            "may not be reversible with Ctrl+Z, so a copy is saved first if auto-save" & vbCrLf
    h = h & "is on.  Review the result before sharing." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode shows the metadata currently in the document." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Metadata Scrubber"
End Sub
Private Sub cmdPreview_Click()
    PreviewFromPanel
End Sub
Public Sub PreviewFromPanel()
    chkPreviewOnly.Value = True
    cmdRun_Click
End Sub
Private Sub cmdReset_Click()
    UserForm_Initialize
    chkPreviewOnly.Value = False
End Sub
Public Sub RunAfterPreview()
    chkPreviewOnly.Value = False
    cmdRun_Click
End Sub
Private Sub cmdRun_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected. Please remove protection before running cleanup.", vbExclamation, "Document Protected": Exit Sub
    End If
    If Not GuardBeforeCleanup("Metadata Scrubber") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim doProps As Boolean: doProps = chkProperties.Value
    Dim doPersonal As Boolean: doPersonal = chkPersonalInfo.Value
    Dim doComments As Boolean: doComments = chkComments.Value
    Dim doRevisions As Boolean: doRevisions = chkRevisions.Value
    If Not (doProps Or doPersonal Or doComments Or doRevisions) Then MsgBox "Nothing selected to remove.", vbInformation: Exit Sub
    If previewOnly Then
        Dim info As String
        info = "Current document metadata:" & vbCrLf & vbCrLf
        On Error Resume Next
        info = info & "Author: " & ActiveDocument.BuiltInDocumentProperties("Author").Value & vbCrLf
        info = info & "Company: " & ActiveDocument.BuiltInDocumentProperties("Company").Value & vbCrLf
        info = info & "Last author: " & ActiveDocument.BuiltInDocumentProperties("Last Author").Value & vbCrLf
        On Error GoTo 0
        info = info & "Comments: " & ActiveDocument.Comments.Count & vbCrLf
        info = info & "Tracked changes: " & ActiveDocument.Revisions.Count & vbCrLf & vbCrLf
        info = info & "Run without Preview to remove the selected items."
        ShowPreviewActions Me, "Metadata Scrubber", info
        Exit Sub
    End If
    MarkCleanupStart "Metadata Scrubber"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Metadata Scrubber"
    On Error GoTo RunErr
    Dim results As Collection: Set results = New Collection
    If doComments Then
        Dim cc As Long: cc = ActiveDocument.Comments.Count
        ActiveDocument.DeleteAllComments
        results.Add "Comments removed: " & cc
    End If
    If doRevisions Then
        Dim rc As Long: rc = ActiveDocument.Revisions.Count
        ActiveDocument.Revisions.AcceptAll
        results.Add "Tracked changes accepted: " & rc
    End If
    If doProps Then
        On Error Resume Next
        ActiveDocument.RemoveDocumentInformation wdRDIDocumentProperties
        On Error GoTo RunErr
        results.Add "Document properties cleared"
    End If
    If doPersonal Then
        On Error Resume Next
        ActiveDocument.RemoveDocumentInformation wdRDIRemovePersonalInformation
        On Error GoTo RunErr
        results.Add "Personal information removed"
    End If
    ShowCleanupReport "Metadata Scrubber", results
    undoRec.EndCustomRecord
    MarkCleanupEnd
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
