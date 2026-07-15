Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  MetaData
'  Strips identifying information such as author, company, and save
'  identity before a document is shared.  A focused subset of Word's
'  Document Inspector for document-level privacy cleanup.
'  Always operates on the whole document.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    chkProperties.Value = True
    chkPersonalInfo.Value = True
    fraScopeSelection.Enabled = False
    optScopeDocument.Value = True
    optScopeDocument.Caption = "Entire document (always)"
    optScopeSelection.Enabled = False
    chkProperties.Caption = "Clear document properties (Title, Author, Company, etc.)"
    chkPersonalInfo.Caption = "Clear personal information (Last Saved By, revision authors)"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub chkProperties_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkPersonalInfo_Click(): LayoutCleanupToolForm Me: End Sub
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
Private Sub cmdRun_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected. Please remove protection before running cleanup.", vbExclamation, "Document Protected": Exit Sub
    End If
    If Not GuardBeforeCleanup("MetaData") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim doProps As Boolean: doProps = chkProperties.Value
    Dim doPersonal As Boolean: doPersonal = chkPersonalInfo.Value
    If Not (doProps Or doPersonal) Then MsgBox "Nothing selected to remove.", vbInformation: Exit Sub
    Dim propertySignals As Variant
    Dim personalSignals As Variant
    propertySignals = Array("Title", "Subject", "Author", "Company", "Manager", "Category")
    personalSignals = Array("Author", "Last Author")
    If previewOnly Then
        Dim previewProperties As Long
        Dim previewPersonal As Long
        If doProps Then previewProperties = CountPopulatedMetadataValues(propertySignals)
        If doPersonal Then previewPersonal = CountPopulatedMetadataValues(personalSignals)
        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Document properties", previewProperties, True, (Not doProps)
        AddPreviewSummaryRow previewRows, "Personal information", previewPersonal, True, (Not doPersonal)
        ShowPreviewActionsSummary Me, "MetaData", previewRows
        Exit Sub
    End If
    MarkCleanupStart "MetaData"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - MetaData"
    On Error GoTo RunErr
    Dim propsBefore As Long
    Dim propsAfter As Long
    Dim personalBefore As Long
    Dim personalAfter As Long
    propsBefore = CountPopulatedMetadataValues(propertySignals)
    personalBefore = CountPopulatedMetadataValues(personalSignals)
    If doProps Then
        On Error Resume Next
        ActiveDocument.RemoveDocumentInformation wdRDIDocumentProperties
        On Error GoTo RunErr
        propsAfter = CountPopulatedMetadataValues(propertySignals)
    End If
    If doPersonal Then
        On Error Resume Next
        ActiveDocument.RemoveDocumentInformation wdRDIRemovePersonalInformation
        On Error GoTo RunErr
        personalAfter = CountPopulatedMetadataValues(personalSignals)
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

Private Function MetadataPropertySnapshot(ByVal propertyName As String) As String
    On Error Resume Next
    Dim rawValue As Variant
    rawValue = ActiveDocument.BuiltInDocumentProperties(propertyName).Value
    If Err.Number <> 0 Then
        Err.Clear
        MetadataPropertySnapshot = "(unavailable)"
        Exit Function
    End If
    If IsNull(rawValue) Then
        MetadataPropertySnapshot = "(blank)"
    ElseIf Len(Trim$(CStr(rawValue))) = 0 Then
        MetadataPropertySnapshot = "(blank)"
    Else
        MetadataPropertySnapshot = CStr(rawValue)
    End If
End Function

Private Function CountPopulatedMetadataValues(ByVal propertyNames As Variant) As Long
    Dim item As Variant
    Dim total As Long
    For Each item In propertyNames
        If MetadataPropertySnapshot(CStr(item)) <> "(blank)" And MetadataPropertySnapshot(CStr(item)) <> "(unavailable)" Then
            total = total + 1
        End If
    Next item
    CountPopulatedMetadataValues = total
End Function
