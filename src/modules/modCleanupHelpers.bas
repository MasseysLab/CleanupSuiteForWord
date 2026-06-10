Option Explicit
' Shared helper functions for cleanup forms
Public Sub RemoveAllHighlighting(Optional scopeRange As Range = Nothing)
    Dim hlRange As Range
    If scopeRange Is Nothing Then
        Set hlRange = ActiveDocument.Content
    Else
        Set hlRange = scopeRange
    End If
    With hlRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Format = True
        .Highlight = True
        .Text = ""
        .Replacement.Text = ""
        .Replacement.Highlight = False
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Public Function GuardBeforeCleanup(toolName As String) As Boolean
    GuardBeforeCleanup = False
    If ActiveDocument.Path = "" Then
        ' Document has never been saved -- Word has nothing to recover
        Dim msg As String
        msg = "This document has not been saved yet." & vbCrLf & vbCrLf & _
              "Saving first lets Word's built-in recovery protect you if anything goes wrong." & vbCrLf & vbCrLf & _
              "Save now before running " & toolName & "?"
        Dim ans As Integer
        ans = MsgBox(msg, vbExclamation + vbYesNoCancel, "Unsaved Document")
        If ans = vbCancel Then Exit Function
        If ans = vbYes Then
            On Error Resume Next
            ActiveDocument.Save
            If Err.Number <> 0 Then
                MsgBox "Save failed. Please save manually before running cleanup.", vbCritical, "Save Failed"
                Exit Function
            End If
            On Error GoTo 0
        End If
    ElseIf GetAutoSaveSetting() And Not IsRunningInMacroHostDocument() Then
        ' Document is saved and auto-save is ON -- save silently before proceeding
        On Error Resume Next
        ActiveDocument.Save
        On Error GoTo 0
    End If
    GuardBeforeCleanup = True
End Function
Public Function IsRunningInMacroHostDocument() As Boolean
    On Error Resume Next
    IsRunningInMacroHostDocument = (ActiveDocument.FullName = ThisDocument.FullName)
    On Error GoTo 0
End Function
Public Sub MarkCleanupStart(toolName As String)
    On Error Resume Next
    If IsRunningInMacroHostDocument() Then Exit Sub
    ActiveDocument.CustomDocumentProperties("CleanupSuiteInProgress").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteInProgress", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=toolName
    If Len(ActiveDocument.Path) > 0 Then ActiveDocument.Save
    On Error GoTo 0
End Sub
Public Sub MarkCleanupEnd()
    On Error Resume Next
    If IsRunningInMacroHostDocument() Then Exit Sub
    ActiveDocument.CustomDocumentProperties("CleanupSuiteInProgress").Delete
    On Error GoTo 0
End Sub
Public Function GetAutoSaveSetting() As Boolean
    ' Returns True (auto-save ON) by default; False only if explicitly disabled
    On Error Resume Next
    Dim prop As Object
    Set prop = ActiveDocument.CustomDocumentProperties("CleanupSuiteAutoSave")
    If Err.Number <> 0 Then
        GetAutoSaveSetting = True
        Err.Clear
    Else
        GetAutoSaveSetting = (prop.Value = "True")
    End If
    On Error GoTo 0
End Function
Public Sub SetAutoSaveSetting(enabled As Boolean)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("CleanupSuiteAutoSave").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteAutoSave", _
        LinkToContent:=False, Type:=msoPropertyTypeString, _
        Value:=IIf(enabled, "True", "False")
    On Error GoTo 0
End Sub
Public Function GetTargetRange() As Range
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        Set GetTargetRange = Selection.Range
    Else
        Set GetTargetRange = ActiveDocument.Content
    End If
End Function
Public Function ParagraphsInRange(scopeRange As Range) As Collection
    Dim col As Collection: Set col = New Collection
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= scopeRange.Start And p.Range.End <= scopeRange.End Then col.Add p
    Next p
    Set ParagraphsInRange = col
End Function
Public Sub ShowCleanupReport(reportTitle As String, results As Collection)
    If results.Count = 0 Then
        MsgBox "No items were changed.", vbInformation, reportTitle: Exit Sub
    End If
    Dim reportText As String: reportText = reportTitle & vbCrLf & vbCrLf
    Dim lineItem As Variant
    For Each lineItem In results
        reportText = reportText & "  " & lineItem & vbCrLf
    Next lineItem
    MsgBox reportText, vbInformation, reportTitle
End Sub
