Option Explicit
' --------------------------------------------------------------
'  Style Cleanup
'  Removes unused custom styles and optionally remaps common variant styles
'  (Normal (Web), Body Text 2, etc.) to the Normal style.  Built-in Word
'  styles are never deleted.  Always operates on the whole document.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    chkRemoveUnused.Value = True
    chkRemapVariants.Value = False
    fraScopeSelection.Enabled = False
    optScopeDocument.Value = True
    optScopeDocument.Caption = "Entire document (always)"
    optScopeSelection.Enabled = False
    chkRemoveUnused.Caption = "Remove unused styles"
    chkRemapVariants.Caption = "Remap style variants to canonical equivalents"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Style Cleanup" & vbCrLf & vbCrLf
    h = h & "Removes clutter from the document style list.  Documents assembled from " & _
            "many sources accumulate dozens of unused or duplicate styles." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Remove unused styles    --  deletes custom styles that no text uses." & vbCrLf
    h = h & "                              Built-in Word styles are never deleted." & vbCrLf
    h = h & "  Remap common variants   --  reassigns text in variant styles such as" & vbCrLf
    h = h & "                              Normal (Web), Body Text 2, and Body Text 3 " & _
            "                              to the Normal style, then removes them." & vbCrLf
    h = h & vbCrLf
    h = h & "NOTE: This tool always works on the whole document.  Usage is detected " & _
            "in the main document body; a style used only inside a header, footer, or" & vbCrLf
    h = h & "text box may be treated as unused." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode reports how many unused custom styles were found." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Style Cleanup"
End Sub
Private Function StyleInUse(st As Style) As Boolean
    Dim fr As Range: Set fr = ActiveDocument.Content
    With fr.Find
        .ClearFormatting
        .Style = st
        .Text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Execute
        StyleInUse = .Found
    End With
End Function
Private Function RemapStyle(fromName As String, toName As String) As Long
    On Error Resume Next
    Dim st As Style: Set st = ActiveDocument.Styles(fromName)
    If st Is Nothing Then RemapStyle = 0: Exit Function
    With ActiveDocument.Content.Find
        .ClearFormatting: .Style = ActiveDocument.Styles(fromName)
        .Replacement.ClearFormatting: .Replacement.Style = ActiveDocument.Styles(toName)
        .Text = "": .Replacement.Text = ""
        .Forward = True: .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
    If Not st.BuiltIn Then st.Delete
    RemapStyle = 1
    On Error GoTo 0
End Function
Private Sub cmdRun_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected. Please remove protection before running cleanup.", vbExclamation, "Document Protected": Exit Sub
    End If
    If Not GuardBeforeCleanup("Style Cleanup") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim doUnused As Boolean: doUnused = chkRemoveUnused.Value
    Dim doRemap As Boolean: doRemap = chkRemapVariants.Value
    If Not (doUnused Or doRemap) Then MsgBox "Nothing selected.", vbInformation: Exit Sub
    Dim st As Style
    If previewOnly Then
        Dim preCount As Long: preCount = 0
        For Each st In ActiveDocument.Styles
            If Not st.BuiltIn Then
                If Not StyleInUse(st) Then preCount = preCount + 1
            End If
        Next st
        MsgBox "Preview complete. " & preCount & " unused custom styles found.", vbInformation
        Unload Me: Exit Sub
    End If
    MarkCleanupStart "Style Cleanup"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Style Cleanup"
    On Error GoTo RunErr
    Dim results As Collection: Set results = New Collection
    Dim cntRemap As Long: cntRemap = 0
    Dim cntUnused As Long: cntUnused = 0
    If doRemap Then
        Dim variants As Variant, vi As Long
        variants = Array("Normal (Web)", "Body Text 2", "Body Text 3", "Body Text Indent", "Body Text Indent 2")
        For vi = LBound(variants) To UBound(variants)
            cntRemap = cntRemap + RemapStyle(CStr(variants(vi)), "Normal")
        Next vi
        results.Add "Variant styles remapped to Normal: " & cntRemap
    End If
    If doUnused Then
        Dim i As Long
        For i = ActiveDocument.Styles.Count To 1 Step -1
            Set st = ActiveDocument.Styles(i)
            If Not st.BuiltIn Then
                If Not StyleInUse(st) Then
                    On Error Resume Next
                    st.Delete
                    If Err.Number = 0 Then cntUnused = cntUnused + 1
                    Err.Clear
                    On Error GoTo RunErr
                End If
            End If
        Next i
        results.Add "Unused custom styles removed: " & cntUnused
    End If
    ShowCleanupReport "Style Cleanup", results
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
