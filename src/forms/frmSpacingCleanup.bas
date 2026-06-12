Option Explicit
' --------------------------------------------------------------
'  Spacing Fixer
'  Corrects spacing inconsistencies throughout the document: collapses
'  double (and multiple) spaces, trims leading and trailing spaces from
'  paragraphs, removes spaces before punctuation, normalizes spaces after
'  punctuation to a single space, and collapses runs of excess blank lines.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optAll.GroupName = "SpacMode"
    optDoubleSpaces.GroupName = "SpacMode"
    optTrim.GroupName = "SpacMode"
    optCustom.GroupName = "SpacMode"
    optScopeDocument.GroupName = "SpacScope"
    optScopeSelection.GroupName = "SpacScope"
    optAll.Value = True: UpdateCustomVisibility: ClearCustomChecks
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
    optAll.Caption = "All spacing fixes"
    optDoubleSpaces.Caption = "Double spaces only"
    optTrim.Caption = "Trim leading / trailing spaces"
    optCustom.Caption = "Custom (choose below)"
    chkDoubleSpaces.Caption = "Collapse double (and multiple) spaces"
    chkTrimSpaces.Caption = "Remove leading and trailing spaces"
    chkSpaceBeforePunct.Caption = "Remove space before punctuation  ( , . : ; ! ? )"
    chkNormalizeAfterPunct.Caption = "Normalize spacing after punctuation"
    chkExtraBlankLines.Caption = "Collapse extra blank lines between paragraphs"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Spacing Fixer" & vbCrLf & vbCrLf
    h = h & "Corrects the most common spacing inconsistencies in assembled documents." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Double spaces       --  Collapses two or more consecutive spaces to one." & vbCrLf
    h = h & "  Trim paragraphs     --  Removes leading and trailing spaces from each" & vbCrLf
    h = h & "                          paragraph." & vbCrLf
    h = h & "  Space before punct  --  Removes any space immediately before . , ; : ! ?" & vbCrLf
    h = h & "  Space after punct   --  Ensures exactly one space follows  . ? ! , ; :" & vbCrLf
    h = h & "  Extra blank lines   --  Collapses three or more consecutive blank" & vbCrLf
    h = h & "                          paragraphs down to two." & vbCrLf
    h = h & vbCrLf
    h = h & "Each option counts how many changes it makes and reports them at the end." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights problem areas without making changes." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Spacing Fixer"
End Sub
Private Sub UpdateCustomVisibility(): fraCustom.Visible = optCustom.Value: End Sub
Private Sub ClearCustomChecks()
    chkDoubleSpaces.Value = False: chkTrimSpaces.Value = False: chkSpaceBeforePunct.Value = False
    chkNormalizeAfterPunct.Value = False: chkExtraBlankLines.Value = False
End Sub
Private Sub cmdSelectAll_Click()
    If fraCustom.Visible Then
        chkDoubleSpaces.Value = True
        chkTrimSpaces.Value = True
        chkSpaceBeforePunct.Value = True
        chkNormalizeAfterPunct.Value = True
        chkExtraBlankLines.Value = True
    End If
End Sub
Private Sub cmdDeselectAll_Click()
    If fraCustom.Visible Then
        ClearCustomChecks
    End If
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
    If Not GuardBeforeCleanup("Spacing Fixer") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doDouble As Boolean, doTrim As Boolean, doBeforePunct As Boolean, doAfterPunct As Boolean, doBlank As Boolean
    If optAll.Value Then doDouble = True: doTrim = True: doBeforePunct = True: doAfterPunct = True: doBlank = True
    If optDoubleSpaces.Value Then doDouble = True
    If optTrim.Value Then doTrim = True
    If optCustom.Value Then
        doDouble = chkDoubleSpaces.Value: doTrim = chkTrimSpaces.Value: doBeforePunct = chkSpaceBeforePunct.Value
        doAfterPunct = chkNormalizeAfterPunct.Value: doBlank = chkExtraBlankLines.Value
    End If
    If Not (doDouble Or doTrim Or doBeforePunct Or doAfterPunct Or doBlank) Then MsgBox "No spacing options selected.", vbInformation: Exit Sub
    ' Highlight preview matches
    If doDouble Then
        With targetRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "  "
            .Replacement.Text = "^&"
            .Replacement.Highlight = True
            .Execute Replace:=wdReplaceAll
        End With
    End If
    If doBeforePunct Then
        With targetRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = " ([.,;:!?])"
            .MatchWildcards = True
            .Replacement.Text = "^&"
            .Replacement.Highlight = True
            .Execute Replace:=wdReplaceAll
            .MatchWildcards = False
        End With
    End If
    If doAfterPunct Then
        Dim punctPatterns: punctPatterns = Array(".  ","?  ","!  ",",  ",";  ",":  ")
        Dim punct: For Each punct In punctPatterns: With targetRange.Find: .ClearFormatting: .Replacement.ClearFormatting: .Text = punct: .Replacement.Text = "^&": .Replacement.Highlight = True: .Execute Replace:=wdReplaceAll: End With: Next punct
    End If
    If doBlank Then
        With targetRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "^p^p^p"
            .Replacement.Text = "^&"
            .Replacement.Highlight = True
            .Execute Replace:=wdReplaceAll
        End With
    End If
    If doTrim Then
        With targetRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "^p "
            .Replacement.Text = "^&"
            .Replacement.Highlight = True
            .Execute Replace:=wdReplaceAll
            .Text = " ^p"
            .Replacement.Text = "^&"
            .Execute Replace:=wdReplaceAll
        End With
    End If
    If previewOnly Then
        ShowPreviewActions Me, "Spacing Fixer", "Preview complete. Matches highlighted."
        Exit Sub
    End If
    MarkCleanupStart "Spacing Fixer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Spacing Fixer"
    On Error GoTo RunErr
    ' Perform replacements with counting
    Dim cntDouble As Long, cntTrim As Long, cntBefore As Long, cntAfter As Long, cntBlank As Long
    If doDouble Then
        With targetRange.Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = "  ": .Replacement.Text = " ": .Forward = True: .Wrap = wdFindStop
            Do While .Execute(Replace:=wdReplaceOne): cntDouble = cntDouble + 1: Loop
        End With
        ' Mop-up pass: a triple-space becomes a double-space after the loop above; repeat until clean
        Do
            With targetRange.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "  "
                .Replacement.Text = " "
                If Not .Execute(Replace:=wdReplaceAll) Then Exit Do
            End With
        Loop While targetRange.Find.Execute(FindText:="  ", Forward:=True)
    End If
    If doTrim Then
        Dim para As Paragraph
        For Each para In ActiveDocument.Paragraphs
            If para.Range.Start >= targetRange.Start And para.Range.End <= targetRange.End Then
                Dim paraText As String: paraText = para.Range.Text
                If Right$(paraText, 1) = vbCr Then paraText = Left$(paraText, Len(paraText) - 1)
                Dim trimmedText As String: trimmedText = Trim$(paraText)
                If trimmedText <> paraText Then para.Range.Text = trimmedText & vbCr: cntTrim = cntTrim + 1
            End If
        Next para
    End If
    If doBeforePunct Then
        With targetRange.Find: .ClearFormatting: .Replacement.ClearFormatting: .Text = " ([.,;:!?])": .Replacement.Text = "\1": .MatchWildcards = True: Do While .Execute(Replace:=wdReplaceOne): cntBefore = cntBefore + 1: Loop: .MatchWildcards = False: End With
    End If
    If doAfterPunct Then
        Dim punctPatterns: punctPatterns = Array(".  ","?  ","!  ",",  ",";  ",":  ")
        Dim pattern: For Each pattern In punctPatterns: With targetRange.Find: .ClearFormatting: .Replacement.ClearFormatting: .Text = pattern: .Replacement.Text = Left$(pattern, 1) & " ": .MatchWildcards = False: Do While .Execute(Replace:=wdReplaceOne): cntAfter = cntAfter + 1: Loop: End With: Next pattern
    End If
    If doBlank Then
        With targetRange.Find: .ClearFormatting: .Replacement.ClearFormatting: .Text = "^p^p^p": .Replacement.Text = "^p^p": Do While .Execute(Replace:=wdReplaceOne): cntBlank = cntBlank + 1: Loop: End With
        Do
            With targetRange.Find
                .ClearFormatting
                .Replacement.ClearFormatting
                .Text = "^p^p^p"
                .Replacement.Text = "^p^p"
                If Not .Execute(Replace:=wdReplaceAll) Then Exit Do
            End With
        Loop While targetRange.Find.Execute(FindText:="^p^p^p", Forward:=True)
    End If
    RemoveAllHighlighting targetRange
    Dim results As Collection: Set results = New Collection
    If doDouble      Then results.Add "Double spaces removed:         " & cntDouble
    If doTrim        Then results.Add "Leading/trailing trims:        " & cntTrim
    If doBeforePunct Then results.Add "Spaces before punctuation:     " & cntBefore
    If doAfterPunct  Then results.Add "Spaces after punctuation fixed: " & cntAfter
    If doBlank       Then results.Add "Extra blank lines collapsed:    " & cntBlank
    ShowCleanupReport "Spacing Cleanup", results
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
