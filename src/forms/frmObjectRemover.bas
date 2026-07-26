Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Remove Objects & Elements
'  Deletes embedded objects/elements that arrive as clutter: pictures,
'  text boxes, frames (text kept), horizontal lines, HTML/ActiveX controls,
'  hidden text, and (with confirmation) whole tables.  Every option is OFF
'  by default.  Operates on the main document body only -- header/footer
'  objects are left untouched.  Some deletions may not be fully reversible.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    chkPictures.Value = False
    chkTextBoxes.Value = False
    chkFrames.Value = False
    chkHorizontalLines.Value = False
    chkHtmlControls.Value = False
    chkHiddenText.Value = False
    chkTables.Value = False
    chkPictures.Caption = "Pictures (inline and floating images)"
    chkTextBoxes.Caption = "Text boxes (and their contents)"
    chkFrames.Caption = "Frames (remove frame, keep the text)"
    chkHorizontalLines.Caption = "Horizontal lines"
    chkHtmlControls.Caption = "HTML / ActiveX form controls"
    chkHiddenText.Caption = "Hidden text"
    chkTables.Caption = "Tables -- DELETES tables and all their contents"
    optScopeDocument.GroupName = "ObjScope"
    optScopeSelection.GroupName = "ObjScope"
    optScopeDocument.Value = True
    optScopeDocument.Caption = "Entire document (always)"
    optScopeSelection.Enabled = False
    fraScopeSelection.Enabled = False
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Function ProcessPictures(doDelete As Boolean) As Long
    On Error Resume Next
    Dim c As Long, k As Long
    For k = ActiveDocument.InlineShapes.Count To 1 Step -1
        Select Case ActiveDocument.InlineShapes(k).Type
            Case wdInlineShapePicture, wdInlineShapeLinkedPicture
                c = c + 1
                If doDelete Then
                    ActiveDocument.InlineShapes(k).Delete
                Else
                    ApplyPreviewMinimalMarker ActiveDocument.InlineShapes(k).Range
                End If
        End Select
    Next k
    For k = ActiveDocument.Shapes.Count To 1 Step -1
        Select Case ActiveDocument.Shapes(k).Type
            Case msoPicture, msoLinkedPicture
                c = c + 1
                If doDelete Then
                    ActiveDocument.Shapes(k).Delete
                Else
                    ApplyPreviewMinimalMarker ActiveDocument.Shapes(k).Anchor
                End If
        End Select
    Next k
    ProcessPictures = c
End Function
Private Function ProcessTextBoxes(doDelete As Boolean) As Long
    On Error Resume Next
    Dim c As Long, k As Long
    For k = ActiveDocument.Shapes.Count To 1 Step -1
        If ActiveDocument.Shapes(k).Type = msoTextBox Then
            c = c + 1
            If doDelete Then
                ActiveDocument.Shapes(k).Delete
            Else
                ApplyPreviewMinimalMarker ActiveDocument.Shapes(k).Anchor
            End If
        End If
    Next k
    ProcessTextBoxes = c
End Function
Private Function ProcessFrames(doDelete As Boolean) As Long
    On Error Resume Next
    Dim c As Long, k As Long
    c = ActiveDocument.Frames.Count
    If doDelete Then
        For k = ActiveDocument.Frames.Count To 1 Step -1
            ActiveDocument.Frames(k).Delete
        Next k
    Else
        For k = ActiveDocument.Frames.Count To 1 Step -1
            ApplyPreviewMinimalMarker ActiveDocument.Frames(k).Range, True
        Next k
    End If
    ProcessFrames = c
End Function
Private Function ProcessHorizontalLines(doDelete As Boolean) As Long
    On Error Resume Next
    Dim c As Long, k As Long
    For k = ActiveDocument.InlineShapes.Count To 1 Step -1
        If ActiveDocument.InlineShapes(k).Type = wdInlineShapeHorizontalLine Then
            c = c + 1
            If doDelete Then
                ActiveDocument.InlineShapes(k).Delete
            Else
                ApplyPreviewMinimalMarker ActiveDocument.InlineShapes(k).Range
            End If
        End If
    Next k
    ProcessHorizontalLines = c
End Function
Private Function ProcessHtmlControls(doDelete As Boolean) As Long
    On Error Resume Next
    Dim c As Long, k As Long
    For k = ActiveDocument.ContentControls.Count To 1 Step -1
        c = c + 1
        If doDelete Then
            ActiveDocument.ContentControls(k).Delete True
        Else
            ApplyPreviewMinimalMarker ActiveDocument.ContentControls(k).Range, True
        End If
    Next k
    For k = ActiveDocument.InlineShapes.Count To 1 Step -1
        If ActiveDocument.InlineShapes(k).Type = wdInlineShapeOLEControlObject Then
            c = c + 1
            If doDelete Then
                ActiveDocument.InlineShapes(k).Delete
            Else
                ApplyPreviewMinimalMarker ActiveDocument.InlineShapes(k).Range
            End If
        End If
    Next k
    ProcessHtmlControls = c
End Function
Private Function ProcessTables(doDelete As Boolean) As Long
    On Error Resume Next
    Dim c As Long, k As Long
    c = ActiveDocument.Tables.Count
    If doDelete Then
        For k = ActiveDocument.Tables.Count To 1 Step -1
            ActiveDocument.Tables(k).Delete
        Next k
    Else
        For k = ActiveDocument.Tables.Count To 1 Step -1
            ApplyPreviewMinimalMarker ActiveDocument.Tables(k).Range, True
        Next k
    End If
    ProcessTables = c
End Function
Private Function ProcessHiddenText(doDelete As Boolean) As Long
    On Error Resume Next
    Dim rng As Range: Set rng = ActiveDocument.Content.Duplicate
    rng.Find.ClearFormatting
    rng.Find.Text = ""
    rng.Find.Font.Hidden = True
    rng.Find.Forward = True
    rng.Find.Wrap = wdFindStop
    Dim found As Boolean: found = rng.Find.Execute
    If found And Not doDelete Then ApplyPreviewMinimalMarker rng, True
    If found And doDelete Then
        Dim r2 As Range: Set r2 = ActiveDocument.Content.Duplicate
        r2.Find.ClearFormatting
        r2.Find.Text = ""
        r2.Find.Font.Hidden = True
        r2.Find.Replacement.ClearFormatting
        r2.Find.Replacement.Text = ""
        r2.Find.Execute Replace:=wdReplaceAll
    End If
    If found Then ProcessHiddenText = 1 Else ProcessHiddenText = 0
End Function
Private Sub chkPictures_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkTextBoxes_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkFrames_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkHorizontalLines_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkHtmlControls_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkHiddenText_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkTables_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub cmdRiskPlacementObjectPictures_Click(): ShowToolRiskChoiceExplanation "Remove pictures", "Removes content", "Deletes inline and floating images." : End Sub
Private Sub cmdRiskPlacementObjectTextBoxes_Click(): ShowToolRiskChoiceExplanation "Remove text boxes", "Removes content", "Deletes text boxes and their contents." : End Sub
Private Sub cmdRiskPlacementObjectFrames_Click(): ShowToolRiskChoiceExplanation "Remove frames", "Structure change", "Removes frame containers while keeping their text." : End Sub
Private Sub cmdRiskPlacementObjectLines_Click(): ShowToolRiskChoiceExplanation "Remove horizontal lines", "Removes content", "Deletes horizontal line objects." : End Sub
Private Sub cmdRiskPlacementObjectControls_Click(): ShowToolRiskChoiceExplanation "Remove form controls", "Removes content", "Deletes HTML or ActiveX form controls." : End Sub
Private Sub cmdRiskPlacementObjectHiddenText_Click(): ShowToolRiskChoiceExplanation "Remove hidden text", "Removes content", "Deletes text marked hidden." : End Sub
Private Sub cmdRiskPlacementObjectTables_Click(): ShowToolRiskChoiceExplanation "Remove tables", "Removes content", "Deletes tables and everything inside them." : End Sub
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
    If Not GuardBeforeCleanup("Remove Objects") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim doPic As Boolean: doPic = chkPictures.Value
    Dim doTxt As Boolean: doTxt = chkTextBoxes.Value
    Dim doFra As Boolean: doFra = chkFrames.Value
    Dim doHL As Boolean: doHL = chkHorizontalLines.Value
    Dim doHtml As Boolean: doHtml = chkHtmlControls.Value
    Dim doHid As Boolean: doHid = chkHiddenText.Value
    Dim doTab As Boolean: doTab = chkTables.Value
    If Not (doPic Or doTxt Or doFra Or doHL Or doHtml Or doHid Or doTab) Then MsgBox "Select at least one element type to remove.", vbInformation: Exit Sub
    If previewOnly Then
        Dim previewPictures As Long
        Dim previewTextBoxes As Long
        Dim previewFrames As Long
        Dim previewHorizontalLines As Long
        Dim previewFormControls As Long
        Dim previewHiddenText As Long
        Dim previewTables As Long
        If doPic Then previewPictures = ProcessPictures(False)
        If doTxt Then previewTextBoxes = ProcessTextBoxes(False)
        If doFra Then previewFrames = ProcessFrames(False)
        If doHL Then previewHorizontalLines = ProcessHorizontalLines(False)
        If doHtml Then previewFormControls = ProcessHtmlControls(False)
        If doHid Then previewHiddenText = ProcessHiddenText(False)
        If doTab Then previewTables = ProcessTables(False)
        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Pictures", previewPictures, False, (Not doPic)
        AddPreviewSummaryRow previewRows, "Text boxes", previewTextBoxes, False, (Not doTxt)
        AddPreviewSummaryRow previewRows, "Frames", previewFrames, False, (Not doFra)
        AddPreviewSummaryRow previewRows, "Horizontal lines", previewHorizontalLines, False, (Not doHL)
        AddPreviewSummaryRow previewRows, "Form controls", previewFormControls, False, (Not doHtml)
        AddPreviewSummaryRow previewRows, "Hidden text", previewHiddenText, False, (Not doHid)
        AddPreviewSummaryRow previewRows, "Tables", previewTables, False, (Not doTab)
        ShowPreviewActionsSummary Me, "Object Remover", previewRows
        Exit Sub
    End If
    If doTab Then
        Dim tcount As Long: tcount = ActiveDocument.Tables.Count
        If tcount > 0 Then
            If MsgBox("This will permanently DELETE " & tcount & " table(s) and ALL their contents." & vbCrLf & vbCrLf & "(Different from Table Cleaner's 'Convert table to text', which keeps the content.)" & vbCrLf & vbCrLf & "Continue?", vbExclamation + vbYesNo, "Delete Tables") = vbNo Then Exit Sub
        End If
    End If
    MarkCleanupStart "Remove Objects"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Remove Objects"
    On Error GoTo RunErr
    Application.ScreenUpdating = False
    If doPic Then UpdateCleanupProgress "Removing pictures": ProcessPictures True
    If doTxt Then UpdateCleanupProgress "Removing text boxes": ProcessTextBoxes True
    If doFra Then UpdateCleanupProgress "Removing frames": ProcessFrames True
    If doHL Then UpdateCleanupProgress "Removing horizontal lines": ProcessHorizontalLines True
    If doHtml Then UpdateCleanupProgress "Removing form controls": ProcessHtmlControls True
    If doTab Then UpdateCleanupProgress "Removing tables": ProcessTables True
    If doHid Then
        UpdateCleanupProgress "Removing hidden text"
        ProcessHiddenText True
    End If
    Application.ScreenUpdating = True
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
    Application.ScreenUpdating = True
    On Error Resume Next: undoRec.EndCustomRecord: On Error GoTo 0
    MarkCleanupEnd
    Dim errMsg As String
    errMsg = "An unexpected error occurred: " & originalErrNumber & " - " & originalErrDescription
    errMsg = errMsg & vbCrLf & vbCrLf & _
             "Some changes may have been made to the document." & vbCrLf & _
             "Use Word's Undo command (Ctrl+Z) after closing this message if you want to roll them back."
    MsgBox errMsg, vbCritical, "Cleanup Error"
End Sub
