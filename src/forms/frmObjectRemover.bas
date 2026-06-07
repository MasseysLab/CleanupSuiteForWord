Option Explicit
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
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Remove Objects & Elements" & vbCrLf & vbCrLf
    h = h & "Deletes embedded objects and elements that often arrive as clutter from" & vbCrLf
    h = h & "web pages, PDFs, or email -- so you can reduce a document to clean text." & vbCrLf
    h = h & vbCrLf
    h = h & "WHAT EACH OPTION REMOVES" & vbCrLf
    h = h & "  Pictures         --  inline and floating images." & vbCrLf
    h = h & "  Text boxes       --  floating text boxes AND the text inside them." & vbCrLf
    h = h & "  Frames           --  the frame container; the text inside is kept." & vbCrLf
    h = h & "  Horizontal lines --  the inserted horizontal-line graphic." & vbCrLf
    h = h & "  HTML/ActiveX     --  form controls (buttons, input boxes) from web paste." & vbCrLf
    h = h & "  Hidden text      --  text formatted as hidden." & vbCrLf
    h = h & "  Tables           --  DELETES whole tables and all their contents." & vbCrLf
    h = h & vbCrLf
    h = h & "IMPORTANT CAUTIONS" & vbCrLf
    h = h & "  - Every option is OFF by default; tick only what you want removed." & vbCrLf
    h = h & "  - Removing tables is destructive and is NOT the same as Table Cleaner's" & vbCrLf
    h = h & "    'Convert table to text', which keeps the content.  You will be asked" & vbCrLf
    h = h & "    to confirm before any tables are deleted." & vbCrLf
    h = h & "  - Deleting pictures and shapes may not always be fully reversible by" & vbCrLf
    h = h & "    Undo, depending on your Word version.  Work on a copy if unsure." & vbCrLf
    h = h & "  - Works on the main document body only; objects in headers and footers" & vbCrLf
    h = h & "    (such as logos) are left untouched." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode reports a count of what would be removed (it cannot" & vbCrLf
    h = h & "highlight an object that is about to be deleted)." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Remove Objects & Elements"
End Sub
Private Function ProcessPictures(doDelete As Boolean) As Long
    On Error Resume Next
    Dim c As Long, k As Long
    For k = ActiveDocument.InlineShapes.Count To 1 Step -1
        Select Case ActiveDocument.InlineShapes(k).Type
            Case wdInlineShapePicture, wdInlineShapeLinkedPicture
                c = c + 1
                If doDelete Then ActiveDocument.InlineShapes(k).Delete
        End Select
    Next k
    For k = ActiveDocument.Shapes.Count To 1 Step -1
        Select Case ActiveDocument.Shapes(k).Type
            Case msoPicture, msoLinkedPicture
                c = c + 1
                If doDelete Then ActiveDocument.Shapes(k).Delete
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
            If doDelete Then ActiveDocument.Shapes(k).Delete
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
    End If
    ProcessFrames = c
End Function
Private Function ProcessHorizontalLines(doDelete As Boolean) As Long
    On Error Resume Next
    Dim c As Long, k As Long
    For k = ActiveDocument.InlineShapes.Count To 1 Step -1
        If ActiveDocument.InlineShapes(k).Type = wdInlineShapeHorizontalLine Then
            c = c + 1
            If doDelete Then ActiveDocument.InlineShapes(k).Delete
        End If
    Next k
    ProcessHorizontalLines = c
End Function
Private Function ProcessHtmlControls(doDelete As Boolean) As Long
    On Error Resume Next
    Dim c As Long, k As Long
    For k = ActiveDocument.InlineShapes.Count To 1 Step -1
        If ActiveDocument.InlineShapes(k).Type = wdInlineShapeOLEControlObject Then
            c = c + 1
            If doDelete Then ActiveDocument.InlineShapes(k).Delete
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
        Dim msg As String
        msg = "Preview -- elements that would be removed (entire document):"
        If doPic Then msg = msg & vbCrLf & "Pictures: " & ProcessPictures(False)
        If doTxt Then msg = msg & vbCrLf & "Text boxes: " & ProcessTextBoxes(False)
        If doFra Then msg = msg & vbCrLf & "Frames: " & ProcessFrames(False)
        If doHL Then msg = msg & vbCrLf & "Horizontal lines: " & ProcessHorizontalLines(False)
        If doHtml Then msg = msg & vbCrLf & "HTML/ActiveX controls: " & ProcessHtmlControls(False)
        If doTab Then msg = msg & vbCrLf & "Tables: " & ProcessTables(False)
        If doHid Then msg = msg & vbCrLf & "Hidden text present: " & IIf(ProcessHiddenText(False) > 0, "Yes", "No")
        MsgBox msg, vbInformation, "Remove Objects -- Preview"
        Unload Me: Exit Sub
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
    Dim results As Collection: Set results = New Collection
    If doPic Then results.Add "Pictures removed: " & ProcessPictures(True)
    If doTxt Then results.Add "Text boxes removed: " & ProcessTextBoxes(True)
    If doFra Then results.Add "Frames removed (text kept): " & ProcessFrames(True)
    If doHL Then results.Add "Horizontal lines removed: " & ProcessHorizontalLines(True)
    If doHtml Then results.Add "HTML/ActiveX controls removed: " & ProcessHtmlControls(True)
    If doTab Then results.Add "Tables removed: " & ProcessTables(True)
    If doHid Then
        If ProcessHiddenText(True) > 0 Then
            results.Add "Hidden text removed"
        Else
            results.Add "No hidden text found"
        End If
    End If
    Application.ScreenUpdating = True
    ShowCleanupReport "Remove Objects", results
    undoRec.EndCustomRecord
    MarkCleanupEnd
    Unload Me
    Exit Sub
RunErr:
    Application.ScreenUpdating = True
    On Error Resume Next: undoRec.EndCustomRecord: On Error GoTo 0
    MarkCleanupEnd
    Dim errMsg As String
    errMsg = "An unexpected error occurred: " & Err.Number & " - " & Err.Description
    errMsg = errMsg & vbCrLf & vbCrLf & "Some changes may have been made to the document." & vbCrLf & "Note: deletion of pictures or shapes may not always be fully reversible." & vbCrLf & "Would you like to undo all changes made before the error?"
    If MsgBox(errMsg, vbCritical + vbYesNo, "Cleanup Error") = vbYes Then
        On Error Resume Next: Application.Undo: On Error GoTo 0
    End If
End Sub