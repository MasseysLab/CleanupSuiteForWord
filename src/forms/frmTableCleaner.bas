Option Explicit
' --------------------------------------------------------------
'  Table Cleaner
'  Removes empty rows and columns from tables, normalises cell padding
'  to Word defaults (0.04" top/bottom, 0.08" left/right), and optionally
'  strips direct table-level formatting, resets or removes border styles, and
'  can convert tables to plain tab-separated text.  Operates on all tables
'  within the target scope.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    chkRemoveEmptyRows.Value = True
    chkRemoveEmptyCols.Value = True
    chkNormalizePadding.Value = True
    chkStripDirectFormat.Value = False
    chkNormalizeBorders.Value = False
    chkRemoveBorders.Value = False
    chkConvertToText.Value = False
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
    chkRemoveEmptyRows.Caption = "Remove empty rows"
    chkRemoveEmptyCols.Caption = "Remove empty columns"
    chkNormalizePadding.Caption = "Normalize cell padding"
    chkStripDirectFormat.Caption = "Strip direct formatting from cells"
    chkNormalizeBorders.Caption = "Normalize border styles"
    chkRemoveBorders.Caption = "Remove all borders"
    chkConvertToText.Caption = "Convert single-column tables to plain text"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Table Cleaner" & vbCrLf & vbCrLf
    h = h & "Removes structural clutter from tables and normalises their formatting." & vbCrLf
    h = h & "Operates on all tables within the target scope." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Remove empty rows     --  Deletes any row where every cell contains" & vbCrLf
    h = h & "                            no text (a lone paragraph mark is not text)." & vbCrLf
    h = h & "  Remove empty columns  --  Deletes any column where every cell is empty." & vbCrLf
    h = h & "  Normalise padding     --  Sets all cell padding to Word defaults:" & vbCrLf
    h = h & "                            0.04 inch top/bottom,  0.08 inch left/right." & vbCrLf
    h = h & "  Strip direct format   --  Clears direct formatting overrides from the" & vbCrLf
    h = h & "                            table, restoring the table's applied style." & vbCrLf
    h = h & "  Normalise borders     --  Enables borders on the table using the" & vbCrLf
    h = h & "                            current table style." & vbCrLf
    h = h & "  Remove all borders    --  Turns off every border on the table." & vbCrLf
    h = h & "  Convert table to text --  Replaces the table with tab-separated text" & vbCrLf
    h = h & "                            (done last; other table options run first)." & vbCrLf
    h = h & vbCrLf
    h = h & "TIP: Run 'Remove empty rows/columns' before 'Strip direct format' for" & vbCrLf
    h = h & "the cleanest results.  Use Preview first to see which rows are empty." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to tables in the selection." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Table Cleaner"
End Sub
Private Function IsRowEmpty(rw As Row) As Boolean
    Dim cel As Cell
    For Each cel In rw.Cells
        If Len(Trim$(cel.Range.Text)) > 2 Then IsRowEmpty = False: Exit Function
    Next cel
    IsRowEmpty = True
End Function
Private Function IsColEmpty(tbl As Table, colIdx As Long) As Boolean
    Dim r As Long
    For r = 1 To tbl.Rows.Count
        On Error Resume Next
        Dim cel As Cell: Set cel = tbl.Cell(r, colIdx)
        On Error GoTo 0
        If Not cel Is Nothing Then
            If Len(Trim$(cel.Range.Text)) > 2 Then IsColEmpty = False: Exit Function
        End If
    Next r
    IsColEmpty = True
End Function
Private Sub cmdRun_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected. Please remove protection before running cleanup.", vbExclamation, "Document Protected": Exit Sub
    End If
    If Not GuardBeforeCleanup("Table Cleaner") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doEmptyRows As Boolean: doEmptyRows = chkRemoveEmptyRows.Value
    Dim doEmptyCols As Boolean: doEmptyCols = chkRemoveEmptyCols.Value
    Dim doPadding As Boolean: doPadding = chkNormalizePadding.Value
    Dim doDirectFmt As Boolean: doDirectFmt = chkStripDirectFormat.Value
    Dim doBorders As Boolean: doBorders = chkNormalizeBorders.Value
    Dim doRemoveBorders As Boolean: doRemoveBorders = chkRemoveBorders.Value
    Dim doConvertText As Boolean: doConvertText = chkConvertToText.Value
    If Not (doEmptyRows Or doEmptyCols Or doPadding Or doDirectFmt Or doBorders Or doRemoveBorders Or doConvertText) Then MsgBox "No table options selected.", vbInformation: Exit Sub
    ' Collect tables in scope
    Dim tbl As Table, tablesInScope As Collection
    Set tablesInScope = New Collection
    For Each tbl In ActiveDocument.Tables
        If tbl.Range.Start >= targetRange.Start And tbl.Range.End <= targetRange.End Then tablesInScope.Add tbl
    Next tbl
    If tablesInScope.Count = 0 Then MsgBox "No tables found in the selected scope.", vbInformation: Unload Me: Exit Sub
    ' Preview: highlight empty rows and cells
    Dim cnt As Long: cnt = 0
    If previewOnly Then
        Dim t As Variant
        For Each t In tablesInScope
            Set tbl = t
            If doEmptyRows Then
                Dim rw As Row
                For Each rw In tbl.Rows
                    If IsRowEmpty(rw) Then rw.Range.HighlightColorIndex = wdYellow: cnt = cnt + 1
                Next rw
            End If
        Next t
        MsgBox "Preview complete. " & cnt & " empty rows highlighted across " & tablesInScope.Count & " tables.", vbInformation
        Unload Me: Exit Sub
    End If
    MarkCleanupStart "Table Cleaner"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Table Cleaner"
    On Error GoTo RunErr
    Dim cntRows As Long, cntCols As Long, cntTables As Long, cntConverted As Long
    Dim tVar As Variant
    For Each tVar In tablesInScope
        Set tbl = tVar
        If doEmptyRows Then
            Dim ri As Long
            For ri = tbl.Rows.Count To 1 Step -1
                If IsRowEmpty(tbl.Rows(ri)) Then tbl.Rows(ri).Delete: cntRows = cntRows + 1
            Next ri
        End If
        If doEmptyCols Then
            Dim ci As Long
            For ci = tbl.Columns.Count To 1 Step -1
                If IsColEmpty(tbl, ci) Then tbl.Columns(ci).Delete: cntCols = cntCols + 1
            Next ci
        End If
        If doPadding Then
            tbl.TopPadding = InchesToPoints(0.04)
            tbl.BottomPadding = InchesToPoints(0.04)
            tbl.LeftPadding = InchesToPoints(0.08)
            tbl.RightPadding = InchesToPoints(0.08)
        End If
        If doDirectFmt Then
            tbl.Range.ClearFormatting
        End If
        If doBorders Then
            tbl.Borders.Enable = True
        End If
        If doRemoveBorders Then
            tbl.Borders.Enable = False
        End If
        cntTables = cntTables + 1
        If doConvertText Then
            tbl.ConvertToText Separator:=wdSeparateByTabs
            cntConverted = cntConverted + 1
        End If
    Next tVar
    Dim results As Collection: Set results = New Collection
    results.Add "Tables processed: " & cntTables
    If doEmptyRows Then results.Add "Empty rows removed: " & cntRows
    If doEmptyCols Then results.Add "Empty columns removed: " & cntCols
    If doConvertText Then results.Add "Tables converted to text: " & cntConverted
    ShowCleanupReport "Table Cleaner", results
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
