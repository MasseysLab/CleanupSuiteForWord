Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
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
    RefreshScopeSelectionState Me, True
    chkRemoveEmptyRows.Caption = "Remove empty rows"
    chkRemoveEmptyCols.Caption = "Remove empty columns"
    chkNormalizePadding.Caption = "Normalize cell padding"
    chkStripDirectFormat.Caption = "Strip direct formatting from cells"
    chkNormalizeBorders.Caption = "Normalize border styles"
    chkRemoveBorders.Caption = "Remove all borders"
    chkConvertToText.Caption = "Convert single-column tables to plain text"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
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
Private Sub chkRemoveEmptyRows_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkRemoveEmptyCols_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkNormalizePadding_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkStripDirectFormat_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkNormalizeBorders_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkRemoveBorders_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkConvertToText_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub cmdRiskPlacementTableRows_Click(): ShowToolRiskChoiceExplanation "Remove empty rows", "Structure change", "Deletes rows with no visible cell content." : End Sub
Private Sub cmdRiskPlacementTableCols_Click(): ShowToolRiskChoiceExplanation "Remove empty columns", "Structure change", "Deletes columns with no visible cell content." : End Sub
Private Sub cmdRiskPlacementTablePadding_Click(): ShowToolRiskChoiceExplanation "Normalize cell padding", "Formatting change", "Resets table cell margins to the suite defaults." : End Sub
Private Sub cmdRiskPlacementTableStripFormat_Click(): ShowToolRiskChoiceExplanation "Strip cell formatting", "Formatting change", "Removes direct formatting inside table cells." : End Sub
Private Sub cmdRiskPlacementTableBorders_Click(): ShowToolRiskChoiceExplanation "Normalize borders", "Formatting change", "Turns table borders back on using standard border behavior." : End Sub
Private Sub cmdRiskPlacementTableRemoveBorders_Click(): ShowToolRiskChoiceExplanation "Remove borders", "Formatting change", "Removes visible table borders." : End Sub
Private Sub cmdRiskPlacementTableConvertToText_Click(): ShowToolRiskChoiceExplanation "Convert table to text", "Structure change", "Replaces single-column tables with plain text." : End Sub
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
    If previewOnly Then
        Dim cnt As Long
        Dim previewEmptyRows As Long
        Dim previewCols As Long
        Dim previewPadding As Long
        Dim previewDirectFmt As Long
        Dim previewNormalizeBorders As Long
        Dim previewRemoveBorders As Long
        Dim previewConvertText As Long
        Dim t As Variant
        For Each t In tablesInScope
            Set tbl = t
            If doEmptyRows Then
                Dim previewRow As Row
                For Each previewRow In tbl.Rows
                    If IsRowEmpty(previewRow) Then previewEmptyRows = previewEmptyRows + 1
                Next previewRow
            End If
            If doEmptyCols Then
                Dim previewColIndex As Long
                For previewColIndex = 1 To tbl.Columns.Count
                    If IsColEmpty(tbl, previewColIndex) Then previewCols = previewCols + 1
                Next previewColIndex
            End If
            If doPadding Then previewPadding = previewPadding + 1
            If doDirectFmt Then previewDirectFmt = previewDirectFmt + 1
            If doBorders Then previewNormalizeBorders = previewNormalizeBorders + 1
            If doRemoveBorders Then previewRemoveBorders = previewRemoveBorders + 1
            If doConvertText And tbl.Columns.Count = 1 Then previewConvertText = previewConvertText + 1
        Next t

        For Each t In tablesInScope
            Set tbl = t
            If doEmptyRows Then
                Dim rw As Row
                For Each rw In tbl.Rows
                    If IsRowEmpty(rw) Then ApplyPreviewShading rw.Range: cnt = cnt + 1
                Next rw
            End If
            If doEmptyCols Then
                Dim previewShadeCol As Long
                Dim previewShadeRow As Long
                For previewShadeCol = 1 To tbl.Columns.Count
                    If IsColEmpty(tbl, previewShadeCol) Then
                        For previewShadeRow = 1 To tbl.Rows.Count
                            On Error Resume Next
                            ApplyPreviewShading tbl.Cell(previewShadeRow, previewShadeCol).Range
                            On Error GoTo 0
                        Next previewShadeRow
                    End If
                Next previewShadeCol
            End If
        Next t
        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Empty rows", previewEmptyRows, False, (Not doEmptyRows)
        AddPreviewSummaryRow previewRows, "Empty columns", previewCols, False, (Not doEmptyCols)
        AddPreviewSummaryRow previewRows, "Cell padding", previewPadding, True, (Not doPadding)
        AddPreviewSummaryRow previewRows, "Direct cell formatting", previewDirectFmt, True, (Not doDirectFmt)
        AddPreviewSummaryRow previewRows, "Border normalization", previewNormalizeBorders, True, (Not doBorders)
        AddPreviewSummaryRow previewRows, "Border removal", previewRemoveBorders, True, (Not doRemoveBorders)
        AddPreviewSummaryRow previewRows, "Tables to text", previewConvertText, True, (Not doConvertText)
        ShowPreviewActionsSummary Me, "Table Cleaner", previewRows
        Exit Sub
    End If
    MarkCleanupStart "Table Cleaner"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Table Cleaner"
    On Error GoTo RunErr
    Dim cntRows As Long, cntCols As Long, cntTables As Long, cntConverted As Long
    Dim tVar As Variant
    For Each tVar In tablesInScope
        UpdateCleanupProgress "Cleaning tables", cntTables + 1, tablesInScope.Count
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
            tbl.Range.Font.Reset
            tbl.Range.ParagraphFormat.Reset
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
