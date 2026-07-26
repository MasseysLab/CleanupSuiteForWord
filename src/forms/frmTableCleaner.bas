Option Explicit
Private mStructuralPreviewCompleted As Boolean
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
    mStructuralPreviewCompleted = False
    chkRemoveEmptyRows.Value = False
    chkRemoveEmptyCols.Value = False
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
Private Function CollectEligibleEmptyRowIndexes(ByVal sourceTable As Table, ByRef protectedOrSkipped As Long) As Collection
    Dim indexes As Collection
    Dim rowIndex As Long
    Dim tableRow As Row

    Set indexes = New Collection
    If Not CleanupTableStructureIsSupported(sourceTable) Then
        protectedOrSkipped = protectedOrSkipped + 1
        Set CollectEligibleEmptyRowIndexes = indexes
        Exit Function
    End If

    For rowIndex = 1 To sourceTable.Rows.Count
        Set tableRow = sourceTable.Rows(rowIndex)
        If Not CleanupTableRowHasMeaningfulContent(tableRow) Then
            If sourceTable.Rows.Count <= 1 Or tableRow.HeadingFormat <> 0 Then
                protectedOrSkipped = protectedOrSkipped + 1
            Else
                indexes.Add rowIndex
            End If
        End If
    Next rowIndex

    If indexes.Count >= sourceTable.Rows.Count Then
        protectedOrSkipped = protectedOrSkipped + indexes.Count
        Set indexes = New Collection
    End If
    Set CollectEligibleEmptyRowIndexes = indexes
End Function

Private Function CollectEligibleEmptyColumnIndexes(ByVal sourceTable As Table, ByRef protectedOrSkipped As Long) As Collection
    Dim indexes As Collection
    Dim columnIndex As Long

    Set indexes = New Collection
    If Not CleanupTableStructureIsSupported(sourceTable) Then
        protectedOrSkipped = protectedOrSkipped + 1
        Set CollectEligibleEmptyColumnIndexes = indexes
        Exit Function
    End If

    For columnIndex = 1 To sourceTable.Columns.Count
        If Not CleanupTableColumnHasMeaningfulContent(sourceTable, columnIndex) Then
            If sourceTable.Columns.Count <= 1 Then
                protectedOrSkipped = protectedOrSkipped + 1
            Else
                indexes.Add columnIndex
            End If
        End If
    Next columnIndex

    If indexes.Count >= sourceTable.Columns.Count Then
        protectedOrSkipped = protectedOrSkipped + indexes.Count
        Set indexes = New Collection
    End If
    Set CollectEligibleEmptyColumnIndexes = indexes
End Function

Private Function TableRowStillEligible(ByVal sourceTable As Table, ByVal rowIndex As Long) As Boolean
    On Error GoTo SafeExit
    If Not CleanupTableStructureIsSupported(sourceTable) Then Exit Function
    If sourceTable.Rows.Count <= 1 Then Exit Function
    If rowIndex < 1 Or rowIndex > sourceTable.Rows.Count Then Exit Function
    If sourceTable.Rows(rowIndex).HeadingFormat <> 0 Then Exit Function
    TableRowStillEligible = Not CleanupTableRowHasMeaningfulContent(sourceTable.Rows(rowIndex))
SafeExit:
End Function

Private Function TableColumnStillEligible(ByVal sourceTable As Table, ByVal columnIndex As Long) As Boolean
    On Error GoTo SafeExit
    If Not CleanupTableStructureIsSupported(sourceTable) Then Exit Function
    If sourceTable.Columns.Count <= 1 Then Exit Function
    If columnIndex < 1 Or columnIndex > sourceTable.Columns.Count Then Exit Function
    TableColumnStillEligible = Not CleanupTableColumnHasMeaningfulContent(sourceTable, columnIndex)
SafeExit:
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
    Dim partiallyScopedTables As Long
    Set tablesInScope = New Collection
    For Each tbl In ActiveDocument.Tables
        If tbl.Range.Start >= targetRange.Start And tbl.Range.End <= targetRange.End Then
            tablesInScope.Add tbl
        ElseIf tbl.Range.End > targetRange.Start And tbl.Range.Start < targetRange.End Then
            partiallyScopedTables = partiallyScopedTables + 1
        End If
    Next tbl
    If tablesInScope.Count = 0 Then MsgBox "No tables found in the selected scope.", vbInformation: Unload Me: Exit Sub
    If previewOnly Then
        Dim previewEmptyRows As Long
        Dim previewProtectedRows As Long
        Dim previewCols As Long
        Dim previewProtectedCols As Long
        Dim previewPadding As Long
        Dim previewDirectFmt As Long
        Dim previewNormalizeBorders As Long
        Dim previewRemoveBorders As Long
        Dim previewConvertText As Long
        Dim previewSkippedConvert As Long
        Dim previewProtectedConvert As Long
        Dim previewMarkTable As Boolean
        Dim t As Variant
        For Each t In tablesInScope
            Set tbl = t
            If doEmptyRows Then
                Dim rowIndexes As Collection
                Set rowIndexes = CollectEligibleEmptyRowIndexes(tbl, previewProtectedRows)
                previewEmptyRows = previewEmptyRows + rowIndexes.Count
                Dim previewRowCandidate As Long
                For previewRowCandidate = 1 To rowIndexes.Count
                    ApplyPreviewShading tbl.Rows(CLng(rowIndexes(previewRowCandidate))).Range
                Next previewRowCandidate
            End If
            If doEmptyCols Then
                Dim columnIndexes As Collection
                Set columnIndexes = CollectEligibleEmptyColumnIndexes(tbl, previewProtectedCols)
                previewCols = previewCols + columnIndexes.Count
                Dim previewColumnCandidate As Long
                Dim previewColumnRow As Long
                For previewColumnCandidate = 1 To columnIndexes.Count
                    For previewColumnRow = 1 To tbl.Rows.Count
                        On Error Resume Next
                        ApplyPreviewShading tbl.Cell(previewColumnRow, CLng(columnIndexes(previewColumnCandidate))).Range
                        On Error GoTo 0
                    Next previewColumnRow
                Next previewColumnCandidate
            End If
            If doPadding Then previewPadding = previewPadding + 1
            If doDirectFmt Then previewDirectFmt = previewDirectFmt + 1
            If doBorders Then previewNormalizeBorders = previewNormalizeBorders + 1
            If doRemoveBorders Then previewRemoveBorders = previewRemoveBorders + 1
            If doConvertText Then
                If Not CleanupTableStructureIsSupported(tbl) Then
                    previewProtectedConvert = previewProtectedConvert + 1
                ElseIf tbl.Columns.Count = 1 Then
                    previewConvertText = previewConvertText + 1
                Else
                    previewSkippedConvert = previewSkippedConvert + 1
                End If
            End If
            previewMarkTable = doPadding Or doDirectFmt Or doBorders Or doRemoveBorders
            If doConvertText And tbl.Columns.Count = 1 And CleanupTableStructureIsSupported(tbl) Then
                previewMarkTable = True
            End If
            If previewMarkTable Then
                ApplyPreviewMinimalMarker tbl.Range, True
            End If
        Next t
        Dim previewRows As Collection
        Set previewRows = NewPreviewSummaryRows()
        AddPreviewSummaryRow previewRows, "Empty rows", previewEmptyRows, False, (Not doEmptyRows)
        AddPreviewSummaryRow previewRows, "Protected/skipped rows", previewProtectedRows, True, (Not doEmptyRows)
        AddPreviewSummaryRow previewRows, "Empty columns", previewCols, False, (Not doEmptyCols)
        AddPreviewSummaryRow previewRows, "Protected/skipped columns", previewProtectedCols, True, (Not doEmptyCols)
        AddPreviewSummaryRow previewRows, "Cell padding", previewPadding, False, (Not doPadding)
        AddPreviewSummaryRow previewRows, "Direct cell formatting", previewDirectFmt, False, (Not doDirectFmt)
        AddPreviewSummaryRow previewRows, "Border normalization", previewNormalizeBorders, False, (Not doBorders)
        AddPreviewSummaryRow previewRows, "Border removal", previewRemoveBorders, False, (Not doRemoveBorders)
        AddPreviewSummaryRow previewRows, "Tables to text", previewConvertText, False, (Not doConvertText)
        AddPreviewSummaryRow previewRows, "Skipped multi-column tables", previewSkippedConvert, True, (Not doConvertText)
        AddPreviewSummaryRow previewRows, "Protected/unsupported conversions", previewProtectedConvert, True, (Not doConvertText)
        AddPreviewSummaryRow previewRows, "Partially scoped tables", partiallyScopedTables, True
        mStructuralPreviewCompleted = True
        ShowPreviewActionsSummary Me, "Table Cleaner", previewRows
        Exit Sub
    End If

    If (doEmptyRows Or doEmptyCols Or doConvertText) And Not mStructuralPreviewCompleted Then
        MsgBox "Preview these structural table choices before applying them.", vbInformation, "Preview Required"
        Exit Sub
    End If

    Dim structuralRowCount As Long
    Dim structuralColumnCount As Long
    Dim structuralConvertCount As Long
    Dim applyProtectedRows As Long
    Dim applyProtectedColumns As Long
    Dim applyRowIndexes As Collection
    Dim applyColumnIndexes As Collection
    Dim structuralTable As Variant
    For Each structuralTable In tablesInScope
        Set tbl = structuralTable
        If doEmptyRows Then
            Set applyRowIndexes = CollectEligibleEmptyRowIndexes(tbl, applyProtectedRows)
            structuralRowCount = structuralRowCount + applyRowIndexes.Count
        End If
        If doEmptyCols Then
            Set applyColumnIndexes = CollectEligibleEmptyColumnIndexes(tbl, applyProtectedColumns)
            structuralColumnCount = structuralColumnCount + applyColumnIndexes.Count
        End If
        If doConvertText And tbl.Columns.Count = 1 And CleanupTableStructureIsSupported(tbl) Then
            structuralConvertCount = structuralConvertCount + 1
        End If
    Next structuralTable

    If structuralRowCount + structuralColumnCount + structuralConvertCount > 0 Then
        Dim structuralMessage As String
        structuralMessage = "Apply these structural table changes?" & vbCrLf & vbCrLf & _
                            structuralRowCount & " empty row(s)" & vbCrLf & _
                            structuralColumnCount & " empty column(s)" & vbCrLf & _
                            structuralConvertCount & " single-column table conversion(s)" & vbCrLf & vbCrLf & _
                            "Protected or unsupported structures will be left unchanged."
        If MsgBox(structuralMessage, vbExclamation + vbYesNo + vbDefaultButton2, "Confirm Table Structure Changes") <> vbYes Then Exit Sub
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
        Dim convertThisTable As Boolean
        convertThisTable = (doConvertText And tbl.Columns.Count = 1 And CleanupTableStructureIsSupported(tbl))
        If doEmptyRows Then
            Set applyRowIndexes = CollectEligibleEmptyRowIndexes(tbl, applyProtectedRows)
            Dim rowCandidateIndex As Long
            For rowCandidateIndex = applyRowIndexes.Count To 1 Step -1
                If TableRowStillEligible(tbl, CLng(applyRowIndexes(rowCandidateIndex))) Then
                    tbl.Rows(CLng(applyRowIndexes(rowCandidateIndex))).Delete
                    cntRows = cntRows + 1
                End If
            Next rowCandidateIndex
        End If
        If doEmptyCols Then
            Set applyColumnIndexes = CollectEligibleEmptyColumnIndexes(tbl, applyProtectedColumns)
            Dim columnCandidateIndex As Long
            For columnCandidateIndex = applyColumnIndexes.Count To 1 Step -1
                If TableColumnStillEligible(tbl, CLng(applyColumnIndexes(columnCandidateIndex))) Then
                    tbl.Columns(CLng(applyColumnIndexes(columnCandidateIndex))).Delete
                    cntCols = cntCols + 1
                End If
            Next columnCandidateIndex
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
        If convertThisTable Then
            tbl.ConvertToText Separator:=wdSeparateByTabs
            cntConverted = cntConverted + 1
        End If
    Next tVar
    MarkCleanupEnd
    MarkCleanupToolApplied
    undoRec.EndCustomRecord
    Unload Me
    Exit Sub
RunErr:
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    On Error Resume Next
    MarkCleanupEnd
    undoRec.EndCustomRecord
    On Error GoTo 0
    Dim errMsg As String
    errMsg = "An unexpected error occurred: " & originalErrNumber & " - " & originalErrDescription
    errMsg = errMsg & vbCrLf & vbCrLf & _
             "Some changes may have been made to the document." & vbCrLf & _
             "Use Word's Undo command (Ctrl+Z) after closing this message if you want to roll them back."
    MsgBox errMsg, vbCritical, "Cleanup Error"
End Sub
