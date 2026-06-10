Option Explicit
' --------------------------------------------------------------
'  Break Normalizer
'  Collapses consecutive section and page breaks -- a common problem in
'  documents assembled from multiple files -- and optionally converts all
'  section breaks to a chosen type: Next Page, Continuous, Even Page,
'  or Odd Page.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optConvertNextPage.GroupName = "BreakMode"
    optConvertContinuous.GroupName = "BreakMode"
    optConvertEvenPage.GroupName = "BreakMode"
    optConvertOddPage.GroupName = "BreakMode"
    optScopeDocument.GroupName = "BreakScope"
    optScopeSelection.GroupName = "BreakScope"
    chkCollapseSectionBreaks.Value = True
    chkCollapsePageBreaks.Value = True
    chkConvertSectionBreaks.Value = False
    fraConvertTo.Enabled = False
    optConvertContinuous.Value = True
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
    chkCollapseSectionBreaks.Caption = "Collapse consecutive section breaks"
    chkCollapsePageBreaks.Caption = "Collapse consecutive page breaks"
    chkConvertSectionBreaks.Caption = "Convert section breaks to page breaks"
    optConvertNextPage.Caption = "Next Page"
    optConvertContinuous.Caption = "Continuous"
    optConvertEvenPage.Caption = "Even Page"
    optConvertOddPage.Caption = "Odd Page"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Break Normalizer" & vbCrLf & vbCrLf
    h = h & "Fixes problems with section and page breaks.  Most common in documents" & vbCrLf
    h = h & "assembled from multiple files, where each source file may have had its" & vbCrLf
    h = h & "own trailing section break." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Collapse section breaks  --  Reduces consecutive section breaks to one." & vbCrLf
    h = h & "  Collapse page breaks     --  Reduces consecutive manual page breaks to one." & vbCrLf
    h = h & "  Convert section breaks   --  Changes all section break types in scope to" & vbCrLf
    h = h & "                               the type you choose below." & vbCrLf
    h = h & vbCrLf
    h = h & "CONVERSION TYPES" & vbCrLf
    h = h & "  Next Page   --  Section starts at the top of the next page." & vbCrLf
    h = h & "  Continuous  --  Section continues on the same page (no forced break)." & vbCrLf
    h = h & "  Even Page   --  Section starts on the next even-numbered page." & vbCrLf
    h = h & "  Odd Page    --  Section starts on the next odd-numbered page." & vbCrLf
    h = h & vbCrLf
    h = h & "NOTE: The first section of a document is never converted -- it has no" & vbCrLf
    h = h & "preceding break to change." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights consecutive break sequences in yellow." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Break Normalizer"
End Sub
Private Sub chkConvertSectionBreaks_Click()
    fraConvertTo.Enabled = chkConvertSectionBreaks.Value
End Sub
Private Sub cmdRun_Click()
    If ActiveDocument.ProtectionType <> wdNoProtection Then
        MsgBox "This document is protected. Please remove protection before running cleanup.", vbExclamation, "Document Protected": Exit Sub
    End If
    If Not GuardBeforeCleanup("Break Normalizer") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doCollSect As Boolean: doCollSect = chkCollapseSectionBreaks.Value
    Dim doCollPage As Boolean: doCollPage = chkCollapsePageBreaks.Value
    Dim doConvert As Boolean: doConvert = chkConvertSectionBreaks.Value
    If Not (doCollSect Or doCollPage Or doConvert) Then MsgBox "No break options selected.", vbInformation: Exit Sub
    ' Determine target section start type for conversion
    Dim targetSectStart As Long: targetSectStart = wdSectionContinuous
    If optConvertNextPage.Value Then targetSectStart = wdSectionNewPage
    If optConvertContinuous.Value Then targetSectStart = wdSectionContinuous
    If optConvertEvenPage.Value Then targetSectStart = wdSectionEvenPage
    If optConvertOddPage.Value Then targetSectStart = wdSectionOddPage
    ' Preview: highlight consecutive section/page breaks
    Dim cnt As Long: cnt = 0
    If doCollSect Or doCollPage Then
        With targetRange.Find
            .ClearFormatting: .Replacement.ClearFormatting
            If doCollSect Then .Text = "^b^b": .Replacement.Text = "^&": .Replacement.Highlight = True: .Execute Replace:=wdReplaceAll: cnt = cnt + 1
            If doCollPage Then .Text = "^m^m": .Replacement.Text = "^&": .Replacement.Highlight = True: .Execute Replace:=wdReplaceAll: cnt = cnt + 1
        End With
    End If
    MsgBox "Preview complete. Break sequences highlighted.", vbInformation
    If previewOnly Then Unload Me: Exit Sub
    MarkCleanupStart "Break Normalizer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Break Normalizer"
    On Error GoTo RunErr
    Dim cntSect As Long, cntPage As Long, cntConv As Long
    If doCollSect Then
        Do
            With targetRange.Find: .ClearFormatting: .Replacement.ClearFormatting: .Text = "^b^b": .Replacement.Text = "^b": Do While .Execute(Replace:=wdReplaceOne): cntSect = cntSect + 1: Loop: End With
        Loop While targetRange.Find.Execute(FindText:="^b^b", Forward:=True)
    End If
    If doCollPage Then
        Do
            With targetRange.Find: .ClearFormatting: .Replacement.ClearFormatting: .Text = "^m^m": .Replacement.Text = "^m": Do While .Execute(Replace:=wdReplaceOne): cntPage = cntPage + 1: Loop: End With
        Loop While targetRange.Find.Execute(FindText:="^m^m", Forward:=True)
    End If
    If doConvert Then
        Dim sec As Section
        For Each sec In ActiveDocument.Sections
            If sec.Range.Start >= targetRange.Start And sec.Range.End <= targetRange.End Then
                If sec.Index > 1 Then
                    sec.PageSetup.SectionStart = targetSectStart
                    cntConv = cntConv + 1
                End If
            End If
        Next sec
    End If
    RemoveAllHighlighting targetRange
    Dim results As Collection: Set results = New Collection
    If doCollSect Then results.Add "Consecutive section breaks collapsed: " & cntSect
    If doCollPage Then results.Add "Consecutive page breaks collapsed: " & cntPage
    If doConvert Then results.Add "Section break types converted: " & cntConv
    ShowCleanupReport "Break Normalizer", results
    undoRec.EndCustomRecord
    MarkCleanupEnd
    Unload Me
    Exit Sub
RunErr:
    On Error Resume Next: undoRec.EndCustomRecord: On Error GoTo 0
    MarkCleanupEnd
    Dim errMsg As String
    errMsg = "An unexpected error occurred: " & Err.Number & " - " & Err.Description
    errMsg = errMsg & vbCrLf & vbCrLf & "Some changes may have been made to the document." & vbCrLf & "Would you like to undo all changes made before the error?"
    If MsgBox(errMsg, vbCritical + vbYesNo, "Cleanup Error") = vbYes Then
        On Error Resume Next: Application.Undo: On Error GoTo 0
    End If
End Sub