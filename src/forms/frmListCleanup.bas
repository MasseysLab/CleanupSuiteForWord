Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  List Normalizer
'  Normalizes list formatting: converts hyphen/asterisk lines to proper
'  Word bullet lists, standardizes existing bullet styles, converts
'  manually-typed numbered lists (1. 2. 3.) to Word auto-numbering,
'  and fixes list indentation to the standard 0.25-inch depth.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optAll.GroupName = "ListMode"
    optBullets.GroupName = "ListMode"
    optNumbering.GroupName = "ListMode"
    optIndent.GroupName = "ListMode"
    optCustom.GroupName = "ListMode"
    optScopeDocument.GroupName = "ListScope"
    optScopeSelection.GroupName = "ListScope"
    optAll.Value = True: UpdateCustomVisibility: ClearCustomChecks
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
    optAll.Caption = "All list fixes"
    optBullets.Caption = "Convert manual bullet lists"
    optNumbering.Caption = "Convert manual numbered lists"
    optIndent.Caption = "Fix list indentation"
    optCustom.Caption = "Custom (choose below)"
    chkNormalizeBullets.Caption = "Convert manual hyphens / asterisks to List Bullet style"
    chkNormalizeNumbering.Caption = "Convert manual numbers to List Number style"
    chkFixIndent.Caption = "Fix list indentation"
    chkHyphenToBullets.Caption = "Convert hyphen-dash lines to bullets"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub UpdateCustomVisibility(): fraCustom.Visible = False: End Sub
Private Sub optAll_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub optBullets_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub optNumbering_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub optIndent_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub optCustom_Click(): UpdateCustomVisibility: LayoutCleanupToolForm Me: End Sub
Private Sub ClearCustomChecks(): chkNormalizeBullets.Value = False: chkNormalizeNumbering.Value = False: chkFixIndent.Value = False: chkHyphenToBullets.Value = False: End Sub
Private Sub cmdSelectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    chkNormalizeBullets.Value = True
    chkNormalizeNumbering.Value = True
    chkFixIndent.Value = True
    chkHyphenToBullets.Value = True
    LayoutCleanupToolForm Me
End Sub
Private Sub cmdDeselectAll_Click()
    optCustom.Value = True
    UpdateCustomVisibility
    ClearCustomChecks
    LayoutCleanupToolForm Me
End Sub
Private Sub chkNormalizeBullets_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkNormalizeNumbering_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkFixIndent_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub chkHyphenToBullets_Click(): LayoutCleanupToolForm Me: End Sub
Private Sub cmdPreview_Click()
    PreviewFromPanel
End Sub
Public Sub PreviewFromPanel()
    chkPreviewOnly.Value = True
    cmdRun_Click
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
    If Not GuardBeforeCleanup("List Normalizer") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doBullets As Boolean, doNumbering As Boolean, doIndent As Boolean, doHyphen As Boolean
    If optAll.Value Then doBullets = True: doNumbering = True: doIndent = True: doHyphen = True
    If optBullets.Value Then doBullets = True
    If optNumbering.Value Then doNumbering = True
    If optIndent.Value Then doIndent = True
    If optCustom.Value Then doBullets = chkNormalizeBullets.Value: doNumbering = chkNormalizeNumbering.Value: doIndent = chkFixIndent.Value: doHyphen = chkHyphenToBullets.Value
    If Not (doBullets Or doNumbering Or doIndent Or doHyphen) Then MsgBox "No list options selected.", vbInformation: Exit Sub
    ' Preview: highlight candidate paragraphs
    Dim p As Paragraph, txt As String, changed As Long: changed = 0
    For Each p In ActiveDocument.Paragraphs
        txt = Trim$(p.Range.Text)
        If Right$(txt, 1) = vbCr Then txt = Left$(txt, Len(txt) - 1)
        If doHyphen And (Left$(txt, 2) = "- " Or Left$(txt, 2) = "* ") Then p.Range.HighlightColorIndex = wdYellow: changed = changed + 1
        If doNumbering And txt Like "[0-9]*.*" Then p.Range.HighlightColorIndex = wdYellow: changed = changed + 1
        If doBullets And (Left$(txt, 2) = Chr(149) & " " Or Left$(txt, 2) = "* ") Then p.Range.HighlightColorIndex = wdYellow: changed = changed + 1
    Next p
    If previewOnly Then
        ShowPreviewActions Me, "List Normalizer", "Preview complete. " & changed & " paragraphs highlighted."
        Exit Sub
    End If
    MarkCleanupStart "List Normalizer"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - List Normalizer"
    On Error GoTo RunErr
    Dim cntHyphen As Long, cntBullets As Long, cntNumbering As Long, cntIndent As Long
    If doHyphen    Then cntHyphen    = ConvertHyphenListsToBullets(targetRange)
    If doBullets   Then cntBullets   = NormalizeBullets(targetRange)
    If doNumbering Then cntNumbering = NormalizeNumbering(targetRange)
    If doIndent    Then cntIndent    = FixListIndentation(targetRange)
    RemoveAllHighlighting targetRange
    Dim results As Collection: Set results = New Collection
    If doHyphen    Then results.Add "Hyphen lists converted to bullets: " & cntHyphen
    If doBullets   Then results.Add "Bullet styles normalized:          " & cntBullets
    If doNumbering Then results.Add "Numbered items normalized:         " & cntNumbering
    If doIndent    Then results.Add "List indents fixed:                " & cntIndent
    ShowCleanupReport "List Normalization", results
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
Private Function ConvertHyphenListsToBullets(scopeRange As Range) As Long
    Dim p As Paragraph, txt As String, cnt As Long
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= scopeRange.Start And p.Range.End <= scopeRange.End Then
            txt = Trim$(p.Range.Text)
            If Right$(txt, 1) = vbCr Then txt = Left$(txt, Len(txt) - 1)
            If Left$(txt, 2) = "- " Or Left$(txt, 2) = "* " Then
                p.Range.Text = Mid$(txt, 3) & vbCr
                p.Range.ListFormat.ApplyBulletDefault
                cnt = cnt + 1
            End If
        End If
    Next p
    ConvertHyphenListsToBullets = cnt
End Function
Private Function NormalizeBullets(scopeRange As Range) As Long
    Dim p As Paragraph, cnt As Long
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= scopeRange.Start And p.Range.End <= scopeRange.End Then
            If p.Range.ListFormat.ListType <> wdListNoNumbering Then
                If p.Range.ListFormat.ListType = wdListBullet Then
                    p.Range.ListFormat.ApplyBulletDefault
                    cnt = cnt + 1
                End If
            End If
        End If
    Next p
    NormalizeBullets = cnt
End Function
Private Function NormalizeNumbering(scopeRange As Range) As Long
    Dim p As Paragraph, txt As String, cnt As Long
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= scopeRange.Start And p.Range.End <= scopeRange.End Then
            txt = Trim$(p.Range.Text)
            If Right$(txt, 1) = vbCr Then txt = Left$(txt, Len(txt) - 1)
            If txt Like "[0-9]*.*" Then
                p.Range.ListFormat.ApplyNumberDefault
                Dim pos As Long: pos = InStr(txt, "." )
                If pos > 0 Then p.Range.Text = LTrim$(Mid$(txt, pos + 1)) & vbCr
                cnt = cnt + 1
            End If
        End If
    Next p
    NormalizeNumbering = cnt
End Function
Private Function FixListIndentation(scopeRange As Range) As Long
    Dim p As Paragraph, cnt As Long
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= scopeRange.Start And p.Range.End <= scopeRange.End Then
            If p.Range.ListFormat.ListType <> wdListNoNumbering Then
                p.LeftIndent = InchesToPoints(0.25)
                cnt = cnt + 1
            End If
        End If
    Next p
    FixListIndentation = cnt
End Function
