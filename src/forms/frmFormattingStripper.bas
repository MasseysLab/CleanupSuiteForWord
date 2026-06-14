Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub
' --------------------------------------------------------------
'  Direct Formatting Stripper
'  Removes manual (direct) character and paragraph formatting while
'  PRESERVING the paragraph style.  Unlike Word built-in Clear Formatting,
'  which resets text to the Normal style, this tool returns each paragraph
'  to exactly what its applied style specifies.  Three emphasis modes
'  balance completeness against speed.  Drop caps are preserved by default;
'  untick 'Preserve drop caps' to reset them too.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    chkResetChar.Value = True
    chkResetPara.Value = True
    chkPreserveHighlight.Value = True
    chkPreserveDropCaps.Value = True
    chkPreserveHighlight.Caption = "Preserve highlighting"
    chkPreserveDropCaps.Caption = "Preserve drop caps"
    optEmphQuick.GroupName = "StripEmphasis"
    optEmphThorough.GroupName = "StripEmphasis"
    optEmphStrip.GroupName = "StripEmphasis"
    optEmphQuick.Value = True
    lblSpeedWarning.Caption = "Quick mode is fast and preserves emphasis that applies to whole paragraphs."
    optScopeDocument.GroupName = "StripScope"
    optScopeSelection.GroupName = "StripScope"
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
    chkResetChar.Caption = "Strip direct character formatting"
    chkResetPara.Caption = "Strip direct paragraph formatting"
    optEmphQuick.Caption = "Quick  (preserve paragraph-level emphasis)"
    optEmphThorough.Caption = "Thorough  (preserve word-level emphasis)"
    optEmphStrip.Caption = "Strip  (remove all direct formatting)"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub optEmphQuick_Click()
    lblSpeedWarning.Caption = "Quick mode is fast and preserves emphasis that applies to whole paragraphs."
End Sub
Private Sub optEmphThorough_Click()
    lblSpeedWarning.Caption = "Thorough mode examines every run individually and preserves word-level emphasis.  On a long document this may take a minute or more."
End Sub
Private Sub optEmphStrip_Click()
    lblSpeedWarning.Caption = "Strip mode removes all direct formatting, including bold and italic emphasis."
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Direct Formatting Stripper" & vbCrLf & vbCrLf
    h = h & "Removes manual formatting overrides and returns text to what its " & _
            "paragraph style defines.  Unlike Word Clear Formatting (which resets" & vbCrLf
    h = h & "text to the Normal style), this tool keeps your styles intact and strips " & _
            "only the direct formatting layered on top." & vbCrLf
    h = h & vbCrLf
    h = h & "WHAT TO RESET" & vbCrLf
    h = h & "  Reset character formatting  --  font face, size, colour, spacing." & vbCrLf
    h = h & "  Reset paragraph formatting  --  indents, line spacing, alignment." & vbCrLf
    h = h & vbCrLf
    h = h & "EMPHASIS HANDLING" & vbCrLf
    h = h & "  Quick clean      --  Fast.  Preserves bold/italic that applies to a" & vbCrLf
    h = h & "                       whole paragraph.  Best for most documents." & vbCrLf
    h = h & "  Thorough clean   --  Preserves bold/italic on individual words." & vbCrLf
    h = h & "                       Slower on large documents (examines every run)." & vbCrLf
    h = h & "  Strip everything --  Removes all emphasis too, for a full reset." & vbCrLf
    h = h & vbCrLf
    h = h & "  Preserve highlighting keeps text highlight colours through the reset." & vbCrLf
    h = h & "  Preserve drop caps keeps decorative dropped capitals; untick to reset " & _
            "  them to normal body text as part of the strip." & vbCrLf
    h = h & vbCrLf
    h = h & "NOTE: Quick mode restores emphasis only when it is uniform across the " & _
            "whole paragraph.  For a paragraph with one bold word among normal text," & vbCrLf
    h = h & "use Thorough mode to preserve that word precisely." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights paragraphs carrying direct formatting." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Direct Formatting Stripper"
End Sub
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
    If Not GuardBeforeCleanup("Formatting Stripper") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doResetChar As Boolean: doResetChar = chkResetChar.Value
    Dim doResetPara As Boolean: doResetPara = chkResetPara.Value
    Dim preserveHL As Boolean: preserveHL = chkPreserveHighlight.Value
    Dim preserveDC As Boolean: preserveDC = chkPreserveDropCaps.Value
    If Not (doResetChar Or doResetPara) Then MsgBox "Nothing selected to reset. Tick at least one reset option.", vbInformation: Exit Sub
    Dim p As Paragraph
    If previewOnly Then
        Dim pcnt As Long: pcnt = 0
        For Each p In ActiveDocument.Paragraphs
            If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
                If HasDirectFormatting(p, doResetChar, doResetPara) Then p.Range.HighlightColorIndex = wdYellow: pcnt = pcnt + 1
            End If
        Next p
        ShowPreviewActions Me, "Formatting Stripper", "Preview complete. " & pcnt & " paragraphs with direct formatting highlighted."
        Exit Sub
    End If
    Dim emphMode As Integer: emphMode = 0
    If optEmphThorough.Value Then emphMode = 1
    If optEmphStrip.Value Then emphMode = 2
    If emphMode = 1 Then
        Dim totalChars As Long: totalChars = 0
        For Each p In ActiveDocument.Paragraphs
            If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then totalChars = totalChars + p.Range.Characters.Count
        Next p
        If totalChars > 20000 Then
            If MsgBox("This document has " & totalChars & " characters in scope." & vbCrLf & "Thorough cleaning may take a while. Continue?", vbQuestion + vbYesNo, "Thorough Clean") = vbNo Then Exit Sub
        End If
    End If
    MarkCleanupStart "Formatting Stripper"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Formatting Stripper"
    On Error GoTo RunErr
    Application.ScreenUpdating = False
    Dim cntReset As Long: cntReset = 0
    Dim processed As Long: processed = 0
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
            ApplyStripToParagraph p, doResetChar, doResetPara, emphMode, preserveHL, preserveDC
            cntReset = cntReset + 1
            processed = processed + 1
            If emphMode = 1 And (processed Mod 25 = 0) Then Application.StatusBar = "Formatting Stripper: " & processed & " paragraphs processed..."
        End If
    Next p
    Application.ScreenUpdating = True
    Application.StatusBar = False
    Dim results As Collection: Set results = New Collection
    results.Add "Paragraphs processed: " & cntReset
    If doResetChar Then results.Add "Character formatting reset"
    If doResetPara Then results.Add "Paragraph formatting reset"
    Select Case emphMode
        Case 0: results.Add "Emphasis: whole-paragraph preserved"
        Case 1: results.Add "Emphasis: word-level preserved"
        Case 2: results.Add "Emphasis: removed"
    End Select
    ShowCleanupReport "Formatting Stripper", results
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
    Application.StatusBar = False
    On Error Resume Next: undoRec.EndCustomRecord: On Error GoTo 0
    MarkCleanupEnd
    Dim errMsg As String
    errMsg = "An unexpected error occurred: " & originalErrNumber & " - " & originalErrDescription
    errMsg = errMsg & vbCrLf & vbCrLf & _
             "Some changes may have been made to the document." & vbCrLf & _
             "Use Word's Undo command (Ctrl+Z) after closing this message if you want to roll them back."
    MsgBox errMsg, vbCritical, "Cleanup Error"
End Sub
Private Function HasDirectFormatting(p As Paragraph, doChar As Boolean, doPara As Boolean) As Boolean
    HasDirectFormatting = False
    On Error Resume Next
    Dim st As Style
    Set st = ActiveDocument.Styles(p.Style)
    If st Is Nothing Then Exit Function
    If doChar Then
        If p.Range.Font.Name <> st.Font.Name Then HasDirectFormatting = True
        If p.Range.Font.Size <> st.Font.Size And p.Range.Font.Size > 0 Then HasDirectFormatting = True
        If p.Range.Font.Bold <> st.Font.Bold Then HasDirectFormatting = True
        If p.Range.Font.Italic <> st.Font.Italic Then HasDirectFormatting = True
        If p.Range.Font.Color <> st.Font.Color Then HasDirectFormatting = True
    End If
    If doPara And Not HasDirectFormatting Then
        If p.LeftIndent <> st.ParagraphFormat.LeftIndent Then HasDirectFormatting = True
        If p.SpaceBefore <> st.ParagraphFormat.SpaceBefore Then HasDirectFormatting = True
        If p.SpaceAfter <> st.ParagraphFormat.SpaceAfter Then HasDirectFormatting = True
        If p.Alignment <> st.ParagraphFormat.Alignment Then HasDirectFormatting = True
    End If
    On Error GoTo 0
End Function
Private Sub ApplyStripToParagraph(p As Paragraph, doChar As Boolean, doPara As Boolean, emphMode As Integer, preserveHL As Boolean, preserveDropCaps As Boolean)
    If Not preserveDropCaps Then
        On Error Resume Next
        p.DropCap.Position = wdDropNone
        On Error GoTo 0
    End If
    If Not doChar Then
        If doPara Then p.Range.ClearParagraphDirectFormatting
        Exit Sub
    End If
    If emphMode = 2 Then
        p.Range.ClearCharacterDirectFormatting
        If doPara Then p.Range.ClearParagraphDirectFormatting
    ElseIf emphMode = 0 Then
        Dim wasBold As Variant, wasItalic As Variant, savedHL As Variant
        wasBold = p.Range.Font.Bold
        wasItalic = p.Range.Font.Italic
        savedHL = p.Range.HighlightColorIndex
        p.Range.ClearCharacterDirectFormatting
        If doPara Then p.Range.ClearParagraphDirectFormatting
        If wasBold = True Then p.Range.Font.Bold = True
        If wasItalic = True Then p.Range.Font.Italic = True
        If preserveHL And savedHL <> wdUndefined And savedHL <> wdNoHighlight Then p.Range.HighlightColorIndex = savedHL
    Else
        Dim n As Long: n = p.Range.Characters.Count
        If n < 1 Then Exit Sub
        Dim bArr() As Boolean, iArr() As Boolean, hArr() As Long
        ReDim bArr(1 To n)
        ReDim iArr(1 To n)
        ReDim hArr(1 To n)
        Dim ch As Range, idx As Long: idx = 0
        For Each ch In p.Range.Characters
            idx = idx + 1
            If idx <= n Then
                bArr(idx) = (ch.Font.Bold = True)
                iArr(idx) = (ch.Font.Italic = True)
                hArr(idx) = ch.HighlightColorIndex
            End If
        Next ch
        Dim baseStart As Long: baseStart = p.Range.Start
        p.Range.ClearCharacterDirectFormatting
        If doPara Then p.Range.ClearParagraphDirectFormatting
        ApplyRuns baseStart, bArr, iArr, hArr, n, preserveHL
    End If
End Sub
Private Sub ApplyRuns(baseStart As Long, bArr() As Boolean, iArr() As Boolean, hArr() As Long, n As Long, preserveHL As Boolean)
    Dim i As Long
    i = 1
    Do While i <= n
        If bArr(i) Then
            Dim jb As Long: jb = i
            Do While jb < n
                If Not bArr(jb + 1) Then Exit Do
                jb = jb + 1
            Loop
            ActiveDocument.Range(baseStart + i - 1, baseStart + jb).Font.Bold = True
            i = jb + 1
        Else
            i = i + 1
        End If
    Loop
    i = 1
    Do While i <= n
        If iArr(i) Then
            Dim ji As Long: ji = i
            Do While ji < n
                If Not iArr(ji + 1) Then Exit Do
                ji = ji + 1
            Loop
            ActiveDocument.Range(baseStart + i - 1, baseStart + ji).Font.Italic = True
            i = ji + 1
        Else
            i = i + 1
        End If
    Loop
    If preserveHL Then
        i = 1
        Do While i <= n
            If hArr(i) <> wdNoHighlight And hArr(i) <> wdUndefined Then
                Dim jh As Long: jh = i
                Do While jh < n
                    If hArr(jh + 1) <> hArr(i) Then Exit Do
                    jh = jh + 1
                Loop
                ActiveDocument.Range(baseStart + i - 1, baseStart + jh).HighlightColorIndex = hArr(i)
                i = jh + 1
            Else
                i = i + 1
            End If
        Loop
    End If
End Sub
