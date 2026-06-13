Option Explicit
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then ReturnToMainAfterToolClose
End Sub
' --------------------------------------------------------------
'  Duplicate Paragraph Detector
'  Finds repeated paragraphs and highlights them, or removes all but the
'  first occurrence.  Three matching methods: exact (case-insensitive),
'  normalized (ignores punctuation and spacing), and fuzzy (groups
'  paragraphs that share a chosen percentage of their words).  Fuzzy
'  matching uses global single-linkage clustering, so the groups it finds
'  do not depend on the order paragraphs appear in the document.
' --------------------------------------------------------------
Private Sub UserForm_Initialize()
    optHighlightOnly.GroupName = "DupAction"
    optRemoveDupes.GroupName = "DupAction"
    optMatchExact.GroupName = "DupMatch"
    optMatchNormalized.GroupName = "DupMatch"
    optMatchFuzzy.GroupName = "DupMatch"
    optFuzzyLoose.GroupName = "DupThreshold"
    optFuzzyMedium.GroupName = "DupThreshold"
    optFuzzyStrict.GroupName = "DupThreshold"
    optScopeDocument.GroupName = "DupScope"
    optScopeSelection.GroupName = "DupScope"
    optHighlightOnly.Value = True
    optMatchExact.Value = True
    optFuzzyMedium.Value = True
    lblFuzzyWarning.Caption = "Exact and normalized matching are fast.  Fuzzy matching compares every paragraph and is slower on large documents."
    optScopeDocument.Caption = "Entire document"
    optScopeDocument.Value = True
    optScopeSelection.Caption = "Selected text only"
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        optScopeSelection.Enabled = True
    Else
        optScopeSelection.Enabled = False
    End If
    optHighlightOnly.Caption = "Highlight duplicates (preview only)"
    optRemoveDupes.Caption = "Remove duplicate paragraphs"
    optMatchExact.Caption = "Exact match"
    optMatchNormalized.Caption = "Normalized match (ignore punctuation and spaces)"
    optMatchFuzzy.Caption = "Fuzzy match (word overlap)"
    optFuzzyLoose.Caption = "Loose  (50% word overlap)"
    optFuzzyMedium.Caption = "Medium  (70% word overlap)"
    optFuzzyStrict.Caption = "Strict  (90% word overlap)"
    chkPreviewOnly.Caption = "Preview only (highlight, do not change)"
    LayoutCleanupToolForm Me
End Sub
Private Sub optMatchExact_Click()
    lblFuzzyWarning.Caption = "Exact match: paragraphs must be identical apart from letter case."
End Sub
Private Sub optMatchNormalized_Click()
    lblFuzzyWarning.Caption = "Normalized match: ignores punctuation, extra spaces, and letter case."
End Sub
Private Sub optMatchFuzzy_Click()
    lblFuzzyWarning.Caption = "Fuzzy match: groups paragraphs sharing the chosen percentage of words.  Slower on large documents."
End Sub
Private Sub cmdHelp_Click()
    Dim h As String
    h = "Duplicate Paragraph Detector" & vbCrLf & vbCrLf
    h = h & "Finds paragraphs that repeat in the document and highlights them, or " & _
            "removes every occurrence except the first.  Blank paragraphs are ignored." & vbCrLf
    h = h & vbCrLf
    h = h & "MATCHING METHOD" & vbCrLf
    h = h & "  Exact            --  identical apart from letter case." & vbCrLf
    h = h & "  Ignore punctuation & spacing  --  duplicates even if they differ only" & vbCrLf
    h = h & "                       in punctuation, extra spaces, or letter case." & vbCrLf
    h = h & "  Fuzzy (similar wording)  --  flags paragraphs that share a chosen" & vbCrLf
    h = h & "                       percentage of their words; catches near-duplicates " & _
            "                       that were lightly edited.  Slower on large documents." & vbCrLf
    h = h & vbCrLf
    h = h & "FUZZY SENSITIVITY (used only by fuzzy matching)" & vbCrLf
    h = h & "  Loose (60%)   --  more matches, including loosely similar paragraphs." & vbCrLf
    h = h & "  Medium (80%)  --  a balanced default." & vbCrLf
    h = h & "  Strict (90%)  --  fewer matches; only very close paragraphs." & vbCrLf
    h = h & vbCrLf
    h = h & "ACTION" & vbCrLf
    h = h & "  Highlight only     --  marks duplicates in yellow for review." & vbCrLf
    h = h & "  Remove duplicates  --  deletes every occurrence except the first." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to search within the selection only." & vbCrLf
    h = h & vbCrLf
    h = h & "Similar paragraphs are grouped together (a paragraph linked to a match, " & _
            "which is linked to another, joins the same group), and the first in each" & vbCrLf
    h = h & "group is kept.  Run with Preview first to see what would be affected." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Duplicate Paragraph Detector"
End Sub
Private Function NormalizeText(ByVal t As String, ByVal stripPunct As Boolean) As String
    Dim r As String
    r = t
    If Len(r) > 0 Then
        If Right$(r, 1) = vbCr Then r = Left$(r, Len(r) - 1)
    End If
    r = LCase$(Trim$(r))
    If stripPunct Then
        Dim i As Long, ch As String, out As String
        out = ""
        For i = 1 To Len(r)
            ch = Mid$(r, i, 1)
            If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then
                out = out & ch
            ElseIf ch = " " Then
                If Len(out) > 0 Then
                    If Right$(out, 1) <> " " Then out = out & " "
                End If
            End If
        Next i
        r = Trim$(out)
    End If
    NormalizeText = r
End Function
Private Function BuildWordSet(ByVal t As String) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Dim norm As String: norm = NormalizeText(t, True)
    If Len(norm) = 0 Then Set BuildWordSet = d: Exit Function
    Dim parts() As String: parts = Split(norm, " ")
    Dim i As Long, w As String
    For i = LBound(parts) To UBound(parts)
        w = parts(i)
        If Len(w) > 0 Then
            If Not d.Exists(w) Then d.Add w, 1
        End If
    Next i
    Set BuildWordSet = d
End Function
Private Function JaccardSimilarity(ByVal a As Object, ByVal b As Object) As Double
    If a.Count = 0 And b.Count = 0 Then JaccardSimilarity = 1: Exit Function
    If a.Count = 0 Or b.Count = 0 Then JaccardSimilarity = 0: Exit Function
    Dim shared As Long: shared = 0
    Dim k As Variant
    If a.Count <= b.Count Then
        For Each k In a.Keys
            If b.Exists(k) Then shared = shared + 1
        Next k
    Else
        For Each k In b.Keys
            If a.Exists(k) Then shared = shared + 1
        Next k
    End If
    JaccardSimilarity = shared / (a.Count + b.Count - shared)
End Function
Private Function FindRoot(ByRef par() As Long, ByVal x As Long) As Long
    Do While par(x) <> x
        par(x) = par(par(x))
        x = par(x)
    Loop
    FindRoot = x
End Function
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
    If Not GuardBeforeCleanup("Duplicate Detector") Then Unload Me: Exit Sub
    Dim previewOnly As Boolean: previewOnly = chkPreviewOnly.Value
    Dim targetRange As Range
    If optScopeSelection.Value And optScopeSelection.Enabled Then
        Set targetRange = GetTargetRange()
    Else
        Set targetRange = ActiveDocument.Content
    End If
    Dim doRemove As Boolean: doRemove = optRemoveDupes.Value
    Dim matchMode As Integer: matchMode = 0
    If optMatchNormalized.Value Then matchMode = 1
    If optMatchFuzzy.Value Then matchMode = 2
    Dim threshold As Double: threshold = 0.8
    If optFuzzyLoose.Value Then threshold = 0.6
    If optFuzzyStrict.Value Then threshold = 0.9
    Dim paras As Collection: Set paras = New Collection
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= targetRange.Start And p.Range.End <= targetRange.End Then
            If Len(Trim$(p.Range.Text)) > 1 Then paras.Add p
        End If
    Next p
    If matchMode = 2 And paras.Count > 1500 Then
        If MsgBox("This scope has " & paras.Count & " paragraphs. Fuzzy matching compares them all and may take a while. Continue?", vbQuestion + vbYesNo, "Fuzzy Match") = vbNo Then Exit Sub
    End If
    MarkCleanupStart "Duplicate Detector"
    Dim undoRec As UndoRecord
    Set undoRec = Application.UndoRecord
    undoRec.StartCustomRecord "Cleanup Suite - Duplicate Detector"
    On Error GoTo RunErr
    Application.ScreenUpdating = False
    Dim dupCount As Long: dupCount = 0
    Dim dupStarts As Collection: Set dupStarts = New Collection
    Dim i As Long, j As Long
    If matchMode <= 1 Then
        Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
        dict.CompareMode = 1
        Dim key As String
        For i = 1 To paras.Count
            Set p = paras(i)
            If matchMode = 0 Then
                key = NormalizeText(p.Range.Text, False)
            Else
                key = NormalizeText(p.Range.Text, True)
            End If
            If Len(key) > 0 Then
                If dict.Exists(key) Then
                    p.Range.HighlightColorIndex = wdYellow
                    dupStarts.Add p.Range.Start
                    dupCount = dupCount + 1
                Else
                    dict.Add key, i
                End If
            End If
        Next i
    Else
        Dim n As Long: n = paras.Count
        If n > 0 Then
            Dim wsArr() As Object
            ReDim wsArr(1 To n)
            For i = 1 To n
                Set wsArr(i) = BuildWordSet(paras(i).Range.Text)
            Next i
            Dim par() As Long
            ReDim par(1 To n)
            For i = 1 To n
                par(i) = i
            Next i
            Dim ca As Long, cb As Long, mn As Long, mx As Long, ra As Long, rb As Long
            For i = 1 To n - 1
                ca = wsArr(i).Count
                If ca > 0 Then
                    For j = i + 1 To n
                        cb = wsArr(j).Count
                        If cb > 0 Then
                            If ca <= cb Then
                                mn = ca: mx = cb
                            Else
                                mn = cb: mx = ca
                            End If
                            If mn / mx >= threshold Then
                                If JaccardSimilarity(wsArr(i), wsArr(j)) >= threshold Then
                                    ra = FindRoot(par, i)
                                    rb = FindRoot(par, j)
                                    If ra <> rb Then par(ra) = rb
                                End If
                            End If
                        End If
                    Next j
                End If
                If (i Mod 25 = 0) Then Application.StatusBar = "Duplicate Detector: comparing " & i & " of " & n & "..."
            Next i
            Dim rootMin As Object: Set rootMin = CreateObject("Scripting.Dictionary")
            Dim rt As String
            For i = 1 To n
                rt = CStr(FindRoot(par, i))
                If rootMin.Exists(rt) Then
                    If i < rootMin(rt) Then rootMin(rt) = i
                Else
                    rootMin.Add rt, i
                End If
            Next i
            For i = 1 To n
                rt = CStr(FindRoot(par, i))
                If i <> rootMin(rt) Then
                    paras(i).Range.HighlightColorIndex = wdYellow
                    dupStarts.Add paras(i).Range.Start
                    dupCount = dupCount + 1
                End If
            Next i
        End If
    End If
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If previewOnly Then
        undoRec.EndCustomRecord
        MarkCleanupEnd
        ShowPreviewActions Me, "Duplicate Detector", "Preview complete. " & dupCount & " duplicate paragraphs highlighted."
        Exit Sub
    End If
    Dim results As Collection: Set results = New Collection
    If doRemove Then
        ' Remove ONLY paragraphs detected as duplicates, identified by the
        ' start positions captured during detection -- never by highlight
        ' colour. This leaves any pre-existing user highlighting untouched.
        Dim removed As Long: removed = 0
        Dim dn As Long: dn = dupStarts.Count
        If dn > 0 Then
            Dim ds() As Long, di As Long
            ReDim ds(1 To dn)
            For di = 1 To dn: ds(di) = dupStarts(di): Next di
            ' sort start positions descending so each deletion never shifts
            ' the positions of paragraphs still queued for removal
            Dim a As Long, b As Long, tmpS As Long
            For a = 1 To dn - 1
                For b = a + 1 To dn
                    If ds(b) > ds(a) Then tmpS = ds(a): ds(a) = ds(b): ds(b) = tmpS
                Next b
            Next a
            For di = 1 To dn
                ActiveDocument.Range(ds(di), ds(di)).Paragraphs(1).Range.Delete
                removed = removed + 1
            Next di
        End If
        results.Add "Duplicate paragraphs removed: " & removed
    Else
        results.Add "Duplicate paragraphs highlighted: " & dupCount
    End If
    Select Case matchMode
        Case 0: results.Add "Match method: exact (case-insensitive)"
        Case 1: results.Add "Match method: normalized"
        Case 2: results.Add "Match method: fuzzy (" & CLng(threshold * 100) & "%, grouped)"
    End Select
    ShowCleanupReport "Duplicate Detector", results
    undoRec.EndCustomRecord
    MarkCleanupEnd
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
