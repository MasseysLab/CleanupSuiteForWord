Option Explicit
' Shared helper functions for cleanup forms
Public gPreviewActionPanel As Object
Private gPreviewActionPanelHasPosition As Boolean
Private gPreviewActionPanelLeft As Single
Private gPreviewActionPanelTop As Single
Private gCleanupToolExitReason As String
Private Const CLEANUP_TOOL_EXIT_APPLY As String = "apply"
Private Const CLEANUP_TOOL_EXIT_CLOSE As String = "close"
Public Const CLEANUP_SUITE_USER_MANUAL_PDF_URL As String = "https://raw.githubusercontent.com/MasseysLab/CleanupSuiteForWord/v0.8.0/documents/CleanupSuite_Human_Friendly_User_Manual_v0.8.0.pdf"

Public Sub OpenCleanupUserManualPdf()
    On Error GoTo OpenErr
    ActiveDocument.FollowHyperlink Address:=CLEANUP_SUITE_USER_MANUAL_PDF_URL, NewWindow:=True
    Exit Sub
OpenErr:
    MsgBox "Could not open the PDF User Manual." & vbCrLf & vbCrLf & _
           "You can open it manually at:" & vbCrLf & _
           CLEANUP_SUITE_USER_MANUAL_PDF_URL, vbExclamation, "Cleanup Suite User Manual"
End Sub

Public Sub ShowCleanupToolHelp(ByVal toolKey As String)
    MsgBox CleanupHelpText(toolKey), vbInformation, CleanupHelpTitle(toolKey) & " Help"
End Sub

Public Function CleanupHelpTitle(ByVal toolKey As String) As String
    Select Case toolKey
        Case "Unicode": CleanupHelpTitle = "Invisible Unicode Cleaner"
        Case "Punctuation": CleanupHelpTitle = "Punctuation Normalizer"
        Case "Spacing": CleanupHelpTitle = "Spacing Fixer"
        Case "Capitalization": CleanupHelpTitle = "Capitalization Fixer"
        Case "List": CleanupHelpTitle = "List Normalizer"
        Case "Paragraph": CleanupHelpTitle = "Paragraph Structure Fixer"
        Case "Duplicate": CleanupHelpTitle = "Duplicate Paragraph Detector"
        Case "Font": CleanupHelpTitle = "Font Normalizer"
        Case "Table": CleanupHelpTitle = "Table Cleaner"
        Case "Break": CleanupHelpTitle = "Break Normalizer"
        Case "DocumentTrim": CleanupHelpTitle = "Document Trim"
        Case "Formatting": CleanupHelpTitle = "Formatting Stripper"
        Case "Hyperlink": CleanupHelpTitle = "Hyperlink Remover"
        Case "SoftReturn": CleanupHelpTitle = "Soft Return Converter"
        Case "Metadata": CleanupHelpTitle = "Metadata Scrubber"
        Case "Style": CleanupHelpTitle = "Style Cleanup"
        Case "Footnote": CleanupHelpTitle = "Footnote / Endnote Remover"
        Case "HeaderFooter": CleanupHelpTitle = "Header / Footer Standardizer"
        Case "Object": CleanupHelpTitle = "Object Remover"
        Case Else: CleanupHelpTitle = "Cleanup Suite"
    End Select
End Function

Public Function CleanupHelpText(ByVal toolKey As String) As String
    Dim h As String

    Select Case toolKey
        Case "Unicode"
            AddHelpLine h, "Invisible Unicode Cleaner"
            AddHelpLine h, ""
            AddHelpLine h, "Removes invisible or problem characters that can break search, sorting, comparison, exports, and automated text processing."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- Non-breaking spaces: regular-looking spaces that prevent line breaks."
            AddHelpLine h, "- Zero-width characters: invisible characters that can split words or affect searching."
            AddHelpLine h, "- Byte order marks, soft hyphens, and non-breaking hyphens: hidden or special characters that often arrive from web, PDF, or cross-platform text."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to limit cleanup to the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights matches in yellow without changing the document."

        Case "Punctuation"
            AddHelpLine h, "Punctuation Normalizer"
            AddHelpLine h, ""
            AddHelpLine h, "Replaces typographic punctuation with plain-text equivalents for cleaner export, comparison, and reuse."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- All punctuation types: quotes, dashes, and ellipses."
            AddHelpLine h, "- Quotes only: curly double and single quotes become straight quotes."
            AddHelpLine h, "- Dashes only: em dashes and en dashes become hyphens."
            AddHelpLine h, "- Ellipses only: ellipsis characters become three dots."
            AddHelpLine h, "- Custom: choose individual punctuation types."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to limit cleanup to the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights matches without changing the document."

        Case "Spacing"
            AddHelpLine h, "Spacing Fixer"
            AddHelpLine h, ""
            AddHelpLine h, "Corrects common spacing problems in documents assembled from copied text, templates, OCR, PDFs, and multiple files."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- All spacing fixes: runs the full spacing cleanup set."
            AddHelpLine h, "- Double spaces only: collapses repeated spaces."
            AddHelpLine h, "- Trim leading / trailing spaces: removes extra spaces at paragraph edges."
            AddHelpLine h, "- Custom: choose double spaces, paragraph trimming, punctuation spacing, and extra blank-line cleanup."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to limit cleanup to the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights problem areas without changing the document."

        Case "Capitalization"
            AddHelpLine h, "Capitalization Fixer"
            AddHelpLine h, ""
            AddHelpLine h, "Applies consistent capitalization to paragraph text in the selected scope."
            AddHelpLine h, ""
            AddHelpLine h, "Modes:"
            AddHelpLine h, "- Fix sentence starts: recommended for general text. It capitalizes likely sentence starts after . ? ! while leaving existing casing alone where possible."
            AddHelpLine h, "- Sentence case: capitalizes the first character and lowercases the rest."
            AddHelpLine h, "- Title case: capitalizes the first letter of every word."
            AddHelpLine h, "- UPPERCASE / lowercase: converts affected paragraph text to that case."
            AddHelpLine h, "- Custom: runs selected choices from top to bottom. Later choices can override earlier choices."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening this tool to enable Selected text only."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights paragraphs that would change."

        Case "List"
            AddHelpLine h, "List Normalizer"
            AddHelpLine h, ""
            AddHelpLine h, "Cleans manual bullets, numbering, and list indentation so list paragraphs behave like Word lists."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- All list fixes: runs the full list cleanup set."
            AddHelpLine h, "- Convert manual bullet lists: changes typed bullet markers into list formatting."
            AddHelpLine h, "- Convert manual numbered lists: changes typed numbers into list formatting."
            AddHelpLine h, "- Fix list indentation: standardizes list paragraph indents."
            AddHelpLine h, "- Custom: choose the list fixes individually."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to limit cleanup to the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights list paragraphs that would be affected."

        Case "Paragraph"
            AddHelpLine h, "Paragraph Structure Fixer"
            AddHelpLine h, ""
            AddHelpLine h, "Cleans empty paragraphs, paragraph spacing, soft returns, and manual indentation."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- All paragraph fixes: runs the full paragraph cleanup set."
            AddHelpLine h, "- Remove empty paragraphs: deletes paragraphs that contain no visible text."
            AddHelpLine h, "- Normalize paragraph spacing: resets direct spacing overrides."
            AddHelpLine h, "- Custom: choose empty paragraph cleanup, soft-return conversion, spacing cleanup, and indent cleanup."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to limit cleanup to the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights affected paragraphs before changes are made."

        Case "Duplicate"
            AddHelpLine h, "Duplicate Paragraph Detector"
            AddHelpLine h, ""
            AddHelpLine h, "Finds repeated or near-repeated paragraphs before you decide whether to keep or remove them."
            AddHelpLine h, ""
            AddHelpLine h, "Choices:"
            AddHelpLine h, "- Highlight duplicates: marks repeated paragraphs for review."
            AddHelpLine h, "- Remove duplicate paragraphs: keeps the first occurrence and removes later matches."
            AddHelpLine h, "- Exact match: fastest and strictest."
            AddHelpLine h, "- Normalized match: ignores punctuation and spacing differences."
            AddHelpLine h, "- Fuzzy match: compares word overlap and is slower on large documents."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to search only within the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview first is strongly recommended before removing duplicate paragraphs."

        Case "Font"
            AddHelpLine h, "Font Normalizer"
            AddHelpLine h, ""
            AddHelpLine h, "Resets direct font overrides so text follows the font settings defined by its paragraph style."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- Normalize font face: returns text to the document default."
            AddHelpLine h, "- Normalize font size: returns text to the style-defined size."
            AddHelpLine h, "- Normalize bold / italic overrides: clears direct emphasis settings."
            AddHelpLine h, "- Remove direct font color overrides: returns text color to the style."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to limit cleanup to the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights paragraphs with direct font overrides."

        Case "Table"
            AddHelpLine h, "Table Cleaner"
            AddHelpLine h, ""
            AddHelpLine h, "Cleans table structure and formatting while keeping table content unless you explicitly convert a table to text."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- Remove empty rows or columns: deletes table rows or columns with no visible content."
            AddHelpLine h, "- Normalize cell padding: standardizes table cell margins."
            AddHelpLine h, "- Strip direct formatting from cells: clears manual formatting overrides."
            AddHelpLine h, "- Normalize or remove borders: standardizes table borders."
            AddHelpLine h, "- Convert single-column tables to plain text: replaces simple one-column tables with text."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to limit cleanup to tables in the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview reports the table cleanup that would run."

        Case "Break"
            AddHelpLine h, "Break Normalizer"
            AddHelpLine h, ""
            AddHelpLine h, "Fixes consecutive page and section breaks, most often found in documents assembled from several source files."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- Collapse consecutive section breaks: reduces repeated section breaks to one."
            AddHelpLine h, "- Collapse consecutive page breaks: reduces repeated manual page breaks to one."
            AddHelpLine h, "- Convert section breaks to page breaks: changes section breaks to the selected break type."
            AddHelpLine h, ""
            AddHelpLine h, "Conversion types:"
            AddHelpLine h, "- Next Page, Continuous, Even Page, or Odd Page."
            AddHelpLine h, "- The first section of a document is never converted because it has no preceding break."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to limit cleanup to the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights consecutive break sequences."

        Case "DocumentTrim"
            AddHelpLine h, "Document Trim"
            AddHelpLine h, ""
            AddHelpLine h, "Removes extra empty paragraphs from the end of the document while leaving Word's required final paragraph mark."
            AddHelpLine h, ""
            AddHelpLine h, "Why this matters:"
            AddHelpLine h, "Documents assembled from multiple files or templates often end with blank paragraphs. Those blanks can create unwanted printed space and may affect headers, footers, and page numbering in some layouts."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "This tool always works on the full document because the end of the document is a document-level location."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights the trailing paragraphs that would be removed."

        Case "Formatting"
            AddHelpLine h, "Formatting Stripper"
            AddHelpLine h, ""
            AddHelpLine h, "Removes direct formatting while preserving the document's style structure."
            AddHelpLine h, ""
            AddHelpLine h, "Choices:"
            AddHelpLine h, "- Strip direct character formatting: resets manual font-level overrides."
            AddHelpLine h, "- Strip direct paragraph formatting: resets manual paragraph-level overrides."
            AddHelpLine h, "- Quick: preserves emphasis that applies to whole paragraphs."
            AddHelpLine h, "- Thorough: preserves emphasis on individual words, but takes longer."
            AddHelpLine h, "- Strip: removes all direct formatting."
            AddHelpLine h, "- Preserve highlighting and drop caps: keep those visual elements during cleanup."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to limit cleanup to the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights paragraphs carrying direct formatting."

        Case "Hyperlink"
            AddHelpLine h, "Hyperlink Remover"
            AddHelpLine h, ""
            AddHelpLine h, "Removes hyperlink targets while keeping the visible text."
            AddHelpLine h, ""
            AddHelpLine h, "Option:"
            AddHelpLine h, "- Also remove hyperlink character style: clears the blue underline after the link is removed."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to limit cleanup to the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights hyperlinks that would be removed."

        Case "SoftReturn"
            AddHelpLine h, "Soft Return Converter"
            AddHelpLine h, ""
            AddHelpLine h, "Converts soft returns made with Shift+Enter to paragraph marks, or converts paragraph marks back to soft returns."
            AddHelpLine h, ""
            AddHelpLine h, "Choices:"
            AddHelpLine h, "- Convert soft returns to paragraph marks: the usual fix for PDF or web text imports."
            AddHelpLine h, "- Convert paragraph marks to soft returns: useful only when you intentionally want one paragraph split across lines."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to limit cleanup to the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights breaks that would be converted."

        Case "Metadata"
            AddHelpLine h, "Metadata Scrubber"
            AddHelpLine h, ""
            AddHelpLine h, "Removes document metadata and review artifacts before a document is shared."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- Clear document properties: removes built-in properties such as title, author, company, and manager."
            AddHelpLine h, "- Clear personal information: removes or anonymizes revision author information where Word allows it."
            AddHelpLine h, "- Remove all comments: deletes comments from the document."
            AddHelpLine h, "- Accept tracked changes and clear revision marks: finalizes the current revision state."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "This tool always works on the full document because metadata is document-level information."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview shows current document metadata and review counts."

        Case "Style"
            AddHelpLine h, "Style Cleanup"
            AddHelpLine h, ""
            AddHelpLine h, "Removes unused custom styles and can remap common style variants to cleaner equivalents."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- Remove unused styles: deletes custom styles that are not used in the main document body."
            AddHelpLine h, "- Remap style variants: moves common duplicate style variants back to canonical styles."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "This tool always works on the full document because style definitions belong to the document."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview reports how many unused styles or remappable variants were found."

        Case "Footnote"
            AddHelpLine h, "Footnote / Endnote Remover"
            AddHelpLine h, ""
            AddHelpLine h, "Removes footnotes, endnotes, or both. It can optionally keep note text inline in brackets."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- Remove footnotes: deletes footnote references and note text."
            AddHelpLine h, "- Remove endnotes: deletes endnote references and note text."
            AddHelpLine h, "- Keep note text inline: inserts note text at the reference point before removing the note."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to remove only notes referenced in the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview reports how many notes would be removed."

        Case "HeaderFooter"
            AddHelpLine h, "Header / Footer Standardizer"
            AddHelpLine h, ""
            AddHelpLine h, "Standardizes header and footer formatting across the full document, or clears header and footer content."
            AddHelpLine h, ""
            AddHelpLine h, "Modes:"
            AddHelpLine h, "- Standardize formatting: applies selected formatting cleanup to header and footer areas."
            AddHelpLine h, "- Clear all header/footer content: removes header and footer text."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- Include headers / Include footers: choose which areas to process."
            AddHelpLine h, "- Reset font: uses the document default font and size."
            AddHelpLine h, "- Remove paragraph spacing: sets space before and after to zero."
            AddHelpLine h, "- Set alignment: applies left, center, or right alignment."
            AddHelpLine h, "- Unlink sections: breaks Link to Previous so each section can be processed independently."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "This tool always works on the full document. Page numbers, dates, and other fields are not intentionally changed."

        Case "Object"
            AddHelpLine h, "Object Remover"
            AddHelpLine h, ""
            AddHelpLine h, "Removes embedded objects and hidden content that often arrive from web pages, PDFs, email, or copied documents."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- Pictures: removes inline and floating images."
            AddHelpLine h, "- Text boxes: removes text boxes and their contents."
            AddHelpLine h, "- Frames: removes frame containers while keeping the text."
            AddHelpLine h, "- Horizontal lines: removes inserted horizontal-line graphics."
            AddHelpLine h, "- HTML / ActiveX form controls: removes form controls from pasted web content."
            AddHelpLine h, "- Hidden text: removes text formatted as hidden."
            AddHelpLine h, "- Tables: deletes whole tables and all their contents after confirmation."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "This tool always works on the full document body. Header and footer objects are left alone."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview reports counts for objects that would be removed."

        Case Else
            AddHelpLine h, "Cleanup Suite"
            AddHelpLine h, ""
            AddHelpLine h, "No help text is available for this tool yet."
    End Select

    CleanupHelpText = h
End Function

Private Sub AddHelpLine(ByRef textValue As String, ByVal lineValue As String)
    If Len(textValue) > 0 Then textValue = textValue & vbCrLf
    textValue = textValue & lineValue
End Sub

Public Sub RememberPreviewActionPanelPosition(panelLeft As Single, panelTop As Single)
    gPreviewActionPanelLeft = panelLeft
    gPreviewActionPanelTop = panelTop
    gPreviewActionPanelHasPosition = True
End Sub

Private Sub ApplyPreviewActionPanelPosition(ByVal panel As Object)
    On Error Resume Next
    If gPreviewActionPanelHasPosition Then
        panel.StartUpPosition = 0
        panel.Left = gPreviewActionPanelLeft
        panel.Top = gPreviewActionPanelTop
    End If
    On Error GoTo 0
End Sub

Public Sub ShowPreviewActions(ByVal sourceForm As Object, toolName As String, summaryText As String)
    On Error GoTo Fallback
    sourceForm.Hide
    Set gPreviewActionPanel = New frmPreviewActions
    gPreviewActionPanel.Configure sourceForm, toolName, summaryText
    ApplyPreviewActionPanelPosition gPreviewActionPanel
    gPreviewActionPanel.Show
    Exit Sub
Fallback:
    MsgBox summaryText & vbCrLf & vbCrLf & _
           "Preview is still visible in the document. Run again without Preview to apply, or clear highlighting manually.", _
           vbInformation, toolName & " Preview"
End Sub

Public Sub RemoveAllHighlighting(Optional scopeRange As Range = Nothing)
    Dim hlRange As Range
    If scopeRange Is Nothing Then
        Set hlRange = ActiveDocument.Content
    Else
        Set hlRange = scopeRange
    End If
    With hlRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Format = True
        .Highlight = True
        .Text = ""
        .Replacement.Text = ""
        .Replacement.Highlight = False
        .Forward = True
        .Wrap = wdFindStop
        .Execute Replace:=wdReplaceAll
    End With
End Sub
Public Function DocumentHasVisibleContent(Optional scopeRange As Range = Nothing) As Boolean
    On Error GoTo SafeExit

    Dim textValue As String
    If scopeRange Is Nothing Then
        textValue = ActiveDocument.Content.Text
    Else
        textValue = scopeRange.Text
    End If

    textValue = Replace(textValue, vbCr, "")
    textValue = Replace(textValue, vbLf, "")
    textValue = Replace(textValue, vbTab, "")
    textValue = Replace(textValue, Chr$(11), "")
    textValue = Replace(textValue, Chr$(12), "")
    textValue = Replace(textValue, ChrW$(160), "")
    If Len(Trim$(textValue)) > 0 Then
        DocumentHasVisibleContent = True
        Exit Function
    End If

    If Not scopeRange Is Nothing Then Exit Function
    If ActiveDocument.Tables.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
    If ActiveDocument.InlineShapes.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
    If ActiveDocument.Shapes.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
    If ActiveDocument.Hyperlinks.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
    If ActiveDocument.Footnotes.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
    If ActiveDocument.Endnotes.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
    If ActiveDocument.Comments.Count > 0 Then DocumentHasVisibleContent = True: Exit Function
SafeExit:
End Function
Public Function GuardBeforeCleanup(toolName As String) As Boolean
    GuardBeforeCleanup = False
    If ActiveDocument.Path = "" Then
        ' Document has never been saved, so Word has nothing to recover.
        Dim msg As String
        msg = "This document has not been saved yet." & vbCrLf & vbCrLf & _
              "Saving first lets Word's built-in recovery protect you if anything goes wrong." & vbCrLf & vbCrLf & _
              "Save now before running " & toolName & "?"
        Dim ans As Integer
        ans = MsgBox(msg, vbExclamation + vbYesNoCancel, "Unsaved Document")
        If ans = vbCancel Then Exit Function
        If ans = vbYes Then
            On Error Resume Next
            ActiveDocument.Save
            If Err.Number <> 0 Then
                MsgBox "Save failed. Please save manually before running cleanup.", vbCritical, "Save Failed"
                Exit Function
            End If
            On Error GoTo 0
        End If
    ElseIf GetAutoSaveSetting() And Not IsRunningInMacroHostDocument() Then
        ' Document is saved and auto-save is ON, so save silently before proceeding.
        On Error Resume Next
        ActiveDocument.Save
        On Error GoTo 0
    End If
    GuardBeforeCleanup = True
End Function
Public Function IsRunningInMacroHostDocument() As Boolean
    On Error Resume Next
    IsRunningInMacroHostDocument = (ActiveDocument.FullName = ThisDocument.FullName)
    On Error GoTo 0
End Function
Public Sub MarkCleanupStart(toolName As String)
    On Error Resume Next
    If IsRunningInMacroHostDocument() Then Exit Sub
    ActiveDocument.CustomDocumentProperties("CleanupSuiteInProgress").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteInProgress", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=toolName
    If Len(ActiveDocument.Path) > 0 Then ActiveDocument.Save
    On Error GoTo 0
End Sub
Public Sub MarkCleanupEnd()
    On Error Resume Next
    If IsRunningInMacroHostDocument() Then Exit Sub
    ActiveDocument.CustomDocumentProperties("CleanupSuiteInProgress").Delete
    On Error GoTo 0
End Sub
Public Function GetAutoSaveSetting() As Boolean
    ' Returns True (auto-save ON) by default; False only if explicitly disabled
    On Error Resume Next
    Dim prop As Object
    Set prop = ActiveDocument.CustomDocumentProperties("CleanupSuiteAutoSave")
    If Err.Number <> 0 Then
        GetAutoSaveSetting = True
        Err.Clear
    Else
        GetAutoSaveSetting = (prop.Value = "True")
    End If
    On Error GoTo 0
End Function
Public Sub SetAutoSaveSetting(enabled As Boolean)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("CleanupSuiteAutoSave").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteAutoSave", _
        LinkToContent:=False, Type:=msoPropertyTypeString, _
        Value:=IIf(enabled, "True", "False")
    On Error GoTo 0
End Sub
Public Function GetReturnToMainAfterApplySetting() As Boolean
    ' Returns True by default so Apply keeps the current workflow unless changed.
    On Error Resume Next
    Dim prop As Object
    Set prop = ActiveDocument.CustomDocumentProperties("CleanupSuiteReturnToMainAfterApply")
    If Err.Number <> 0 Then
        GetReturnToMainAfterApplySetting = True
        Err.Clear
    Else
        GetReturnToMainAfterApplySetting = (prop.Value = "True")
    End If
    On Error GoTo 0
End Function
Public Sub SetReturnToMainAfterApplySetting(enabled As Boolean)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("CleanupSuiteReturnToMainAfterApply").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteReturnToMainAfterApply", _
        LinkToContent:=False, Type:=msoPropertyTypeString, _
        Value:=IIf(enabled, "True", "False")
    On Error GoTo 0
End Sub
Public Function GetShowCompletionReviewAfterApplySetting() As Boolean
    ' Returns True by default so users see the completion summary unless they opt out.
    On Error Resume Next
    Dim prop As Object
    Set prop = ActiveDocument.CustomDocumentProperties("CleanupSuiteShowCompletionReviewAfterApply")
    If Err.Number <> 0 Then
        GetShowCompletionReviewAfterApplySetting = True
        Err.Clear
    Else
        GetShowCompletionReviewAfterApplySetting = (prop.Value = "True")
    End If
    On Error GoTo 0
End Function
Public Sub SetShowCompletionReviewAfterApplySetting(enabled As Boolean)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("CleanupSuiteShowCompletionReviewAfterApply").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteShowCompletionReviewAfterApply", _
        LinkToContent:=False, Type:=msoPropertyTypeString, _
        Value:=IIf(enabled, "True", "False")
    On Error GoTo 0
End Sub
Public Function GetReturnToMainAfterCloseSetting() As Boolean
    ' Returns True by default so closing a tool menu restores the launcher.
    On Error Resume Next
    Dim prop As Object
    Set prop = ActiveDocument.CustomDocumentProperties("CleanupSuiteReturnToMainAfterClose")
    If Err.Number <> 0 Then
        GetReturnToMainAfterCloseSetting = True
        Err.Clear
    Else
        GetReturnToMainAfterCloseSetting = (prop.Value = "True")
    End If
    On Error GoTo 0
End Function
Public Sub SetReturnToMainAfterCloseSetting(enabled As Boolean)
    On Error Resume Next
    ActiveDocument.CustomDocumentProperties("CleanupSuiteReturnToMainAfterClose").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteReturnToMainAfterClose", _
        LinkToContent:=False, Type:=msoPropertyTypeString, _
        Value:=IIf(enabled, "True", "False")
    On Error GoTo 0
End Sub
Public Sub BeginCleanupToolSession()
    gCleanupToolExitReason = ""
End Sub
Public Sub MarkCleanupToolApplied()
    gCleanupToolExitReason = CLEANUP_TOOL_EXIT_APPLY
End Sub
Public Sub MarkCleanupToolClosedByUser()
    If Len(gCleanupToolExitReason) = 0 Then gCleanupToolExitReason = CLEANUP_TOOL_EXIT_CLOSE
End Sub
Public Function ShouldReturnToLauncherAfterToolSession() As Boolean
    Select Case gCleanupToolExitReason
        Case CLEANUP_TOOL_EXIT_APPLY
            ShouldReturnToLauncherAfterToolSession = GetReturnToMainAfterApplySetting()
        Case CLEANUP_TOOL_EXIT_CLOSE
            ShouldReturnToLauncherAfterToolSession = GetReturnToMainAfterCloseSetting()
        Case Else
            ShouldReturnToLauncherAfterToolSession = False
    End Select
    gCleanupToolExitReason = ""
End Function
Public Sub ReturnToMainAfterToolClose(ByVal closingForm As Object, ByRef Cancel As Integer, ByVal CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        MarkCleanupToolClosedByUser
    End If
End Sub
Public Sub ResetCleanupSuiteDefaults()
    On Error Resume Next
    RemoveAllHighlighting ActiveDocument.Content
    SetAutoSaveSetting True
    SetReturnToMainAfterApplySetting True
    SetShowCompletionReviewAfterApplySetting True
    SetReturnToMainAfterCloseSetting True
    BeginCleanupToolSession
    gPreviewActionPanelHasPosition = False
    If Not gPreviewActionPanel Is Nothing Then Unload gPreviewActionPanel
    Set gPreviewActionPanel = Nothing
    Unload frmPunctuationCleanup
    Unload frmUnicodeCleanup
    Unload frmSpacingCleanup
    Unload frmCapitalizationCleanup
    Unload frmListCleanup
    Unload frmParagraphCleanup
    Unload frmDuplicateDetector
    Unload frmFontNormalizer
    Unload frmTableCleaner
    Unload frmBreakNormalizer
    Unload frmDocumentTrim
    Unload frmFormattingStripper
    Unload frmHyperlinkRemover
    Unload frmSoftReturnConverter
    Unload frmMetadataScrubber
    Unload frmStyleCleanup
    Unload frmFootnoteRemover
    Unload frmHeaderFooterStandardizer
    Unload frmObjectRemover
    On Error GoTo 0
End Sub
Public Sub LayoutCleanupToolForm(ByVal toolForm As Object)
    On Error Resume Next
    Const FORM_W As Single = 420
    Const M As Single = 14
    Const GAP As Single = 6
    Const BH As Single = 24
    Dim contentW As Single
    contentW = FORM_W - (2 * M)

    toolForm.Width = FORM_W + 8
    toolForm.BackColor = RGB(248, 249, 251)
    ApplyMSFormTitleStrategy toolForm, True
    toolForm.Controls("chkPreviewOnly").Visible = False
    toolForm.Controls("fraScopeSelection").Visible = False

    Dim titleLabel As Object
    Dim introLabel As Object
    Set titleLabel = EnsureGuidedLabel(toolForm, "lblGuidedToolName")
    Set introLabel = EnsureGuidedLabel(toolForm, "lblGuidedIntro")

    HideGuidedTitleChrome toolForm

    titleLabel.Caption = ""
    titleLabel.Visible = False
    With titleLabel
        .Move 0, 0, 0, 0
        .Font.Name = "Segoe UI"
        .Font.Size = 10
        .Font.Bold = False
        .ForeColor = RGB(32, 37, 45)
        .BackColor = RGB(248, 249, 251)
        .TextAlign = 0
        .WordWrap = False
        .Visible = False
    End With

    Dim introH As Single
    introH = GuidedIntroHeight(GuidedToolIntro(toolForm.Name))
    With introLabel
        .Caption = GuidedToolIntro(toolForm.Name)
        .Move M, 10, contentW, introH
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = False
        .ForeColor = RGB(88, 101, 116)
        .BackColor = RGB(248, 249, 251)
        .TextAlign = 2
        .WordWrap = True
        .Visible = True
        .ZOrder 0
    End With

    Dim y As Single
    y = 10 + introH + 8
    Dim pendingButton As String
    pendingButton = ""
    Dim pendingChoice As String
    pendingChoice = ""
    Dim ctl As Object
    For Each ctl In toolForm.Controls
        If ShouldSkipGuidedLayoutControl(ctl.Name) Then GoTo NextGuidedControl
        If GuidedControlMatchesToolTitle(toolForm, ctl) Then
            HideGuidedControl ctl
            GoTo NextGuidedControl
        End If
        If ShouldHideGuidedControl(toolForm, ctl.Name) Then
            FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
            FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
            HideGuidedControl ctl
            GoTo NextGuidedControl
        End If
        If Left$(ctl.Name, 3) = "fra" Then
            FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
            FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
            HideGuidedControl ctl
            GoTo NextGuidedControl
        End If

        Select Case ctl.Name
            Case "cmdPreview", "cmdRun", "cmdReset", "cmdHelp"
                ' preview and footer actions are laid out after the choices
            Case Else
                Select Case Left$(ctl.Name, 3)
                    Case "cmd"
                        FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
                        If Len(pendingButton) > 0 Then
                            StyleGuidedButton toolForm.Controls(pendingButton), False
                            StyleGuidedButton ctl, False
                            toolForm.Controls(pendingButton).Move M, y, (contentW - GAP) / 2, BH
                            ctl.Move M + ((contentW - GAP) / 2) + GAP, y, (contentW - GAP) / 2, BH
                            toolForm.Controls(pendingButton).Visible = True
                            ctl.Visible = True
                            y = y + BH + GAP
                            pendingButton = ""
                        Else
                            pendingButton = ctl.Name
                        End If
                    Case "lbl"
                        FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
                        FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
                        StyleGuidedBodyLabel ctl
                        ctl.Move M, y, contentW, GuidedLabelHeight(ctl)
                        ctl.Visible = True
                        y = y + ctl.Height + GAP
                    Case "opt", "chk"
                        FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
                        If Len(pendingChoice) > 0 Then
                            PositionGuidedChoicePair toolForm, pendingChoice, ctl, M, y, contentW, GAP
                        Else
                            pendingChoice = ctl.Name
                        End If
                    Case Else
                        FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
                        FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
                End Select
        End Select
NextGuidedControl:
    Next ctl
    FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
    FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
    y = y + 2

    LayoutGuidedPreviewButton toolForm, M, y, contentW
    LayoutGuidedScopeRow toolForm, M, y, contentW
    LayoutGuidedActionButtons toolForm, M, y, contentW
    toolForm.Height = y + 34
End Sub
Private Function EnsureGuidedLabel(ByVal toolForm As Object, ByVal controlName As String) As Object
    On Error Resume Next
    Set EnsureGuidedLabel = toolForm.Controls(controlName)
    On Error GoTo 0
    If EnsureGuidedLabel Is Nothing Then
        On Error Resume Next
        Set EnsureGuidedLabel = toolForm.Controls.Add("Forms.Label.1", controlName, True)
        On Error GoTo 0
    End If
End Function
Private Function ShouldSkipGuidedLayoutControl(ByVal controlName As String) As Boolean
    Select Case controlName
        Case "lblGuidedToolName", "lblGuidedIntro", "chkPreviewOnly", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdPreview", "cmdRun", "cmdReset", "cmdHelp"
            ShouldSkipGuidedLayoutControl = True
        Case Else
            ShouldSkipGuidedLayoutControl = False
    End Select
End Function
Private Sub HideGuidedControl(ByVal ctl As Object)
    On Error Resume Next
    ctl.Visible = False
    ctl.Move 0, 0, 0, 0
End Sub
Private Sub HideGuidedTitleChrome(ByVal toolForm As Object)
    On Error Resume Next
    Dim ctl As Object
    For Each ctl In toolForm.Controls
        If Left$(LCase$(ctl.Name), 3) <> "lbl" Then GoTo NextGuidedTitleControl
        If ctl.Name = "lblGuidedToolName" Or ctl.Name = "lblGuidedIntro" Then GoTo NextGuidedTitleControl
        If ctl.Name = "lblSpeedWarning" Or ctl.Name = "lblFuzzyWarning" Then GoTo NextGuidedTitleControl
        If IsLegacyTitleControlName(ctl.Name) Or GuidedControlMatchesToolTitle(toolForm, ctl) Or IsLikelyGuidedTitleLabel(ctl) Then
            HideGuidedControl ctl
        End If
NextGuidedTitleControl:
    Next ctl
End Sub
Private Function IsLikelyGuidedTitleLabel(ByVal ctl As Object) As Boolean
    On Error Resume Next
    If Len(Trim$(CStr(ctl.Caption))) = 0 Then Exit Function
    If ctl.Top <= 24 And ctl.Height <= 24 Then IsLikelyGuidedTitleLabel = True
End Function
Private Function GuidedControlMatchesToolTitle(ByVal toolForm As Object, ByVal ctl As Object) As Boolean
    On Error Resume Next
    If ctl.Name = "lblGuidedToolName" Then Exit Function
    If Left$(LCase$(ctl.Name), 3) <> "lbl" Then Exit Function
    GuidedControlMatchesToolTitle = (Trim$(ctl.Caption) = GuidedToolName(toolForm.Name))
End Function
Private Function ShouldHideGuidedControl(ByVal toolForm As Object, ByVal controlName As String) As Boolean
    On Error Resume Next
    ShouldHideGuidedControl = False

    If Not GuidedCustomModeIsActive(toolForm) Then
        Select Case controlName
            Case "chkNBSP", "chkZWSP", "chkZWNJ", "chkZWJ", "chkBOM", "chkSoftHyphen", "chkNBHyphen", _
                 "chkCurlyDouble", "chkCurlySingle", "chkEmDash", "chkEnDash", "chkEllipses", _
                 "chkDoubleSpaces", "chkTrimSpaces", "chkSpaceBeforePunct", "chkNormalizeAfterPunct", "chkExtraBlankLines", _
                 "chkNormalizeBullets", "chkNormalizeNumbering", "chkFixIndent", "chkHyphenToBullets", _
                 "chkRemoveEmpty", "chkCollapseBreaks", "chkNormalizeParaSpacing", _
                 "cmdSelectAll", "cmdDeselectAll"
                ShouldHideGuidedControl = True
        End Select
    End If

    If toolForm.Name = "frmBreakNormalizer" And Not toolForm.Controls("chkConvertSectionBreaks").Value Then
        Select Case controlName
            Case "optConvertNextPage", "optConvertContinuous", "optConvertEvenPage", "optConvertOddPage"
                ShouldHideGuidedControl = True
        End Select
    End If

    If toolForm.Name = "frmDuplicateDetector" And Not toolForm.Controls("optMatchFuzzy").Value Then
        Select Case controlName
            Case "optFuzzyLoose", "optFuzzyMedium", "optFuzzyStrict"
                ShouldHideGuidedControl = True
        End Select
    End If

    If toolForm.Name = "frmHeaderFooterStandardizer" Then
        If toolForm.Controls("optClearAll").Value Then
            Select Case controlName
                Case "chkFont", "chkSpacing", "chkAlignment", "optAlignLeft", "optAlignCenter", "optAlignRight", "chkBreakLinks"
                    ShouldHideGuidedControl = True
            End Select
        ElseIf Not toolForm.Controls("chkAlignment").Value Then
            Select Case controlName
                Case "optAlignLeft", "optAlignCenter", "optAlignRight"
                    ShouldHideGuidedControl = True
            End Select
        End If
    End If
End Function
Private Function GuidedCustomModeIsActive(ByVal toolForm As Object) As Boolean
    On Error Resume Next
    GuidedCustomModeIsActive = toolForm.Controls("optCustom").Value
End Function
Private Sub FlushGuidedPendingChoice(ByVal toolForm As Object, ByRef pendingChoice As String, ByVal M As Single, ByRef y As Single, ByVal contentW As Single, ByVal GAP As Single)
    On Error Resume Next
    If Len(pendingChoice) = 0 Then Exit Sub
    Dim singleChoiceW As Single
    singleChoiceW = GuidedSingleChoiceWidth(toolForm.Controls(pendingChoice), contentW)
    Dim rowH As Single
    rowH = GuidedChoiceRowHeight(toolForm.Controls(pendingChoice))
    StyleGuidedOption toolForm.Controls(pendingChoice)
    toolForm.Controls(pendingChoice).Move M + ((contentW - singleChoiceW) / 2), y, singleChoiceW, rowH
    toolForm.Controls(pendingChoice).Visible = True
    y = y + rowH + GAP
    pendingChoice = ""
End Sub
Private Sub PositionGuidedChoicePair(ByVal toolForm As Object, ByRef pendingChoice As String, ByVal choiceCtl As Object, ByVal M As Single, ByRef y As Single, ByVal contentW As Single, ByVal GAP As Single)
    On Error Resume Next
    Dim choiceW As Single
    choiceW = (contentW - GAP) / 2
    Dim rowH As Single
    rowH = GuidedChoiceRowHeight(toolForm.Controls(pendingChoice), choiceCtl)
    StyleGuidedOption toolForm.Controls(pendingChoice)
    StyleGuidedOption choiceCtl
    toolForm.Controls(pendingChoice).Move M + 8, y, choiceW - 8, rowH
    choiceCtl.Move M + choiceW + GAP + 8, y, choiceW - 8, rowH
    toolForm.Controls(pendingChoice).Visible = True
    choiceCtl.Visible = True
    y = y + rowH + GAP
    pendingChoice = ""
End Sub
Private Function GuidedChoiceRowHeight(ByVal firstControl As Object, Optional ByVal secondControl As Object = Nothing) As Single
    On Error Resume Next
    GuidedChoiceRowHeight = GuidedOptionHeight(firstControl)
    If Not secondControl Is Nothing Then
        If GuidedOptionHeight(secondControl) > GuidedChoiceRowHeight Then GuidedChoiceRowHeight = GuidedOptionHeight(secondControl)
    End If
End Function
Private Function GuidedSingleChoiceWidth(ByVal ctl As Object, ByVal contentW As Single) As Single
    On Error Resume Next
    Dim pairW As Single
    pairW = ((contentW - 6) / 2) - 8
    GuidedSingleChoiceWidth = pairW + 32
    If GuidedOptionHeight(ctl) > 18 Then GuidedSingleChoiceWidth = contentW - 40
    If GuidedSingleChoiceWidth < pairW Then GuidedSingleChoiceWidth = pairW
    If GuidedSingleChoiceWidth > contentW Then GuidedSingleChoiceWidth = contentW
End Function
Private Sub FlushGuidedPendingButton(ByVal toolForm As Object, ByRef pendingButton As String, ByVal M As Single, ByRef y As Single, ByVal contentW As Single, ByVal BH As Single, ByVal GAP As Single)
    On Error Resume Next
    If Len(pendingButton) = 0 Then Exit Sub
    StyleGuidedButton toolForm.Controls(pendingButton), False
    toolForm.Controls(pendingButton).Move M, y, contentW, BH
    toolForm.Controls(pendingButton).Visible = True
    y = y + BH + GAP
    pendingButton = ""
End Sub
Private Sub StyleGuidedBodyLabel(ByVal ctl As Object)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 8.5
    ctl.Font.Bold = False
    ctl.ForeColor = RGB(67, 80, 96)
    ctl.BackColor = RGB(255, 255, 255)
    ctl.BorderStyle = 1
    ctl.SpecialEffect = 0
    ctl.WordWrap = True
End Sub
Private Sub StyleGuidedOption(ByVal ctl As Object)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 8.5
    ctl.Font.Bold = False
    ctl.ForeColor = RGB(32, 37, 45)
    ctl.BackColor = RGB(248, 249, 251)
    ctl.WordWrap = True
    ctl.AutoSize = False
End Sub
Private Sub StyleGuidedButton(ByVal ctl As Object, Optional ByVal primary As Boolean = False)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 9
    ctl.Font.Bold = primary
    ctl.BackColor = RGB(255, 255, 255)
    If primary Then
        ctl.ForeColor = RGB(24, 90, 145)
    Else
        ctl.ForeColor = RGB(32, 37, 45)
    End If
    ctl.Enabled = True
End Sub
Private Function GuidedLabelHeight(ByVal ctl As Object) As Single
    On Error Resume Next
    If Len(ctl.Caption) > 72 Then
        GuidedLabelHeight = 42
    Else
        GuidedLabelHeight = 32
    End If
End Function
Private Function GuidedIntroHeight(ByVal captionText As String) As Single
    If Len(captionText) > 72 Then
        GuidedIntroHeight = 30
    Else
        GuidedIntroHeight = 18
    End If
End Function
Private Function GuidedOptionHeight(ByVal ctl As Object) As Single
    On Error Resume Next
    If Len(ctl.Caption) > 72 Then
        GuidedOptionHeight = 44
    ElseIf Len(ctl.Caption) > 34 Then
        GuidedOptionHeight = 28
    Else
        GuidedOptionHeight = 18
    End If
End Function
Private Sub LayoutGuidedScopeRow(ByVal toolForm As Object, ByVal M As Single, ByRef y As Single, ByVal contentW As Single)
    On Error Resume Next
    Const GAP As Single = 6
    Dim choiceW As Single
    choiceW = (contentW - GAP) / 2
    StyleGuidedOption toolForm.Controls("optScopeDocument")
    StyleGuidedOption toolForm.Controls("optScopeSelection")
    If Not ScopeSelectionIsAvailable(toolForm) Then
        HideGuidedControl toolForm.Controls("optScopeDocument")
        HideGuidedControl toolForm.Controls("optScopeSelection")
        Exit Sub
    End If

    toolForm.Controls("optScopeDocument").Move M + 8, y, choiceW - 8, 18
    toolForm.Controls("optScopeSelection").Move M + choiceW + GAP + 8, y, choiceW - 8, 18
    toolForm.Controls("optScopeDocument").Visible = True
    toolForm.Controls("optScopeSelection").Visible = True
    y = y + 28
End Sub
Private Function ScopeSelectionIsAvailable(ByVal toolForm As Object) As Boolean
    On Error Resume Next
    ScopeSelectionIsAvailable = _
        InStr(1, toolForm.Controls("optScopeDocument").Caption, "(always)", vbTextCompare) = 0 _
        And Len(Trim$(toolForm.Controls("optScopeSelection").Caption)) > 0
End Function
Private Sub LayoutGuidedPreviewButton(ByVal toolForm As Object, ByVal M As Single, ByRef y As Single, ByVal contentW As Single)
    On Error Resume Next
    Const GAP As Single = 6
    Const BH As Single = 24
    StyleGuidedButton toolForm.Controls("cmdPreview"), True
    toolForm.Controls("cmdPreview").Caption = "Preview"
    toolForm.Controls("cmdPreview").Move M, y, contentW, BH
    toolForm.Controls("cmdPreview").Visible = True
    y = y + BH + GAP
End Sub
Private Sub LayoutGuidedActionButtons(ByVal toolForm As Object, ByVal M As Single, ByRef y As Single, ByVal contentW As Single)
    On Error Resume Next
    Const GAP As Single = 6
    Const BH As Single = 24
    Dim thirdW As Single
    thirdW = (contentW - (2 * GAP)) / 3
    StyleGuidedButton toolForm.Controls("cmdRun"), False
    StyleGuidedButton toolForm.Controls("cmdReset"), False
    StyleGuidedButton toolForm.Controls("cmdHelp"), False
    toolForm.Controls("cmdRun").Caption = "Apply"
    toolForm.Controls("cmdReset").Caption = "Reset"
    toolForm.Controls("cmdHelp").Caption = "Help"
    toolForm.Controls("cmdRun").Move M, y, thirdW, BH
    toolForm.Controls("cmdReset").Move M + thirdW + GAP, y, thirdW, BH
    toolForm.Controls("cmdHelp").Move M + (2 * (thirdW + GAP)), y, thirdW, BH
    toolForm.Controls("cmdRun").Visible = True
    toolForm.Controls("cmdReset").Visible = True
    toolForm.Controls("cmdHelp").Visible = True
    y = y + BH + 28
End Sub
Private Function GuidedToolName(ByVal formName As String) As String
    Select Case formName
        Case "frmPunctuationCleanup": GuidedToolName = "Punctuation Normalizer"
        Case "frmUnicodeCleanup": GuidedToolName = "Invisible Unicode Cleaner"
        Case "frmSpacingCleanup": GuidedToolName = "Spacing Fixer"
        Case "frmListCleanup": GuidedToolName = "List Normalizer"
        Case "frmParagraphCleanup": GuidedToolName = "Paragraph Structure Fixer"
        Case "frmDuplicateDetector": GuidedToolName = "Duplicate Paragraph Detector"
        Case "frmFontNormalizer": GuidedToolName = "Font Normalizer"
        Case "frmTableCleaner": GuidedToolName = "Table Cleaner"
        Case "frmBreakNormalizer": GuidedToolName = "Break Normalizer"
        Case "frmDocumentTrim": GuidedToolName = "Document Trim"
        Case "frmFormattingStripper": GuidedToolName = "Formatting Stripper"
        Case "frmHyperlinkRemover": GuidedToolName = "Hyperlink Remover"
        Case "frmSoftReturnConverter": GuidedToolName = "Soft Return Converter"
        Case "frmMetadataScrubber": GuidedToolName = "Metadata Scrubber"
        Case "frmStyleCleanup": GuidedToolName = "Style Cleanup"
        Case "frmFootnoteRemover": GuidedToolName = "Footnote / Endnote Remover"
        Case "frmHeaderFooterStandardizer": GuidedToolName = "Header / Footer Standardizer"
        Case "frmObjectRemover": GuidedToolName = "Object Remover"
        Case Else: GuidedToolName = formName
    End Select
End Function
Private Function GuidedToolIntro(ByVal formName As String) As String
    Select Case formName
        Case "frmPunctuationCleanup": GuidedToolIntro = "Choose which punctuation patterns to normalize."
        Case "frmUnicodeCleanup": GuidedToolIntro = "Choose which invisible or problem characters to clean."
        Case "frmSpacingCleanup": GuidedToolIntro = "Choose which spacing problems to fix."
        Case "frmListCleanup": GuidedToolIntro = "Choose how lists, bullets, numbering, and indents should be cleaned."
        Case "frmParagraphCleanup": GuidedToolIntro = "Choose which paragraph structure issues to clean."
        Case "frmDuplicateDetector": GuidedToolIntro = "Choose how repeated or similar paragraphs should be reviewed."
        Case "frmFontNormalizer": GuidedToolIntro = "Choose which direct font overrides to reset."
        Case "frmTableCleaner": GuidedToolIntro = "Choose which table structure and formatting issues to clean."
        Case "frmBreakNormalizer": GuidedToolIntro = "Choose how page and section breaks should be normalized."
        Case "frmDocumentTrim": GuidedToolIntro = "Remove extra empty paragraphs at the end of the document."
        Case "frmFormattingStripper": GuidedToolIntro = "Choose how direct formatting should be stripped while keeping content."
        Case "frmHyperlinkRemover": GuidedToolIntro = "Choose whether link styling should be removed with the hyperlinks."
        Case "frmSoftReturnConverter": GuidedToolIntro = "Choose which line-break conversion to apply."
        Case "frmMetadataScrubber": GuidedToolIntro = "Choose which document metadata and review artifacts to remove."
        Case "frmStyleCleanup": GuidedToolIntro = "Choose which style cleanup actions to run."
        Case "frmFootnoteRemover": GuidedToolIntro = "Choose which notes to remove and whether to keep their text."
        Case "frmHeaderFooterStandardizer": GuidedToolIntro = "Choose how headers and footers should be standardized or cleared."
        Case "frmObjectRemover": GuidedToolIntro = "Choose which embedded objects or hidden content to remove."
        Case Else: GuidedToolIntro = "Choose the cleanup options for this tool."
    End Select
End Function
Public Sub ApplyMSFormTitleStrategy(ByVal toolForm As Object, Optional bodyOwnsTitle As Boolean = False)
    ' Method 1: if the form body already identifies the workflow, blank the native UserForm caption.
    ' Method 2: otherwise keep the native caption and hide only legacy body-title labels.
    On Error Resume Next
    If bodyOwnsTitle Then toolForm.Caption = ""
    HideLegacyTitleControls toolForm
    On Error GoTo 0
End Sub
Private Sub HideLegacyTitleControls(ByVal toolForm As Object)
    On Error Resume Next
    Dim ctl As Object
    For Each ctl In toolForm.Controls
        If IsLegacyTitleControlName(ctl.Name) Or GuidedControlMatchesToolTitle(toolForm, ctl) Then
            ctl.Visible = False
            ctl.Move 0, 0, 0, 0
        End If
    Next ctl
End Sub
Private Function IsLegacyTitleControlName(controlName As String) As Boolean
    Dim nm As String
    nm = LCase$(controlName)
    IsLegacyTitleControlName = (nm = "lbltitle" Or Left$(nm, 8) = "lbltitle")
End Function
Public Sub LayoutCapitalizationCleanupForm(ByVal toolForm As Object)
    On Error Resume Next
    Const M As Single = 14
    Const GAP As Single = 6
    Const BH As Single = 24
    Const FORM_W As Single = 420
    Dim contentW As Single
    contentW = FORM_W - (2 * M)

    toolForm.Width = FORM_W + 8
    toolForm.BackColor = RGB(248, 249, 251)
    ApplyMSFormTitleStrategy toolForm, True
    toolForm.Controls("chkPreviewOnly").Visible = False
    toolForm.Controls("fraScopeSelection").Visible = False

    With toolForm.Controls("lblIntro")
        .Move M, 10, contentW, 18
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = False
        .ForeColor = RGB(88, 101, 116)
        .BackColor = RGB(248, 249, 251)
        .TextAlign = 2
        .WordWrap = False
        .Visible = True
    End With

    Dim y As Single
    y = 36
    Dim optionW As Single
    optionW = (contentW - GAP) / 2
    PositionGuidedChoice toolForm, "optAll", M, y, optionW
    PositionGuidedChoice toolForm, "optSentence", M + optionW + GAP, y, optionW
    y = y + 24
    PositionGuidedChoice toolForm, "optTitle", M, y, optionW
    PositionGuidedChoice toolForm, "optUpper", M + optionW + GAP, y, optionW
    y = y + 24
    PositionGuidedChoice toolForm, "optLower", M, y, optionW
    PositionGuidedChoice toolForm, "optCustom", M + optionW + GAP, y, optionW
    y = y + 28

    With toolForm.Controls("lblModeSummary")
        .Move M, y, contentW, 34
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .Font.Bold = False
        .ForeColor = RGB(67, 80, 96)
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = 1
        .SpecialEffect = 0
        .WordWrap = True
        .Visible = True
    End With
    y = y + 42

    toolForm.Controls("fraCustom").Visible = False
    If toolForm.Controls("optCustom").Value Then
        PositionGuidedCheck toolForm, "chkSentence", M + 8, y, contentW - 16
        PositionGuidedCheck toolForm, "chkTitle", M + 8, y + 18, contentW - 16
        PositionGuidedCheck toolForm, "chkUpper", M + 8, y + 36, contentW - 16
        PositionGuidedCheck toolForm, "chkLower", M + 8, y + 54, contentW - 16
        PositionGuidedCheck toolForm, "chkSmartSentences", M + 8, y + 72, contentW - 16
        y = y + 94
        With toolForm.Controls("cmdSelectAll")
            .Move M, y, (contentW - GAP) / 2, BH
            .Visible = True
            .ZOrder 0
        End With
        With toolForm.Controls("cmdDeselectAll")
            .Move M + ((contentW - GAP) / 2) + GAP, y, (contentW - GAP) / 2, BH
            .Visible = True
            .ZOrder 0
        End With
        y = y + BH + 12
    Else
        HideGuidedCustom toolForm
    End If

    With toolForm.Controls("optScopeDocument")
        .Move M, y, (contentW - GAP) / 2, 18
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .BackColor = RGB(248, 249, 251)
        .Visible = True
    End With
    With toolForm.Controls("optScopeSelection")
        .Move M + ((contentW - GAP) / 2) + GAP, y, (contentW - GAP) / 2, 18
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .BackColor = RGB(248, 249, 251)
        .Visible = True
    End With
    y = y + 28

    With toolForm.Controls("cmdPreview")
        .Move M, y, contentW, BH
        .Caption = "Preview"
        .Font.Bold = True
        .Visible = True
        .Enabled = True
    End With
    y = y + BH + GAP
    With toolForm.Controls("cmdRun")
        .Move M, y, (contentW - (2 * GAP)) / 3, BH
        .Caption = "Apply"
        .Visible = True
        .Enabled = True
    End With
    With toolForm.Controls("cmdReset")
        .Move M + ((contentW - (2 * GAP)) / 3) + GAP, y, (contentW - (2 * GAP)) / 3, BH
        .Caption = "Reset"
        .Visible = True
        .Enabled = True
    End With
    With toolForm.Controls("cmdHelp")
        .Move M + (2 * (((contentW - (2 * GAP)) / 3) + GAP)), y, (contentW - (2 * GAP)) / 3, BH
        .Caption = "Help"
        .Visible = True
        .Enabled = True
    End With
    y = y + BH + 28

    toolForm.Height = y + 34
End Sub
Private Sub PositionGuidedChoice(ByVal toolForm As Object, nm As String, L As Single, T As Single, W As Single)
    On Error Resume Next
    With toolForm.Controls(nm)
        .Move L, T, W, 18
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .Font.Bold = False
        .BackColor = RGB(248, 249, 251)
        .ForeColor = RGB(32, 37, 45)
        .Visible = True
    End With
End Sub
Private Sub PositionGuidedCheck(ByVal toolForm As Object, nm As String, L As Single, T As Single, W As Single)
    On Error Resume Next
    With toolForm.Controls(nm)
        .Move L, T, W, 16
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .BackColor = RGB(248, 249, 251)
        .Visible = True
        .ZOrder 0
    End With
End Sub
Private Sub HideGuidedCustom(ByVal toolForm As Object)
    On Error Resume Next
    toolForm.Controls("fraCustom").Visible = False
    toolForm.Controls("chkSentence").Visible = False
    toolForm.Controls("chkTitle").Visible = False
    toolForm.Controls("chkUpper").Visible = False
    toolForm.Controls("chkLower").Visible = False
    toolForm.Controls("chkSmartSentences").Visible = False
    toolForm.Controls("cmdSelectAll").Visible = False
    toolForm.Controls("cmdDeselectAll").Visible = False
End Sub
Public Function GetTargetRange() As Range
    If Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        Set GetTargetRange = Selection.Range
    Else
        Set GetTargetRange = ActiveDocument.Content
    End If
End Function
Public Function ParagraphsInRange(scopeRange As Range) As Collection
    Dim col As Collection: Set col = New Collection
    Dim p As Paragraph
    For Each p In ActiveDocument.Paragraphs
        If p.Range.Start >= scopeRange.Start And p.Range.End <= scopeRange.End Then col.Add p
    Next p
    Set ParagraphsInRange = col
End Function
Public Sub ShowCleanupReport(reportTitle As String, results As Collection)
    If Not GetShowCompletionReviewAfterApplySetting() Then Exit Sub
    If results.Count = 0 Then
        MsgBox "No items were changed.", vbInformation, reportTitle: Exit Sub
    End If
    Dim reportText As String: reportText = reportTitle & vbCrLf & vbCrLf
    Dim lineItem As Variant
    For Each lineItem In results
        reportText = reportText & "  " & lineItem & vbCrLf
    Next lineItem
    MsgBox reportText, vbInformation, reportTitle
End Sub
