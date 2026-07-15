Option Explicit
' Shared helper functions for cleanup forms
Public gPreviewActionPanel As Object
Private gPreviewActionPanelHasPosition As Boolean
Private gPreviewActionPanelLeft As Single
Private gPreviewActionPanelTop As Single
Private gPreviewScopeHasRange As Boolean
Private gPreviewScopeStart As Long
Private gPreviewScopeEnd As Long
Private gPreviewScopeDocumentKey As String
Private gCleanupToolExitReason As String
Private gPreviewShadingDocuments As Collection
Private gPreviewShadingStarts As Collection
Private gPreviewShadingEnds As Collection
Private gPreviewShadingKeys As Object
Private gPreviewShadingBackgrounds As Collection
Private gPreviewShadingForegrounds As Collection
Private gPreviewShadingTextures As Collection
Private gPreviewViewSessionActive As Boolean
Private gPreviewViewDocument As Document
Private gPreviewViewWindow As Window
Private gPreviewViewSelection As Range
Private gPreviewViewType As WdViewType
Private gPreviewViewReadingLayout As Boolean
Private gPreviewViewZoom As Long
Private gPreviewViewVerticalScroll As Long
Private gPreviewViewHasVerticalScroll As Boolean
Private gCleanupProgressActive As Boolean
Private gCleanupProgressToolName As String
Private gPreviewActionIndicatorActive As Boolean
Private gPreviewActionIndicatorForm As Object
Private gPreviewActionIndicatorCaption As String
Private gPreviewActionIndicatorEnabled As Boolean
Private gPreviewActionIndicatorBackColor As Long
Private gPreviewActionIndicatorForeColor As Long
Private Const CLEANUP_TOOL_EXIT_APPLY As String = "apply"
Private Const CLEANUP_TOOL_EXIT_CLOSE As String = "close"
Private Const CLEANUP_SETTINGS_APP As String = "CleanupSuiteForWord"
Private Const CLEANUP_SETTINGS_SECTION As String = "GlobalDefaults"
Public Const CLEANUP_SUITE_USER_MANUAL_PDF_URL As String = "https://raw.githubusercontent.com/MasseysLab/CleanupSuiteForWord/main/documents/CleanupSuite_User_Manual.pdf"

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
        Case "Hyperlink": CleanupHelpTitle = "Hyperlink Cleaner"
        Case "SoftReturn": CleanupHelpTitle = "Soft Return Converter"
        Case "MetaDataSuite": CleanupHelpTitle = "MetaDataSuite"
        Case "FinalReview": CleanupHelpTitle = "Final Review"
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
            AddHelpLine h, "Repairs messy capitalization without duplicating Word's broad case-conversion commands."
            AddHelpLine h, ""
            AddHelpLine h, "Repair choices:"
            AddHelpLine h, "- Recommended repair: uses every protection for a dependable general cleanup."
            AddHelpLine h, "- Custom repair: uses only the protections you select."
            AddHelpLine h, ""
            AddHelpLine h, "Advanced protections:"
            AddHelpLine h, "- Abbreviation lists: reduces false sentence starts after common abbreviations."
            AddHelpLine h, "- Acronym protection: keeps terms such as API, PDF, or VBA uppercase."
            AddHelpLine h, "- Name and brand protection: restores known mixed-case names such as iPhone or GitHub."
            AddHelpLine h, "- Heading repair: title-cases likely headings; when cleared, suspected headings are left completely unchanged."
            AddHelpLine h, "- Parenthetical heading repair: optionally title-cases balanced words inside parentheses; it is off by default."
            AddHelpLine h, "- Boundary context: treats bullets, quotes, and other sentence-boundary punctuation more carefully."
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
            AddHelpLine h, "Hyperlink Cleaner"
            AddHelpLine h, ""
            AddHelpLine h, "Choose whether to hide visible hyperlink styling while keeping links, or remove hyperlink targets while keeping visible text."
            AddHelpLine h, ""
            AddHelpLine h, "Choices:"
            AddHelpLine h, "- Hide hyperlink styling, keep links: text looks normal, but the link target remains available."
            AddHelpLine h, "- Remove hyperlinks, keep visible text: deletes the link target and leaves the words."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "Select text before opening the tool to limit cleanup to the selection."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview highlights hyperlinks affected by the selected choice."

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

        Case "MetaDataSuite"
            AddHelpLine h, "MetaDataSuite"
            AddHelpLine h, ""
            AddHelpLine h, "Advanced metadata inspection, editing, clearing, and privacy management for Word documents."
            AddHelpLine h, ""
            AddHelpLine h, "Why it is gated:"
            AddHelpLine h, "- This tool can inspect and act across multiple metadata and package areas."
            AddHelpLine h, "- It is meant for deliberate advanced use, not routine one-click cleanup."
            AddHelpLine h, ""
            AddHelpLine h, "Includes:"
            AddHelpLine h, "- Quick document-property and package checks."
            AddHelpLine h, "- Deeper metadata and crypto-pattern investigation."
            AddHelpLine h, "- Safe-edit and grouped metadata-management workflows."
            AddHelpLine h, "- Clear Sharing Properties for populated document info fields before sharing."
            AddHelpLine h, ""
            AddHelpLine h, "Boundary:"
            AddHelpLine h, "- Comments, tracked changes, and author-removal workflows belong in Final Review."
            AddHelpLine h, ""
            AddHelpLine h, "Undoability:"
            AddHelpLine h, "- Some metadata actions are not fully undoable after save. Read each warning before applying changes."

        Case "FinalReview"
            AddHelpLine h, "Final Review"
            AddHelpLine h, ""
            AddHelpLine h, "Finalizes review markup before a document is shared or archived."
            AddHelpLine h, ""
            AddHelpLine h, "Options:"
            AddHelpLine h, "- Remove all comments: deletes reviewer comments from the document."
            AddHelpLine h, "- Accept tracked changes and clear revision marks: finalizes the current revision state."
            AddHelpLine h, ""
            AddHelpLine h, "Scope:"
            AddHelpLine h, "This tool always works on the full document because comments and tracked changes are document-level review artifacts."
            AddHelpLine h, ""
            AddHelpLine h, "Preview:"
            AddHelpLine h, "Preview shows the current comment count and tracked-change count."

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
    Dim rows As Collection
    Set rows = NewPreviewSummaryRows()
    AddPreviewSummaryRow rows, summaryText, ""
    ShowPreviewActionsSummary sourceForm, toolName, rows
End Sub

Public Function NewPreviewSummaryRows() As Collection
    Set NewPreviewSummaryRows = New Collection
End Function

Public Sub AddPreviewSummaryRow(ByVal rows As Collection, ByVal itemText As String, ByVal countValue As Variant, Optional ByVal isUnhighlighted As Boolean = False, Optional ByVal isInactive As Boolean = False)
    If rows Is Nothing Then Exit Sub
    rows.Add PreviewSummaryRowText(itemText, countValue, isUnhighlighted, isInactive)
End Sub

Private Function PreviewSummaryRowText(ByVal itemText As String, ByVal countValue As Variant, ByVal isUnhighlighted As Boolean, ByVal isInactive As Boolean) As String
    Dim displayText As String
    displayText = CleanPreviewSummaryCell(itemText)
    If isUnhighlighted Then displayText = displayText & " (unhighlighted)"
    PreviewSummaryRowText = displayText & vbTab & CStr(countValue) & vbTab & IIf(isInactive, "inactive", "")
End Function

Private Function CleanPreviewSummaryCell(ByVal textValue As String) As String
    CleanPreviewSummaryCell = Replace(textValue, vbTab, " ")
    CleanPreviewSummaryCell = Replace(CleanPreviewSummaryCell, vbCrLf, " ")
    CleanPreviewSummaryCell = Replace(CleanPreviewSummaryCell, vbCr, " ")
    CleanPreviewSummaryCell = Replace(CleanPreviewSummaryCell, vbLf, " ")
    CleanPreviewSummaryCell = Trim$(CleanPreviewSummaryCell)
End Function

Public Sub ShowPreviewActionsSummary(ByVal sourceForm As Object, ByVal toolName As String, ByVal rows As Collection)
    On Error GoTo Fallback
    EndPreviewActionIndicator
    CapturePreviewScopeRange sourceForm
    sourceForm.Hide
    Set gPreviewActionPanel = New frmPreviewActions
    gPreviewActionPanel.ConfigureSummary sourceForm, toolName, rows
    ApplyPreviewActionPanelPosition gPreviewActionPanel
    gPreviewActionPanel.Show vbModal
    ResumePreviewActionPanelModelessIfRequested
    Exit Sub
Fallback:
    RestoreCleanupSuiteTransientState
    MsgBox PreviewSummaryFallbackText(rows) & vbCrLf & vbCrLf & _
           "Preview is still visible in the document. Run again without Preview to apply, or clear highlighting manually.", _
           vbInformation, toolName & " Preview"
End Sub

Private Sub ResumePreviewActionPanelModelessIfRequested()
    Dim resumeModeless As Boolean

    On Error Resume Next
    If gPreviewActionPanel Is Nothing Then Exit Sub
    resumeModeless = CBool(CallByName(gPreviewActionPanel, "ConsumeModelessResumeRequest", VbMethod))
    If Not resumeModeless Then Exit Sub
    ApplyPreviewActionPanelPosition gPreviewActionPanel
    gPreviewActionPanel.Show vbModeless
    On Error GoTo 0
End Sub

Private Function PreviewSummaryFallbackText(ByVal rows As Collection) As String
    Dim rowText As Variant
    Dim rowParts As Variant
    If rows Is Nothing Then Exit Function
    For Each rowText In rows
        rowParts = Split(CStr(rowText), vbTab)
        If Len(PreviewSummaryFallbackText) > 0 Then PreviewSummaryFallbackText = PreviewSummaryFallbackText & vbCrLf
        If UBound(rowParts) >= 1 Then
            PreviewSummaryFallbackText = PreviewSummaryFallbackText & rowParts(0) & ": " & rowParts(1)
        Else
            PreviewSummaryFallbackText = PreviewSummaryFallbackText & CStr(rowText)
        End If
    Next rowText
End Function

Private Sub CapturePreviewScopeRange(ByVal sourceForm As Object)
    ClearPreviewScopeRange
    On Error Resume Next
    If sourceForm Is Nothing Then Exit Sub
    If Not ScopeSelectionIsAvailable(sourceForm) Then Exit Sub
    If Not sourceForm.Controls("optScopeSelection").Value Then Exit Sub
    If Not sourceForm.Controls("optScopeSelection").Enabled Then Exit Sub
    If Selection.Type <> wdSelectionNormal And Selection.Type <> wdSelectionColumn Then Exit Sub

    gPreviewScopeStart = Selection.Range.Start
    gPreviewScopeEnd = Selection.Range.End
    gPreviewScopeDocumentKey = ActiveDocument.FullName
    If Len(gPreviewScopeDocumentKey) = 0 Then gPreviewScopeDocumentKey = ActiveDocument.Name
    gPreviewScopeHasRange = (gPreviewScopeEnd > gPreviewScopeStart)
End Sub

Public Sub ClearPreviewScopeRange()
    gPreviewScopeHasRange = False
    gPreviewScopeStart = 0
    gPreviewScopeEnd = 0
    gPreviewScopeDocumentKey = ""
End Sub

Public Function BeginPreviewViewSession() As Boolean
    Dim previewSelection As Selection

    RestorePreviewViewSession
    On Error GoTo BeginErr
    If Documents.Count = 0 Then Exit Function
    If Windows.Count = 0 Then Exit Function

    Set gPreviewViewDocument = ActiveDocument
    Set gPreviewViewWindow = ActiveWindow
    Set previewSelection = gPreviewViewWindow.Selection
    Set gPreviewViewSelection = previewSelection.Range.Duplicate
    gPreviewViewType = gPreviewViewWindow.View.Type
    gPreviewViewReadingLayout = gPreviewViewWindow.View.ReadingLayout
    gPreviewViewZoom = gPreviewViewWindow.View.Zoom.Percentage
    On Error Resume Next
    Err.Clear
    gPreviewViewVerticalScroll = gPreviewViewWindow.VerticalPercentScrolled
    gPreviewViewHasVerticalScroll = (Err.Number = 0)
    Err.Clear
    On Error GoTo BeginErr
    gPreviewViewSessionActive = True

    If Not gPreviewViewReadingLayout Then
        gPreviewViewWindow.View.ReadingLayout = True
    End If
    BeginPreviewViewSession = True
    Exit Function
BeginErr:
    RestorePreviewViewSession
End Function

Public Sub RestorePreviewViewSession()
    Dim previewDocumentIsLive As Boolean
    Dim previewWindowIsLive As Boolean
    Dim liveDocument As Document
    Dim liveWindow As Window

    On Error Resume Next
    If Not gPreviewViewSessionActive Then
        ClearPreviewViewSessionState
        Exit Sub
    End If

    ' Clear the active flag first so restoration remains idempotent even if Word raises an error.
    gPreviewViewSessionActive = False
    For Each liveDocument In Application.Documents
        If liveDocument Is gPreviewViewDocument Then
            previewDocumentIsLive = True
            Exit For
        End If
    Next liveDocument
    If previewDocumentIsLive Then
        For Each liveWindow In Application.Windows
            If liveWindow Is gPreviewViewWindow Then
                If liveWindow.Document Is gPreviewViewDocument Then previewWindowIsLive = True
                Exit For
            End If
        Next liveWindow
    End If

    If previewWindowIsLive Then
        With gPreviewViewWindow
            If gPreviewViewReadingLayout Then
                .View.ReadingLayout = True
            Else
                .View.ReadingLayout = False
                .View.Type = gPreviewViewType
            End If
            .View.Zoom.Percentage = gPreviewViewZoom
        End With

        ' Do not steal focus if the user switched documents or windows while previewing.
        If ActiveWindow Is gPreviewViewWindow Then
            If Not gPreviewViewSelection Is Nothing Then gPreviewViewSelection.Select
        End If
        If gPreviewViewHasVerticalScroll Then gPreviewViewWindow.VerticalPercentScrolled = gPreviewViewVerticalScroll
    End If

    ClearPreviewViewSessionState
    On Error GoTo 0
End Sub

Private Sub ClearPreviewViewSessionState()
    gPreviewViewSessionActive = False
    Set gPreviewViewSelection = Nothing
    Set gPreviewViewWindow = Nothing
    Set gPreviewViewDocument = Nothing
    gPreviewViewType = wdPrintView
    gPreviewViewReadingLayout = False
    gPreviewViewZoom = 0
    gPreviewViewVerticalScroll = 0
    gPreviewViewHasVerticalScroll = False
End Sub

Public Function PreviewViewSessionIsActive() As Boolean
    PreviewViewSessionIsActive = gPreviewViewSessionActive
End Function

Public Sub RemovePreviewHighlighting(ByVal previewDocument As Document)
    Dim liveDocument As Document
    Dim previewDocumentIsLive As Boolean

    On Error Resume Next
    RestorePreviewShading
    If previewDocument Is Nothing Then Exit Sub
    For Each liveDocument In Application.Documents
        If liveDocument Is previewDocument Then
            previewDocumentIsLive = True
            Exit For
        End If
    Next liveDocument
    If previewDocumentIsLive Then previewDocument.Content.HighlightColorIndex = wdNoHighlight
    On Error GoTo 0
End Sub

Public Sub BeginCleanupProgress(ByVal toolName As String, Optional ByVal phaseText As String = "Starting")
    EndCleanupProgress
    gCleanupProgressToolName = Trim$(toolName)
    On Error Resume Next
    gCleanupProgressActive = True
    PublishCleanupProgress CleanupProgressText(phaseText)
    On Error GoTo 0
End Sub

Public Sub BeginPreviewActionIndicator(ByVal sourceForm As Object)
    Dim previewButton As Object

    EndPreviewActionIndicator
    On Error GoTo SafeExit
    If sourceForm Is Nothing Then Exit Sub
    Set previewButton = sourceForm.Controls("cmdPreview")
    Set gPreviewActionIndicatorForm = sourceForm
    gPreviewActionIndicatorCaption = CStr(previewButton.Caption)
    gPreviewActionIndicatorEnabled = CBool(previewButton.Enabled)
    gPreviewActionIndicatorBackColor = CLng(previewButton.BackColor)
    gPreviewActionIndicatorForeColor = CLng(previewButton.ForeColor)
    gPreviewActionIndicatorActive = True
    previewButton.Caption = "Previewing..."
    previewButton.Enabled = False
    previewButton.BackColor = RGB(221, 235, 247)
    previewButton.ForeColor = RGB(31, 78, 121)
    sourceForm.Repaint
SafeExit:
End Sub

Public Sub EndPreviewActionIndicator()
    Dim previewButton As Object

    On Error Resume Next
    If gPreviewActionIndicatorActive Then
        Set previewButton = gPreviewActionIndicatorForm.Controls("cmdPreview")
        previewButton.Caption = gPreviewActionIndicatorCaption
        previewButton.Enabled = gPreviewActionIndicatorEnabled
        previewButton.BackColor = gPreviewActionIndicatorBackColor
        previewButton.ForeColor = gPreviewActionIndicatorForeColor
        gPreviewActionIndicatorForm.Repaint
    End If
    gPreviewActionIndicatorActive = False
    Set gPreviewActionIndicatorForm = Nothing
    gPreviewActionIndicatorCaption = ""
    On Error GoTo 0
End Sub

Public Sub UpdateCleanupProgress(ByVal phaseText As String, Optional ByVal currentCount As Long = -1, Optional ByVal totalCount As Long = -1)
    If Not gCleanupProgressActive Then Exit Sub
    On Error Resume Next
    PublishCleanupProgress CleanupProgressText(phaseText, currentCount, totalCount)
    On Error GoTo 0
End Sub

Public Sub EndCleanupProgress()
    On Error Resume Next
    If gCleanupProgressActive Then Application.StatusBar = ""
    If Not gPreviewActionPanel Is Nothing Then CallByName gPreviewActionPanel, "ClearProgress", VbMethod
    gCleanupProgressActive = False
    gCleanupProgressToolName = ""
    On Error GoTo 0
End Sub

Private Sub PublishCleanupProgress(ByVal statusText As String)
    On Error Resume Next
    Application.StatusBar = statusText
    If Not gPreviewActionPanel Is Nothing Then CallByName gPreviewActionPanel, "ShowProgress", VbMethod, statusText
    On Error GoTo 0
End Sub

Private Function CleanupProgressText(ByVal phaseText As String, Optional ByVal currentCount As Long = -1, Optional ByVal totalCount As Long = -1) As String
    Dim statusText As String
    statusText = gCleanupProgressToolName
    If Len(Trim$(phaseText)) > 0 Then
        If Len(statusText) > 0 Then statusText = statusText & ": "
        statusText = statusText & Trim$(phaseText)
    End If
    If currentCount >= 0 Then
        statusText = statusText & " - " & Format$(currentCount, "#,##0")
        If totalCount > 0 Then statusText = statusText & " of " & Format$(totalCount, "#,##0")
    End If
    CleanupProgressText = statusText
End Function

Public Function CleanupProgressIsActive() As Boolean
    CleanupProgressIsActive = gCleanupProgressActive
End Function

Public Sub RestoreCleanupSuiteTransientState()
    EndPreviewActionIndicator
    RestorePreviewViewSession
    RestorePreviewShading
    EndCleanupProgress
End Sub

Public Sub RemoveAllHighlighting(Optional scopeRange As Range = Nothing)
    RestorePreviewShading
    Dim hlRange As Range
    If scopeRange Is Nothing Then
        Set hlRange = ActiveDocument.Content.Duplicate
    Else
        Set hlRange = scopeRange.Duplicate
    End If
    hlRange.HighlightColorIndex = wdNoHighlight
End Sub
Public Sub ApplyPreviewShading(ByVal targetRange As Range, Optional ByVal shadeColor As Long = wdColorBrightGreen)
    Dim paragraphItem As Paragraph
    Dim paragraphRange As Range

    On Error GoTo SafeExit
    For Each paragraphItem In targetRange.Paragraphs
        Set paragraphRange = paragraphItem.Range.Duplicate
        ApplyPreviewShadingToParagraph paragraphRange, shadeColor
    Next paragraphItem
SafeExit:
End Sub

Private Sub ApplyPreviewShadingToParagraph(ByVal paragraphRange As Range, ByVal shadeColor As Long)
    Dim shadingKey As String

    On Error GoTo SafeExit
    If gPreviewShadingDocuments Is Nothing Then
        Set gPreviewShadingDocuments = New Collection
        Set gPreviewShadingStarts = New Collection
        Set gPreviewShadingEnds = New Collection
        Set gPreviewShadingKeys = CreateObject("Scripting.Dictionary")
        Set gPreviewShadingBackgrounds = New Collection
        Set gPreviewShadingForegrounds = New Collection
        Set gPreviewShadingTextures = New Collection
    End If
    shadingKey = PreviewShadingDocumentIdentity(paragraphRange.Document) & "|" & _
                 CStr(paragraphRange.Start) & ":" & CStr(paragraphRange.End)
    If gPreviewShadingKeys.Exists(shadingKey) Then
        paragraphRange.ParagraphFormat.Shading.Texture = wdTextureNone
        paragraphRange.ParagraphFormat.Shading.BackgroundPatternColor = shadeColor
        Exit Sub
    End If
    gPreviewShadingKeys.Add shadingKey, True
    gPreviewShadingDocuments.Add paragraphRange.Document
    gPreviewShadingStarts.Add CLng(paragraphRange.Start)
    gPreviewShadingEnds.Add CLng(paragraphRange.End)
    gPreviewShadingBackgrounds.Add CLng(paragraphRange.ParagraphFormat.Shading.BackgroundPatternColor)
    gPreviewShadingForegrounds.Add CLng(paragraphRange.ParagraphFormat.Shading.ForegroundPatternColor)
    gPreviewShadingTextures.Add CLng(paragraphRange.ParagraphFormat.Shading.Texture)
    paragraphRange.ParagraphFormat.Shading.Texture = wdTextureNone
    paragraphRange.ParagraphFormat.Shading.BackgroundPatternColor = shadeColor
SafeExit:
End Sub

Private Function PreviewShadingDocumentIdentity(ByVal previewDocument As Document) As String
    On Error Resume Next
    PreviewShadingDocumentIdentity = LCase$(previewDocument.FullName)
    If Len(PreviewShadingDocumentIdentity) = 0 Then
        PreviewShadingDocumentIdentity = LCase$(previewDocument.Name)
    End If
    On Error GoTo 0
End Function

Private Sub RestorePreviewShading()
    Dim i As Long
    Dim previewDocument As Document
    Dim liveDocument As Document
    Dim restoreRange As Range
    Dim previewDocumentIsLive As Boolean
    Dim rangeStart As Long
    Dim rangeEnd As Long
    Dim documentEnd As Long
    Dim expectedBackground As Long
    Dim expectedForeground As Long
    Dim expectedTexture As Long
    Dim allLiveRecordsRestored As Boolean

    On Error Resume Next
    allLiveRecordsRestored = True
    If Not gPreviewShadingDocuments Is Nothing Then
        For i = gPreviewShadingDocuments.Count To 1 Step -1
            Set previewDocument = Nothing
            Set previewDocument = gPreviewShadingDocuments(i)
            previewDocumentIsLive = False
            For Each liveDocument In Application.Documents
                If liveDocument Is previewDocument Then
                    previewDocumentIsLive = True
                    Exit For
                End If
            Next liveDocument
            If previewDocumentIsLive Then
                rangeStart = CLng(gPreviewShadingStarts(i))
                rangeEnd = CLng(gPreviewShadingEnds(i))
                documentEnd = previewDocument.Content.End
                If rangeStart < 0 Then rangeStart = 0
                If rangeStart > documentEnd Then rangeStart = documentEnd
                If rangeEnd < rangeStart Then rangeEnd = rangeStart
                If rangeEnd > documentEnd Then rangeEnd = documentEnd
                Set restoreRange = previewDocument.Range(rangeStart, rangeEnd)
                expectedTexture = CLng(gPreviewShadingTextures(i))
                expectedForeground = CLng(gPreviewShadingForegrounds(i))
                expectedBackground = CLng(gPreviewShadingBackgrounds(i))
                Err.Clear
                restoreRange.Shading.Texture = expectedTexture
                restoreRange.Shading.ForegroundPatternColor = expectedForeground
                restoreRange.Shading.BackgroundPatternColor = expectedBackground
                restoreRange.ParagraphFormat.Shading.Texture = expectedTexture
                restoreRange.ParagraphFormat.Shading.ForegroundPatternColor = expectedForeground
                restoreRange.ParagraphFormat.Shading.BackgroundPatternColor = expectedBackground
                If Err.Number <> 0 Then
                    allLiveRecordsRestored = False
                ElseIf CLng(restoreRange.ParagraphFormat.Shading.Texture) <> expectedTexture Or _
                       CLng(restoreRange.ParagraphFormat.Shading.ForegroundPatternColor) <> expectedForeground Or _
                       CLng(restoreRange.ParagraphFormat.Shading.BackgroundPatternColor) <> expectedBackground Then
                    allLiveRecordsRestored = False
                End If
                Err.Clear
            End If
        Next i
    End If
    If allLiveRecordsRestored Then
        Set gPreviewShadingDocuments = Nothing
        Set gPreviewShadingStarts = Nothing
        Set gPreviewShadingEnds = Nothing
        Set gPreviewShadingKeys = Nothing
        Set gPreviewShadingBackgrounds = Nothing
        Set gPreviewShadingForegrounds = Nothing
        Set gPreviewShadingTextures = Nothing
    End If
    On Error GoTo 0
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
    BeginCleanupProgress toolName, "Applying changes"
    If IsRunningInMacroHostDocument() Then Exit Sub
    ActiveDocument.CustomDocumentProperties("CleanupSuiteInProgress").Delete
    Err.Clear
    ActiveDocument.CustomDocumentProperties.Add Name:="CleanupSuiteInProgress", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=toolName
    On Error GoTo 0
End Sub
Public Sub MarkCleanupEnd()
    On Error Resume Next
    RestoreCleanupSuiteTransientState
    If IsRunningInMacroHostDocument() Then Exit Sub
    ActiveDocument.CustomDocumentProperties("CleanupSuiteInProgress").Delete
    On Error GoTo 0
End Sub
Public Function GetAutoSaveSetting() As Boolean
    GetAutoSaveSetting = GetGlobalBooleanSetting("AutoSave", True)
End Function
Public Sub SetAutoSaveSetting(enabled As Boolean)
    SaveGlobalBooleanSetting "AutoSave", enabled
End Sub
Public Function GetReturnToMainAfterApplySetting() As Boolean
    GetReturnToMainAfterApplySetting = GetGlobalBooleanSetting("ReturnToMainAfterApply", True)
End Function
Public Sub SetReturnToMainAfterApplySetting(enabled As Boolean)
    SaveGlobalBooleanSetting "ReturnToMainAfterApply", enabled
End Sub
Public Function GetReturnToMainAfterCloseSetting() As Boolean
    GetReturnToMainAfterCloseSetting = GetGlobalBooleanSetting("ReturnToMainAfterClose", True)
End Function
Public Sub SetReturnToMainAfterCloseSetting(enabled As Boolean)
    SaveGlobalBooleanSetting "ReturnToMainAfterClose", enabled
End Sub
Private Function GetGlobalBooleanSetting(ByVal settingName As String, ByVal defaultValue As Boolean) As Boolean
    Dim defaultText As String
    Dim settingText As String
    defaultText = IIf(defaultValue, "True", "False")
    On Error GoTo UseDefault
    settingText = GetSetting(CLEANUP_SETTINGS_APP, CLEANUP_SETTINGS_SECTION, settingName, defaultText)
    GetGlobalBooleanSetting = (StrComp(Trim$(settingText), "True", vbTextCompare) = 0)
    Exit Function
UseDefault:
    GetGlobalBooleanSetting = defaultValue
End Function
Private Sub SaveGlobalBooleanSetting(ByVal settingName As String, ByVal enabled As Boolean)
    On Error Resume Next
    SaveSetting CLEANUP_SETTINGS_APP, CLEANUP_SETTINGS_SECTION, settingName, IIf(enabled, "True", "False")
    On Error GoTo 0
End Sub
Public Sub BeginCleanupToolSession()
    RestoreCleanupSuiteTransientState
    gCleanupToolExitReason = ""
End Sub
Public Sub MarkCleanupToolApplied()
    RestoreCleanupSuiteTransientState
    gCleanupToolExitReason = CLEANUP_TOOL_EXIT_APPLY
End Sub
Public Sub MarkCleanupToolClosedByUser()
    RestoreCleanupSuiteTransientState
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
Public Sub ReturnToLauncherAfterDetachedToolSession()
    If ShouldReturnToLauncherAfterToolSession() Then ShowCleanupSuiteLauncher
End Sub
Public Sub ResetCleanupSuiteDefaults()
    On Error Resume Next
    RestoreCleanupSuiteTransientState
    RemoveAllHighlighting ActiveDocument.Content
    ClearPreviewScopeRange
    SetAutoSaveSetting True
    SetReturnToMainAfterApplySetting True
    SetReturnToMainAfterCloseSetting True
    SetCleanupSuiteUpdateChecksEnabled True
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
    Unload frmMetaDataSuite
    Unload frmFinalReview
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
    Const GAP As Single = 4
    Const BH As Single = 24
    Dim contentW As Single
    contentW = FORM_W - (2 * M)

    toolForm.Width = FORM_W + 8
    toolForm.BackColor = RGB(248, 249, 251)
    ApplyMSFormTitleStrategy toolForm, True
    RefreshScopeSelectionState toolForm
    toolForm.Controls("chkPreviewOnly").Visible = False
    toolForm.Controls("fraScopeSelection").Visible = False

    Dim introLabel As Object
    RemoveGuidedGeneratedLabel toolForm, "lblGuidedToolName"
    RemoveGuidedGeneratedLabel toolForm, "lblGuidedIntro"
    Set introLabel = EnsureGuidedLabel(toolForm, "lblGuidedIntro")

    HideGuidedTitleChrome toolForm
    ResetGuidedGeneratedChrome toolForm
    LayoutGuidedTitleReset toolForm, M, contentW

    Dim introH As Single
    introH = GuidedIntroHeight(GuidedToolIntro(toolForm.Name))
    With introLabel
        .Caption = GuidedToolIntro(toolForm.Name)
        .Move M, 8, contentW - 52, introH
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = True
        .ForeColor = RGB(88, 101, 116)
        .BackColor = RGB(248, 249, 251)
        .TextAlign = 2
        .WordWrap = True
        .Visible = True
        .ZOrder 0
    End With

    Dim y As Single
    y = 8 + introH + 4

    HideGuidedExternalWarnings toolForm

    If toolForm.Name = "frmHeaderFooterStandardizer" Then
        LayoutHeaderFooterChoices toolForm, M, y, contentW, GAP
        ApplyGuidedRiskPlacementLayout toolForm
        LayoutGuidedInfoBox toolForm, M, y, contentW
        LayoutGuidedPreviewAndScopeRow toolForm, M, y, contentW
        toolForm.Height = y + 26
        Exit Sub
    End If

    If toolForm.Name = "frmCapitalizationCleanup" Then
        LayoutCapitalizationChoices toolForm, M, y, contentW, BH, GAP
        ApplyGuidedRiskPlacementLayout toolForm
        LayoutGuidedInfoBox toolForm, M, y, contentW
        LayoutGuidedPreviewAndScopeRow toolForm, M, y, contentW
        toolForm.Height = y + 26
        Exit Sub
    End If

    Dim pendingButton As String
    pendingButton = ""
    Dim pendingChoice As String
    pendingChoice = ""
    Dim customDividerShown As Boolean
    customDividerShown = False
    Dim ctl As Object
    For Each ctl In toolForm.Controls
        If ShouldSkipGuidedLayoutControl(ctl.Name) Then GoTo NextGuidedControl
        If GuidedControlMatchesToolTitle(toolForm, ctl) Then
            HideGuidedControl ctl
            GoTo NextGuidedControl
        End If
        If ShouldHideGuidedControl(toolForm, ctl.Name) Then
            HideGuidedControl ctl
            GoTo NextGuidedControl
        End If
        If Left$(ctl.Name, 3) = "fra" Then
            FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
            FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
            HideGuidedControl ctl
            GoTo NextGuidedControl
        End If

        If Not customDividerShown Then
            If GuidedCustomModeIsActive(toolForm) And IsGuidedCustomChoiceControl(toolForm.Name, ctl.Name) Then
                FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
                FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
                LayoutGuidedDivider toolForm, M, y, contentW
                customDividerShown = True
            ElseIf toolForm.Name = "frmCapitalizationCleanup" And GuidedCustomModeIsActive(toolForm) And IsCapitalizationAdvancedControl(ctl.Name) Then
                FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
                FlushGuidedPendingButton toolForm, pendingButton, M, y, contentW, BH, GAP
                LayoutGuidedDivider toolForm, M, y, contentW
                customDividerShown = True
            End If
        End If

        Select Case ctl.Name
            Case "cmdPreview", "cmdRun", "cmdReset"
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
    y = y + 1

    LayoutGuidedInfoBox toolForm, M, y, contentW
    LayoutGuidedPreviewAndScopeRow toolForm, M, y, contentW
    toolForm.Height = y + 26
    ApplyGuidedRiskPlacementLayout toolForm
End Sub
Public Sub LayoutCleanupToolFormForReconfigure(ByVal toolForm As Object)
    LayoutCleanupToolForm toolForm
End Sub
Private Sub ApplyGuidedRiskPlacementLayout(ByVal toolForm As Object)
    On Error Resume Next
    Select Case toolForm.Name
        Case "frmUnicodeCleanup"
            LayoutUnicodeRiskChoiceChips toolForm
        Case "frmPunctuationCleanup"
            LayoutPunctuationRiskChoiceChips toolForm
        Case "frmSpacingCleanup"
            LayoutSpacingRiskChoiceChips toolForm
        Case "frmCapitalizationCleanup"
            LayoutCapitalizationRiskChoiceChips toolForm
        Case "frmListCleanup"
            LayoutListRiskChoiceChips toolForm
        Case "frmParagraphCleanup"
            LayoutParagraphRiskChoiceChips toolForm
        Case "frmDuplicateDetector"
            LayoutDuplicateRiskChoiceChips toolForm
        Case "frmFormattingStripper"
            LayoutFormattingRiskChoiceChips toolForm
        Case "frmFontNormalizer"
            LayoutFontRiskChoiceChips toolForm
        Case "frmStyleCleanup"
            LayoutStyleRiskChoiceChips toolForm
        Case "frmTableCleaner"
            LayoutTableRiskChoiceChips toolForm
        Case "frmBreakNormalizer"
            LayoutBreakRiskChoiceChips toolForm
        Case "frmHeaderFooterStandardizer"
            LayoutHeaderFooterRiskChoiceChips toolForm
        Case "frmFootnoteRemover"
            LayoutFootnoteRiskChoiceChips toolForm
        Case "frmObjectRemover"
            LayoutObjectRiskChoiceChips toolForm
        Case "frmDocumentTrim"
            LayoutDocumentTrimRiskSummary toolForm
        Case "frmSoftReturnConverter"
            LayoutSoftReturnRiskChoiceChips toolForm
        Case "frmFinalReview"
            LayoutFinalReviewRiskSummary toolForm
        Case "frmHyperlinkRemover"
            LayoutHyperlinkRiskChoiceChips toolForm
    End Select
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
Private Sub ResetGuidedGeneratedChrome(ByVal toolForm As Object)
    On Error Resume Next
    Dim i As Long
    RemoveGuidedGeneratedLabel toolForm, "lblGuidedToolName"
    HideGuidedGeneratedLabel toolForm, "lblGuidedDivider"
    HideGuidedGeneratedLabel toolForm, "lblGuidedInfoBorder"
    HideGuidedGeneratedLabel toolForm, "lblGuidedInfoMeasure"
    For i = 1 To 20
        HideGuidedGeneratedLabel toolForm, "lblGuidedInfoLine" & CStr(i)
        HideGuidedGeneratedLabel toolForm, "lblGuidedInfoName" & CStr(i)
        HideGuidedGeneratedLabel toolForm, "lblGuidedInfoDetail" & CStr(i)
    Next i
    HideRiskPlacementControls toolForm
End Sub
Private Sub RemoveGuidedGeneratedLabel(ByVal toolForm As Object, ByVal controlName As String)
    On Error Resume Next
    toolForm.Controls.Remove controlName
End Sub
Private Sub HideGuidedGeneratedLabel(ByVal toolForm As Object, ByVal controlName As String)
    On Error Resume Next
    HideGuidedControl toolForm.Controls(controlName)
End Sub
Private Sub HideGuidedExternalWarnings(ByVal toolForm As Object)
    On Error Resume Next
    HideGuidedControl toolForm.Controls("lblFuzzyWarning")
    HideGuidedControl toolForm.Controls("lblSpeedWarning")
    HideGuidedControl toolForm.Controls("lblModeSummary")
End Sub
Private Function ShouldSkipGuidedLayoutControl(ByVal controlName As String) As Boolean
    If InStr(1, controlName, "RiskPlacement", vbTextCompare) > 0 Then
        ShouldSkipGuidedLayoutControl = True
        Exit Function
    End If
    If controlName = "lblTitle" Or controlName = "lblIntro" Or controlName = "lblModeSummary" Or controlName = "lblSpeedWarning" Or controlName = "lblFuzzyWarning" Or controlName = "lblGuidedDivider" Then
        ShouldSkipGuidedLayoutControl = True
        Exit Function
    End If
    If Left$(controlName, 13) = "lblGuidedInfo" Then
        ShouldSkipGuidedLayoutControl = True
        Exit Function
    End If
    Select Case controlName
        Case "lblGuidedIntro", "chkPreviewOnly", "chkHeadingParentheses", "fraScopeSelection", "optScopeDocument", "optScopeSelection", "cmdPreview", "cmdRun", "cmdReset", "cmdHelp"
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
Private Sub HideRiskPlacementControls(ByVal toolForm As Object)
    On Error Resume Next
    Dim ctl As Object
    For Each ctl In toolForm.Controls
        If InStr(1, ctl.Name, "RiskPlacement", vbTextCompare) > 0 Then HideGuidedControl ctl
    Next ctl
End Sub
Private Function EnsureRiskPlacementLabel(ByVal toolForm As Object, ByVal controlName As String) As Object
    On Error Resume Next
    Set EnsureRiskPlacementLabel = toolForm.Controls(controlName)
    On Error GoTo 0
    If EnsureRiskPlacementLabel Is Nothing Then
        On Error Resume Next
        Set EnsureRiskPlacementLabel = toolForm.Controls.Add("Forms.Label.1", controlName, True)
        On Error GoTo 0
    End If
End Function
Private Function EnsureRiskPlacementChip(ByVal toolForm As Object, ByVal controlName As String) As Object
    On Error Resume Next
    Set EnsureRiskPlacementChip = toolForm.Controls(controlName)
    On Error GoTo 0
    If EnsureRiskPlacementChip Is Nothing Then
        On Error Resume Next
        Set EnsureRiskPlacementChip = toolForm.Controls.Add("Forms.CommandButton.1", controlName, True)
        On Error GoTo 0
    End If
End Function
Private Sub StyleToolRiskChip(ByVal chipCtl As Object, ByVal labelText As String)
    On Error Resume Next
    With chipCtl
        .Caption = Replace(labelText, " ", vbCrLf)
        .Font.Name = "Segoe UI"
        .Font.Size = 7
        .Font.Bold = True
        .WordWrap = True
        .TakeFocusOnClick = False
        .BackColor = ToolRiskBackColor(labelText)
        .ForeColor = ToolRiskForeColor(labelText)
        .Visible = True
        .ZOrder 0
    End With
End Sub
Private Function ToolRiskBackColor(ByVal labelText As String) As Long
    Select Case labelText
        Case "Safe cleanup", "Safe": ToolRiskBackColor = RGB(229, 245, 235)
        Case "Preview first", "Inspect first": ToolRiskBackColor = RGB(228, 239, 251)
        Case "Text caution", "Formatting change": ToolRiskBackColor = RGB(255, 246, 218)
        Case "Structure change", "Removes content": ToolRiskBackColor = RGB(255, 235, 218)
        Case "Final action", "Privacy cleanup": ToolRiskBackColor = RGB(255, 226, 226)
        Case Else: ToolRiskBackColor = RGB(240, 242, 245)
    End Select
End Function
Private Function ToolRiskForeColor(ByVal labelText As String) As Long
    Select Case labelText
        Case "Safe cleanup", "Safe": ToolRiskForeColor = RGB(24, 105, 59)
        Case "Preview first", "Inspect first": ToolRiskForeColor = RGB(31, 91, 145)
        Case "Text caution", "Formatting change": ToolRiskForeColor = RGB(132, 89, 0)
        Case "Structure change", "Removes content": ToolRiskForeColor = RGB(154, 73, 24)
        Case "Final action", "Privacy cleanup": ToolRiskForeColor = RGB(171, 40, 40)
        Case Else: ToolRiskForeColor = RGB(84, 91, 104)
    End Select
End Function
Public Sub ShowToolRiskChoiceExplanation(ByVal actionName As String, ByVal riskLabel As String, ByVal reasonText As String)
    Dim msg As String
    msg = actionName & vbCrLf & vbCrLf & _
          "Risk label: " & riskLabel & vbCrLf & vbCrLf & _
          "Why this choice says that:" & vbCrLf & _
          reasonText & vbCrLf & vbCrLf & _
          "Preview first when you are not sure which result you want."
    MsgBox msg, vbInformation, "Why This Risk?"
End Sub
Private Sub ShiftControlsBelow(ByVal toolForm As Object, ByVal cutoffTop As Single, ByVal deltaY As Single)
    On Error Resume Next
    Dim ctl As Object
    For Each ctl In toolForm.Controls
        If ctl.Visible And ctl.Top >= cutoffTop Then
            If InStr(1, ctl.Name, "RiskPlacement", vbTextCompare) = 0 Then ctl.Top = ctl.Top + deltaY
        End If
    Next ctl
    toolForm.Height = toolForm.Height + deltaY
End Sub
Private Sub PlaceRiskChipBesideChoice(ByVal toolForm As Object, ByVal choiceName As String, ByVal chipName As String, ByVal labelText As String)
    On Error Resume Next
    Dim choiceCtl As Object
    Set choiceCtl = toolForm.Controls(choiceName)
    If choiceCtl Is Nothing Or Not choiceCtl.Visible Then Exit Sub

    Const CHIP_W As Single = 58
    Const CHIP_H As Single = 20
    Const GAP As Single = 4
    Dim chipCtl As Object
    Set chipCtl = EnsureRiskPlacementChip(toolForm, chipName)
    StyleToolRiskChip chipCtl, labelText
    If choiceCtl.Width > CHIP_W + 32 Then choiceCtl.Width = choiceCtl.Width - CHIP_W - GAP
    chipCtl.Move choiceCtl.Left + choiceCtl.Width + GAP, choiceCtl.Top, CHIP_W, CHIP_H
End Sub
Public Sub LayoutUnicodeRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "optAll", "cmdRiskPlacementUnicodeAll", "Safe cleanup"
    PlaceRiskChipBesideChoice toolForm, "optNBSP", "cmdRiskPlacementUnicodeNBSP", "Safe cleanup"
    PlaceRiskChipBesideChoice toolForm, "optZeroWidth", "cmdRiskPlacementUnicodeZeroWidth", "Safe cleanup"
    PlaceRiskChipBesideChoice toolForm, "optCustom", "cmdRiskPlacementUnicodeCustom", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "chkNBSP", "cmdRiskPlacementUnicodeChkNBSP", "Safe cleanup"
    PlaceRiskChipBesideChoice toolForm, "chkZWSP", "cmdRiskPlacementUnicodeChkZWSP", "Safe cleanup"
    PlaceRiskChipBesideChoice toolForm, "chkZWNJ", "cmdRiskPlacementUnicodeChkZWNJ", "Safe cleanup"
    PlaceRiskChipBesideChoice toolForm, "chkZWJ", "cmdRiskPlacementUnicodeChkZWJ", "Safe cleanup"
    PlaceRiskChipBesideChoice toolForm, "chkBOM", "cmdRiskPlacementUnicodeChkBOM", "Safe cleanup"
    PlaceRiskChipBesideChoice toolForm, "chkSoftHyphen", "cmdRiskPlacementUnicodeChkSoftHyphen", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkNBHyphen", "cmdRiskPlacementUnicodeChkNBHyphen", "Formatting change"
End Sub
Public Sub LayoutPunctuationRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "optAll", "cmdRiskPlacementPunctuationAll", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "optQuotes", "cmdRiskPlacementPunctuationQuotes", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "optDashes", "cmdRiskPlacementPunctuationDashes", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "optEllipses", "cmdRiskPlacementPunctuationEllipses", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "optCustom", "cmdRiskPlacementPunctuationCustom", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "chkCurlyDouble", "cmdRiskPlacementPunctuationChkCurlyDouble", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "chkCurlySingle", "cmdRiskPlacementPunctuationChkCurlySingle", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "chkEmDash", "cmdRiskPlacementPunctuationChkEmDash", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "chkEnDash", "cmdRiskPlacementPunctuationChkEnDash", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "chkEllipses", "cmdRiskPlacementPunctuationChkEllipses", "Text caution"
End Sub
Public Sub LayoutSpacingRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "optAll", "cmdRiskPlacementSpacingAll", "Safe"
    PlaceRiskChipBesideChoice toolForm, "optDoubleSpaces", "cmdRiskPlacementSpacingDouble", "Safe"
    PlaceRiskChipBesideChoice toolForm, "optTrim", "cmdRiskPlacementSpacingTrim", "Safe"
    PlaceRiskChipBesideChoice toolForm, "optCustom", "cmdRiskPlacementSpacingCustom", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "chkDoubleSpaces", "cmdRiskPlacementSpacingChkDouble", "Safe"
    PlaceRiskChipBesideChoice toolForm, "chkTrimSpaces", "cmdRiskPlacementSpacingChkTrim", "Safe"
    PlaceRiskChipBesideChoice toolForm, "chkSpaceBeforePunct", "cmdRiskPlacementSpacingChkBefore", "Safe"
    PlaceRiskChipBesideChoice toolForm, "chkNormalizeAfterPunct", "cmdRiskPlacementSpacingChkAfter", "Safe"
    PlaceRiskChipBesideChoice toolForm, "chkExtraBlankLines", "cmdRiskPlacementSpacingChkBlank", "Structure change"
End Sub
Public Sub LayoutCapitalizationRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "optAll", "cmdRiskPlacementCapitalizationRecommended", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "optCustom", "cmdRiskPlacementCapitalizationCustom", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "chkSentence", "cmdRiskPlacementCapitalizationChkSentence", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "chkTitle", "cmdRiskPlacementCapitalizationChkTitle", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "chkUpper", "cmdRiskPlacementCapitalizationChkUpper", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "chkLower", "cmdRiskPlacementCapitalizationChkLower", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "chkSmartSentences", "cmdRiskPlacementCapitalizationChkSmartSentences", "Text caution"
    PlaceRiskChipBesideChoice toolForm, "chkHeadingParentheses", "cmdRiskPlacementCapitalizationChkHeadingParentheses", "Text caution"
End Sub
Public Sub LayoutListRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "optAll", "cmdRiskPlacementListAll", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "optBullets", "cmdRiskPlacementListBullets", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "optNumbering", "cmdRiskPlacementListNumbering", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "optIndent", "cmdRiskPlacementListIndent", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "optCustom", "cmdRiskPlacementListCustom", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "chkNormalizeBullets", "cmdRiskPlacementListNormalizeBullets", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkNormalizeNumbering", "cmdRiskPlacementListNormalizeNumbering", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkFixIndent", "cmdRiskPlacementListFixIndent", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkHyphenToBullets", "cmdRiskPlacementListHyphenBullets", "Structure change"
End Sub
Private Function ParagraphChoiceRiskLabel(ByVal choiceName As String) As String
    Select Case choiceName
        Case "chkRemoveEmpty", "optRemoveEmpty": ParagraphChoiceRiskLabel = "Safe cleanup"
        Case "chkCollapseBreaks": ParagraphChoiceRiskLabel = "Structure change"
        Case "chkNormalizeParaSpacing", "optNormalizeSpacing": ParagraphChoiceRiskLabel = "Formatting change"
        Case "chkFixIndent": ParagraphChoiceRiskLabel = "Structure change"
        Case "optAll", "optCustom": ParagraphChoiceRiskLabel = "Structure change"
        Case Else: ParagraphChoiceRiskLabel = "Inspect first"
    End Select
End Function
Public Sub LayoutParagraphRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "optAll", "cmdRiskPlacementParagraphAll", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "optRemoveEmpty", "cmdRiskPlacementParagraphRemoveEmpty", "Safe cleanup"
    PlaceRiskChipBesideChoice toolForm, "optNormalizeSpacing", "cmdRiskPlacementParagraphSpacing", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "optCustom", "cmdRiskPlacementParagraphCustom", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "chkRemoveEmpty", "cmdRiskPlacementParagraphChkRemoveEmpty", "Safe cleanup"
    PlaceRiskChipBesideChoice toolForm, "chkCollapseBreaks", "cmdRiskPlacementParagraphChkBreaks", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkNormalizeParaSpacing", "cmdRiskPlacementParagraphChkSpacing", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkFixIndent", "cmdRiskPlacementParagraphChkIndent", "Structure change"
End Sub
Public Sub LayoutDuplicateRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "optHighlightOnly", "cmdRiskPlacementDuplicatePreview", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "optRemoveDupes", "cmdRiskPlacementDuplicateRemove", "Removes content"
    PlaceRiskChipBesideChoice toolForm, "optMatchExact", "cmdRiskPlacementDuplicateExact", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "optMatchNormalized", "cmdRiskPlacementDuplicateNormalized", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "optMatchFuzzy", "cmdRiskPlacementDuplicateFuzzy", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "optFuzzyLoose", "cmdRiskPlacementDuplicateFuzzyLoose", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "optFuzzyMedium", "cmdRiskPlacementDuplicateFuzzyMedium", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "optFuzzyStrict", "cmdRiskPlacementDuplicateFuzzyStrict", "Inspect first"
End Sub
Public Sub LayoutFormattingRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "chkResetChar", "cmdRiskPlacementFormattingChar", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkResetPara", "cmdRiskPlacementFormattingPara", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "optEmphQuick", "cmdRiskPlacementFormattingQuick", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "optEmphThorough", "cmdRiskPlacementFormattingThorough", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "optEmphStrip", "cmdRiskPlacementFormattingStrip", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkPreserveHighlight", "cmdRiskPlacementFormattingHighlight", "Safe cleanup"
    PlaceRiskChipBesideChoice toolForm, "chkPreserveDropCaps", "cmdRiskPlacementFormattingDropCaps", "Safe cleanup"
End Sub
Public Sub LayoutFontRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "chkFontFace", "cmdRiskPlacementFontFace", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkFontSize", "cmdRiskPlacementFontSize", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkBold", "cmdRiskPlacementFontBold", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkItalic", "cmdRiskPlacementFontItalic", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkFontColor", "cmdRiskPlacementFontColor", "Formatting change"
End Sub
Public Sub LayoutStyleRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "chkRemoveUnused", "cmdRiskPlacementStyleRemoveUnused", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkRemapVariants", "cmdRiskPlacementStyleRemap", "Inspect first"
End Sub
Public Sub LayoutTableRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "chkRemoveEmptyRows", "cmdRiskPlacementTableRows", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkRemoveEmptyCols", "cmdRiskPlacementTableCols", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkNormalizePadding", "cmdRiskPlacementTablePadding", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkStripDirectFormat", "cmdRiskPlacementTableStripFormat", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkNormalizeBorders", "cmdRiskPlacementTableBorders", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkRemoveBorders", "cmdRiskPlacementTableRemoveBorders", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkConvertToText", "cmdRiskPlacementTableConvertToText", "Structure change"
End Sub
Public Sub LayoutBreakRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "chkCollapseSectionBreaks", "cmdRiskPlacementBreakSections", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkCollapsePageBreaks", "cmdRiskPlacementBreakPages", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkConvertSectionBreaks", "cmdRiskPlacementBreakConvert", "Inspect first"
    PlaceRiskChipBesideChoice toolForm, "optConvertNextPage", "cmdRiskPlacementBreakNextPage", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "optConvertContinuous", "cmdRiskPlacementBreakContinuous", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "optConvertEvenPage", "cmdRiskPlacementBreakEvenPage", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "optConvertOddPage", "cmdRiskPlacementBreakOddPage", "Structure change"
End Sub
Public Sub LayoutHeaderFooterRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "optStandardize", "cmdRiskPlacementHeaderFooterStandardize", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "optClearAll", "cmdRiskPlacementHeaderFooterClear", "Removes content"
    PlaceRiskChipBesideChoice toolForm, "chkHeaders", "cmdRiskPlacementHeaderFooterHeaders", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkFooters", "cmdRiskPlacementHeaderFooterFooters", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkFont", "cmdRiskPlacementHeaderFooterFont", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkSpacing", "cmdRiskPlacementHeaderFooterSpacing", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "chkBreakLinks", "cmdRiskPlacementHeaderFooterBreakLinks", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkAlignment", "cmdRiskPlacementHeaderFooterAlignment", "Formatting change"
End Sub
Public Sub LayoutFootnoteRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "chkFootnotes", "cmdRiskPlacementFootnoteFootnotes", "Removes content"
    PlaceRiskChipBesideChoice toolForm, "chkEndnotes", "cmdRiskPlacementFootnoteEndnotes", "Removes content"
    PlaceRiskChipBesideChoice toolForm, "chkKeepTextInline", "cmdRiskPlacementFootnoteKeepText", "Text caution"
End Sub
Public Sub LayoutObjectRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "chkPictures", "cmdRiskPlacementObjectPictures", "Removes content"
    PlaceRiskChipBesideChoice toolForm, "chkTextBoxes", "cmdRiskPlacementObjectTextBoxes", "Removes content"
    PlaceRiskChipBesideChoice toolForm, "chkFrames", "cmdRiskPlacementObjectFrames", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "chkHorizontalLines", "cmdRiskPlacementObjectLines", "Removes content"
    PlaceRiskChipBesideChoice toolForm, "chkHtmlControls", "cmdRiskPlacementObjectControls", "Removes content"
    PlaceRiskChipBesideChoice toolForm, "chkHiddenText", "cmdRiskPlacementObjectHiddenText", "Removes content"
    PlaceRiskChipBesideChoice toolForm, "chkTables", "cmdRiskPlacementObjectTables", "Removes content"
End Sub
Public Sub LayoutDocumentTrimRiskSummary(ByVal toolForm As Object)
    On Error Resume Next
    Dim insertTop As Single
    insertTop = toolForm.Controls("lblGuidedInfoBorder").Top
    If insertTop <= 0 Then insertTop = toolForm.Controls("cmdPreview").Top
    If insertTop <= 0 Then Exit Sub

    Const SUMMARY_H As Single = 20
    ShiftControlsBelow toolForm, insertTop, SUMMARY_H + 4

    Dim summaryCtl As Object
    Set summaryCtl = EnsureRiskPlacementLabel(toolForm, "lblRiskPlacementDocumentTrimSummary")
    With summaryCtl
        .Caption = "Risk label: Safe cleanup. Only trailing empty paragraphs are highlighted or removed."
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .ForeColor = ToolRiskForeColor("Safe cleanup")
        .BackColor = ToolRiskBackColor("Safe cleanup")
        .BackStyle = 1
        .BorderStyle = 1
        .SpecialEffect = 0
        .WordWrap = True
        .Move 14, insertTop, toolForm.Width - 36, SUMMARY_H
        .Visible = True
        .ZOrder 0
    End With
End Sub
Public Sub LayoutSoftReturnRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "optSoftToPara", "cmdRiskPlacementSoftReturnToPara", "Structure change"
    PlaceRiskChipBesideChoice toolForm, "optParaToSoft", "cmdRiskPlacementSoftReturnToSoft", "Structure change"
End Sub
Public Sub LayoutFinalReviewRiskSummary(ByVal toolForm As Object)
    On Error Resume Next
    Dim insertTop As Single
    insertTop = toolForm.Controls("lblGuidedInfoBorder").Top
    If insertTop <= 0 Then insertTop = toolForm.Controls("cmdPreview").Top
    If insertTop <= 0 Then insertTop = toolForm.Controls("cmdRun").Top
    If insertTop <= 0 Then Exit Sub

    Const SUMMARY_H As Single = 34
    ShiftControlsBelow toolForm, insertTop, SUMMARY_H + 4

    Dim summaryText As String
    If toolForm.Controls("chkComments").Value And toolForm.Controls("chkRevisions").Value Then
        summaryText = "Current selection: Removes content + Final action. Comments will be removed and tracked changes accepted."
    ElseIf toolForm.Controls("chkComments").Value Then
        summaryText = "Current selection: Removes content. Comments will be deleted from the document."
    ElseIf toolForm.Controls("chkRevisions").Value Then
        summaryText = "Current selection: Final action. Tracked changes will be accepted and markup review will be finalized."
    Else
        summaryText = "Current selection: Nothing selected yet. Choose a review action to see its risk."
    End If

    Dim summaryCtl As Object
    Set summaryCtl = EnsureRiskPlacementLabel(toolForm, "lblRiskPlacementFinalReviewSummary")
    With summaryCtl
        .Caption = summaryText
        .Font.Name = "Segoe UI"
        .Font.Size = 8
        .Font.Bold = True
        .ForeColor = ToolRiskForeColor(IIf(toolForm.Controls("chkRevisions").Value, "Final action", "Removes content"))
        .BackColor = ToolRiskBackColor(IIf(toolForm.Controls("chkRevisions").Value, "Final action", "Removes content"))
        .BackStyle = 1
        .BorderStyle = 1
        .SpecialEffect = 0
        .WordWrap = True
        .Move 14, insertTop, toolForm.Width - 36, SUMMARY_H
        .Visible = True
        .ZOrder 0
    End With
End Sub
Public Sub LayoutHyperlinkRiskChoiceChips(ByVal toolForm As Object)
    PlaceRiskChipBesideChoice toolForm, "optHideLinks", "cmdRiskPlacementHyperlinkHide", "Formatting change"
    PlaceRiskChipBesideChoice toolForm, "optRemoveLinks", "cmdRiskPlacementHyperlinkRemove", "Removes content"
End Sub
Private Sub HideGuidedTitleChrome(ByVal toolForm As Object)
    On Error Resume Next
    Dim ctl As Object
    For Each ctl In toolForm.Controls
        If Left$(LCase$(ctl.Name), 3) <> "lbl" Then GoTo NextGuidedTitleControl
        If ctl.Name = "lblGuidedIntro" Then GoTo NextGuidedTitleControl
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
    If Left$(LCase$(ctl.Name), 3) <> "lbl" Then Exit Function
    Dim captionText As String
    captionText = Trim$(CStr(ctl.Caption))
    If Len(captionText) = 0 Then Exit Function
    GuidedControlMatchesToolTitle = _
        (captionText = GuidedToolName(toolForm.Name) Or captionText = Trim$(toolForm.Caption))
End Function
Private Function ShouldHideGuidedControl(ByVal toolForm As Object, ByVal controlName As String) As Boolean
    On Error Resume Next
    ShouldHideGuidedControl = False

    Select Case controlName
        Case "cmdSelectAll", "cmdDeselectAll"
            If Not GuidedCustomModeIsActive(toolForm) Then
                ShouldHideGuidedControl = True
            End If
            Exit Function
    End Select

    If Not GuidedCustomModeIsActive(toolForm) Then
        If IsGuidedCustomChoiceControl(toolForm.Name, controlName) Then
            ShouldHideGuidedControl = True
            Exit Function
        End If
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

    If toolForm.Name = "frmCapitalizationCleanup" Then
        Select Case controlName
            Case "chkSentence", "chkTitle", "chkUpper", "chkLower", "chkSmartSentences", "cmdEditExceptions"
                ShouldHideGuidedControl = Not GuidedCustomModeIsActive(toolForm)
                Exit Function
            Case "chkHeadingParentheses"
                ShouldHideGuidedControl = Not GuidedCustomModeIsActive(toolForm) Or Not toolForm.Controls("chkLower").Value
                Exit Function
        End Select
    End If
End Function
Private Function GuidedCustomModeIsActive(ByVal toolForm As Object) As Boolean
    On Error Resume Next
    GuidedCustomModeIsActive = toolForm.Controls("optCustom").Value
End Function
Private Function IsCapitalizationAdvancedControl(ByVal controlName As String) As Boolean
    IsCapitalizationAdvancedControl = (controlName = "chkSentence" Or controlName = "chkTitle" Or controlName = "chkUpper" Or controlName = "chkLower" Or controlName = "chkHeadingParentheses" Or controlName = "chkSmartSentences" Or controlName = "cmdEditExceptions")
End Function

Private Sub LayoutCapitalizationChoices(ByVal toolForm As Object, ByVal M As Single, ByRef y As Single, ByVal contentW As Single, ByVal BH As Single, ByVal GAP As Single)
    On Error Resume Next
    PositionGuidedChoicePairByName toolForm, "optAll", "optCustom", M, y, contentW, GAP

    If Not GuidedCustomModeIsActive(toolForm) Then
        HideGuidedControl toolForm.Controls("chkSentence")
        HideGuidedControl toolForm.Controls("chkTitle")
        HideGuidedControl toolForm.Controls("chkUpper")
        HideGuidedControl toolForm.Controls("chkLower")
        HideGuidedControl toolForm.Controls("chkSmartSentences")
        HideGuidedControl toolForm.Controls("chkHeadingParentheses")
        HideGuidedControl toolForm.Controls("cmdSelectAll")
        HideGuidedControl toolForm.Controls("cmdDeselectAll")
        HideGuidedControl toolForm.Controls("cmdEditExceptions")
        Exit Sub
    End If

    LayoutGuidedDivider toolForm, M, y, contentW
    PositionCapitalizationGridRow toolForm, "chkSentence", "chkTitle", M, y, contentW, GAP
    PositionCapitalizationGridRow toolForm, "chkUpper", "chkLower", M, y, contentW, GAP
    PositionCapitalizationGridRow toolForm, "chkSmartSentences", "chkHeadingParentheses", M, y, contentW, GAP

    StyleGuidedButton toolForm.Controls("cmdSelectAll"), False
    StyleGuidedButton toolForm.Controls("cmdDeselectAll"), False
    toolForm.Controls("cmdSelectAll").Move M, y, (contentW - GAP) / 2, BH
    toolForm.Controls("cmdDeselectAll").Move M + ((contentW - GAP) / 2) + GAP, y, (contentW - GAP) / 2, BH
    toolForm.Controls("cmdSelectAll").Visible = True
    toolForm.Controls("cmdDeselectAll").Visible = True
    y = y + BH + GAP

    StyleGuidedButton toolForm.Controls("cmdEditExceptions"), False
    toolForm.Controls("cmdEditExceptions").Move M, y, contentW, BH
    toolForm.Controls("cmdEditExceptions").Visible = True
    y = y + BH + GAP + 1
End Sub

Private Sub PositionCapitalizationGridRow(ByVal toolForm As Object, ByVal leftName As String, ByVal rightName As String, ByVal M As Single, ByRef y As Single, ByVal contentW As Single, ByVal GAP As Single)
    On Error Resume Next
    Dim leftCtl As Object
    Dim rightCtl As Object
    Dim leftVisible As Boolean
    Dim rightVisible As Boolean
    Dim choiceW As Single
    Dim rowH As Single

    Set leftCtl = toolForm.Controls(leftName)
    Set rightCtl = toolForm.Controls(rightName)
    leftVisible = Not ShouldHideGuidedControl(toolForm, leftName)
    rightVisible = Not ShouldHideGuidedControl(toolForm, rightName)
    If Not leftVisible And Not rightVisible Then Exit Sub

    choiceW = (contentW - GAP) / 2
    If leftVisible Then rowH = GuidedOptionHeight(leftCtl)
    If rightVisible Then
        If GuidedOptionHeight(rightCtl) > rowH Then rowH = GuidedOptionHeight(rightCtl)
    End If

    If leftVisible Then
        StyleGuidedOption leftCtl
        leftCtl.Move M + 8, y, choiceW - 8, rowH
        leftCtl.Visible = True
    Else
        HideGuidedControl leftCtl
    End If

    If rightVisible Then
        StyleGuidedOption rightCtl
        rightCtl.Move M + choiceW + GAP + 8, y, choiceW - 8, rowH
        rightCtl.Visible = True
        rightCtl.Enabled = True
    Else
        HideGuidedControl rightCtl
    End If

    y = y + rowH + GAP
End Sub
Private Sub LayoutHeaderFooterChoices(ByVal toolForm As Object, ByVal M As Single, ByRef y As Single, ByVal contentW As Single, ByVal GAP As Single)
    On Error Resume Next
    HideGuidedControl toolForm.Controls("fraOptions")
    HideGuidedControl toolForm.Controls("fraAlign")

    PositionGuidedChoicePairByName toolForm, "optStandardize", "optClearAll", M, y, contentW, GAP
    PositionGuidedChoicePairByName toolForm, "chkHeaders", "chkFooters", M, y, contentW, GAP

    If toolForm.Controls("optStandardize").Value Then
        PositionGuidedChoicePairByName toolForm, "chkFont", "chkSpacing", M, y, contentW, GAP
        PositionGuidedChoicePairByName toolForm, "chkBreakLinks", "chkAlignment", M, y, contentW, GAP

        If toolForm.Controls("chkAlignment").Value Then
            PositionGuidedChoicePairByName toolForm, "optAlignLeft", "optAlignRight", M, y, contentW, GAP
            PositionGuidedCenteredChoice toolForm, "optAlignCenter", M, y, contentW, GAP
        Else
            HideGuidedControl toolForm.Controls("optAlignLeft")
            HideGuidedControl toolForm.Controls("optAlignRight")
            HideGuidedControl toolForm.Controls("optAlignCenter")
        End If
    Else
        HideGuidedControl toolForm.Controls("chkFont")
        HideGuidedControl toolForm.Controls("chkSpacing")
        HideGuidedControl toolForm.Controls("chkBreakLinks")
        HideGuidedControl toolForm.Controls("chkAlignment")
        HideGuidedControl toolForm.Controls("optAlignLeft")
        HideGuidedControl toolForm.Controls("optAlignRight")
        HideGuidedControl toolForm.Controls("optAlignCenter")
    End If
End Sub
Private Function IsGuidedCustomChoiceControl(ByVal formName As String, ByVal controlName As String) As Boolean
    Select Case formName
        Case "frmPunctuationCleanup"
            IsGuidedCustomChoiceControl = (controlName = "chkCurlyDouble" Or controlName = "chkCurlySingle" Or controlName = "chkEmDash" Or controlName = "chkEnDash" Or controlName = "chkEllipses")
        Case "frmUnicodeCleanup"
            IsGuidedCustomChoiceControl = (controlName = "chkNBSP" Or controlName = "chkZWSP" Or controlName = "chkZWNJ" Or controlName = "chkZWJ" Or controlName = "chkBOM" Or controlName = "chkSoftHyphen" Or controlName = "chkNBHyphen")
        Case "frmSpacingCleanup"
            IsGuidedCustomChoiceControl = (controlName = "chkDoubleSpaces" Or controlName = "chkTrimSpaces" Or controlName = "chkSpaceBeforePunct" Or controlName = "chkNormalizeAfterPunct" Or controlName = "chkExtraBlankLines")
        Case "frmListCleanup"
            IsGuidedCustomChoiceControl = (controlName = "chkNormalizeBullets" Or controlName = "chkNormalizeNumbering" Or controlName = "chkFixIndent" Or controlName = "chkHyphenToBullets")
        Case "frmParagraphCleanup"
            IsGuidedCustomChoiceControl = (controlName = "chkRemoveEmpty" Or controlName = "chkCollapseBreaks" Or controlName = "chkNormalizeParaSpacing" Or controlName = "chkFixIndent")
        Case Else
            IsGuidedCustomChoiceControl = False
    End Select
End Function
Private Sub LayoutGuidedDivider(ByVal toolForm As Object, ByVal M As Single, ByRef y As Single, ByVal contentW As Single)
    On Error Resume Next
    Dim divider As Object
    Set divider = EnsureGuidedLabel(toolForm, "lblGuidedDivider")
    With divider
        .Caption = ""
        .BackColor = RGB(210, 215, 221)
        .BorderStyle = 0
        .SpecialEffect = 0
        .Move M, y + 2, contentW, 1
        .Visible = True
        .ZOrder 0
    End With
    y = y + 5
End Sub
Private Sub PositionGuidedChoicePairByName(ByVal toolForm As Object, ByVal firstName As String, ByVal secondName As String, ByVal M As Single, ByRef y As Single, ByVal contentW As Single, ByVal GAP As Single)
    On Error Resume Next
    Dim firstCtl As Object
    Dim secondCtl As Object
    Set firstCtl = toolForm.Controls(firstName)
    Set secondCtl = toolForm.Controls(secondName)
    Dim pendingChoice As String
    pendingChoice = firstName
    PositionGuidedChoicePair toolForm, pendingChoice, secondCtl, M, y, contentW, GAP
End Sub
Private Sub PositionGuidedCenteredChoice(ByVal toolForm As Object, ByVal controlName As String, ByVal M As Single, ByRef y As Single, ByVal contentW As Single, ByVal GAP As Single)
    On Error Resume Next
    Dim pendingChoice As String
    pendingChoice = controlName
    FlushGuidedPendingChoice toolForm, pendingChoice, M, y, contentW, GAP
End Sub
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
    GuidedSingleChoiceWidth = pairW + 12
    If GuidedOptionHeight(ctl) > 18 Then GuidedSingleChoiceWidth = pairW + 28
    If Len(Trim$(ctl.Caption)) > 42 Then
        GuidedSingleChoiceWidth = contentW - 24
    ElseIf Len(Trim$(ctl.Caption)) > 30 Then
        GuidedSingleChoiceWidth = pairW + 72
    End If
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
        GuidedLabelHeight = 36
    Else
        GuidedLabelHeight = 24
    End If
End Function
Private Function GuidedIntroHeight(ByVal captionText As String) As Single
    If Len(captionText) > 72 Then
        GuidedIntroHeight = 26
    Else
        GuidedIntroHeight = 16
    End If
End Function
Private Sub LayoutGuidedInfoBox(ByVal toolForm As Object, ByVal M As Single, ByRef y As Single, ByVal contentW As Single)
    On Error Resume Next
    Const GAP As Single = 4
    Const BOX_PAD As Single = 10
    Const INNER_GAP As Single = 4
    Dim lines As Collection
    Set lines = GuidedInfoLines(toolForm)
    If lines.Count = 0 Then Exit Sub

    Dim boxH As Single
    boxH = 8
    Dim i As Long
    Dim rawLine As String
    Dim boldPos As Long
    Dim nameText As String
    Dim detailText As String
    Dim lineBold As Boolean
    Dim lineH As Single
    Dim nameW As Single
    Dim nameColW As Single
    Dim detailW As Single
    Dim measureCtl As Object
    nameColW = 72
    For i = 1 To lines.Count
        rawLine = CStr(lines(i))
        nameText = GuidedInfoRecordName(rawLine)
        nameW = GuidedInfoNameWidth(nameText)
        If nameW > nameColW Then nameColW = nameW
    Next i
    If nameColW > 92 Then nameColW = 92
    detailW = contentW - ((BOX_PAD * 2) + nameColW + INNER_GAP)
    Set measureCtl = EnsureGuidedLabel(toolForm, "lblGuidedInfoMeasure")
    For i = 1 To lines.Count
        rawLine = CStr(lines(i))
        nameText = GuidedInfoRecordName(rawLine)
        detailText = GuidedInfoRecordDetail(rawLine)
        boldPos = InStrRev(rawLine, Chr$(30), -1, vbBinaryCompare)
        If boldPos > 0 Then
            lineBold = (Mid$(rawLine, boldPos + 1) = "1")
        Else
            lineBold = False
        End If
        lineH = GuidedInfoLineHeight(toolForm, detailText, detailW, lineBold)
        boxH = boxH + lineH
    Next i
    boxH = boxH + 4
    If boxH < 34 Then boxH = 34

    Dim border As Object
    Set border = EnsureGuidedLabel(toolForm, "lblGuidedInfoBorder")
    With border
        .Caption = ""
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = 1
        .SpecialEffect = 0
        .Move M, y, contentW, boxH
        .Visible = True
    End With

    Dim lineCtl As Object
    Dim nameCtl As Object
    Dim detailCtl As Object
    Dim lineTop As Single
    lineTop = y + 6
    For i = 1 To lines.Count
        rawLine = CStr(lines(i))
        nameText = GuidedInfoRecordName(rawLine)
        detailText = GuidedInfoRecordDetail(rawLine)
        boldPos = InStrRev(rawLine, Chr$(30), -1, vbBinaryCompare)
        If boldPos > 0 Then
            lineBold = (Mid$(rawLine, boldPos + 1) = "1")
        Else
            lineBold = False
        End If
        lineH = GuidedInfoLineHeight(toolForm, detailText, detailW, lineBold)

        Set lineCtl = EnsureGuidedLabel(toolForm, "lblGuidedInfoLine" & CStr(i))
        lineCtl.Visible = False
        lineCtl.Move 0, 0, 0, 0

        Set nameCtl = EnsureGuidedLabel(toolForm, "lblGuidedInfoName" & CStr(i))
        With nameCtl
            .Caption = nameText & ":"
            .Font.Name = "Segoe UI"
            .Font.Size = 7.5
            .Font.Bold = lineBold
            .ForeColor = RGB(67, 80, 96)
            .BackColor = RGB(255, 255, 255)
            .BackStyle = 1
            .TextAlign = 1
            .WordWrap = True
            .BorderStyle = 0
            .SpecialEffect = 0
            .Move M + BOX_PAD, lineTop, nameColW, lineH
            .Visible = True
            .ZOrder 0
        End With

        Set detailCtl = EnsureGuidedLabel(toolForm, "lblGuidedInfoDetail" & CStr(i))
        With detailCtl
            .Caption = detailText
            .Font.Name = "Segoe UI"
            .Font.Size = 7.5
            .Font.Bold = lineBold
            .ForeColor = RGB(67, 80, 96)
            .BackColor = RGB(255, 255, 255)
            .BackStyle = 1
            .TextAlign = 1
            .WordWrap = True
            .BorderStyle = 0
            .SpecialEffect = 0
            .Move M + BOX_PAD + nameColW + INNER_GAP, lineTop, detailW, lineH
            .Visible = True
            .ZOrder 0
        End With
        lineTop = lineTop + lineH
    Next i

    For i = lines.Count + 1 To 20
        HideGuidedControl toolForm.Controls("lblGuidedInfoLine" & CStr(i))
        HideGuidedControl toolForm.Controls("lblGuidedInfoName" & CStr(i))
        HideGuidedControl toolForm.Controls("lblGuidedInfoDetail" & CStr(i))
    Next i
    y = y + boxH + GAP
End Sub
Private Function GuidedInfoNameWidth(ByVal nameText As String) As Single
    GuidedInfoNameWidth = 18 + (Len(nameText) * 3.6)
    If GuidedInfoNameWidth < 56 Then GuidedInfoNameWidth = 56
    If GuidedInfoNameWidth > 110 Then GuidedInfoNameWidth = 110
End Function
Private Function GuidedInfoLineHeight(ByVal toolForm As Object, ByVal detailText As String, ByVal detailWidth As Single, Optional ByVal isBold As Boolean = False) As Single
    Dim measureCtl As Object
    Set measureCtl = EnsureGuidedLabel(toolForm, "lblGuidedInfoMeasure")
    With measureCtl
        .Caption = detailText
        .Font.Name = "Segoe UI"
        .Font.Size = 7.5
        .Font.Bold = isBold
        .ForeColor = RGB(67, 80, 96)
        .BackColor = RGB(255, 255, 255)
        .BackStyle = 1
        .TextAlign = 1
        .WordWrap = True
        .BorderStyle = 0
        .SpecialEffect = 0
        .Visible = False
        .Move -1000, -1000, detailWidth, 11
        .AutoSize = False
        .AutoSize = True
        GuidedInfoLineHeight = measureCtl.Height - 2
        .AutoSize = False
        .Move -1000, -1000, detailWidth, GuidedInfoLineHeight
    End With
    If GuidedInfoLineHeight < 11 Then GuidedInfoLineHeight = 11
End Function
Private Function GuidedInfoLines(ByVal toolForm As Object) As Collection
    On Error Resume Next
    Set GuidedInfoLines = New Collection

    If toolForm.Name = "frmHeaderFooterStandardizer" Then
        GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "optStandardize")
        GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "optClearAll")
        GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "chkHeaders")
        GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "chkFooters")
        GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "chkFont")
        GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "chkSpacing")
        GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "chkBreakLinks")
        GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "chkAlignment")
        GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "optAlignLeft")
        GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "optAlignCenter")
        GuidedInfoLines.Add GuidedInfoLineForControl(toolForm, "optAlignRight")
        Exit Function
    End If

    Dim customActive As Boolean
    customActive = GuidedCustomModeIsActive(toolForm)
    Dim ctl As Object
    For Each ctl In toolForm.Controls
        If Not GuidedChoiceShouldAppearInInfo(toolForm, ctl, customActive) Then GoTo NextInfoControl
        GuidedInfoLines.Add GuidedInfoRecord(toolForm.Name, ctl.Name, ctl.Caption, GuidedControlIsSelected(ctl))
NextInfoControl:
    Next ctl
End Function
Private Function GuidedInfoLineForControl(ByVal toolForm As Object, ByVal controlName As String) As String
    On Error Resume Next
    GuidedInfoLineForControl = GuidedInfoRecord(toolForm.Name, controlName, toolForm.Controls(controlName).Caption, GuidedControlIsSelected(toolForm.Controls(controlName)))
End Function
Private Function GuidedChoiceShouldAppearInInfo(ByVal toolForm As Object, ByVal ctl As Object, ByVal customActive As Boolean) As Boolean
    On Error Resume Next
    GuidedChoiceShouldAppearInInfo = False
    If Left$(ctl.Name, 3) <> "opt" And Left$(ctl.Name, 3) <> "chk" Then Exit Function
    Select Case ctl.Name
        Case "chkPreviewOnly", "optScopeDocument", "optScopeSelection"
            Exit Function
    End Select
    If ShouldHideGuidedControl(toolForm, ctl.Name) Then Exit Function
    If toolForm.Name = "frmCapitalizationCleanup" Then
        If customActive Then
            GuidedChoiceShouldAppearInInfo = IsCapitalizationAdvancedControl(ctl.Name)
        Else
            GuidedChoiceShouldAppearInInfo = (Left$(ctl.Name, 3) = "opt")
        End If
    ElseIf customActive Then
        GuidedChoiceShouldAppearInInfo = IsGuidedCustomChoiceControl(toolForm.Name, ctl.Name)
    Else
        GuidedChoiceShouldAppearInInfo = Not IsGuidedCustomChoiceControl(toolForm.Name, ctl.Name)
    End If
End Function
Private Function GuidedControlIsSelected(ByVal ctl As Object) As Boolean
    On Error Resume Next
    GuidedControlIsSelected = CBool(ctl.Value)
End Function
Private Function GuidedInfoRecord(ByVal formName As String, ByVal controlName As String, ByVal fallbackText As String, ByVal isSelected As Boolean) As String
    GuidedInfoRecord = GuidedInfoDisplayName(formName, controlName, fallbackText) & Chr$(29) & GuidedInfoAdvisory(formName, controlName, fallbackText) & Chr$(30) & IIf(isSelected, "1", "0")
End Function
Private Function GuidedInfoRecordName(ByVal rawLine As String) As String
    Dim namePos As Long
    namePos = InStr(1, rawLine, Chr$(29), vbBinaryCompare)
    If namePos > 0 Then
        GuidedInfoRecordName = Left$(rawLine, namePos - 1)
    Else
        GuidedInfoRecordName = rawLine
    End If
End Function
Private Function GuidedInfoRecordDetail(ByVal rawLine As String) As String
    Dim namePos As Long
    Dim boldPos As Long
    namePos = InStr(1, rawLine, Chr$(29), vbBinaryCompare)
    boldPos = InStrRev(rawLine, Chr$(30), -1, vbBinaryCompare)
    If namePos = 0 Then Exit Function
    If boldPos > namePos Then
        GuidedInfoRecordDetail = Mid$(rawLine, namePos + 1, (boldPos - namePos) - 1)
    Else
        GuidedInfoRecordDetail = Mid$(rawLine, namePos + 1)
    End If
End Function
Private Function GuidedInfoDisplayName(ByVal formName As String, ByVal controlName As String, ByVal fallbackText As String) As String
    Dim key As String
    key = formName & "." & controlName
    Select Case key
        Case "frmBreakNormalizer.chkCollapseSectionBreaks": GuidedInfoDisplayName = "Collapse section breaks"
        Case "frmBreakNormalizer.chkCollapsePageBreaks": GuidedInfoDisplayName = "Collapse page breaks"
        Case "frmBreakNormalizer.chkConvertSectionBreaks": GuidedInfoDisplayName = "Convert section breaks"
        Case "frmCapitalizationCleanup.optAll": GuidedInfoDisplayName = "Recommended repair"
        Case "frmCapitalizationCleanup.optCustom": GuidedInfoDisplayName = "Custom repair"
        Case "frmCapitalizationCleanup.chkSentence": GuidedInfoDisplayName = "Abbreviation recognition"
        Case "frmCapitalizationCleanup.chkTitle": GuidedInfoDisplayName = "Acronym protection"
        Case "frmCapitalizationCleanup.chkUpper": GuidedInfoDisplayName = "Name and brand protection"
        Case "frmCapitalizationCleanup.chkLower": GuidedInfoDisplayName = "Heading repair"
        Case "frmCapitalizationCleanup.chkHeadingParentheses": GuidedInfoDisplayName = "Parenthetical heading repair"
        Case "frmCapitalizationCleanup.chkSmartSentences": GuidedInfoDisplayName = "Sentence context"
        Case "frmCapitalizationCleanup.cmdEditExceptions": GuidedInfoDisplayName = "Custom exceptions"
        Case "frmDuplicateDetector.optHighlightOnly": GuidedInfoDisplayName = "Preview duplicates"
        Case "frmDuplicateDetector.optRemoveDupes": GuidedInfoDisplayName = "Remove duplicates"
        Case "frmDuplicateDetector.optMatchExact": GuidedInfoDisplayName = "Exact match"
        Case "frmDuplicateDetector.optMatchNormalized": GuidedInfoDisplayName = "Normalized match"
        Case "frmDuplicateDetector.optMatchFuzzy": GuidedInfoDisplayName = "Fuzzy match"
        Case "frmDuplicateDetector.optFuzzyLoose": GuidedInfoDisplayName = "Loose"
        Case "frmDuplicateDetector.optFuzzyMedium": GuidedInfoDisplayName = "Medium"
        Case "frmDuplicateDetector.optFuzzyStrict": GuidedInfoDisplayName = "Strict"
        Case "frmFontNormalizer.chkFontFace": GuidedInfoDisplayName = "Font face"
        Case "frmFontNormalizer.chkFontSize": GuidedInfoDisplayName = "Font size"
        Case "frmFontNormalizer.chkBold": GuidedInfoDisplayName = "Bold overrides"
        Case "frmFontNormalizer.chkItalic": GuidedInfoDisplayName = "Italic overrides"
        Case "frmFontNormalizer.chkFontColor": GuidedInfoDisplayName = "Font color overrides"
        Case "frmFootnoteRemover.chkFootnotes": GuidedInfoDisplayName = "Remove footnotes"
        Case "frmFootnoteRemover.chkEndnotes": GuidedInfoDisplayName = "Remove endnotes"
        Case "frmFootnoteRemover.chkKeepTextInline": GuidedInfoDisplayName = "Keep note text"
        Case "frmFormattingStripper.chkResetChar": GuidedInfoDisplayName = "Strip character formatting"
        Case "frmFormattingStripper.chkResetPara": GuidedInfoDisplayName = "Strip paragraph formatting"
        Case "frmFormattingStripper.optEmphQuick": GuidedInfoDisplayName = "Quick"
        Case "frmFormattingStripper.optEmphThorough": GuidedInfoDisplayName = "Thorough"
        Case "frmFormattingStripper.optEmphStrip": GuidedInfoDisplayName = "Strip"
        Case "frmFormattingStripper.chkPreserveHighlight": GuidedInfoDisplayName = "Preserve highlighting"
        Case "frmFormattingStripper.chkPreserveDropCaps": GuidedInfoDisplayName = "Preserve drop caps"
        Case "frmHeaderFooterStandardizer.optStandardize": GuidedInfoDisplayName = "Standardize"
        Case "frmHeaderFooterStandardizer.optClearAll": GuidedInfoDisplayName = "Clear all"
        Case "frmHeaderFooterStandardizer.chkHeaders": GuidedInfoDisplayName = "Headers"
        Case "frmHeaderFooterStandardizer.chkFooters": GuidedInfoDisplayName = "Footers"
        Case "frmHeaderFooterStandardizer.chkFont": GuidedInfoDisplayName = "Reset font"
        Case "frmHeaderFooterStandardizer.chkSpacing": GuidedInfoDisplayName = "Remove spacing"
        Case "frmHeaderFooterStandardizer.chkBreakLinks": GuidedInfoDisplayName = "Unlink sections"
        Case "frmHeaderFooterStandardizer.chkAlignment": GuidedInfoDisplayName = "Set alignment"
        Case "frmHyperlinkRemover.optHideLinks": GuidedInfoDisplayName = "Hide link styling"
        Case "frmHyperlinkRemover.optRemoveLinks": GuidedInfoDisplayName = "Remove hyperlinks"
        Case "frmListCleanup.optAll": GuidedInfoDisplayName = "All list fixes"
        Case "frmListCleanup.optBullets": GuidedInfoDisplayName = "Bullet conversion"
        Case "frmListCleanup.optNumbering": GuidedInfoDisplayName = "Number conversion"
        Case "frmListCleanup.optIndent": GuidedInfoDisplayName = "Indentation only"
        Case "frmListCleanup.optCustom": GuidedInfoDisplayName = "Custom"
        Case "frmListCleanup.chkNormalizeBullets": GuidedInfoDisplayName = "Bullet conversion"
        Case "frmListCleanup.chkNormalizeNumbering": GuidedInfoDisplayName = "Number conversion"
        Case "frmListCleanup.chkFixIndent": GuidedInfoDisplayName = "Indentation"
        Case "frmListCleanup.chkHyphenToBullets": GuidedInfoDisplayName = "Hyphen lines"
        Case "frmFinalReview.chkComments": GuidedInfoDisplayName = "Comments"
        Case "frmFinalReview.chkRevisions": GuidedInfoDisplayName = "Tracked changes"
        Case "frmObjectRemover.chkPictures": GuidedInfoDisplayName = "Pictures"
        Case "frmObjectRemover.chkTextBoxes": GuidedInfoDisplayName = "Text boxes"
        Case "frmObjectRemover.chkFrames": GuidedInfoDisplayName = "Frames"
        Case "frmObjectRemover.chkHorizontalLines": GuidedInfoDisplayName = "Horizontal lines"
        Case "frmObjectRemover.chkHtmlControls": GuidedInfoDisplayName = "Form controls"
        Case "frmObjectRemover.chkHiddenText": GuidedInfoDisplayName = "Hidden text"
        Case "frmObjectRemover.chkTables": GuidedInfoDisplayName = "Tables"
        Case "frmParagraphCleanup.optAll": GuidedInfoDisplayName = "All paragraph fixes"
        Case "frmParagraphCleanup.optRemoveEmpty": GuidedInfoDisplayName = "Empty paragraphs"
        Case "frmParagraphCleanup.optNormalizeSpacing": GuidedInfoDisplayName = "Paragraph spacing"
        Case "frmParagraphCleanup.optCustom": GuidedInfoDisplayName = "Custom"
        Case "frmParagraphCleanup.chkRemoveEmpty": GuidedInfoDisplayName = "Collapse empties"
        Case "frmParagraphCleanup.chkCollapseBreaks": GuidedInfoDisplayName = "Soft returns to paragraphs"
        Case "frmParagraphCleanup.chkNormalizeParaSpacing": GuidedInfoDisplayName = "Space Before/After"
        Case "frmParagraphCleanup.chkFixIndent": GuidedInfoDisplayName = "Leading manual indents"
        Case "frmPunctuationCleanup.optAll": GuidedInfoDisplayName = "All punctuation fixes"
        Case "frmPunctuationCleanup.optQuotes": GuidedInfoDisplayName = "Quotes"
        Case "frmPunctuationCleanup.optDashes": GuidedInfoDisplayName = "Dashes"
        Case "frmPunctuationCleanup.optEllipses": GuidedInfoDisplayName = "Ellipses"
        Case "frmPunctuationCleanup.optCustom": GuidedInfoDisplayName = "Custom"
        Case "frmPunctuationCleanup.chkCurlyDouble": GuidedInfoDisplayName = "Curly double quotes"
        Case "frmPunctuationCleanup.chkCurlySingle": GuidedInfoDisplayName = "Curly single quotes"
        Case "frmPunctuationCleanup.chkEmDash": GuidedInfoDisplayName = "Em dashes"
        Case "frmPunctuationCleanup.chkEnDash": GuidedInfoDisplayName = "En dashes"
        Case "frmPunctuationCleanup.chkEllipses": GuidedInfoDisplayName = "Ellipsis character"
        Case "frmSoftReturnConverter.optSoftToPara": GuidedInfoDisplayName = "Soft to paragraphs"
        Case "frmSoftReturnConverter.optParaToSoft": GuidedInfoDisplayName = "Paragraphs to soft returns"
        Case "frmSpacingCleanup.optAll": GuidedInfoDisplayName = "All spacing fixes"
        Case "frmSpacingCleanup.optDoubleSpaces": GuidedInfoDisplayName = "Double spaces"
        Case "frmSpacingCleanup.optTrim": GuidedInfoDisplayName = "Trim outer spaces"
        Case "frmSpacingCleanup.optCustom": GuidedInfoDisplayName = "Custom"
        Case "frmSpacingCleanup.chkDoubleSpaces": GuidedInfoDisplayName = "Collapse extra spaces"
        Case "frmSpacingCleanup.chkTrimSpaces": GuidedInfoDisplayName = "Trim outer spaces"
        Case "frmSpacingCleanup.chkSpaceBeforePunct": GuidedInfoDisplayName = "Space before punctuation"
        Case "frmSpacingCleanup.chkNormalizeAfterPunct": GuidedInfoDisplayName = "Spacing after punctuation"
        Case "frmSpacingCleanup.chkExtraBlankLines": GuidedInfoDisplayName = "Extra blank lines"
        Case "frmStyleCleanup.chkRemoveUnused": GuidedInfoDisplayName = "Remove unused styles"
        Case "frmStyleCleanup.chkRemapVariants": GuidedInfoDisplayName = "Remap variants"
        Case "frmTableCleaner.chkRemoveEmptyRows": GuidedInfoDisplayName = "Empty rows"
        Case "frmTableCleaner.chkRemoveEmptyCols": GuidedInfoDisplayName = "Empty columns"
        Case "frmTableCleaner.chkNormalizePadding": GuidedInfoDisplayName = "Cell padding"
        Case "frmTableCleaner.chkStripDirectFormat": GuidedInfoDisplayName = "Cell formatting"
        Case "frmTableCleaner.chkNormalizeBorders": GuidedInfoDisplayName = "Border styles"
        Case "frmTableCleaner.chkRemoveBorders": GuidedInfoDisplayName = "Remove borders"
        Case "frmTableCleaner.chkConvertToText": GuidedInfoDisplayName = "Single-column to text"
        Case "frmUnicodeCleanup.optAll": GuidedInfoDisplayName = "All invisible characters"
        Case "frmUnicodeCleanup.optNBSP": GuidedInfoDisplayName = "Nonbreaking spaces"
        Case "frmUnicodeCleanup.optZeroWidth": GuidedInfoDisplayName = "Zero-width characters"
        Case "frmUnicodeCleanup.optCustom": GuidedInfoDisplayName = "Custom"
        Case "frmUnicodeCleanup.chkNBSP": GuidedInfoDisplayName = "Nonbreaking spaces"
        Case "frmUnicodeCleanup.chkZWSP": GuidedInfoDisplayName = "Zero-width spaces"
        Case "frmUnicodeCleanup.chkZWNJ": GuidedInfoDisplayName = "Zero-width nonjoiners"
        Case "frmUnicodeCleanup.chkZWJ": GuidedInfoDisplayName = "Zero-width joiners"
        Case "frmUnicodeCleanup.chkBOM": GuidedInfoDisplayName = "Byte-order marks"
        Case "frmUnicodeCleanup.chkSoftHyphen": GuidedInfoDisplayName = "Soft hyphens"
        Case "frmUnicodeCleanup.chkNBHyphen": GuidedInfoDisplayName = "Nonbreaking hyphens"
        Case Else
            GuidedInfoDisplayName = CleanGuidedInfoCaption(fallbackText)
    End Select
End Function
Private Function GuidedInfoAdvisory(ByVal formName As String, ByVal controlName As String, ByVal fallbackText As String) As String
    Dim key As String
    key = formName & "." & controlName
    Select Case key
        Case "frmBreakNormalizer.chkCollapseSectionBreaks": GuidedInfoAdvisory = "Shrinks repeated section-break stacks to one clean break."
        Case "frmBreakNormalizer.chkCollapsePageBreaks": GuidedInfoAdvisory = "Shrinks repeated page-break stacks to one clean break."
        Case "frmBreakNormalizer.chkConvertSectionBreaks": GuidedInfoAdvisory = "Replaces section boundaries with plain page breaks, removing section-level formatting splits."
        Case "frmBreakNormalizer.optConvertNextPage": GuidedInfoAdvisory = "Keeps a normal next-page start after conversion."
        Case "frmBreakNormalizer.optConvertContinuous": GuidedInfoAdvisory = "Preserves flow without forcing a new page."
        Case "frmBreakNormalizer.optConvertEvenPage": GuidedInfoAdvisory = "Forces converted content to begin on the next even page."
        Case "frmBreakNormalizer.optConvertOddPage": GuidedInfoAdvisory = "Forces converted content to begin on the next odd page."
        Case "frmCapitalizationCleanup.optAll": GuidedInfoAdvisory = "Repairs sentence starts, obvious all-caps damage, headings, acronyms, names, and brands using every protection."
        Case "frmCapitalizationCleanup.optCustom": GuidedInfoAdvisory = "Lets you turn individual capitalization protections on and off."
        Case "frmCapitalizationCleanup.chkSentence": GuidedInfoAdvisory = "Reduces false sentence starts after common abbreviations such as Dr. or Inc."
        Case "frmCapitalizationCleanup.chkTitle": GuidedInfoAdvisory = "Keeps terms such as API, PDF, or VBA uppercase after repair."
        Case "frmCapitalizationCleanup.chkUpper": GuidedInfoAdvisory = "Restores known mixed-case names and brands such as iPhone or GitHub."
        Case "frmCapitalizationCleanup.chkLower": GuidedInfoAdvisory = "Repairs likely headings in title case while preserving protected terms."
        Case "frmCapitalizationCleanup.chkHeadingParentheses": GuidedInfoAdvisory = "Also title-cases balanced parenthetical wording as its own heading segment; it is off by default."
        Case "frmCapitalizationCleanup.chkSmartSentences": GuidedInfoAdvisory = "Recognizes quotes, bullets, and punctuation while finding sentence starts."
        Case "frmCapitalizationCleanup.cmdEditExceptions": GuidedInfoAdvisory = "Adds document-specific names, brands, and acronyms that should keep their exact capitalization."
        Case "frmDuplicateDetector.optHighlightOnly": GuidedInfoAdvisory = "Marks likely duplicates for review without deleting any text."
        Case "frmDuplicateDetector.optRemoveDupes": GuidedInfoAdvisory = "Deletes later duplicates and keeps the first surviving paragraph in each group."
        Case "frmDuplicateDetector.optMatchExact": GuidedInfoAdvisory = "Fastest check; punctuation and spacing must still match."
        Case "frmDuplicateDetector.optMatchNormalized": GuidedInfoAdvisory = "Ignores punctuation and extra spaces to catch cosmetic repeats."
        Case "frmDuplicateDetector.optMatchFuzzy": GuidedInfoAdvisory = "Catches near-duplicates by word overlap, but runs slower on long documents."
        Case "frmDuplicateDetector.optFuzzyLoose": GuidedInfoAdvisory = "Broadest fuzzy setting; finds more candidates but needs more review."
        Case "frmDuplicateDetector.optFuzzyMedium": GuidedInfoAdvisory = "Balanced fuzzy setting for most cleanup passes."
        Case "frmDuplicateDetector.optFuzzyStrict": GuidedInfoAdvisory = "Most conservative fuzzy setting; fewer false matches, but easier to miss loose repeats."
        Case "frmFontNormalizer.chkFontFace": GuidedInfoAdvisory = "Returns direct font-name overrides to the document default."
        Case "frmFontNormalizer.chkFontSize": GuidedInfoAdvisory = "Returns direct size overrides to the active style's size."
        Case "frmFontNormalizer.chkBold": GuidedInfoAdvisory = "Removes manual bolding that does not come from the style."
        Case "frmFontNormalizer.chkItalic": GuidedInfoAdvisory = "Removes manual italics that do not come from the style."
        Case "frmFontNormalizer.chkFontColor": GuidedInfoAdvisory = "Clears manual font colors so styled colors can lead again."
        Case "frmFootnoteRemover.chkFootnotes": GuidedInfoAdvisory = "Deletes footnotes in scope, including their reference marks."
        Case "frmFootnoteRemover.chkEndnotes": GuidedInfoAdvisory = "Deletes endnotes in scope, including their reference marks."
        Case "frmFootnoteRemover.chkKeepTextInline": GuidedInfoAdvisory = "Keeps each note's wording in brackets before removing the note itself."
        Case "frmFormattingStripper.chkResetChar": GuidedInfoAdvisory = "Removes direct font-level overrides while leaving styles in charge."
        Case "frmFormattingStripper.chkResetPara": GuidedInfoAdvisory = "Removes direct paragraph formatting such as indents, spacing, and alignment."
        Case "frmFormattingStripper.optEmphQuick": GuidedInfoAdvisory = "Fastest pass; preserves emphasis that applies to whole paragraphs."
        Case "frmFormattingStripper.optEmphThorough": GuidedInfoAdvisory = "Checks runs one by one; slower, but safer for mixed emphasis."
        Case "frmFormattingStripper.optEmphStrip": GuidedInfoAdvisory = "Removes direct emphasis too, including manual bold and italic."
        Case "frmFormattingStripper.chkPreserveHighlight": GuidedInfoAdvisory = "Leaves highlight colors in place while other direct formatting is stripped."
        Case "frmFormattingStripper.chkPreserveDropCaps": GuidedInfoAdvisory = "Keeps decorative drop caps instead of flattening them."
        Case "frmHeaderFooterStandardizer.optStandardize": GuidedInfoAdvisory = "Keeps the content but cleans formatting across the chosen header and footer areas."
        Case "frmHeaderFooterStandardizer.optClearAll": GuidedInfoAdvisory = "Deletes the chosen header and footer content outright."
        Case "frmHeaderFooterStandardizer.chkHeaders": GuidedInfoAdvisory = "Applies the chosen action to headers."
        Case "frmHeaderFooterStandardizer.chkFooters": GuidedInfoAdvisory = "Applies the chosen action to footers."
        Case "frmHeaderFooterStandardizer.chkFont": GuidedInfoAdvisory = "Returns header and footer text to the document's default font."
        Case "frmHeaderFooterStandardizer.chkSpacing": GuidedInfoAdvisory = "Clears paragraph spacing so header/footer lines sit tightly again."
        Case "frmHeaderFooterStandardizer.chkBreakLinks": GuidedInfoAdvisory = "Lets each section stand on its own instead of inheriting from the previous one."
        Case "frmHeaderFooterStandardizer.chkAlignment": GuidedInfoAdvisory = "Turns on the left, center, or right alignment choices below."
        Case "frmHeaderFooterStandardizer.optAlignLeft": GuidedInfoAdvisory = "Sets the chosen header or footer paragraphs to left alignment."
        Case "frmHeaderFooterStandardizer.optAlignCenter": GuidedInfoAdvisory = "Sets the chosen header or footer paragraphs to centered alignment."
        Case "frmHeaderFooterStandardizer.optAlignRight": GuidedInfoAdvisory = "Sets the chosen header or footer paragraphs to right alignment."
        Case "frmHyperlinkRemover.optHideLinks": GuidedInfoAdvisory = "Keeps hyperlink targets available while making the visible text look normal."
        Case "frmHyperlinkRemover.optRemoveLinks": GuidedInfoAdvisory = "Removes hyperlink targets from the document while keeping the visible words."
        Case "frmListCleanup.optAll": GuidedInfoAdvisory = "Runs bullet conversion, number conversion, and indentation cleanup together."
        Case "frmListCleanup.optBullets": GuidedInfoAdvisory = "Converts manual bullet-like lines without changing numbered lists."
        Case "frmListCleanup.optNumbering": GuidedInfoAdvisory = "Converts typed number patterns into real Word numbering."
        Case "frmListCleanup.optIndent": GuidedInfoAdvisory = "Leaves list content alone and only corrects indent depth."
        Case "frmListCleanup.optCustom": GuidedInfoAdvisory = "Shows the individual list repairs so you can mix only the ones you need."
        Case "frmListCleanup.chkNormalizeBullets": GuidedInfoAdvisory = "Turns manual bullets into the document's standard Word bullet form."
        Case "frmListCleanup.chkNormalizeNumbering": GuidedInfoAdvisory = "Turns typed numbering into Word-managed numbered lists."
        Case "frmListCleanup.chkFixIndent": GuidedInfoAdvisory = "Straightens list left indents without changing bullet or number text."
        Case "frmListCleanup.chkHyphenToBullets": GuidedInfoAdvisory = "Promotes plain hyphen-led lines into proper bullet items."
        Case "frmFinalReview.chkComments": GuidedInfoAdvisory = "Deletes reviewer comments from the document."
        Case "frmFinalReview.chkRevisions": GuidedInfoAdvisory = "Accepts tracked changes first, so markup cannot be reviewed afterward."
        Case "frmObjectRemover.chkTables": GuidedInfoAdvisory = "Deletes whole tables and everything inside them, not just table formatting."
        Case "frmParagraphCleanup.optAll": GuidedInfoAdvisory = "Runs the tool's full paragraph cleanup pass."
        Case "frmParagraphCleanup.optCustom": GuidedInfoAdvisory = "Shows the individual paragraph repairs so you can combine them."
        Case "frmPunctuationCleanup.optAll": GuidedInfoAdvisory = "Runs quote, dash, and ellipsis cleanup in one pass."
        Case "frmPunctuationCleanup.optQuotes": GuidedInfoAdvisory = "Targets curly quotes and apostrophes without touching dashes or ellipses."
        Case "frmPunctuationCleanup.optDashes": GuidedInfoAdvisory = "Targets em and en dashes without changing quotes or ellipses."
        Case "frmPunctuationCleanup.optEllipses": GuidedInfoAdvisory = "Targets only the ellipsis character and leaves other punctuation alone."
        Case "frmPunctuationCleanup.optCustom": GuidedInfoAdvisory = "Shows the individual punctuation conversions so you can mix them."
        Case "frmPunctuationCleanup.chkCurlyDouble": GuidedInfoAdvisory = "Replaces curly double quotes with plain straight double quotes."
        Case "frmPunctuationCleanup.chkCurlySingle": GuidedInfoAdvisory = "Replaces curly single quotes and apostrophes with straight apostrophes."
        Case "frmPunctuationCleanup.chkEmDash": GuidedInfoAdvisory = "Replaces em dashes with a plain keyboard hyphen."
        Case "frmPunctuationCleanup.chkEnDash": GuidedInfoAdvisory = "Replaces en dashes with a plain keyboard hyphen."
        Case "frmPunctuationCleanup.chkEllipses": GuidedInfoAdvisory = "Replaces the single ellipsis character with three plain periods."
        Case "frmSoftReturnConverter.optSoftToPara": GuidedInfoAdvisory = "Turns manual line breaks into real paragraphs for cleaner editing and reflow."
        Case "frmSoftReturnConverter.optParaToSoft": GuidedInfoAdvisory = "Keeps visible line breaks without starting new paragraphs."
        Case "frmSpacingCleanup.optAll": GuidedInfoAdvisory = "Runs every spacing repair this tool offers in one pass."
        Case "frmSpacingCleanup.optDoubleSpaces": GuidedInfoAdvisory = "Targets repeated spaces between words and leaves edge spacing alone."
        Case "frmSpacingCleanup.optTrim": GuidedInfoAdvisory = "Targets leading and trailing spaces without changing doubled inner spaces."
        Case "frmSpacingCleanup.optCustom": GuidedInfoAdvisory = "Shows the individual spacing repairs so you can combine them."
        Case "frmSpacingCleanup.chkDoubleSpaces": GuidedInfoAdvisory = "Collapses repeated spaces so only one clean space remains."
        Case "frmSpacingCleanup.chkTrimSpaces": GuidedInfoAdvisory = "Removes stray spaces at the start or end of paragraphs."
        Case "frmSpacingCleanup.chkSpaceBeforePunct": GuidedInfoAdvisory = "Deletes stray spaces that appear before punctuation marks."
        Case "frmSpacingCleanup.chkNormalizeAfterPunct": GuidedInfoAdvisory = "Resets uneven post-punctuation spacing to a single clean space."
        Case "frmSpacingCleanup.chkExtraBlankLines": GuidedInfoAdvisory = "Collapses repeated blank paragraphs between blocks of text."
        Case "frmUnicodeCleanup.optAll": GuidedInfoAdvisory = "Sweeps the common invisible troublemakers that interfere with search, wrap, and cleanup."
        Case "frmUnicodeCleanup.optNBSP": GuidedInfoAdvisory = "Targets only nonbreaking spaces and leaves zero-width characters alone."
        Case "frmUnicodeCleanup.optZeroWidth": GuidedInfoAdvisory = "Targets zero-width characters and leaves visible spacing characters alone."
        Case "frmUnicodeCleanup.optCustom": GuidedInfoAdvisory = "Shows the individual invisible-character cleanups so you can combine them."
        Case "frmUnicodeCleanup.chkNBSP": GuidedInfoAdvisory = "Replaces nonbreaking spaces with normal spaces."
        Case "frmUnicodeCleanup.chkZWSP": GuidedInfoAdvisory = "Removes hidden zero-width spaces that split searches and word counts."
        Case "frmUnicodeCleanup.chkZWNJ": GuidedInfoAdvisory = "Removes zero-width nonjoiners often copied from web or PDF text."
        Case "frmUnicodeCleanup.chkZWJ": GuidedInfoAdvisory = "Removes zero-width joiners that can hide inside pasted text."
        Case "frmUnicodeCleanup.chkBOM": GuidedInfoAdvisory = "Removes stray byte-order marks embedded inside document text."
        Case "frmUnicodeCleanup.chkSoftHyphen": GuidedInfoAdvisory = "Removes soft hyphens that can disrupt clean searching and export."
        Case "frmUnicodeCleanup.chkNBHyphen": GuidedInfoAdvisory = "Replaces nonbreaking hyphens with standard editable hyphens."
        Case "frmStyleCleanup.chkRemoveUnused": GuidedInfoAdvisory = "Deletes custom styles that are not currently applied."
        Case "frmStyleCleanup.chkRemapVariants": GuidedInfoAdvisory = "Moves content into your canonical styles instead of leaving lookalike variants."
        Case Else
            GuidedInfoAdvisory = GuidedGenericAdvisory(CleanGuidedInfoCaption(fallbackText))
    End Select
End Function
Private Function GuidedGenericAdvisory(ByVal labelText As String) As String
    If Len(labelText) = 0 Then Exit Function
    If Left$(labelText, 6) = "Clear " Then
        GuidedGenericAdvisory = "Removes this metadata or review content from the file."
    ElseIf Left$(labelText, 7) = "Remove " Then
        GuidedGenericAdvisory = "Removes matching content wherever this cleanup finds it."
    ElseIf Left$(labelText, 10) = "Normalize " Then
        GuidedGenericAdvisory = "Resets direct variations to a cleaner, consistent form."
    ElseIf Left$(labelText, 8) = "Convert " Then
        GuidedGenericAdvisory = "Converts matching content into the target Word-friendly form."
    ElseIf Left$(labelText, 9) = "Collapse " Then
        GuidedGenericAdvisory = "Merges repeats so only one clean instance remains."
    ElseIf Left$(labelText, 10) = "Preserve " Then
        GuidedGenericAdvisory = "Leaves this formatting in place while other cleanup runs."
    ElseIf Left$(labelText, 8) = "Include " Then
        GuidedGenericAdvisory = "Applies the chosen action to this area."
    ElseIf Left$(labelText, 4) = "Fix " Then
        GuidedGenericAdvisory = "Corrects this issue without changing unrelated text."
    ElseIf Left$(labelText, 4) = "Set " Then
        GuidedGenericAdvisory = "Applies the selected setting when the tool runs."
    ElseIf Left$(labelText, 6) = "Also, " Then
        GuidedGenericAdvisory = "Adds this extra cleanup after the main removal step."
    ElseIf labelText = "Left" Or labelText = "Center" Or labelText = "Right" Then
        GuidedGenericAdvisory = "Uses this alignment when alignment cleanup is enabled."
    ElseIf labelText = "Custom" Then
        GuidedGenericAdvisory = "Shows the individual actions so you can combine them."
    Else
        GuidedGenericAdvisory = "Applies this cleanup when the tool runs."
    End If
End Function
Private Function CleanGuidedInfoCaption(ByVal captionText As String) As String
    CleanGuidedInfoCaption = Replace(captionText, vbCrLf, " ")
    CleanGuidedInfoCaption = Replace(CleanGuidedInfoCaption, vbCr, " ")
    CleanGuidedInfoCaption = Replace(CleanGuidedInfoCaption, vbLf, " ")
    Do While InStr(CleanGuidedInfoCaption, "  ") > 0
        CleanGuidedInfoCaption = Replace(CleanGuidedInfoCaption, "  ", " ")
    Loop
    CleanGuidedInfoCaption = Trim$(CleanGuidedInfoCaption)
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
Private Sub LayoutGuidedTitleReset(ByVal toolForm As Object, ByVal M As Single, ByVal contentW As Single)
    On Error Resume Next
    Dim resetW As Single
    Dim resetH As Single
    Dim resetLeft As Single
    Dim resetTop As Single
    resetW = 50
    resetH = 18
    resetLeft = M + contentW - resetW
    resetTop = 8

    StyleGuidedButton toolForm.Controls("cmdReset"), False
    toolForm.Controls("cmdReset").Caption = "Reset"
    toolForm.Controls("cmdReset").Font.Size = 7.5
    toolForm.Controls("cmdReset").Font.Bold = False
    toolForm.Controls("cmdReset").Move resetLeft, resetTop, resetW, resetH
    toolForm.Controls("cmdReset").Visible = True
    HideGuidedControl toolForm.Controls("cmdRun")
    HideGuidedControl toolForm.Controls("cmdHelp")
End Sub
Public Sub RefreshScopeSelectionState(ByVal toolForm As Object, Optional ByVal preferSelectionWhenAvailable As Boolean = False)
    On Error Resume Next
    If Not ScopeSelectionIsAvailable(toolForm) Then Exit Sub

    Dim selectionAvailable As Boolean
    selectionAvailable = (Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn)
    toolForm.Controls("optScopeSelection").Enabled = selectionAvailable

    If selectionAvailable Then
        If preferSelectionWhenAvailable Then
            toolForm.Controls("optScopeSelection").Value = True
        ElseIf Not toolForm.Controls("optScopeSelection").Value And Not toolForm.Controls("optScopeDocument").Value Then
            toolForm.Controls("optScopeDocument").Value = True
        End If
    ElseIf toolForm.Controls("optScopeSelection").Value Then
        toolForm.Controls("optScopeDocument").Value = True
    End If
End Sub
Private Function ScopeSelectionIsAvailable(ByVal toolForm As Object) As Boolean
    On Error Resume Next
    ScopeSelectionIsAvailable = _
        InStr(1, toolForm.Controls("optScopeDocument").Caption, "(always)", vbTextCompare) = 0 _
        And Len(Trim$(toolForm.Controls("optScopeSelection").Caption)) > 0
End Function
Private Sub LayoutGuidedPreviewAndScopeRow(ByVal toolForm As Object, ByVal M As Single, ByRef y As Single, ByVal contentW As Single)
    On Error Resume Next
    StyleGuidedButton toolForm.Controls("cmdPreview"), True
    toolForm.Controls("cmdPreview").Caption = "Preview"
    HideGuidedControl toolForm.Controls("cmdRun")
    HideGuidedControl toolForm.Controls("cmdHelp")

    If ScopeSelectionIsAvailable(toolForm) Then
        Dim scopeW As Single
        Dim previewW As Single
        Dim previewLeft As Single
        Dim scopeLeft As Single
        scopeW = 154
        previewW = contentW - scopeW - 10
        scopeLeft = M
        previewLeft = M + scopeW + 10

        StyleGuidedOption toolForm.Controls("optScopeSelection")
        StyleGuidedOption toolForm.Controls("optScopeDocument")
        toolForm.Controls("optScopeSelection").Caption = "Selected text only"
        toolForm.Controls("optScopeDocument").Caption = "Entire document"
        toolForm.Controls("optScopeSelection").Move scopeLeft, y, scopeW, 18
        toolForm.Controls("optScopeDocument").Move scopeLeft, y + 20, scopeW, 18
        toolForm.Controls("optScopeSelection").Visible = True
        toolForm.Controls("optScopeDocument").Visible = True
        toolForm.Controls("cmdPreview").Move previewLeft, y, previewW, 42
        toolForm.Controls("cmdPreview").Visible = True
        y = y + 48
    Else
        HideGuidedControl toolForm.Controls("optScopeDocument")
        HideGuidedControl toolForm.Controls("optScopeSelection")
        toolForm.Controls("cmdPreview").Move M, y, contentW, 30
        toolForm.Controls("cmdPreview").Visible = True
        y = y + 36
    End If
End Sub
Private Function GuidedToolName(ByVal formName As String) As String
    Select Case formName
        Case "frmCapitalizationCleanup": GuidedToolName = "Capitalization Fixer"
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
        Case "frmHyperlinkRemover": GuidedToolName = "Hyperlink Cleaner"
        Case "frmSoftReturnConverter": GuidedToolName = "Soft Return Converter"
        Case "frmFinalReview": GuidedToolName = "Final Review"
        Case "frmStyleCleanup": GuidedToolName = "Style Cleanup"
        Case "frmFootnoteRemover": GuidedToolName = "Footnote / Endnote Remover"
        Case "frmHeaderFooterStandardizer": GuidedToolName = "Header / Footer Standardizer"
        Case "frmObjectRemover": GuidedToolName = "Object Remover"
        Case Else: GuidedToolName = formName
    End Select
End Function
Private Function GuidedToolIntro(ByVal formName As String) As String
    Select Case formName
        Case "frmCapitalizationCleanup": GuidedToolIntro = "Use the recommended repair, or customize its protections."
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
        Case "frmHyperlinkRemover": GuidedToolIntro = "Choose whether to hide hyperlink styling or remove hyperlink targets."
        Case "frmSoftReturnConverter": GuidedToolIntro = "Choose which line-break conversion to apply."
        Case "frmFinalReview": GuidedToolIntro = "Choose which review markup should be finalized before sharing."
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
    If bodyOwnsTitle Then
        toolForm.Caption = ""
    Else
        HideLegacyTitleControls toolForm
    End If
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
Public Function GetTargetRange() As Range
    If gPreviewScopeHasRange Then
        If ActiveDocument.FullName = gPreviewScopeDocumentKey Or ActiveDocument.Name = gPreviewScopeDocumentKey Then
            Set GetTargetRange = ActiveDocument.Range(gPreviewScopeStart, gPreviewScopeEnd)
        Else
            Set GetTargetRange = ActiveDocument.Content
        End If
    ElseIf Selection.Type = wdSelectionNormal Or Selection.Type = wdSelectionColumn Then
        Set GetTargetRange = Selection.Range
    Else
        Set GetTargetRange = ActiveDocument.Content
    End If
End Function
Public Function ParagraphContainedInRange(ByVal paragraphItem As Paragraph, ByVal scopeRange As Range) As Boolean
    Dim paragraphRange As Range
    Set paragraphRange = paragraphItem.Range

    ParagraphContainedInRange = (paragraphRange.Start >= scopeRange.Start And _
                                 paragraphRange.End <= scopeRange.End)
End Function

Public Function ParagraphOverlapsRange(ByVal paragraphItem As Paragraph, ByVal scopeRange As Range) As Boolean
    Dim paragraphRange As Range
    Set paragraphRange = paragraphItem.Range

    ParagraphOverlapsRange = (paragraphRange.End > scopeRange.Start And _
                              paragraphRange.Start < scopeRange.End)
End Function

Public Function ParagraphOverlapsRangeOutsideTable(ByVal paragraphItem As Paragraph, ByVal scopeRange As Range) As Boolean
    If paragraphItem.Range.Information(wdWithInTable) Then Exit Function
    ParagraphOverlapsRangeOutsideTable = ParagraphOverlapsRange(paragraphItem, scopeRange)
End Function

Public Function ParagraphBodyText(ByVal paragraphItem As Paragraph) As String
    Dim fullText As String
    fullText = paragraphItem.Range.Text
    ParagraphBodyText = Left$(fullText, ParagraphBodyLength(fullText))
End Function

Public Function ParagraphBodyRange(ByVal paragraphItem As Paragraph) As Range
    Dim bodyRange As Range
    Dim fullText As String
    Set bodyRange = paragraphItem.Range.Duplicate
    fullText = bodyRange.Text
    bodyRange.End = bodyRange.Start + ParagraphBodyLength(fullText)
    Set ParagraphBodyRange = bodyRange
End Function

Public Sub ReplaceParagraphBodyText(ByVal paragraphItem As Paragraph, ByVal replacementText As String)
    Dim bodyRange As Range
    Set bodyRange = ParagraphBodyRange(paragraphItem)
    bodyRange.Text = replacementText
End Sub

Private Function ParagraphBodyLength(ByVal fullText As String) As Long
    Dim bodyLength As Long
    bodyLength = Len(fullText)

    Do While bodyLength > 0
        If Mid$(fullText, bodyLength, 1) = vbCr Or Mid$(fullText, bodyLength, 1) = Chr$(7) Then
            bodyLength = bodyLength - 1
        Else
            Exit Do
        End If
    Loop

    ParagraphBodyLength = bodyLength
End Function

Public Function CountPreviewFindMatches(ByVal searchRange As Range, ByVal findText As String) As Long
    Dim matchRange As Range
    Set matchRange = searchRange.Duplicate

    With matchRange.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False

        Do While .Execute
            CountPreviewFindMatches = CountPreviewFindMatches + 1
            matchRange.Collapse wdCollapseEnd
        Loop
    End With
End Function
