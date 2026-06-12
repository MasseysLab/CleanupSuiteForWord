Option Explicit
Private Sub UserForm_Initialize()
    chkReturnToMainAfterApply.Caption = "Return to main menu after apply"
    chkReturnToMainAfterApply.Value = GetReturnToMainAfterApplySetting()
    chkAutoSave.Caption = "Auto-save before running each tool"
    chkAutoSave.Value = GetAutoSaveSetting()
End Sub
Private Sub chkReturnToMainAfterApply_Click()
    SetReturnToMainAfterApplySetting chkReturnToMainAfterApply.Value
End Sub
Private Sub chkAutoSave_Click()
    If Not chkAutoSave.Value Then
        Dim msg As String
        msg = "Auto-save is your safety net." & vbCrLf & vbCrLf & _
              "Without it, if something goes wrong during cleanup, " & _
              "Word's recovery file may not have a clean pre-run version to fall back to." & vbCrLf & vbCrLf & _
              "Are you sure you want to turn auto-save off?"
        If MsgBox(msg, vbExclamation + vbYesNo, "Turn Off Auto-Save?") = vbNo Then
            chkAutoSave.Value = True
            Exit Sub
        End If
    End If
    SetAutoSaveSetting chkAutoSave.Value
End Sub
Private Sub cmdResetAll_Click()
    ResetCleanupSuiteDefaults
    chkReturnToMainAfterApply.Value = GetReturnToMainAfterApplySetting()
    chkAutoSave.Value = GetAutoSaveSetting()
    MsgBox "All Cleanup Suite tool settings have been reset to defaults.", vbInformation, "Cleanup Suite"
End Sub
Private Sub cmdUnicode_Click(): Me.Hide: frmUnicodeCleanup.Show: End Sub
Private Sub cmdPunctuation_Click(): Me.Hide: frmPunctuationCleanup.Show: End Sub
Private Sub cmdSpacing_Click(): Me.Hide: frmSpacingCleanup.Show: End Sub
Private Sub cmdCapitalization_Click(): Me.Hide: frmCapitalizationCleanup.Show: End Sub
Private Sub cmdList_Click(): Me.Hide: frmListCleanup.Show: End Sub
Private Sub cmdParagraph_Click(): Me.Hide: frmParagraphCleanup.Show: End Sub
Private Sub cmdDuplicate_Click(): Me.Hide: frmDuplicateDetector.Show: End Sub
Private Sub cmdFontNorm_Click(): Me.Hide: frmFontNormalizer.Show: End Sub
Private Sub cmdTableClean_Click(): Me.Hide: frmTableCleaner.Show: End Sub
Private Sub cmdBreakNorm_Click(): Me.Hide: frmBreakNormalizer.Show: End Sub
Private Sub cmdDocTrim_Click(): Me.Hide: frmDocumentTrim.Show: End Sub
Private Sub cmdFormatStrip_Click(): Me.Hide: frmFormattingStripper.Show: End Sub
Private Sub cmdHyperlink_Click(): Me.Hide: frmHyperlinkRemover.Show: End Sub
Private Sub cmdSoftReturn_Click(): Me.Hide: frmSoftReturnConverter.Show: End Sub
Private Sub cmdMetadata_Click(): Me.Hide: frmMetadataScrubber.Show: End Sub
Private Sub cmdStyleClean_Click(): Me.Hide: frmStyleCleanup.Show: End Sub
Private Sub cmdFootnote_Click(): Me.Hide: frmFootnoteRemover.Show: End Sub
Private Sub cmdHeaderFooter_Click(): Me.Hide: frmHeaderFooterStandardizer.Show: End Sub
Private Sub cmdObjectRemover_Click(): Me.Hide: frmObjectRemover.Show: End Sub
Private Sub cmdHelpUnicode_Click()
    Dim h As String
    h = "Invisible Unicode Cleaner" & vbCrLf & vbCrLf
    h = h & "Removes Unicode characters that are invisible on screen but cause problems " & _
            "for search, sorting, comparison, and automated text processing." & vbCrLf
    h = h & "These characters accumulate from web copy-paste, PDF export, and " & _
            "cross-platform file exchange." & vbCrLf
    h = h & vbCrLf
    h = h & "CHARACTER TYPES" & vbCrLf
    h = h & "  NBSP  (U+00A0)  Non-Breaking Space: a space that prevents line breaks." & vbCrLf
    h = h & "  ZWSP  (U+200B)  Zero-Width Space: completely invisible; affects word break." & vbCrLf
    h = h & "  ZWNJ  (U+200C)  Zero-Width Non-Joiner: controls ligature/joining behaviour." & vbCrLf
    h = h & "  ZWJ   (U+200D)  Zero-Width Joiner: forces joined rendering." & vbCrLf
    h = h & "  BOM   (U+FEFF)  Byte Order Mark: encoding marker; never belongs in text." & vbCrLf
    h = h & "  SHY   (U+00AD)  Soft Hyphen: invisible optional hyphen." & vbCrLf
    h = h & "  NBHY  (U+2011)  Non-Breaking Hyphen: a hyphen that prevents line breaks." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights matches in yellow without changing anything." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Invisible Unicode Cleaner"
End Sub
Private Sub cmdHelpPunct_Click()
    Dim h As String
    h = "Punctuation Normalizer" & vbCrLf & vbCrLf
    h = h & "Replaces typographic ('smart') punctuation with plain-text equivalents." & vbCrLf
    h = h & "Use this before exporting to systems that do not support Unicode punctuation, " & _
            "or before running text-comparison or diff tools." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  All          --  Converts all punctuation types listed below." & vbCrLf
    h = h & "  Quotes only  --  Curly double and single quotes become straight" & vbCrLf
    h = h & "  Dashes only  --  Em dash and en dash  -->  hyphen (-)" & vbCrLf
    h = h & "  Ellipses     --  Ellipsis character (...)  -->  three dots (...)" & vbCrLf
    h = h & "  Custom       --  Tick individual characters below." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights matches in yellow without changing anything." & vbCrLf
    h = h & "Run again without Preview to apply." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Punctuation Normalizer"
End Sub
Private Sub cmdHelpSpacing_Click()
    Dim h As String
    h = "Spacing Fixer" & vbCrLf & vbCrLf
    h = h & "Corrects the most common spacing inconsistencies in assembled documents." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Double spaces       --  Collapses two or more consecutive spaces to one." & vbCrLf
    h = h & "  Trim paragraphs     --  Removes leading and trailing spaces from each" & vbCrLf
    h = h & "                          paragraph." & vbCrLf
    h = h & "  Space before punct  --  Removes any space immediately before . , ; : ! ?" & vbCrLf
    h = h & "  Space after punct   --  Ensures exactly one space follows  . ? ! , ; :" & vbCrLf
    h = h & "  Extra blank lines   --  Collapses three or more consecutive blank" & vbCrLf
    h = h & "                          paragraphs down to two." & vbCrLf
    h = h & vbCrLf
    h = h & "Each option counts how many changes it makes and reports them at the end." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights problem areas without making changes." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Spacing Fixer"
End Sub
Private Sub cmdHelpCap_Click()
    Dim h As String
    h = "Capitalization Fixer" & vbCrLf & vbCrLf
    h = h & "Applies consistent capitalisation to every paragraph in the target scope." & vbCrLf
    h = h & vbCrLf
    h = h & "MODES" & vbCrLf
    h = h & "  All (Smart)   --  Default. Capitalises the first letter after sentence-ending" & vbCrLf
    h = h & "                    punctuation (. ? !).  Best choice for general text." & vbCrLf
    h = h & "  Sentence case --  First letter of paragraph capitalised; rest lowercased." & vbCrLf
    h = h & "  Title Case    --  First letter of every word capitalised." & vbCrLf
    h = h & "  UPPERCASE     --  Every letter uppercased." & vbCrLf
    h = h & "  lowercase     --  Every letter lowercased." & vbCrLf
    h = h & "  Custom        --  Combine modes using the checkboxes.  If more than one " & _
            "is selected, each is applied in sequence -- the last" & vbCrLf
    h = h & "                    one applied wins." & vbCrLf
    h = h & vbCrLf
    h = h & "NOTE: Capitalisation is applied to the full text content of each paragraph, " & _
            "including any directly formatted runs within it." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights paragraphs that would change." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Capitalization Fixer"
End Sub
Private Sub cmdHelpList_Click()
    Dim h As String
    h = "List Normalizer" & vbCrLf & vbCrLf
    h = h & "Standardises list formatting throughout the document." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Normalise bullets    --  Resets all bullet-list paragraphs to the default" & vbCrLf
    h = h & "                           Word bullet style, removing inconsistent characters." & vbCrLf
    h = h & "  Normalise numbering  --  Converts manually typed numbered paragraphs" & vbCrLf
    h = h & "                           (e.g.  1. Item) to proper Word auto-numbering." & vbCrLf
    h = h & "  Fix list indentation --  Sets list paragraph indentation to 0.25 inch." & vbCrLf
    h = h & "  Convert hyphens      --  Converts lines beginning with '- ' or '* ' to" & vbCrLf
    h = h & "                           proper Word bullet-list paragraphs." & vbCrLf
    h = h & vbCrLf
    h = h & "TIP: Run 'Convert hyphens' first if you have many hyphen-lists, then " & _
            "'Normalise bullets' to make them all consistent." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights list paragraphs that would be affected." & vbCrLf
    MsgBox h, vbInformation, "Help  --  List Normalizer"
End Sub
Private Sub cmdHelpPara_Click()
    Dim h As String
    h = "Paragraph Structure Fixer" & vbCrLf & vbCrLf
    h = h & "Cleans up paragraph structure and direct spacing overrides." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Remove empty paragraphs  --  Deletes completely empty paragraphs." & vbCrLf
    h = h & "  Collapse extra breaks    --  Reduces runs of 3+ consecutive paragraph breaks" & vbCrLf
    h = h & "                               to exactly 2 (a standard visual separator)." & vbCrLf
    h = h & "  Normalise para spacing   --  Resets SpaceBefore and SpaceAfter to zero for" & vbCrLf
    h = h & "                               paragraphs that have direct spacing overrides." & vbCrLf
    h = h & "                               Does not affect spacing set by the style itself." & vbCrLf
    h = h & "  Fix indent               --  Resets left indent to zero for paragraphs with" & vbCrLf
    h = h & "                               direct indent overrides." & vbCrLf
    h = h & vbCrLf
    h = h & "TIP: Run after assembling a document from multiple sources or templates, " & _
            "where inconsistent spacing between paragraphs is most common." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights empty paragraphs before any changes are made." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Paragraph Structure Fixer"
End Sub
Private Sub cmdHelpDuplicate_Click()
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
Private Sub cmdHelpFont_Click()
    Dim h As String
    h = "Font Normalizer" & vbCrLf & vbCrLf
    h = h & "Resets direct font overrides to match the style defined for each paragraph, " & _
            "eliminating the font chaos that accumulates from copying and pasting text" & vbCrLf
    h = h & "between documents, emails, and web pages." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS  (choose which properties to reset)" & vbCrLf
    h = h & "  Font face   --  Resets the typeface (e.g. Arial, Times New Roman) to the" & vbCrLf
    h = h & "                  style's defined font." & vbCrLf
    h = h & "  Font size   --  Resets the point size to the style's defined size." & vbCrLf
    h = h & "  Bold        --  Resets bold to the style's setting." & vbCrLf
    h = h & "  Italic      --  Resets italic to the style's setting." & vbCrLf
    h = h & "  Font colour --  Resets text colour to the style's defined colour." & vbCrLf
    h = h & vbCrLf
    h = h & "TIP: Start with Face and Size only.  Bold, Italic, and Colour are opt-in " & _
            "because those properties are more likely to be intentional (e.g. a bolded" & vbCrLf
    h = h & "phrase or a deliberately coloured heading within a paragraph)." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights paragraphs with direct font overrides." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Font Normalizer"
End Sub
Private Sub cmdHelpTable_Click()
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
    h = h & "  Convert table to text --  Replaces the table with tab-separated text." & vbCrLf
    h = h & vbCrLf
    h = h & "TIP: Run 'Remove empty rows/columns' before 'Strip direct format' for " & _
            "the cleanest results.  Use Preview first to see which rows are empty." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to tables in the selection." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Table Cleaner"
End Sub
Private Sub cmdHelpBreak_Click()
    Dim h As String
    h = "Break Normalizer" & vbCrLf & vbCrLf
    h = h & "Fixes problems with section and page breaks.  Most common in documents " & _
            "assembled from multiple files, where each source file may have had its" & vbCrLf
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
    h = h & "NOTE: The first section of a document is never converted -- it has no " & _
            "preceding break to change." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights consecutive break sequences in yellow." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Break Normalizer"
End Sub
Private Sub cmdHelpTrim_Click()
    Dim h As String
    h = "Document Trim" & vbCrLf & vbCrLf
    h = h & "Removes trailing empty paragraphs from the end of the document, " & _
            "leaving exactly one final empty paragraph (which Word requires to exist)." & vbCrLf
    h = h & vbCrLf
    h = h & "WHY THIS MATTERS" & vbCrLf
    h = h & "  Documents assembled from multiple files, or saved from templates, " & _
            "  often end with many blank paragraphs.  These appear as unwanted blank" & vbCrLf
    h = h & "  space at the end of the printed document and can cause unexpected " & _
            "  behaviour with headers, footers, and page numbering in some layouts." & vbCrLf
    h = h & vbCrLf
    h = h & "NOTE: This tool always operates on the full document.  Scope selection " & _
            "is disabled because 'end of document' is inherently document-level --" & vbCrLf
    h = h & "trimming to the end of a selection would not be meaningful." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights the trailing paragraphs that would be removed." & vbCrLf
    h = h & "Run again without Preview to apply." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Document Trim"
End Sub
Private Sub cmdHelpFormat_Click()
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
    h = h & vbCrLf
    h = h & "NOTE: Quick mode restores emphasis only when it is uniform across the " & _
            "whole paragraph.  For a paragraph with one bold word among normal text," & vbCrLf
    h = h & "use Thorough mode to preserve that word precisely." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights paragraphs carrying direct formatting." & vbCrLf
    h = h & "  Preserve drop caps keeps dropped capitals; untick to reset them." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Direct Formatting Stripper"
End Sub
Private Sub cmdHelpHyperlink_Click()
    Dim h As String
    h = "Hyperlink Remover" & vbCrLf & vbCrLf
    h = h & "Removes hyperlinks while keeping the visible text.  Pasting from web " & _
            "pages embeds links throughout a document; this strips the links but" & vbCrLf
    h = h & "leaves the words exactly as they are." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Clear hyperlink formatting  --  also removes the blue underlined look" & vbCrLf
    h = h & "                                  (resets to the default character font)." & vbCrLf
    h = h & "                                  Turn off to keep the blue/underline." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights every hyperlink that would be removed." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Hyperlink Remover"
End Sub
Private Sub cmdHelpSoftReturn_Click()
    Dim h As String
    h = "Soft Return Converter" & vbCrLf & vbCrLf
    h = h & "Converts soft returns (line breaks made with Shift+Enter) to real " & _
            "paragraph marks, or the reverse.  Documents imported from PDFs and web" & vbCrLf
    h = h & "pages are often full of soft returns that look like paragraph breaks but " & _
            "are not -- which quietly breaks spell-check, styles, and spacing." & vbCrLf
    h = h & vbCrLf
    h = h & "DIRECTION" & vbCrLf
    h = h & "  Soft returns to paragraphs  --  the usual fix for imported text." & vbCrLf
    h = h & "  Paragraphs to soft returns  --  the reverse (less common)." & vbCrLf
    h = h & vbCrLf
    h = h & "CAUTION: Converting paragraphs to soft returns merges the affected text " & _
            "into a single paragraph.  Use the reverse only on small selections." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text before opening to limit changes to the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode highlights the breaks that would be converted." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Soft Return Converter"
End Sub
Private Sub cmdHelpMetadata_Click()
    Dim h As String
    h = "Metadata Scrubber" & vbCrLf & vbCrLf
    h = h & "Strips identifying information from the document before you share it " & _
            "externally -- a one-click version of Word built-in Document Inspector." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Document properties  --  author, company, manager, title, and other" & vbCrLf
    h = h & "                           built-in properties." & vbCrLf
    h = h & "  Personal information --  author names attached to comments and tracked" & vbCrLf
    h = h & "                           changes (anonymises them)." & vbCrLf
    h = h & "  Delete all comments  --  removes every comment from the document." & vbCrLf
    h = h & "  Accept all tracked changes  --  finalises the document and removes the" & vbCrLf
    h = h & "                                  revision history." & vbCrLf
    h = h & vbCrLf
    h = h & "NOTE: This tool always works on the whole document.  Metadata removal " & _
            "may not be reversible with Ctrl+Z, so a copy is saved first if auto-save" & vbCrLf
    h = h & "is on.  Review the result before sharing." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode shows the metadata currently in the document." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Metadata Scrubber"
End Sub
Private Sub cmdHelpStyle_Click()
    Dim h As String
    h = "Style Cleanup" & vbCrLf & vbCrLf
    h = h & "Removes clutter from the document style list.  Documents assembled from " & _
            "many sources accumulate dozens of unused or duplicate styles." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Remove unused styles    --  deletes custom styles that no text uses." & vbCrLf
    h = h & "                              Built-in Word styles are never deleted." & vbCrLf
    h = h & "  Remap common variants   --  reassigns text in variant styles such as" & vbCrLf
    h = h & "                              Normal (Web), Body Text 2, and Body Text 3 " & _
            "                              to the Normal style, then removes them." & vbCrLf
    h = h & vbCrLf
    h = h & "NOTE: This tool always works on the whole document.  Usage is detected " & _
            "in the main document body; a style used only inside a header, footer, or" & vbCrLf
    h = h & "text box may be treated as unused." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview mode reports how many unused custom styles were found." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Style Cleanup"
End Sub
Private Sub cmdHelpFootnote_Click()
    Dim h As String
    h = "Footnote / Endnote Remover" & vbCrLf & vbCrLf
    h = h & "Deletes footnotes and/or endnotes -- the reference mark and the note " & _
            "text -- to produce a clean text extract." & vbCrLf
    h = h & vbCrLf
    h = h & "OPTIONS" & vbCrLf
    h = h & "  Remove footnotes / Remove endnotes  --  choose which to delete." & vbCrLf
    h = h & "  Keep note text inline  --  inserts each note's text at the reference" & vbCrLf
    h = h & "                            point in [brackets] before deleting the note." & vbCrLf
    h = h & vbCrLf
    h = h & "Word renumbers any remaining notes automatically." & vbCrLf
    h = h & vbCrLf
    h = h & "SCOPE" & vbCrLf
    h = h & "  Select text first to remove only the notes referenced in the selection." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview highlights the reference marks that would be removed." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Footnote / Endnote Remover"
End Sub
Private Sub cmdHelpHeaderFooter_Click()
    Dim h As String
    h = "Header / Footer Standardizer" & vbCrLf & vbCrLf
    h = h & "Makes header/footer formatting consistent across the whole document, " & _
            "or clears all header/footer content." & vbCrLf
    h = h & vbCrLf
    h = h & "MODE" & vbCrLf
    h = h & "  Standardize formatting  --  reset font to the document default, remove " & _
            "paragraph spacing, and/or set alignment." & vbCrLf
    h = h & "  Clear all content       --  remove all header/footer text." & vbCrLf
    h = h & vbCrLf
    h = h & "  Include headers/footers selects which areas to process; each section " & _
            "  has primary, first-page, and even-page slots." & vbCrLf
    h = h & "  Unlink sections breaks 'Link to Previous' so each is set on its own." & vbCrLf
    h = h & vbCrLf
    h = h & "NOTE: page numbers and other fields are left untouched.  Works on the " & _
            "whole document." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Header / Footer Standardizer"
End Sub
Private Sub cmdHelpObject_Click()
    Dim h As String
    h = "Remove Objects & Elements" & vbCrLf & vbCrLf
    h = h & "Deletes clutter from web/PDF/email paste so you can reduce a document " & _
            "to clean text." & vbCrLf
    h = h & vbCrLf
    h = h & "Removes (each OFF by default): pictures, text boxes (with contents), " & _
            "frames (text kept), horizontal lines, HTML/ActiveX controls, hidden" & vbCrLf
    h = h & "text, and -- with a confirmation -- whole tables and their contents." & vbCrLf
    h = h & vbCrLf
    h = h & "CAUTIONS" & vbCrLf
    h = h & "  - Deleting tables is destructive (unlike Table Cleaner convert-to-text)." & vbCrLf
    h = h & "  - Picture/shape deletion may not always be fully undoable." & vbCrLf
    h = h & "  - Body only; header/footer objects (e.g. logos) are left alone." & vbCrLf
    h = h & vbCrLf
    h = h & "Preview reports a count of what would be removed." & vbCrLf
    MsgBox h, vbInformation, "Help  --  Remove Objects & Elements"
End Sub
