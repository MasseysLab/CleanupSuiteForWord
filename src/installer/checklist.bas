A. If any forms are missing or controls failed to create programmatically:
   1. Open the VBA editor (Alt+F11).
   2. For each missing UserForm: Insert -> UserForm. In the Properties window set (Name) exactly as listed.
   3. Add controls using the Toolbox: OptionButton, Frame, CheckBox, CommandButton, TextBox. Use the exact control names used by the code (see below).
   4. After adding controls, open the form's code window and ensure the code is present (the installer injected code). If not, paste the code from the installer source into the form's code window.

B. Exact control names (copy/paste into the designer):
   frmCleanupSuiteLauncher: cmdUnicode, cmdPunctuation, cmdSpacing, cmdCapitalization, cmdList, cmdParagraph
   frmPunctuationCleanup: optAll, optQuotes, optDashes, optEllipses, optCustom, fraCustom, chkCurlyDouble, chkCurlySingle, chkEmDash, chkEnDash, chkEllipses, chkPreviewOnly, cmdSelectAll, cmdDeselectAll, cmdRun
   frmUnicodeCleanup: optAll, optNBSP, optZeroWidth, optCustom, fraCustom, chkNBSP, chkZWSP, chkZWNJ, chkZWJ, chkBOM, chkSoftHyphen, chkNBHyphen, chkPreviewOnly, cmdSelectAll, cmdDeselectAll, cmdRun
   frmSpacingCleanup: optAll, optDoubleSpaces, optTrim, optCustom, fraCustom, chkDoubleSpaces, chkTrimSpaces, chkSpaceBeforePunct, chkNormalizeAfterPunct, chkExtraBlankLines, chkPreviewOnly, cmdSelectAll, cmdDeselectAll, cmdRun
   frmCapitalizationCleanup: optAll, optSentence, optTitle, optUpper, optLower, optCustom, fraCustom, chkSentence, chkTitle, chkUpper, chkLower, chkSmartSentences, chkPreviewOnly, cmdSelectAll, cmdDeselectAll, cmdRun
   frmListCleanup: optAll, optBullets, optNumbering, optIndent, optCustom, fraCustom, chkNormalizeBullets, chkNormalizeNumbering, chkFixIndent, chkHyphenToBullets, chkPreviewOnly, cmdSelectAll, cmdDeselectAll, cmdRun
   frmParagraphCleanup: optAll, optRemoveEmpty, optNormalizeSpacing, optCustom, fraCustom, chkRemoveEmpty, chkCollapseBreaks, chkNormalizeParaSpacing, chkFixIndent, chkPreviewOnly, cmdSelectAll, cmdDeselectAll, cmdRun

C. After adding controls and verifying code: Debug -> Compile. Then run ShowCleanupSuiteLauncher and test each tool in Preview Only mode on a copy of a document.
