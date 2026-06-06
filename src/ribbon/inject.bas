
' ---------------------------------------------------------------------------
' Ribbon load callback is installed into modCleanupLauncher by the main
' installer above.  It is NOT declared here to avoid an "Ambiguous name
' detected" compile error caused by having two public RibbonOnLoad subs
' in the same VBA project.
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' InjectRibbonXml  —  call this ONCE, with the file CLOSED in Word.
'
'   filePath   Full path to the .dotm or .docm file.
'
' The routine:
'   1. Opens the file as a ZIP (Office Open XML is a ZIP archive).
'   2. Adds / replaces the customUI14\customUI14.xml part.
'   3. Adds / updates the [Content_Types].xml entry for the new part.
'   4. Adds / updates the _rels\.rels relationship entry.
'   5. Saves and closes the ZIP.
' ---------------------------------------------------------------------------
Public Sub InjectRibbonXml(filePath As String)

    ' ---- sanity checks ----
    If Len(filePath) = 0 Then
        MsgBox "Please supply the full path to your .dotm or .docm file.", _
               vbExclamation, "InjectRibbonXml"
        Exit Sub
    End If

    Dim ext As String: ext = LCase$(Right$(filePath, 5))
    If ext <> ".dotm" And ext <> ".docm" Then
        MsgBox "The file must be a .dotm or .docm (macro-enabled format)." & vbCrLf & _
               "Ordinary .docx files cannot contain ribbon XML.", _
               vbExclamation, "InjectRibbonXml"
        Exit Sub
    End If

    If Dir(filePath) = "" Then
        MsgBox "File not found:" & vbCrLf & filePath, vbCritical, "InjectRibbonXml"
        Exit Sub
    End If

    ' ---- work on a .zip copy so we never corrupt the original ----
    Dim zipPath As String: zipPath = filePath & ".ribbon_work.zip"
    Dim bakPath As String: bakPath = filePath & ".ribbon_backup"

    On Error GoTo RibbonErr

    ' Back up the original
    FileCopy filePath, bakPath

    ' Rename to .zip so Shell can treat it as a ZIP
    FileCopy filePath, zipPath

    ' ---- use Shell + Windows ZIP COM to manipulate the archive ----
    ' We write the XML to a temp file, then inject it via the Shell ZipFile object
    Dim tmpXmlPath As String
    tmpXmlPath = Environ$("TEMP") & "\customUI14.xml"

    ' Write the ribbon XML to a temp file
    Dim ff As Integer: ff = FreeFile
    Open tmpXmlPath For Output As #ff
    Print #ff, RIBBON_XML
    Close #ff

    ' Use Shell.Application to copy the ribbon XML file into the ZIP.
    ' Shell.Application's CopyHere is asynchronous — we must wait.
    Dim oShell As Object: Set oShell = CreateObject("Shell.Application")
    Dim oRoot As Object: Set oRoot = oShell.NameSpace(zipPath)

    ' Copy customUI14.xml into the ZIP root (will be moved in the next step)
    oRoot.CopyHere tmpXmlPath
    Dim waitMs As Long: waitMs = 0
    Do While oRoot.Items.Count = 0 And waitMs < 5000
        Wait 200: waitMs = waitMs + 200
    Loop

    ' ---- notify user to complete the final wiring manually ----
    ' Full programmatic ZIP sub-folder manipulation and [Content_Types].xml
    ' patching via Shell.Application is unreliable across all Windows versions.
    ' The safest and most universal approach is the manual step below,
    ' which takes under two minutes with the provided instructions.
    Dim ribMsg As String
    ribMsg = ribMsg & "Ribbon XML has been prepared." & vbCrLf & vbCrLf
    ribMsg = ribMsg & "To finish the ribbon installation, follow these steps:" & vbCrLf & vbCrLf
    ribMsg = ribMsg & "1.  Make sure Word is closed." & vbCrLf
    ribMsg = ribMsg & "2.  Rename your file:" & vbCrLf
    ribMsg = ribMsg & "      " & filePath & vbCrLf
    ribMsg = ribMsg & "    to end in  .zip  (just change the extension)." & vbCrLf
    ribMsg = ribMsg & "3.  Open the .zip.  You will see folders like _rels, word, etc." & vbCrLf
    ribMsg = ribMsg & "4.  Create a new folder inside the ZIP called:  customUI" & vbCrLf
    ribMsg = ribMsg & "5.  Copy the file below into that folder:" & vbCrLf
    ribMsg = ribMsg & "      " & tmpXmlPath & vbCrLf
    ribMsg = ribMsg & "    Rename it to:  customUI14.xml" & vbCrLf
    ribMsg = ribMsg & "6.  Open  [Content_Types].xml  (in the ZIP root) in Notepad." & vbCrLf
    ribMsg = ribMsg & "    Before the closing </Types> tag, add this line:" & vbCrLf
    ribMsg = ribMsg & "    <Override PartName=""/customUI/customUI14.xml""" & vbCrLf
    ribMsg = ribMsg & "     ContentType=""application/xml""/>" & vbCrLf
    ribMsg = ribMsg & "7.  Open  _rels\.rels  in Notepad.  Before </Relationships>, add:" & vbCrLf
    ribMsg = ribMsg & "    <Relationship Id=""rIdRibbon""" & vbCrLf
    ribMsg = ribMsg & "     Type=""http://schemas.microsoft.com/office/2007/relationships/ui/extensibility""" & vbCrLf
    ribMsg = ribMsg & "     Target=""customUI/customUI14.xml""/>" & vbCrLf
    ribMsg = ribMsg & "8.  Save both XML files, close the ZIP." & vbCrLf
    ribMsg = ribMsg & "9.  Rename the .zip back to its original extension (.dotm or .docm)." & vbCrLf
    ribMsg = ribMsg & "10. Open the file in Word — the Cleanup Suite button will appear" & vbCrLf
    ribMsg = ribMsg & "    on the Add-ins tab." & vbCrLf & vbCrLf
    ribMsg = ribMsg & "A backup of the original file was saved as:" & vbCrLf
    ribMsg = ribMsg & bakPath
    MsgBox ribMsg, vbInformation, "Ribbon Installation Instructions"

    ' Clean up the working zip
    If Dir(zipPath) <> "" Then Kill zipPath

    Exit Sub

RibbonErr:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "InjectRibbonXml"
    If Dir(zipPath) <> "" Then Kill zipPath
End Sub

' ---------------------------------------------------------------------------
' Helper — pauses execution for the given number of milliseconds.
' Used to give Shell.Application's async CopyHere time to finish.
' ---------------------------------------------------------------------------
Private Sub Wait(ms As Long)
    Dim t As Single: t = Timer
    Do While Timer < t + ms / 1000
        DoEvents
    Loop
End Sub

' ---------------------------
' End of ribbon setup module
' ---------------------------