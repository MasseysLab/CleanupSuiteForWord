
' ---------------------------------------------------------------------------
' Ribbon load callback is installed into modCleanupLauncher by the main
' installer above.  It is NOT declared here to avoid an "Ambiguous name
' detected" compile error caused by having two public RibbonOnLoad subs
' in the same VBA project.
' ---------------------------------------------------------------------------

' ---------------------------------------------------------------------------
' InjectRibbonXml -- call this ONCE, with the file CLOSED in Word.
'
'   filePath   Full path to the .dotm or .docm file.
'
' The routine:
'   1. Opens the file as a ZIP (Office Open XML is a ZIP archive).
'   2. Adds / replaces the /customUI/customUI14.xml part.
'   3. Adds / updates the [Content_Types].xml entry for the new part.
'   4. Adds / updates the _rels\.rels relationship entry.
'   5. Saves and closes the ZIP.
' ---------------------------------------------------------------------------
Public Sub InjectRibbonXml(filePath As String)
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

    On Error GoTo RibbonErr

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim tempRoot As String
    tempRoot = Environ$("TEMP") & "\CleanupSuiteRibbon_" & Format$(Now, "yyyymmdd_hhnnss")

    Dim bakPath As String
    bakPath = filePath & ".ribbon_backup"

    Dim zipPath As String
    zipPath = filePath & ".ribbon_work.zip"

    If fso.FileExists(bakPath) Then fso.DeleteFile bakPath, True
    If fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True
    If fso.FolderExists(tempRoot) Then fso.DeleteFolder tempRoot, True

    fso.CreateFolder tempRoot
    fso.CreateFolder tempRoot & "\customUI"

    FileCopy filePath, bakPath
    FileCopy filePath, zipPath

    Dim oShell As Object
    Set oShell = CreateObject("Shell.Application")

    Dim zipNs As Object
    Set zipNs = oShell.NameSpace(zipPath)
    If zipNs Is Nothing Then Err.Raise vbObjectError + 3101, "InjectRibbonXml", "Could not open the Office file as a ZIP package."

    oShell.NameSpace(tempRoot).CopyHere zipNs.Items, 16
    WaitForShellCopy tempRoot & "\[Content_Types].xml", 15000

    WriteTextFileUtf8 tempRoot & "\customUI\customUI14.xml", RIBBON_XML

    Dim contentTypesPath As String
    contentTypesPath = tempRoot & "\[Content_Types].xml"
    Dim contentTypesXml As String
    contentTypesXml = ReadTextFileUtf8(contentTypesPath)
    contentTypesXml = RemoveXmlElementByAttribute(contentTypesXml, "Override", "PartName", "/customUI/customUI14.xml")
    contentTypesXml = InsertBeforeClosingTag(contentTypesXml, "Types", _
        "  <Override PartName=""/customUI/customUI14.xml"" ContentType=""application/xml""/>" & vbCrLf)
    WriteTextFileUtf8 contentTypesPath, contentTypesXml

    Dim relsPath As String
    relsPath = tempRoot & "\_rels\.rels"
    Dim relsXml As String
    relsXml = ReadTextFileUtf8(relsPath)
    relsXml = RemoveXmlElementByAttribute(relsXml, "Relationship", "Target", "customUI/customUI14.xml")
    relsXml = RemoveXmlElementByAttribute(relsXml, "Relationship", "Target", "/customUI/customUI14.xml")
    relsXml = InsertBeforeClosingTag(relsXml, "Relationships", _
        "  <Relationship Id=""" & NextRelationshipId(relsXml) & """ Type=""http://schemas.microsoft.com/office/2007/relationships/ui/extensibility"" Target=""customUI/customUI14.xml""/>" & vbCrLf)
    WriteTextFileUtf8 relsPath, relsXml

    fso.DeleteFile zipPath, True
    CreateEmptyZip zipPath
    Set zipNs = oShell.NameSpace(zipPath)
    If zipNs Is Nothing Then Err.Raise vbObjectError + 3102, "InjectRibbonXml", "Could not create the replacement ZIP package."

    zipNs.CopyHere fso.GetFolder(tempRoot).Items, 16
    WaitForZipItems zipNs, fso.GetFolder(tempRoot).Files.Count + fso.GetFolder(tempRoot).SubFolders.Count, 20000
    Wait 3000

    fso.DeleteFile filePath, True
    fso.MoveFile zipPath, filePath
    fso.DeleteFolder tempRoot, True

    MsgBox "Ribbon XML was installed successfully." & vbCrLf & vbCrLf & _
           "A backup of the original file was saved as:" & vbCrLf & bakPath & vbCrLf & vbCrLf & _
           "Re-open the file in Word and use the Cleanup Suite button on the Add-ins tab.", _
           vbInformation, "InjectRibbonXml"
    Exit Sub

RibbonErr:
    On Error Resume Next
    If Not fso Is Nothing Then
        If Len(zipPath) > 0 And fso.FileExists(zipPath) Then fso.DeleteFile zipPath, True
        If Len(tempRoot) > 0 And fso.FolderExists(tempRoot) Then fso.DeleteFolder tempRoot, True
        If Len(bakPath) > 0 And fso.FileExists(bakPath) And fso.FileExists(filePath) = False Then
            fso.CopyFile bakPath, filePath, True
        End If
    End If
    MsgBox "Ribbon injection failed." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Your original file was left unchanged or restored from backup.", _
           vbCritical, "InjectRibbonXml"
End Sub

' ---------------------------------------------------------------------------
' Helpers for ZIP rebuilding and package XML patching.
' ---------------------------------------------------------------------------
Private Sub Wait(ms As Long)
    Dim t As Single: t = Timer
    Do While Timer < t + ms / 1000
        DoEvents
    Loop
End Sub

Private Sub WaitForShellCopy(requiredPath As String, timeoutMs As Long)
    Dim waited As Long
    Do While Dir(requiredPath) = "" And waited < timeoutMs
        Wait 200
        waited = waited + 200
    Loop
    If Dir(requiredPath) = "" Then Err.Raise vbObjectError + 3103, "InjectRibbonXml", "Timed out while unpacking the Office file."
End Sub

Private Sub WaitForZipItems(zipNs As Object, expectedTopLevelItems As Long, timeoutMs As Long)
    Dim waited As Long
    Do While zipNs.Items.Count < expectedTopLevelItems And waited < timeoutMs
        Wait 250
        waited = waited + 250
    Loop
    If zipNs.Items.Count < expectedTopLevelItems Then Err.Raise vbObjectError + 3104, "InjectRibbonXml", "Timed out while rebuilding the Office file."
End Sub

Private Sub CreateEmptyZip(zipPath As String)
    Dim emptyZip(0 To 21) As Byte
    emptyZip(0) = 80
    emptyZip(1) = 75
    emptyZip(2) = 5
    emptyZip(3) = 6

    Dim ff As Integer
    ff = FreeFile
    Open zipPath For Binary As #ff
    Put #ff, , emptyZip
    Close #ff
End Sub

Private Function ReadTextFileUtf8(filePath As String) As String
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile filePath
    ReadTextFileUtf8 = stm.ReadText
    stm.Close
End Function

Private Sub WriteTextFileUtf8(filePath As String, text As String)
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2
    stm.Charset = "utf-8"
    stm.Open
    stm.WriteText text
    stm.SaveToFile filePath, 2
    stm.Close
End Sub

Private Function InsertBeforeClosingTag(xmlText As String, tagName As String, insertion As String) As String
    Dim closeTag As String
    closeTag = "</" & tagName & ">"

    Dim p As Long
    p = InStrRev(xmlText, closeTag, -1, vbTextCompare)
    If p = 0 Then Err.Raise vbObjectError + 3105, "InjectRibbonXml", "Could not find closing tag " & closeTag & "."

    InsertBeforeClosingTag = Left$(xmlText, p - 1) & insertion & Mid$(xmlText, p)
End Function

Private Function RemoveXmlElementByAttribute(xmlText As String, elementName As String, attrName As String, attrValue As String) As String
    Dim searchValue As String
    searchValue = attrName & "=""" & attrValue & """"

    Dim p As Long
    p = InStr(1, xmlText, searchValue, vbTextCompare)

    Do While p > 0
        Dim startTag As Long
        startTag = InStrRev(Left$(xmlText, p), "<" & elementName, -1, vbTextCompare)
        If startTag = 0 Then Exit Do

        Dim endTag As Long
        endTag = InStr(p, xmlText, ">")
        If endTag = 0 Then Exit Do

        xmlText = Left$(xmlText, startTag - 1) & Mid$(xmlText, endTag + 1)
        p = InStr(1, xmlText, searchValue, vbTextCompare)
    Loop

    RemoveXmlElementByAttribute = xmlText
End Function

Private Function NextRelationshipId(relsXml As String) As String
    Dim n As Long
    n = 1
    Do While InStr(1, relsXml, "Id=""rIdCleanupSuite" & CStr(n) & """", vbTextCompare) > 0
        n = n + 1
    Loop
    NextRelationshipId = "rIdCleanupSuite" & CStr(n)
End Function

' ---------------------------
' End of ribbon setup module
' ---------------------------
