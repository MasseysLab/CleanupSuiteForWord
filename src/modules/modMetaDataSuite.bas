Attribute VB_Name = "modMetaDataSuite"
Option Explicit

' ------------------------------------------------------------
' VBA translation of the DOCX metadata audit workflow.
' Intended for use in Word or Excel VBA.
'
' Notes:
' - This version is a practical VBA approximation of the Python tool.
' - It uses PowerShell's Expand-Archive to unpack the DOCX file.
' - XML parsing is done with MSXML.
' - Hash/encoding/guessing checks are pattern-based.
' ------------------------------------------------------------

Private Const STATUS_LEGEND As String = _
    "CRY=Cryptography Analysis | DIC=Dictionary Attack | ENC=Encoding Detection | ENT=Entropy Analysis | " & _
    "ERR=Error Found | FLD=Field Mutation | GES=Guessing Algorithm | HAS=Hash Detection | " & _
    "IMG=Image Metadata | OBJ=Object Detection | PAS=Password Guessing | PKG=Package Analysis | " & _
    "PRM=Permutation Scan | REL=Relationship Check | TIM=Time Tracking | XML=XML Parsing"

Private Const HASH_PATTERN_LIST As String = _
    "^[A-Fa-f0-9]{32}$|^[A-Fa-f0-9]{40}$|^[A-Fa-f0-9]{64}$|" & _
    "^\$2[aby]\$\d{2}\$[./A-Za-z0-9]{53}$|^\$6\$[./A-Za-z0-9]{16}\$[./A-Za-z0-9]{86}$"

Private Const BASE64_PATTERN As String = "^[A-Za-z0-9+/]+={0,2}$"
Private Const HEX_PATTERN As String = "^[A-Fa-f0-9]+$"
Private Const GUID_PATTERN As String = "\b[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\b"
Private Const URL_PATTERN As String = "https?://[^\s'""]+|file://[^\s'""]+|mailto:[^\s'""]+"
Private Const DATE_PATTERN As String = "\b\d{4}-\d{2}-\d{2}([T ]\d{2}:\d{2}(:\d{2})?)?\b"
Private Const META_WORDS As String = "author|creator|created|modified|last|revision|rsid|template|comment|company|manager|subject|title|keyword|category|user|owner|guid|uuid|url|path|mail|application|appversion|docsecurity|totaltime|contentstatus"

Private Type AuditReport
    FilePath As String
    IsZip As Boolean
    DocPropsCore As Collection
    DocPropsApp As Collection
    DocPropsCustom As Collection
    Relationships As Collection
    XmlFindings As Collection
    CryptoFindings As Collection
    GuessFindings As Collection
    PackageParts As Collection
    EmbeddedFiles As Collection
    StatusLegend As String
    ExtractedRoot As String
End Type

Private Type MetadataDashboardItem
    GroupKey As String
    DisplayName As String
    StorageLocation As String
    CurrentValue As String
    PrimaryAction As String
    SecondaryAction As String
    CapabilityNote As String
    IsPossiblyEncrypted As Boolean
End Type

Private Const METADATA_ROLE_VARIABLE As String = "MetadataToolkitRole"
Private Const METADATA_ROLE_RESULTS As String = "results"
Private Const DASHBOARD_ACTION_EDIT As String = "Available in Safe Edit Current"
Private Const DASHBOARD_ACTION_SAVE_OPEN As String = "Read-only finding"
Private Const DASHBOARD_ACTION_DECRYPT As String = "Further investigation available"
Private Const AUTHOR_CLEAR_MINIMAL As String = "minimal"
Private Const AUTHOR_CLEAR_NORMAL As String = "normal"
Private Const AUTHOR_CLEAR_BROAD As String = "broad"
Private Const WD_DIALOG_FILE_SUMMARY_INFO As Long = 86
Private Const MSO_PROPERTY_TYPE_NUMBER As Long = 1
Private Const MSO_PROPERTY_TYPE_BOOLEAN As Long = 2
Private Const MSO_PROPERTY_TYPE_DATE As Long = 3
Private Const MSO_PROPERTY_TYPE_STRING As Long = 4
Private Const MSO_PROPERTY_TYPE_FLOAT As Long = 5

Private mDashboardTargetPath As String
Private mDashboardForm As Object
Private mDashboardQuickReport As AuditReport
Private mDashboardHasQuickReport As Boolean
Private mDashboardMetadataFindings As Collection
Private mDashboardCryptoFindings As Collection
Private mDashboardGuessFindings As Collection
Private mDashboardMetadataInvestigated As Boolean
Private mDashboardCryptoInvestigated As Boolean
Private mDashboardShowSuiteMetadata As Boolean
Private mDashboardIndicatorStep As Long
Private mLastAuditSourcePath As String
Private mLastAuditReportText As String
Private mLastAuditTimestamp As Date
Private mSafeEditDocument As Object
Private mSafeEditWasSaved As Boolean
Private mSafeEditLinkedCustomCount As Long
Private mSafeEditManagedCustomCount As Long

Public Sub ShowMetaDataSuiteDashboard()
    On Error GoTo DashboardError
    BeginCleanupProgress "MetaDataSuite", "Ready"
    mDashboardTargetPath = PeekPreferredDocxTarget()
    ResetDashboardState
    If Not mDashboardForm Is Nothing Then
        Unload mDashboardForm
        Set mDashboardForm = Nothing
    End If
    Set mDashboardForm = New frmMetaDataSuite
    mDashboardForm.Show
    Exit Sub

DashboardError:
    EndCleanupProgress
    MsgBox "Unable to open MetaDataSuite: " & Err.Description, vbCritical, "MetaDataSuite"
End Sub

Public Sub ShowModernMetadataDashboard()
    ShowMetaDataSuiteDashboard
End Sub

Public Sub RunDocxMetadataAudit()
    Dim filePath As String

    filePath = ResolvePreferredDocxTarget(vbNullString, False)
    If filePath = vbNullString Then Exit Sub

    RunDocxMetadataAuditForFile filePath
End Sub

Public Sub RunDocxMetadataAuditAnotherDocument()
    Dim filePath As String

    filePath = PickDocxFile()
    If filePath = vbNullString Then Exit Sub

    RunDocxMetadataAuditForFile filePath
End Sub

Public Sub AuditCurrentDocumentFromLauncher()
    Dim filePath As String

    filePath = ResolvePreferredDocxTarget(mDashboardTargetPath, False)
    If filePath = vbNullString Then Exit Sub

    mDashboardTargetPath = filePath
    RunDocxMetadataAuditForFile filePath
End Sub

Public Sub SafeEditCurrentDocumentFromLauncher()
    Dim filePath As String

    filePath = ResolvePreferredDocxTarget(mDashboardTargetPath, False)
    If filePath = vbNullString Then Exit Sub

    mDashboardTargetPath = filePath
    RunDocxMetadataSafeEditorForFile filePath
End Sub

Public Sub GroupEditCurrentDocumentFromLauncher()
    Dim filePath As String

    filePath = ResolvePreferredDocxTarget(mDashboardTargetPath, False)
    If filePath = vbNullString Then Exit Sub

    mDashboardTargetPath = filePath
    RunDocxMetadataGroupEditorForFile filePath
End Sub

Public Sub CloseMetadataToolkitLauncher()
    EndCleanupProgress
End Sub

Public Sub CloseMetadataDashboard(ByVal frm As Object)
    On Error Resume Next
    Unload frm
    Set mDashboardForm = Nothing
    EndCleanupProgress
    On Error GoTo 0
End Sub

Public Sub InitializeMetadataDashboard(ByVal frm As Object)
    ConfigureDashboardDefaults frm
    If Len(mDashboardTargetPath) = 0 Then
        mDashboardTargetPath = PeekPreferredDocxTarget()
    End If

    If Len(mDashboardTargetPath) > 0 Then
        RefreshMetadataDashboardQuickScan frm
    Else
        UpdateDashboardDisplay frm
        SetDashboardOverallState frm, "WAIT", "Open or save a Word document, or choose Audit Another Document."
    End If
End Sub

Private Sub ConfigureDashboardChrome(ByVal frm As Object)
    If frm Is Nothing Then Exit Sub

    On Error Resume Next
    frm.Caption = ""
    frm.BackColor = RGB(248, 249, 251)
    frm.lblTitle.Caption = "MetaDataSuite"
    frm.lblTitle.Font.Bold = True
    frm.lblTitle.Font.Size = 16
    frm.lblTitle.ForeColor = RGB(32, 37, 45)
    frm.lblTitle.BackColor = RGB(248, 249, 251)

    frm.cmdRefreshCurrent.Caption = "Refresh Quick Checks"
    frm.cmdAuditAnother.Caption = "Audit Another Document"
    frm.cmdOpenWordReport.Caption = "Open Word Report"
    frm.cmdSafeEditCurrent.Caption = "Safe Edit Current"
    frm.cmdGroupEditCurrent.Caption = "Group Edit Current"
    frm.cmdClearSharingProperties.Caption = "Clear Sharing Properties"
    frm.cmdClose.Caption = "Close"
    frm.cmdExpandAll.Caption = "Expand All"
    frm.cmdMetadataInvestigate.Caption = "Investigate"
    frm.cmdCryptoInvestigate.Caption = "Investigate"
    frm.chkShowSuiteMetadata.Caption = "Show CleanupSuite metadata"

    frm.lblCoreHeader.Caption = "Core Properties"
    frm.lblAppHeader.Caption = "Application Properties"
    frm.lblCustomHeader.Caption = "Custom Properties"
    frm.lblPackageHeader.Caption = "Package / Embedded Items"
    frm.lblMetadataHeader.Caption = "Metadata Findings"
    frm.lblCryptoHeader.Caption = "Crypto / Guessing Findings"

    frm.chkCoreDetails.Caption = "Show details"
    frm.chkAppDetails.Caption = "Show details"
    frm.chkCustomDetails.Caption = "Show details"
    frm.chkPackageDetails.Caption = "Show details"
    frm.chkMetadataDetails.Caption = "Show details"
    frm.chkCryptoDetails.Caption = "Show details"

    frm.lblCoreIndicator.Caption = "WAIT"
    frm.lblAppIndicator.Caption = "WAIT"
    frm.lblCustomIndicator.Caption = "WAIT"
    frm.lblPackageIndicator.Caption = "WAIT"
    frm.lblMetadataIndicator.Caption = "WAIT"
    frm.lblCryptoIndicator.Caption = "WAIT"

    frm.lblCoreHeader.Font.Bold = True
    frm.lblAppHeader.Font.Bold = True
    frm.lblCustomHeader.Font.Bold = True
    frm.lblPackageHeader.Font.Bold = True
    frm.lblMetadataHeader.Font.Bold = True
    frm.lblCryptoHeader.Font.Bold = True

    StyleMetadataDashboardButton frm.cmdRefreshCurrent, False
    StyleMetadataDashboardButton frm.cmdAuditAnother, False
    StyleMetadataDashboardButton frm.cmdOpenWordReport, False
    StyleMetadataDashboardButton frm.cmdSafeEditCurrent, False
    StyleMetadataDashboardButton frm.cmdGroupEditCurrent, False
    StyleMetadataDashboardButton frm.cmdClearSharingProperties, False
    StyleMetadataDashboardButton frm.cmdClose, False
    StyleMetadataDashboardButton frm.cmdExpandAll, True
    StyleMetadataDashboardButton frm.cmdMetadataInvestigate, False
    StyleMetadataDashboardButton frm.cmdCryptoInvestigate, False

    StyleMetadataDashboardCheckBox frm.chkShowSuiteMetadata
    StyleMetadataDashboardCheckBox frm.chkCoreDetails
    StyleMetadataDashboardCheckBox frm.chkAppDetails
    StyleMetadataDashboardCheckBox frm.chkCustomDetails
    StyleMetadataDashboardCheckBox frm.chkPackageDetails
    StyleMetadataDashboardCheckBox frm.chkMetadataDetails
    StyleMetadataDashboardCheckBox frm.chkCryptoDetails

    StyleMetadataDashboardSectionHeader frm.lblCoreHeader
    StyleMetadataDashboardSectionHeader frm.lblAppHeader
    StyleMetadataDashboardSectionHeader frm.lblCustomHeader
    StyleMetadataDashboardSectionHeader frm.lblPackageHeader
    StyleMetadataDashboardSectionHeader frm.lblMetadataHeader
    StyleMetadataDashboardSectionHeader frm.lblCryptoHeader

    StyleMetadataDashboardSummary frm.lblCoreSummary
    StyleMetadataDashboardSummary frm.lblAppSummary
    StyleMetadataDashboardSummary frm.lblCustomSummary
    StyleMetadataDashboardSummary frm.lblPackageSummary
    StyleMetadataDashboardSummary frm.lblMetadataSummary
    StyleMetadataDashboardSummary frm.lblCryptoSummary

    StyleMetadataDashboardIndicator frm.lblCoreIndicator
    StyleMetadataDashboardIndicator frm.lblAppIndicator
    StyleMetadataDashboardIndicator frm.lblCustomIndicator
    StyleMetadataDashboardIndicator frm.lblPackageIndicator
    StyleMetadataDashboardIndicator frm.lblMetadataIndicator
    StyleMetadataDashboardIndicator frm.lblCryptoIndicator

    StyleMetadataDashboardDetails frm.txtCoreDetails
    StyleMetadataDashboardDetails frm.txtAppDetails
    StyleMetadataDashboardDetails frm.txtCustomDetails
    StyleMetadataDashboardDetails frm.txtPackageDetails
    StyleMetadataDashboardDetails frm.txtMetadataDetails
    StyleMetadataDashboardDetails frm.txtCryptoDetails
    On Error GoTo 0
End Sub

Private Sub StyleMetadataDashboardButton(ByVal ctl As Object, Optional ByVal primary As Boolean = False)
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
End Sub

Private Sub StyleMetadataDashboardCheckBox(ByVal ctl As Object)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 8.5
    ctl.Font.Bold = False
    ctl.ForeColor = RGB(32, 37, 45)
    ctl.BackColor = RGB(248, 249, 251)
End Sub

Private Sub StyleMetadataDashboardSectionHeader(ByVal ctl As Object)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 11
    ctl.Font.Bold = True
    ctl.ForeColor = RGB(32, 37, 45)
    ctl.BackColor = RGB(248, 249, 251)
End Sub

Private Sub StyleMetadataDashboardSummary(ByVal ctl As Object)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 8.5
    ctl.Font.Bold = False
    ctl.ForeColor = RGB(67, 80, 96)
    ctl.BackColor = RGB(248, 249, 251)
    ctl.WordWrap = True
End Sub

Private Sub StyleMetadataDashboardIndicator(ByVal ctl As Object)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 8.5
    ctl.Font.Bold = True
    ctl.ForeColor = RGB(88, 101, 116)
    ctl.BackColor = RGB(248, 249, 251)
End Sub

Private Sub StyleMetadataDashboardDetails(ByVal ctl As Object)
    On Error Resume Next
    ctl.Font.Name = "Segoe UI"
    ctl.Font.Size = 8.5
    ctl.Font.Bold = False
    ctl.ForeColor = RGB(67, 80, 96)
    ctl.BackColor = RGB(255, 255, 255)
    ctl.BorderStyle = 1
    ctl.SpecialEffect = 0
End Sub

Public Sub SetMetadataDashboardShowSuiteMetadata(ByVal frm As Object, ByVal showSuiteMetadata As Boolean)
    mDashboardShowSuiteMetadata = showSuiteMetadata
    If Not frm Is Nothing Then
        UpdateDashboardDisplay frm
        LayoutMetadataDashboard frm
    End If
End Sub

Private Sub ConfigureDashboardDetailTextBoxes(ByVal frm As Object)
    ConfigureDashboardDetailTextBox frm, "txtCoreDetails"
    ConfigureDashboardDetailTextBox frm, "txtAppDetails"
    ConfigureDashboardDetailTextBox frm, "txtCustomDetails"
    ConfigureDashboardDetailTextBox frm, "txtPackageDetails"
    ConfigureDashboardDetailTextBox frm, "txtMetadataDetails"
    ConfigureDashboardDetailTextBox frm, "txtCryptoDetails"
End Sub

Private Sub ConfigureDashboardDetailTextBox(ByVal frm As Object, ByVal controlName As String)
    Dim txt As Object

    If frm Is Nothing Then Exit Sub

    On Error Resume Next
    Set txt = CallByName(frm, controlName, VbGet)
    If txt Is Nothing Then Exit Sub

    txt.MultiLine = True
    txt.EnterKeyBehavior = True
    txt.Locked = True
    txt.WordWrap = True
    txt.ScrollBars = fmScrollBarsVertical
    txt.Visible = False
    On Error GoTo 0
End Sub

Private Sub HideLegacyAuthorClearPanel(ByVal frm As Object)
    If frm Is Nothing Then Exit Sub

    On Error Resume Next
    frm.lblAuthorClearBorderTop.Visible = False
    frm.lblAuthorClearBorderBottom.Visible = False
    frm.lblAuthorClearBorderLeft.Visible = False
    frm.lblAuthorClearBorderRight.Visible = False
    frm.lblAuthorClearTitle.Visible = False
    frm.lblAuthorClearSubtitle.Visible = False

    frm.lblAuthorClearBorderTop.Move 0, 0, 0, 0
    frm.lblAuthorClearBorderBottom.Move 0, 0, 0, 0
    frm.lblAuthorClearBorderLeft.Move 0, 0, 0, 0
    frm.lblAuthorClearBorderRight.Move 0, 0, 0, 0
    frm.lblAuthorClearTitle.Move 0, 0, 0, 0
    frm.lblAuthorClearSubtitle.Move 0, 0, 0, 0

    frm.cmdClearAuthorMinimal.Visible = False
    frm.cmdClearAuthorNormal.Visible = False
    frm.cmdClearAuthorBroad.Visible = False
    frm.cmdClearAuthorMinimal.Move 0, 0, 0, 0
    frm.cmdClearAuthorNormal.Move 0, 0, 0, 0
    frm.cmdClearAuthorBroad.Move 0, 0, 0, 0
    On Error GoTo 0
End Sub

Public Sub RefreshMetadataDashboardQuickScan(ByVal frm As Object)
    Dim filePath As String

    filePath = ResolvePreferredDocxTarget(vbNullString, False)
    If Len(filePath) = 0 And Len(mDashboardTargetPath) > 0 Then
        If FileExists(mDashboardTargetPath) And IsSupportedDocxPath(mDashboardTargetPath) Then
            filePath = mDashboardTargetPath
        End If
    End If
    If Len(filePath) = 0 Then
        mDashboardTargetPath = vbNullString
        ResetDashboardState
        UpdateDashboardDisplay frm
        SetDashboardOverallState frm, "WAIT", "No saved current document is ready yet."
        Exit Sub
    End If

    mDashboardTargetPath = filePath
    ResetDashboardState
    UpdateDashboardDisplay frm
    SetDashboardOverallState frm, "RUN", "Running quick checks..."

    QuickAuditDocx filePath, mDashboardQuickReport, frm
    mDashboardHasQuickReport = True
    UpdateDashboardDisplay frm
    SetDashboardOverallState frm, "DONE", "Quick checks complete."
End Sub

Public Sub PickMetadataDashboardTarget(ByVal frm As Object)
    Dim filePath As String

    filePath = PickDocxFile()
    If Len(filePath) = 0 Then Exit Sub

    mDashboardTargetPath = filePath
    RefreshMetadataDashboardQuickScan frm
End Sub

Public Sub ApplyMetadataDashboardVisibility(ByVal frm As Object)
    On Error Resume Next
    frm.txtCoreDetails.Visible = frm.chkCoreDetails.Value
    frm.txtAppDetails.Visible = frm.chkAppDetails.Value
    frm.txtCustomDetails.Visible = frm.chkCustomDetails.Value
    frm.txtPackageDetails.Visible = frm.chkPackageDetails.Value
    frm.txtMetadataDetails.Visible = frm.chkMetadataDetails.Value
    frm.txtCryptoDetails.Visible = frm.chkCryptoDetails.Value
    LayoutMetadataDashboard frm
    frm.Repaint
    On Error GoTo 0
End Sub

Public Sub InvestigateMetadataDashboardGroup(ByVal frm As Object, ByVal groupKey As String)
    Dim filePath As String

    filePath = mDashboardTargetPath
    If Len(filePath) = 0 Then
        filePath = ResolvePreferredDocxTarget(vbNullString, False)
        If Len(filePath) = 0 Then Exit Sub
        mDashboardTargetPath = filePath
    End If

    Select Case LCase$(groupKey)
        Case "metadata"
            SetDashboardOverallState frm, "RUN", "Running deeper metadata checks..."
            SetDashboardGroupState frm, "metadata", "RUN", "Investigating metadata findings..."
            Set mDashboardMetadataFindings = ExtractMetadataFindings(filePath, frm, "metadata")
            mDashboardMetadataInvestigated = True
            SetDashboardGroupState frm, "metadata", "DONE", "Metadata investigation complete."
        Case "crypto"
            SetDashboardOverallState frm, "RUN", "Running crypto and guessing checks..."
            SetDashboardGroupState frm, "crypto", "RUN", "Investigating crypto and guessing patterns..."
            ExtractCryptoGuessFindings filePath, mDashboardCryptoFindings, mDashboardGuessFindings, frm, "crypto"
            mDashboardCryptoInvestigated = True
            SetDashboardGroupState frm, "crypto", "DONE", "Crypto investigation complete."
        Case Else
            Exit Sub
    End Select

    UpdateDashboardDisplay frm
    SetDashboardOverallState frm, "DONE", "Requested investigation finished."
End Sub

Public Sub OpenMetadataDashboardWordReport(ByVal frm As Object)
    Dim filePath As String

    filePath = mDashboardTargetPath
    If Len(filePath) = 0 Then
        filePath = ResolvePreferredDocxTarget(vbNullString, False)
        If Len(filePath) = 0 Then Exit Sub
        mDashboardTargetPath = filePath
    End If

    SetDashboardOverallState frm, "RUN", "Opening full Word report..."
    RunDocxMetadataAuditForFile filePath
    SetDashboardOverallState frm, "DONE", "Full Word report opened."
End Sub

Public Sub ExpandMetadataDashboardAllGroups(ByVal frm As Object)
    If frm Is Nothing Then Exit Sub

    On Error Resume Next
    frm.chkCoreDetails.Value = True
    frm.chkAppDetails.Value = True
    frm.chkCustomDetails.Value = True
    frm.chkPackageDetails.Value = True
    frm.chkMetadataDetails.Value = True
    frm.chkCryptoDetails.Value = True
    ApplyMetadataDashboardVisibility frm
    On Error GoTo 0
End Sub

Public Sub ShowAuthorMetadataClearWorkflow(ByVal frm As Object, ByVal removalLevel As String)
    Dim filePath As String
    Dim normalizedLevel As String
    Dim previewText As String
    Dim answer As VbMsgBoxResult
    Dim broadAnswer As String

    normalizedLevel = NormalizeAuthorMetadataRemovalLevel(removalLevel)
    If Len(normalizedLevel) = 0 Then Exit Sub

    filePath = mDashboardTargetPath
    If Len(filePath) = 0 Then
        filePath = ResolvePreferredDocxTarget(vbNullString, False)
        If Len(filePath) = 0 Then Exit Sub
        mDashboardTargetPath = filePath
    End If

    previewText = BuildAuthorMetadataRemovalPreview(filePath, normalizedLevel)
    answer = MsgBox("The " & UCase$(Left$(normalizedLevel, 1)) & Mid$(normalizedLevel, 2) & _
                    " author metadata cleanup would delete or clear:" & vbCrLf & vbCrLf & _
                    previewText & vbCrLf & vbCrLf & _
                    "Are you sure you want all of this deleted?", _
                    vbYesNo + vbQuestion + vbDefaultButton2, "Clear Author Metadata")
    If answer <> vbYes Then Exit Sub

    If normalizedLevel = AUTHOR_CLEAR_BROAD Then
        broadAnswer = InputBox("Broad cleanup is more likely to remove metadata you may still want." & vbCrLf & vbCrLf & _
                               "Type [y]es to continue, or press Enter for [N]o.", _
                               "Confirm Broad Author Metadata Cleanup", "N")
        broadAnswer = LCase$(Trim$(broadAnswer))
        If broadAnswer <> "y" And broadAnswer <> "yes" Then Exit Sub
    End If

    ApplyAuthorMetadataRemoval filePath, normalizedLevel
    ResetDashboardState
    RefreshMetadataDashboardQuickScan frm
    MsgBox "Selected author metadata cleanup has been applied and the document was saved.", _
           vbInformation, "Clear Author Metadata"
End Sub

Public Function PreviewFinalReviewAuthorCleanup(ByVal removalLevel As String) As String
    Dim filePath As String
    filePath = ResolvePreferredDocxTarget(vbNullString, False)
    If Len(filePath) = 0 Then
        PreviewFinalReviewAuthorCleanup = "No saved current document is available for author cleanup."
        Exit Function
    End If
    PreviewFinalReviewAuthorCleanup = BuildAuthorMetadataRemovalPreview(filePath, NormalizeAuthorMetadataRemovalLevel(removalLevel))
End Function

Public Sub RunFinalReviewAuthorCleanup(ByVal removalLevel As String)
    Dim filePath As String
    Dim normalizedLevel As String

    normalizedLevel = NormalizeAuthorMetadataRemovalLevel(removalLevel)
    If Len(normalizedLevel) = 0 Then Exit Sub

    filePath = ResolvePreferredDocxTarget(vbNullString, False)
    If Len(filePath) = 0 Then
        MsgBox "No saved current document is available for author cleanup.", vbInformation, "Final Review"
        Exit Sub
    End If

    ApplyAuthorMetadataRemoval filePath, normalizedLevel
End Sub

Public Sub SafeEditMetadataDashboardTarget(ByVal frm As Object)
    Dim filePath As String

    filePath = ResolvePreferredDocxTarget(mDashboardTargetPath, False)
    If Len(filePath) = 0 Then Exit Sub

    mDashboardTargetPath = filePath
    RunDocxMetadataSafeEditorForFile filePath
End Sub

Public Sub GroupEditMetadataDashboardTarget(ByVal frm As Object)
    Dim filePath As String

    filePath = ResolvePreferredDocxTarget(mDashboardTargetPath, False)
    If Len(filePath) = 0 Then Exit Sub

    mDashboardTargetPath = filePath
    RunDocxMetadataGroupEditorForFile filePath
End Sub

Public Sub RunMetadataDashboardSmokeTest(Optional ByVal explicitTargetPath As String = "", _
                                         Optional ByVal outputPath As String = "", _
                                         Optional ByVal runMetadataInvestigation As Boolean = False, _
                                         Optional ByVal runCryptoInvestigation As Boolean = False)
    Dim filePath As String
    Dim finalOutputPath As String
    Dim frm As Object
    Dim smokeStage As String

    On Error GoTo SmokeTestError

    smokeStage = "resolve-target"
    If Len(Trim$(explicitTargetPath)) > 0 Then
        filePath = Trim$(explicitTargetPath)
    Else
        filePath = ResolvePreferredDocxTarget(vbNullString, False)
    End If

    If Len(filePath) = 0 Then
        Err.Raise vbObjectError + 1200, "RunMetadataDashboardSmokeTest", "No target document was available for the smoke test."
    End If

    finalOutputPath = Trim$(outputPath)
    If Len(finalOutputPath) = 0 Then
        finalOutputPath = CreateTempFolder() & "\metadata-dashboard-smoke.txt"
    End If

    smokeStage = "reset-state"
    mDashboardTargetPath = filePath
    ResetDashboardState
    smokeStage = "create-form"
    Set frm = New frmMetaDataSuite
    smokeStage = "configure-form"
    ConfigureDashboardDefaults frm
    smokeStage = "quick-scan"
    RefreshMetadataDashboardQuickScan frm

    If runMetadataInvestigation Then
        smokeStage = "metadata-investigation"
        InvestigateMetadataDashboardGroup frm, "metadata"
    End If
    If runCryptoInvestigation Then
        smokeStage = "crypto-investigation"
        InvestigateMetadataDashboardGroup frm, "crypto"
    End If

    smokeStage = "write-output"
    WriteTextFile finalOutputPath, BuildMetadataDashboardSmokeText(filePath)
    smokeStage = "close-form"
    CloseMetadataDashboard frm
    Exit Sub

SmokeTestError:
    On Error Resume Next
    If Not frm Is Nothing Then CloseMetadataDashboard frm
    If Len(Trim$(outputPath)) > 0 Then
        WriteTextFile outputPath, "ERROR [" & smokeStage & "] " & CStr(Err.Number) & " | " & Err.Source & " | " & Err.Description
    End If
    On Error GoTo 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub RefreshLastAuditResult()
    If Len(mLastAuditSourcePath) = 0 Then
        MsgBox "Run an audit first so there is a document to refresh.", vbInformation, "DOCX Metadata Audit"
        Exit Sub
    End If

    RunDocxMetadataAuditForFile mLastAuditSourcePath
End Sub

Public Sub SaveLastAuditReportAsText()
    Dim outputPath As String

    If Len(mLastAuditSourcePath) = 0 Or Len(mLastAuditReportText) = 0 Then
        MsgBox "There is no in-memory audit report to export yet.", vbInformation, "DOCX Metadata Audit"
        Exit Sub
    End If

    outputPath = InputBox("Choose where to save the text report.", "Save Audit Report As Text", GetOutputPath(mLastAuditSourcePath))
    If Len(Trim$(outputPath)) = 0 Then Exit Sub

    WriteTextFile outputPath, mLastAuditReportText
    MsgBox "Audit report saved to:" & vbCrLf & outputPath, vbInformation, "DOCX Metadata Audit"
End Sub

Public Sub SafeEditLastAuditedDocument()
    If Len(mLastAuditSourcePath) > 0 And FileExists(mLastAuditSourcePath) Then
        RunDocxMetadataSafeEditorForFile mLastAuditSourcePath
    Else
        RunDocxMetadataSafeEditor
    End If
End Sub

Public Sub GroupEditLastAuditedDocument()
    If Len(mLastAuditSourcePath) > 0 And FileExists(mLastAuditSourcePath) Then
        RunDocxMetadataGroupEditorForFile mLastAuditSourcePath
    Else
        RunDocxMetadataGroupEditor
    End If
End Sub

Private Sub RunDocxMetadataAuditForFile(ByVal filePath As String)
    Dim report As AuditReport
    Dim reportText As String

    On Error GoTo AuditError
    BeginCleanupProgress "MetaDataSuite", "Preparing DOCX audit"

    report = AuditDocx(filePath)
    reportText = BuildReportText(report)
    mLastAuditSourcePath = filePath
    mLastAuditReportText = reportText
    mLastAuditTimestamp = Now

    RenderAuditReportInWord report, reportText

    UpdateCleanupProgress "Audit finished"
    MsgBox "Audit complete." & vbCrLf & vbCrLf & _
           "Results are now open in Word." & vbCrLf & _
           "Use the Save Audit Report As Text button there if you need a file export.", _
           vbInformation, "DOCX Metadata Audit"

    EndCleanupProgress
    Debug.Print reportText
    Exit Sub

AuditError:
    EndCleanupProgress
    MsgBox "Audit failed: " & Err.Description, vbCritical, "DOCX Metadata Audit"
End Sub

Public Sub RunDocxMetadataEditor()
    Dim filePath As String
    Dim tempDir As String
    Dim result As VbMsgBoxResult

    filePath = PickDocxFile()
    If filePath = vbNullString Then Exit Sub

    tempDir = CreateTempFolder()
    If Not UnzipDocx(filePath, tempDir) Then
        MsgBox "Unable to extract the selected DOCX file.", vbExclamation, "DOCX Metadata Editor"
        Exit Sub
    End If

    Shell "explorer.exe """ & tempDir & """", vbNormalFocus
    result = MsgBox("The DOCX package has been extracted. Edit the XML/property files as needed, then click Yes to rebuild the document.", _
                   vbYesNo + vbQuestion, "DOCX Metadata Editor")

    If result = vbYes Then
        If RepackDocx(tempDir, filePath) Then
            MsgBox "The edited DOCX has been rebuilt successfully." & vbCrLf & vbCrLf & _
                   "Updated file:" & vbCrLf & filePath, vbInformation, "DOCX Metadata Editor"
        Else
            MsgBox "The DOCX could not be rebuilt. Please check the extracted package and try again.", vbExclamation, "DOCX Metadata Editor"
        End If
    End If

    On Error Resume Next
    DeleteFolder tempDir
    On Error GoTo 0
End Sub

Public Sub RunDocxMetadataGroupEditor()
    Dim filePath As String
    filePath = ResolvePreferredDocxTarget(vbNullString, False)
    If filePath = vbNullString Then Exit Sub
    RunDocxMetadataGroupEditorForFile filePath
End Sub

Private Sub RunDocxMetadataGroupEditorForFile(ByVal filePath As String)
    Dim tempDir As String

    tempDir = CreateTempFolder()
    If Not UnzipDocx(filePath, tempDir) Then
        MsgBox "Unable to extract the selected DOCX file.", vbExclamation, "DOCX Group Metadata Editor"
        Exit Sub
    End If

    If IsOpenDocumentPath(filePath) Or IsRunningMacroHostPath(filePath) Then
        RunPackageMetadataGroupViewOnlyWorkflow tempDir, filePath, _
            "Group Edit Current is view-only for open documents."
    Else
        RunPackageMetadataGroupWorkflow tempDir, filePath
    End If
End Sub

Public Sub RunDocxMetadataSafeEditor()
    Dim filePath As String
    filePath = ResolvePreferredDocxTarget(vbNullString, False)
    If filePath = vbNullString Then Exit Sub
    RunDocxMetadataSafeEditorForFile filePath
End Sub

Private Sub RunDocxMetadataSafeEditorForFile(ByVal filePath As String)
    Dim backupPath As String

    backupPath = GetBackupPath(filePath)
    If CopyFile(filePath, backupPath) Then
        MsgBox "Backup created before editing:" & vbCrLf & backupPath, vbInformation, "Safe Metadata Editor"
    Else
        MsgBox "Backup could not be created, but you can still continue.", vbExclamation, "Safe Metadata Editor"
    End If

    RunWordSafeMetadataDialog filePath
End Sub

Private Sub ApplyCorePropertyUpdates(ByVal doc As Object, ByVal groupChoice As String)
    Dim payload As String
    Dim pairs() As String
    Dim pair() As String
    Dim i As Long

    payload = InputBox("Enter core metadata updates as Field=Value pairs separated by semicolons." & vbCrLf & _
                       "Example: Title=My Title; Subject=Quarterly; Author=Chris", _
                       "Core Metadata", "")
    If Len(Trim$(payload)) = 0 Then Exit Sub

    pairs = Split(payload, ";")
    For i = LBound(pairs) To UBound(pairs)
        pair = Split(Trim$(pairs(i)), "=", 2)
        If UBound(pair) >= 1 Then
            Select Case LCase$(Trim$(pair(0)))
                Case "title"
                    doc.BuiltInDocumentProperties("Title") = Trim$(pair(1))
                Case "subject"
                    doc.BuiltInDocumentProperties("Subject") = Trim$(pair(1))
                Case "author"
                    doc.BuiltInDocumentProperties("Author") = Trim$(pair(1))
                Case "keywords"
                    doc.BuiltInDocumentProperties("Keywords") = Trim$(pair(1))
                Case "comments"
                    doc.BuiltInDocumentProperties("Comments") = Trim$(pair(1))
                Case "category"
                    doc.BuiltInDocumentProperties("Category") = Trim$(pair(1))
                Case "company"
                    doc.BuiltInDocumentProperties("Company") = Trim$(pair(1))
                Case "manager"
                    doc.BuiltInDocumentProperties("Manager") = Trim$(pair(1))
                Case "lastauthor"
                    doc.BuiltInDocumentProperties("Last Author") = Trim$(pair(1))
            End Select
        End If
    Next i
End Sub

Private Sub ApplyAppPropertyUpdates(ByVal doc As Object, ByVal groupChoice As String)
    Dim payload As String
    Dim pairs() As String
    Dim pair() As String
    Dim i As Long

    payload = InputBox("Enter app metadata updates as Field=Value pairs separated by semicolons." & vbCrLf & _
                       "Example: Application=Microsoft Word; Company=ACME; Revision=3", _
                       "App Metadata", "")
    If Len(Trim$(payload)) = 0 Then Exit Sub

    pairs = Split(payload, ";")
    For i = LBound(pairs) To UBound(pairs)
        pair = Split(Trim$(pairs(i)), "=", 2)
        If UBound(pair) >= 1 Then
            Select Case LCase$(Trim$(pair(0)))
                Case "application"
                    doc.BuiltInDocumentProperties("Application") = Trim$(pair(1))
                Case "company"
                    doc.BuiltInDocumentProperties("Company") = Trim$(pair(1))
                Case "manager"
                    doc.BuiltInDocumentProperties("Manager") = Trim$(pair(1))
                Case "revision"
                    doc.BuiltInDocumentProperties("Revision Number") = Trim$(pair(1))
                Case "docsecurity"
                    doc.BuiltInDocumentProperties("Security") = Trim$(pair(1))
            End Select
        End If
    Next i
End Sub

Private Sub ApplyCustomPropertyUpdates(ByVal doc As Object)
    Dim payload As String
    Dim pairs() As String
    Dim pair() As String
    Dim i As Long
    Dim prop As Object

    payload = InputBox("Enter custom properties as Name=Value pairs separated by semicolons." & vbCrLf & _
                       "Example: Project=Alpha; Owner=Chris", _
                       "Custom Metadata", "")
    If Len(Trim$(payload)) = 0 Then Exit Sub

    pairs = Split(payload, ";")
    For i = LBound(pairs) To UBound(pairs)
        pair = Split(Trim$(pairs(i)), "=", 2)
        If UBound(pair) >= 1 Then
            On Error Resume Next
            Set prop = doc.CustomDocumentProperties(Trim$(pair(0)))
            If Err.Number <> 0 Then
                doc.CustomDocumentProperties.Add Name:=Trim$(pair(0)), LinkToContent:=False, Type:=1, Value:=Trim$(pair(1))
            Else
                prop.Value = Trim$(pair(1))
            End If
            On Error GoTo 0
        End If
    Next i
End Sub

Private Function GetBackupPath(ByVal filePath As String) As String
    Dim fso As Object
    Dim baseName As String
    Dim ext As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    baseName = fso.GetBaseName(filePath)
    ext = fso.GetExtensionName(filePath)
    GetBackupPath = fso.GetParentFolderName(filePath) & "\" & baseName & _
                    ".metadata-backup-" & Format$(Now, "yyyymmdd-hhnnss") & "." & ext
End Function

Private Function CopyFile(ByVal sourcePath As String, ByVal targetPath As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(targetPath) Then fso.DeleteFile targetPath, True
    fso.CopyFile sourcePath, targetPath, True
    CopyFile = fso.FileExists(targetPath)
End Function

Private Sub RunWordSafeMetadataDialog(ByVal filePath As String)
    Dim doc As Object
    Dim closeWhenDone As Boolean
    Dim frm As frmSafeMetadataEditor
    Dim originalErrNumber As Long
    Dim originalErrDescription As String

    On Error GoTo SafeEditorError

    Set doc = ResolveDocumentForMetadataEdit(filePath, closeWhenDone, False)
    If doc Is Nothing Then Exit Sub

    doc.Activate
    Set mSafeEditDocument = doc
    mSafeEditWasSaved = False
    Set frm = New frmSafeMetadataEditor
    frm.Show

    If closeWhenDone Then
        On Error Resume Next
        doc.Close mSafeEditWasSaved
        On Error GoTo SafeEditorError
    End If
    Set frm = Nothing
    Set mSafeEditDocument = Nothing
    Exit Sub

SafeEditorError:
    originalErrNumber = Err.Number
    originalErrDescription = Err.Description
    If closeWhenDone Then
        On Error Resume Next
        If Not doc Is Nothing Then doc.Close False
        On Error GoTo 0
    End If
    Set frm = Nothing
    Set mSafeEditDocument = Nothing
    MsgBox "The safe editor encountered an error (" & originalErrNumber & "): " & _
           originalErrDescription, vbCritical, "Safe Metadata Editor"
End Sub

Public Sub InitializeSafeMetadataEditor(ByVal frm As Object)
    If frm Is Nothing Then Exit Sub
    If mSafeEditDocument Is Nothing Then
        MsgBox "The document for safe metadata editing is no longer available.", _
               vbExclamation, "Safe Metadata Editor"
        Unload frm
        Exit Sub
    End If

    ConfigureSafeMetadataEditorChrome frm
    LoadSafeBuiltInProperty frm.txtTitle, "Title"
    LoadSafeBuiltInProperty frm.txtSubject, "Subject"
    LoadSafeBuiltInProperty frm.txtAuthor, "Author"
    LoadSafeBuiltInProperty frm.txtKeywords, "Keywords"
    LoadSafeBuiltInProperty frm.txtComments, "Comments"
    LoadSafeBuiltInProperty frm.txtCategory, "Category"
    LoadSafeBuiltInProperty frm.txtCompany, "Company"
    LoadSafeBuiltInProperty frm.txtManager, "Manager"
    LoadSafeBuiltInProperty frm.txtLastAuthor, "Last Author"
    LoadSafeBuiltInProperty frm.txtContentStatus, "Content Status"
    LoadSafeBuiltInProperty frm.txtContentType, "Content Type"
    LoadSafeBuiltInProperty frm.txtLanguage, "Language"
    LoadSafeCustomProperties frm
End Sub

Private Sub ConfigureSafeMetadataEditorChrome(ByVal frm As Object)
    Const SAFE_EDITOR_W As Single = 760
    Const SAFE_EDITOR_H As Single = 510

    On Error Resume Next
    frm.Caption = "Safe Metadata Editor"
    ' Keep this inside the launcher's proven 1024 x 768 screen envelope.
    ' These are VBA form points, not screen pixels.
    frm.Width = SAFE_EDITOR_W
    frm.Height = SAFE_EDITOR_H
    frm.BackColor = RGB(248, 249, 251)
    frm.ScrollBars = 0

    ConfigureSafeEditorHeading frm.lblEditorTitle, "Safe Metadata Editor", 24, 10, 700, 24, 18
    ConfigureSafeEditorTextLabel frm.lblDocumentPath, mSafeEditDocument.FullName, 24, 34, 700, 16
    frm.lblDocumentPath.ForeColor = RGB(88, 101, 116)

    ConfigureSafeEditorSection frm.lblSummaryHeader, "Document summary", 24, 54, 700, 18
    ConfigureSafeEditorField frm.lblTitle, frm.txtTitle, "Title", 24, 76, 340
    ConfigureSafeEditorField frm.lblCategory, frm.txtCategory, "Category", 386, 76, 338
    ConfigureSafeEditorField frm.lblSubject, frm.txtSubject, "Subject", 24, 102, 340
    ConfigureSafeEditorField frm.lblCompany, frm.txtCompany, "Company", 386, 102, 338
    ConfigureSafeEditorField frm.lblAuthor, frm.txtAuthor, "Author", 24, 128, 340
    ConfigureSafeEditorField frm.lblManager, frm.txtManager, "Manager", 386, 128, 338
    ConfigureSafeEditorField frm.lblKeywords, frm.txtKeywords, "Keywords", 24, 154, 700
    ConfigureSafeEditorField frm.lblComments, frm.txtComments, "Comments", 24, 180, 700
    frm.txtComments.Height = 36
    frm.txtComments.MultiLine = True
    frm.txtComments.EnterKeyBehavior = True
    frm.txtComments.ScrollBars = 2

    ConfigureSafeEditorSection frm.lblStatusHeader, "Attribution and classification", 24, 222, 700, 18
    ConfigureSafeEditorField frm.lblLastAuthor, frm.txtLastAuthor, "Last author", 24, 244, 340
    ConfigureSafeEditorField frm.lblContentStatus, frm.txtContentStatus, "Content status", 386, 244, 338
    ConfigureSafeEditorField frm.lblContentType, frm.txtContentType, "Content type", 24, 270, 340
    ConfigureSafeEditorField frm.lblLanguage, frm.txtLanguage, "Language", 386, 270, 338

    ConfigureSafeEditorSection frm.lblCustomHeader, "Custom document properties", 24, 296, 700, 18
    ConfigureSafeEditorTextLabel frm.lblCustomHint, _
        "Unlinked user properties are editable here. Linked and CleanupSuite-managed properties remain audit-only.", _
        24, 314, 700, 24
    frm.lblCustomHint.ForeColor = RGB(88, 101, 116)
    frm.lblCustomHint.WordWrap = True
    ConfigureSafeEditorTextLabel frm.lblCustomNameHeader, "Name", 24, 340, 150, 14
    ConfigureSafeEditorTextLabel frm.lblCustomTypeHeader, "Type", 179, 340, 70, 14
    ConfigureSafeEditorTextLabel frm.lblCustomValueHeader, "Value", 254, 340, 300, 14

    With frm.lstCustomProperties
        .Move 24, 356, 530, 56
        .ColumnCount = 5
        .ColumnWidths = "150 pt;70 pt;300 pt;0 pt;0 pt"
        .IntegralHeight = False
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
    End With

    ConfigureSafeEditorField frm.lblCustomName, frm.txtCustomName, "Name", 572, 354, 152
    ConfigureSafeEditorTextLabel frm.lblCustomType, "Type", 572, 378, 42, 18
    With frm.cboCustomType
        .Move 618, 376, 106, 22
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Clear
        .AddItem "Text"
        .AddItem "Number"
        .AddItem "Decimal"
        .AddItem "Yes/No"
        .AddItem "Date"
        .ListIndex = 0
        .Style = 2
    End With
    ConfigureSafeEditorField frm.lblCustomValue, frm.txtCustomValue, "Value", 572, 402, 152

    ConfigureSafeEditorButton frm.cmdCustomAddUpdate, "Add / Update", 572, 424, 96, 22, False
    ConfigureSafeEditorButton frm.cmdCustomDelete, "Delete", 674, 424, 50, 22, False
    ConfigureSafeEditorTextLabel frm.lblAuditOnlyNote, _
        "Audit-only: timestamps, revision counters, template and hyperlink base, security/signatures, package relationships, linked properties, and application-generated statistics.", _
        24, 414, 530, 30
    frm.lblAuditOnlyNote.ForeColor = RGB(88, 101, 116)
    frm.lblAuditOnlyNote.WordWrap = True

    ConfigureSafeEditorButton frm.cmdSave, "Save safe metadata", 456, 448, 168, 28, True
    ConfigureSafeEditorButton frm.cmdCancel, "Cancel", 634, 448, 90, 28, False
    On Error GoTo 0
End Sub

Private Sub ConfigureSafeEditorHeading(ByVal ctl As Object, ByVal captionText As String, _
                                       ByVal leftValue As Single, ByVal topValue As Single, _
                                       ByVal widthValue As Single, ByVal heightValue As Single, _
                                       ByVal fontSize As Single)
    ConfigureSafeEditorTextLabel ctl, captionText, leftValue, topValue, widthValue, heightValue
    ctl.Font.Size = fontSize
    ctl.Font.Bold = True
    ctl.ForeColor = RGB(32, 37, 45)
End Sub

Private Sub ConfigureSafeEditorSection(ByVal ctl As Object, ByVal captionText As String, _
                                       ByVal leftValue As Single, ByVal topValue As Single, _
                                       ByVal widthValue As Single, ByVal heightValue As Single)
    ConfigureSafeEditorHeading ctl, captionText, leftValue, topValue, widthValue, heightValue, 11
    ctl.ForeColor = RGB(24, 90, 145)
End Sub

Private Sub ConfigureSafeEditorTextLabel(ByVal ctl As Object, ByVal captionText As String, _
                                         ByVal leftValue As Single, ByVal topValue As Single, _
                                         ByVal widthValue As Single, ByVal heightValue As Single)
    With ctl
        .Move leftValue, topValue, widthValue, heightValue
        .Caption = captionText
        .BackStyle = 0
        .Font.Name = "Segoe UI"
        .Font.Size = 8.5
        .ForeColor = RGB(32, 37, 45)
    End With
End Sub

Private Sub ConfigureSafeEditorField(ByVal labelControl As Object, ByVal textControl As Object, _
                                     ByVal captionText As String, ByVal leftValue As Single, _
                                     ByVal topValue As Single, ByVal widthValue As Single)
    ConfigureSafeEditorTextLabel labelControl, captionText, leftValue, topValue, 76, 18
    With textControl
        .Move leftValue + 80, topValue - 2, widthValue - 80, 22
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .BackColor = RGB(255, 255, 255)
        .ForeColor = RGB(32, 37, 45)
        .BorderStyle = 1
        .SpecialEffect = 0
    End With
End Sub

Private Sub ConfigureSafeEditorButton(ByVal ctl As Object, ByVal captionText As String, _
                                      ByVal leftValue As Single, ByVal topValue As Single, _
                                      ByVal widthValue As Single, ByVal heightValue As Single, _
                                      ByVal isPrimary As Boolean)
    With ctl
        .Move leftValue, topValue, widthValue, heightValue
        .Caption = captionText
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Bold = isPrimary
        .WordWrap = True
        .BackColor = IIf(isPrimary, RGB(221, 237, 250), RGB(255, 255, 255))
        .ForeColor = IIf(isPrimary, RGB(24, 90, 145), RGB(32, 37, 45))
    End With
End Sub

Private Sub LoadSafeBuiltInProperty(ByVal ctl As Object, ByVal propertyName As String)
    Dim valueText As String
    Dim isAvailable As Boolean

    valueText = ReadSafeBuiltInProperty(propertyName, isAvailable)
    ctl.Text = valueText
    ctl.Enabled = isAvailable
    If Not isAvailable Then
        ctl.Text = "(not available in this file format)"
        ctl.BackColor = RGB(238, 240, 243)
    End If
End Sub

Private Function ReadSafeBuiltInProperty(ByVal propertyName As String, ByRef isAvailable As Boolean) As String
    Dim rawValue As Variant

    On Error GoTo Unavailable
    rawValue = mSafeEditDocument.BuiltInDocumentProperties(propertyName).Value
    isAvailable = True
    If IsNull(rawValue) Then
        ReadSafeBuiltInProperty = vbNullString
    Else
        ReadSafeBuiltInProperty = CStr(rawValue)
    End If
    Exit Function

Unavailable:
    Err.Clear
    isAvailable = False
    ReadSafeBuiltInProperty = vbNullString
End Function

Private Sub LoadSafeCustomProperties(ByVal frm As Object)
    Dim prop As Object
    Dim rowIndex As Long

    mSafeEditLinkedCustomCount = 0
    mSafeEditManagedCustomCount = 0
    frm.lstCustomProperties.Clear

    On Error Resume Next
    For Each prop In mSafeEditDocument.CustomDocumentProperties
        If IsCleanupSuiteManagedProperty(CStr(prop.Name)) Then
            mSafeEditManagedCustomCount = mSafeEditManagedCustomCount + 1
        ElseIf SafeCustomPropertyIsLinked(prop) Then
            mSafeEditLinkedCustomCount = mSafeEditLinkedCustomCount + 1
        Else
            frm.lstCustomProperties.AddItem CStr(prop.Name)
            rowIndex = frm.lstCustomProperties.ListCount - 1
            frm.lstCustomProperties.List(rowIndex, 1) = SafeCustomTypeName(CLng(prop.Type))
            frm.lstCustomProperties.List(rowIndex, 2) = CStr(prop.Value)
            frm.lstCustomProperties.List(rowIndex, 3) = CStr(prop.Name)
            frm.lstCustomProperties.List(rowIndex, 4) = CStr(prop.Type)
        End If
    Next prop
    On Error GoTo 0

    frm.lblCustomHint.Caption = "Unlinked user properties are editable here. " & _
        CStr(mSafeEditLinkedCustomCount) & " linked and " & CStr(mSafeEditManagedCustomCount) & _
        " CleanupSuite-managed properties remain audit-only."
End Sub

Public Sub LoadSelectedSafeCustomProperty(ByVal frm As Object)
    Dim rowIndex As Long
    rowIndex = frm.lstCustomProperties.ListIndex
    If rowIndex < 0 Then Exit Sub

    frm.txtCustomName.Text = CStr(frm.lstCustomProperties.List(rowIndex, 0))
    frm.txtCustomValue.Text = CStr(frm.lstCustomProperties.List(rowIndex, 2))
    SelectSafeCustomType frm.cboCustomType, CStr(frm.lstCustomProperties.List(rowIndex, 1))
End Sub

Public Sub AddOrUpdateSafeCustomProperty(ByVal frm As Object)
    Dim propertyName As String
    Dim propertyValue As String
    Dim propertyType As String
    Dim rowIndex As Long
    Dim existingIndex As Long
    Dim originalName As String
    Dim originalType As String
    Dim validationMessage As String

    propertyName = Trim$(frm.txtCustomName.Text)
    propertyValue = frm.txtCustomValue.Text
    propertyType = CStr(frm.cboCustomType.Value)
    If Not ValidateSafeCustomProperty(propertyName, propertyType, propertyValue, validationMessage) Then
        MsgBox validationMessage, vbExclamation, "Safe Metadata Editor"
        Exit Sub
    End If
    If IsCleanupSuiteManagedProperty(propertyName) Then
        MsgBox "CleanupSuite-managed properties are operational settings and remain audit-only.", _
               vbExclamation, "Safe Metadata Editor"
        Exit Sub
    End If

    rowIndex = frm.lstCustomProperties.ListIndex
    existingIndex = FindSafeCustomRow(frm, propertyName)
    If existingIndex >= 0 And existingIndex <> rowIndex Then
        MsgBox "A custom property with that name is already listed.", vbExclamation, "Safe Metadata Editor"
        Exit Sub
    End If
    If SafeCustomNameConflictsWithProtectedProperty(propertyName) Then
        MsgBox "That name belongs to a linked or CleanupSuite-managed property and remains audit-only.", _
               vbExclamation, "Safe Metadata Editor"
        Exit Sub
    End If

    If rowIndex >= 0 Then
        originalName = CStr(frm.lstCustomProperties.List(rowIndex, 3))
        originalType = CStr(frm.lstCustomProperties.List(rowIndex, 4))
    Else
        frm.lstCustomProperties.AddItem propertyName
        rowIndex = frm.lstCustomProperties.ListCount - 1
    End If

    frm.lstCustomProperties.List(rowIndex, 0) = propertyName
    frm.lstCustomProperties.List(rowIndex, 1) = propertyType
    frm.lstCustomProperties.List(rowIndex, 2) = propertyValue
    frm.lstCustomProperties.List(rowIndex, 3) = originalName
    frm.lstCustomProperties.List(rowIndex, 4) = originalType
    ClearSafeCustomEditor frm
End Sub

Public Sub DeleteSelectedSafeCustomProperty(ByVal frm As Object)
    Dim rowIndex As Long
    rowIndex = frm.lstCustomProperties.ListIndex
    If rowIndex < 0 Then
        MsgBox "Select a custom property first.", vbInformation, "Safe Metadata Editor"
        Exit Sub
    End If

    frm.lstCustomProperties.RemoveItem rowIndex
    ClearSafeCustomEditor frm
End Sub

Private Sub ClearSafeCustomEditor(ByVal frm As Object)
    frm.lstCustomProperties.ListIndex = -1
    frm.txtCustomName.Text = vbNullString
    frm.txtCustomValue.Text = vbNullString
    frm.cboCustomType.ListIndex = 0
End Sub

Private Function FindSafeCustomRow(ByVal frm As Object, ByVal propertyName As String) As Long
    Dim i As Long
    FindSafeCustomRow = -1
    For i = 0 To frm.lstCustomProperties.ListCount - 1
        If StrComp(CStr(frm.lstCustomProperties.List(i, 0)), propertyName, vbTextCompare) = 0 Then
            FindSafeCustomRow = i
            Exit Function
        End If
    Next i
End Function

Private Sub SelectSafeCustomType(ByVal cbo As Object, ByVal typeName As String)
    Dim i As Long
    For i = 0 To cbo.ListCount - 1
        If StrComp(CStr(cbo.List(i)), typeName, vbTextCompare) = 0 Then
            cbo.ListIndex = i
            Exit Sub
        End If
    Next i
    cbo.ListIndex = 0
End Sub

Private Function SafeCustomTypeName(ByVal typeCode As Long) As String
    Select Case typeCode
        Case MSO_PROPERTY_TYPE_NUMBER: SafeCustomTypeName = "Number"
        Case MSO_PROPERTY_TYPE_BOOLEAN: SafeCustomTypeName = "Yes/No"
        Case MSO_PROPERTY_TYPE_DATE: SafeCustomTypeName = "Date"
        Case MSO_PROPERTY_TYPE_FLOAT: SafeCustomTypeName = "Decimal"
        Case Else: SafeCustomTypeName = "Text"
    End Select
End Function

Private Function SafeCustomTypeCode(ByVal typeName As String) As Long
    Select Case LCase$(Trim$(typeName))
        Case "number": SafeCustomTypeCode = MSO_PROPERTY_TYPE_NUMBER
        Case "yes/no": SafeCustomTypeCode = MSO_PROPERTY_TYPE_BOOLEAN
        Case "date": SafeCustomTypeCode = MSO_PROPERTY_TYPE_DATE
        Case "decimal": SafeCustomTypeCode = MSO_PROPERTY_TYPE_FLOAT
        Case Else: SafeCustomTypeCode = MSO_PROPERTY_TYPE_STRING
    End Select
End Function

Private Function SafeCustomTypedValue(ByVal typeName As String, ByVal valueText As String) As Variant
    Select Case SafeCustomTypeCode(typeName)
        Case MSO_PROPERTY_TYPE_NUMBER
            SafeCustomTypedValue = CLng(valueText)
        Case MSO_PROPERTY_TYPE_BOOLEAN
            SafeCustomTypedValue = (LCase$(Trim$(valueText)) = "true" Or _
                                    LCase$(Trim$(valueText)) = "yes" Or Trim$(valueText) = "1")
        Case MSO_PROPERTY_TYPE_DATE
            SafeCustomTypedValue = CDate(valueText)
        Case MSO_PROPERTY_TYPE_FLOAT
            SafeCustomTypedValue = CDbl(valueText)
        Case Else
            SafeCustomTypedValue = valueText
    End Select
End Function

Private Function ValidateSafeCustomProperty(ByVal propertyName As String, ByVal typeName As String, _
                                            ByVal valueText As String, ByRef validationMessage As String) As Boolean
    If Len(propertyName) = 0 Then
        validationMessage = "Enter a custom property name."
        Exit Function
    End If
    If Len(propertyName) > 255 Then
        validationMessage = "Custom property names cannot exceed 255 characters."
        Exit Function
    End If

    Select Case SafeCustomTypeCode(typeName)
        Case MSO_PROPERTY_TYPE_NUMBER
            If Not IsNumeric(valueText) Or CDbl(valueText) <> Fix(CDbl(valueText)) Then
                validationMessage = "Enter a whole number for a Number property."
                Exit Function
            End If
        Case MSO_PROPERTY_TYPE_FLOAT
            If Not IsNumeric(valueText) Then
                validationMessage = "Enter a numeric value for a Decimal property."
                Exit Function
            End If
        Case MSO_PROPERTY_TYPE_BOOLEAN
            Select Case LCase$(Trim$(valueText))
                Case "true", "false", "yes", "no", "1", "0"
                Case Else
                    validationMessage = "Enter Yes, No, True, False, 1, or 0 for a Yes/No property."
                    Exit Function
            End Select
        Case MSO_PROPERTY_TYPE_DATE
            If Not IsDate(valueText) Then
                validationMessage = "Enter a recognizable date for a Date property."
                Exit Function
            End If
    End Select

    ValidateSafeCustomProperty = True
End Function

Public Sub SaveSafeMetadataEditor(ByVal frm As Object)
    Dim validationMessage As String
    Dim errorText As String
    Dim i As Long
    Dim rowIsValid As Boolean

    If mSafeEditDocument Is Nothing Then
        MsgBox "The document for safe metadata editing is no longer available.", _
               vbExclamation, "Safe Metadata Editor"
        Exit Sub
    End If

    For i = 0 To frm.lstCustomProperties.ListCount - 1
        rowIsValid = ValidateSafeCustomProperty(CStr(frm.lstCustomProperties.List(i, 0)), _
                                                CStr(frm.lstCustomProperties.List(i, 1)), _
                                                CStr(frm.lstCustomProperties.List(i, 2)), validationMessage)
        If Not rowIsValid Then
            MsgBox "Custom property '" & CStr(frm.lstCustomProperties.List(i, 0)) & _
                   "' is not valid:" & vbCrLf & validationMessage, vbExclamation, "Safe Metadata Editor"
            Exit Sub
        End If
    Next i

    On Error GoTo SaveError
    ApplySafeBuiltInProperty frm.txtTitle, "Title", errorText
    ApplySafeBuiltInProperty frm.txtSubject, "Subject", errorText
    ApplySafeBuiltInProperty frm.txtAuthor, "Author", errorText
    ApplySafeBuiltInProperty frm.txtKeywords, "Keywords", errorText
    ApplySafeBuiltInProperty frm.txtComments, "Comments", errorText
    ApplySafeBuiltInProperty frm.txtCategory, "Category", errorText
    ApplySafeBuiltInProperty frm.txtCompany, "Company", errorText
    ApplySafeBuiltInProperty frm.txtManager, "Manager", errorText
    ApplySafeBuiltInProperty frm.txtLastAuthor, "Last Author", errorText
    ApplySafeBuiltInProperty frm.txtContentStatus, "Content Status", errorText
    ApplySafeBuiltInProperty frm.txtContentType, "Content Type", errorText
    ApplySafeBuiltInProperty frm.txtLanguage, "Language", errorText
    ApplySafeCustomPropertyChanges frm, errorText

    If Len(errorText) > 0 Then
        MsgBox "Some safe metadata fields could not be saved:" & vbCrLf & vbCrLf & errorText & vbCrLf & _
               "The pre-edit backup is available if you want to restore the original.", _
               vbExclamation, "Safe Metadata Editor"
        Exit Sub
    End If

    mSafeEditDocument.Save
    mSafeEditWasSaved = True
    MsgBox "All safe editable metadata fields were saved.", vbInformation, "Safe Metadata Editor"
    Unload frm
    Exit Sub

SaveError:
    MsgBox "The safe editor could not save the document (" & Err.Number & "): " & _
           Err.Description & vbCrLf & vbCrLf & _
           "The pre-edit backup is available if you want to restore the original.", _
           vbCritical, "Safe Metadata Editor"
End Sub

Private Sub ApplySafeBuiltInProperty(ByVal ctl As Object, ByVal propertyName As String, ByRef errorText As String)
    Dim originalErrNumber As Long
    Dim originalErrDescription As String
    If Not ctl.Enabled Then Exit Sub

    On Error Resume Next
    mSafeEditDocument.BuiltInDocumentProperties(propertyName).Value = ctl.Text
    If Err.Number <> 0 Then
        originalErrNumber = Err.Number
        originalErrDescription = Err.Description
        errorText = errorText & "- " & propertyName & " (" & originalErrNumber & "): " & _
                    originalErrDescription & vbCrLf
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Private Sub ApplySafeCustomPropertyChanges(ByVal frm As Object, ByRef errorText As String)
    Dim prop As Object
    Dim i As Long
    Dim propertyIndex As Long
    Dim originalName As String
    Dim originalType As Long
    Dim currentType As Long
    Dim typedValue As Variant

    On Error Resume Next
    For propertyIndex = mSafeEditDocument.CustomDocumentProperties.Count To 1 Step -1
        Set prop = mSafeEditDocument.CustomDocumentProperties(propertyIndex)
        If SafeCustomPropertyIsEditable(prop) Then
            If Not SafeCustomOriginalNameIsListed(frm, CStr(prop.Name)) Then prop.Delete
        End If
        Err.Clear
    Next propertyIndex
    On Error GoTo 0

    For i = 0 To frm.lstCustomProperties.ListCount - 1
        originalName = CStr(frm.lstCustomProperties.List(i, 3))
        originalType = 0
        If Len(CStr(frm.lstCustomProperties.List(i, 4))) > 0 Then _
            originalType = CLng(frm.lstCustomProperties.List(i, 4))
        currentType = SafeCustomTypeCode(CStr(frm.lstCustomProperties.List(i, 1)))
        typedValue = SafeCustomTypedValue(CStr(frm.lstCustomProperties.List(i, 1)), _
                                          CStr(frm.lstCustomProperties.List(i, 2)))

        On Error Resume Next
        Err.Clear
        If Len(originalName) > 0 Then
            If StrComp(originalName, CStr(frm.lstCustomProperties.List(i, 0)), vbTextCompare) = 0 Then
                If originalType = currentType Then
                    mSafeEditDocument.CustomDocumentProperties(originalName).Value = typedValue
                Else
                    mSafeEditDocument.CustomDocumentProperties(originalName).Delete
                    Err.Clear
                    mSafeEditDocument.CustomDocumentProperties.Add _
                        Name:=CStr(frm.lstCustomProperties.List(i, 0)), _
                        LinkToContent:=False, Type:=currentType, Value:=typedValue
                End If
            Else
                mSafeEditDocument.CustomDocumentProperties(originalName).Delete
                Err.Clear
                mSafeEditDocument.CustomDocumentProperties.Add _
                    Name:=CStr(frm.lstCustomProperties.List(i, 0)), _
                    LinkToContent:=False, Type:=currentType, Value:=typedValue
            End If
        Else
            mSafeEditDocument.CustomDocumentProperties.Add _
                Name:=CStr(frm.lstCustomProperties.List(i, 0)), _
                LinkToContent:=False, Type:=currentType, Value:=typedValue
        End If
        If Err.Number <> 0 Then
            errorText = errorText & "- Custom property " & _
                        CStr(frm.lstCustomProperties.List(i, 0)) & " (" & Err.Number & "): " & _
                        Err.Description & vbCrLf
            Err.Clear
        End If
        On Error GoTo 0
    Next i
End Sub

Private Function SafeCustomPropertyIsLinked(ByVal prop As Object) As Boolean
    On Error GoTo TreatAsLinked
    SafeCustomPropertyIsLinked = CBool(prop.LinkToContent)
    Exit Function

TreatAsLinked:
    Err.Clear
    SafeCustomPropertyIsLinked = True
End Function

Private Function SafeCustomPropertyIsEditable(ByVal prop As Object) As Boolean
    On Error GoTo NotEditable
    If prop Is Nothing Then Exit Function
    If IsCleanupSuiteManagedProperty(CStr(prop.Name)) Then Exit Function
    If SafeCustomPropertyIsLinked(prop) Then Exit Function
    SafeCustomPropertyIsEditable = True
    Exit Function

NotEditable:
    Err.Clear
End Function

Private Function SafeCustomNameConflictsWithProtectedProperty(ByVal propertyName As String) As Boolean
    Dim prop As Object

    On Error Resume Next
    Set prop = mSafeEditDocument.CustomDocumentProperties(propertyName)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0

    If IsCleanupSuiteManagedProperty(CStr(prop.Name)) Then
        SafeCustomNameConflictsWithProtectedProperty = True
    ElseIf SafeCustomPropertyIsLinked(prop) Then
        SafeCustomNameConflictsWithProtectedProperty = True
    End If
End Function

Private Function SafeCustomOriginalNameIsListed(ByVal frm As Object, ByVal propertyName As String) As Boolean
    Dim i As Long
    For i = 0 To frm.lstCustomProperties.ListCount - 1
        If StrComp(CStr(frm.lstCustomProperties.List(i, 3)), propertyName, vbTextCompare) = 0 Then
            SafeCustomOriginalNameIsListed = True
            Exit Function
        End If
    Next i
End Function

Public Sub CancelSafeMetadataEditor(ByVal frm As Object)
    mSafeEditWasSaved = False
    Unload frm
End Sub

Public Sub NoteSafeMetadataEditorCancelled()
    mSafeEditWasSaved = False
End Sub

Private Sub RunPackageMetadataGroupWorkflow(ByVal tempDir As String, ByVal filePath As String)
    Dim result As VbMsgBoxResult

    WritePackageMetadataEditGuide tempDir, filePath
    Shell "explorer.exe """ & tempDir & """", vbNormalFocus

    result = MsgBox("The current document package is open in a working folder." & vbCrLf & vbCrLf & _
                    "Broad group editing works here at the package level." & vbCrLf & _
                    "Edit the metadata files you want, then click Yes to rebuild the document." & vbCrLf & _
                    "Click No to leave the original document unchanged.", _
                    vbYesNo + vbQuestion + vbDefaultButton2, "DOCX Group Metadata Editor")

    If result = vbYes Then
        If RepackDocx(tempDir, filePath) Then
            MsgBox "The edited metadata package was rebuilt into the document successfully.", _
                   vbInformation, "DOCX Group Metadata Editor"
        Else
            MsgBox "The package edits could not be rebuilt into the document.", _
                   vbExclamation, "DOCX Group Metadata Editor"
        End If
    End If

    On Error Resume Next
    DeleteFolder tempDir
    On Error GoTo 0
End Sub

Private Sub WritePackageMetadataEditGuide(ByVal tempDir As String, ByVal filePath As String)
    Dim guideText As String
    Dim guidePath As String

    guidePath = tempDir & "\00_METADATA_EDIT_GUIDE.txt"
    guideText = "DOCX Group Metadata Editor" & vbCrLf & _
                "Source document:" & vbCrLf & filePath & vbCrLf & vbCrLf & _
                "Broad package-level editing happens in this extracted working folder." & vbCrLf & _
                "The original document does not change until you confirm a rebuild." & vbCrLf & vbCrLf & _
                "Common metadata files:" & vbCrLf & _
                "  docProps\core.xml   = title, subject, creator, keywords, description" & vbCrLf & _
                "  docProps\app.xml    = application, company, manager, revision, security" & vbCrLf & _
                "  docProps\custom.xml = custom document properties" & vbCrLf & _
                "  _rels and word\*.rels = relationships" & vbCrLf & _
                "  word\comments.xml, word\settings.xml, and other word\*.xml files may contain metadata too" & vbCrLf & vbCrLf & _
                "Suggested workflow:" & vbCrLf & _
                "  1. Edit the files you want in an external editor." & vbCrLf & _
                "  2. Return to Word." & vbCrLf & _
                "  3. Click Yes when asked to rebuild the document." & vbCrLf & _
                "  4. Click No to discard this working folder and leave the document unchanged."

    WriteTextFile guidePath, guideText
End Sub

Private Sub AutoEditMetadataGroup(ByVal tempDir As String, ByVal filePath As String, ByVal groupChoice As String)
    Dim corePath As String
    Dim appPath As String
    Dim customPath As String
    Dim edited As Boolean

    corePath = tempDir & "\docProps\core.xml"
    appPath = tempDir & "\docProps\app.xml"
    customPath = tempDir & "\docProps\custom.xml"

    edited = False

    If groupChoice = "core" Or groupChoice = "all" Then
        If FileExists(corePath) Then
            If PromptAndUpdateCoreProperties(corePath) Then edited = True
        End If
    End If

    If groupChoice = "app" Or groupChoice = "all" Then
        If FileExists(appPath) Then
            If PromptAndUpdateAppProperties(appPath) Then edited = True
        End If
    End If

    If groupChoice = "custom" Or groupChoice = "all" Then
        If FileExists(customPath) Then
            If PromptAndUpdateCustomProperties(customPath) Then edited = True
        End If
    End If

    If edited Then
        If RepackDocx(tempDir, filePath) Then
            MsgBox "Metadata updates were saved successfully.", vbInformation, "DOCX Group Metadata Editor"
        Else
            MsgBox "Metadata updates were made, but the DOCX could not be rebuilt.", vbExclamation, "DOCX Group Metadata Editor"
        End If
    Else
        MsgBox "No supported metadata group was updated.", vbInformation, "DOCX Group Metadata Editor"
    End If
End Sub

Private Function PromptAndUpdateCoreProperties(ByVal path As String) As Boolean
    Dim titleText As String
    Dim subjectText As String
    Dim creatorText As String
    Dim keywordsText As String
    Dim descriptionText As String

    titleText = PromptValue("Core title", GetXmlTextNode(path, "title"))
    subjectText = PromptValue("Core subject", GetXmlTextNode(path, "subject"))
    creatorText = PromptValue("Core creator", GetXmlTextNode(path, "creator"))
    keywordsText = PromptValue("Core keywords", GetXmlTextNode(path, "keywords"))
    descriptionText = PromptValue("Core description", GetXmlTextNode(path, "description"))

    If Len(titleText) > 0 Or Len(subjectText) > 0 Or Len(creatorText) > 0 Or Len(keywordsText) > 0 Or Len(descriptionText) > 0 Then
        SetXmlTextNode path, "title", titleText
        SetXmlTextNode path, "subject", subjectText
        SetXmlTextNode path, "creator", creatorText
        SetXmlTextNode path, "keywords", keywordsText
        SetXmlTextNode path, "description", descriptionText
        PromptAndUpdateCoreProperties = True
    Else
        PromptAndUpdateCoreProperties = False
    End If
End Function

Private Function PromptAndUpdateAppProperties(ByVal path As String) As Boolean
    Dim appText As String
    Dim companyText As String
    Dim managerText As String
    Dim revisionText As String
    Dim securityText As String

    appText = PromptValue("Application name", GetXmlTextNode(path, "Application"))
    companyText = PromptValue("Company", GetXmlTextNode(path, "Company"))
    managerText = PromptValue("Manager", GetXmlTextNode(path, "Manager"))
    revisionText = PromptValue("Revision number", GetXmlTextNode(path, "Revision"))
    securityText = PromptValue("Security value", GetXmlTextNode(path, "DocSecurity"))

    If Len(appText) > 0 Or Len(companyText) > 0 Or Len(managerText) > 0 Or Len(revisionText) > 0 Or Len(securityText) > 0 Then
        SetXmlTextNode path, "Application", appText
        SetXmlTextNode path, "Company", companyText
        SetXmlTextNode path, "Manager", managerText
        SetXmlTextNode path, "Revision", revisionText
        SetXmlTextNode path, "DocSecurity", securityText
        PromptAndUpdateAppProperties = True
    Else
        PromptAndUpdateAppProperties = False
    End If
End Function

Private Function PromptAndUpdateCustomProperties(ByVal path As String) As Boolean
    Dim inputText As String
    Dim pairs() As String
    Dim pair() As String
    Dim i As Long
    Dim namePart As String
    Dim valuePart As String
    Dim updated As Boolean
    Dim xmlDoc As Object
    Dim node As Object
    Dim nodes As Object

    inputText = InputBox("Enter custom metadata as name=value pairs separated by semicolons." & vbCrLf & _
                         "Example: Project=Alpha; Owner=Chris", "Custom Metadata", "")
    If Len(Trim$(inputText)) = 0 Then
        PromptAndUpdateCustomProperties = False
        Exit Function
    End If

    pairs = Split(inputText, ";")
    updated = False

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load path
    If xmlDoc.parseError.ErrorCode <> 0 Then Exit Function

    For i = LBound(pairs) To UBound(pairs)
        pair = Split(Trim$(pairs(i)), "=", 2)
        If UBound(pair) >= 1 Then
            namePart = Trim$(pair(0))
            valuePart = Trim$(pair(1))
            If Len(namePart) > 0 Then
                Set nodes = xmlDoc.SelectNodes("//*[local-name()='property']")
                For Each node In nodes
                    If LCase$(node.getAttribute("name")) = LCase$(namePart) Then
                        node.Text = valuePart
                        updated = True
                        Exit For
                    End If
                Next node
                If Not updated Then
                    Dim newNode As Object
                    Set newNode = xmlDoc.createElement("property")
                    newNode.setAttribute "name", namePart
                    newNode.Text = valuePart
                    xmlDoc.documentElement.appendChild newNode
                    updated = True
                End If
            End If
        End If
    Next i

    If updated Then
        xmlDoc.Save path
        PromptAndUpdateCustomProperties = True
    Else
        PromptAndUpdateCustomProperties = False
    End If
End Function

Private Function PromptValue(ByVal label As String, ByVal defaultValue As String) As String
    Dim result As String
    result = InputBox(label, "Metadata Edit", defaultValue)
    If IsNull(result) Then
        PromptValue = defaultValue
    Else
        PromptValue = result
    End If
End Function

Private Function GetXmlTextNode(ByVal path As String, ByVal localName As String) As String
    Dim xmlDoc As Object
    Dim node As Object

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load path
    If xmlDoc.parseError.ErrorCode <> 0 Then Exit Function

    Set node = xmlDoc.selectSingleNode("//*[local-name()='" & localName & "']")
    If Not node Is Nothing Then
        GetXmlTextNode = Trim$(node.Text)
    End If
End Function

Private Sub SetXmlTextNode(ByVal path As String, ByVal localName As String, ByVal newValue As String)
    Dim xmlDoc As Object
    Dim node As Object

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load path
    If xmlDoc.parseError.ErrorCode <> 0 Then Exit Sub

    Set node = xmlDoc.selectSingleNode("//*[local-name()='" & localName & "']")
    If node Is Nothing Then
        Set node = xmlDoc.createElement(localName)
        xmlDoc.documentElement.appendChild node
    End If
    node.Text = newValue
    xmlDoc.Save path
End Sub

Public Function AuditDocx(ByVal filePath As String) As AuditReport
    Dim tempDir As String
    Dim xmlFiles As Collection
    Dim xmlFile As Variant
    Dim fso As Object
    Dim report As AuditReport
    Dim partName As String
    Dim xmlText As String
    Dim partData As String
    Dim findings As Collection

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Err.Raise vbObjectError + 1001, , "File does not exist"

    report.FilePath = filePath
    report.IsZip = True
    report.StatusLegend = STATUS_LEGEND

    Set report.DocPropsCore = New Collection
    Set report.DocPropsApp = New Collection
    Set report.DocPropsCustom = New Collection
    Set report.Relationships = New Collection
    Set report.XmlFindings = New Collection
    Set report.CryptoFindings = New Collection
    Set report.GuessFindings = New Collection
    Set report.PackageParts = New Collection
    Set report.EmbeddedFiles = New Collection

    tempDir = CreateTempFolder()
    report.ExtractedRoot = tempDir

    If Not UnzipDocx(filePath, tempDir) Then
        report.IsZip = False
        AuditDocx = report
        Exit Function
    End If

    ' Package parts
    For Each xmlFile In GetAllFiles(tempDir)
        partName = Replace(xmlFile, tempDir & Application.PathSeparator, "")
        report.PackageParts.Add partName
    Next xmlFile

    ' Parse core/app/custom props
    If FileExists(tempDir & "\docProps\core.xml") Then
        ParseDocPropsFile tempDir & "\docProps\core.xml", report.DocPropsCore, "core"
    End If
    If FileExists(tempDir & "\docProps\app.xml") Then
        ParseDocPropsFile tempDir & "\docProps\app.xml", report.DocPropsApp, "app"
    End If
    If FileExists(tempDir & "\docProps\custom.xml") Then
        ParseCustomDocProps tempDir & "\docProps\custom.xml", report.DocPropsCustom
    End If

    ' Parse relationships
    ParseRelationshipsFile tempDir & "\_rels\.rels", report.Relationships
    ParseRelationshipsFile tempDir & "\word\_rels\document.xml.rels", report.Relationships
    ParseRelationshipsFile tempDir & "\customXml\_rels\item1.xml.rels", report.Relationships

    ' Parse XML parts and gather metadata-like values
    Set xmlFiles = New Collection
    CollectXmlFiles tempDir, xmlFiles

    For Each xmlFile In xmlFiles
        partData = ReadTextFile(xmlFile)
        If Len(partData) > 0 Then
            xmlText = partData
            ScanXmlMetadata xmlFile, xmlText, report.XmlFindings
            ScanCryptoText xmlFile, xmlText, report.CryptoFindings
            ScanGuessing xmlFile, xmlText, report.GuessFindings
        End If
    Next xmlFile

    ' Embedded/nested files
    AddEmbeddedFileFindings tempDir, report.EmbeddedFiles

    ' Cleanup temp folder (best effort)
    On Error Resume Next
    DeleteFolder tempDir
    On Error GoTo 0

    AuditDocx = report
End Function

Private Sub QuickAuditDocx(ByVal filePath As String, ByRef report As AuditReport, Optional ByVal frm As Object)
    Dim tempDir As String
    Dim xmlFile As Variant
    Dim fso As Object
    Dim partName As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(filePath) Then Err.Raise vbObjectError + 1001, , "File does not exist"

    report.FilePath = filePath
    report.IsZip = True
    report.StatusLegend = STATUS_LEGEND

    Set report.DocPropsCore = New Collection
    Set report.DocPropsApp = New Collection
    Set report.DocPropsCustom = New Collection
    Set report.Relationships = New Collection
    Set report.XmlFindings = New Collection
    Set report.CryptoFindings = New Collection
    Set report.GuessFindings = New Collection
    Set report.PackageParts = New Collection
    Set report.EmbeddedFiles = New Collection

    tempDir = CreateTempFolder()
    report.ExtractedRoot = tempDir

    SetDashboardGroupState frm, "core", "RUN", "Scanning core properties..."
    SetDashboardGroupState frm, "app", "RUN", "Scanning application properties..."
    SetDashboardGroupState frm, "custom", "RUN", "Scanning custom properties..."
    SetDashboardGroupState frm, "package", "RUN", "Scanning package structure..."

    If Not UnzipDocx(filePath, tempDir) Then
        report.IsZip = False
        Exit Sub
    End If

    PulseDashboardState frm, "package", "Scanning package structure"
    For Each xmlFile In GetAllFiles(tempDir)
        partName = Replace(xmlFile, tempDir & Application.PathSeparator, "")
        report.PackageParts.Add partName
    Next xmlFile

    If FileExists(tempDir & "\docProps\core.xml") Then
        PulseDashboardState frm, "core", "Scanning core properties"
        ParseDocPropsFile tempDir & "\docProps\core.xml", report.DocPropsCore, "core"
    End If
    SetDashboardGroupState frm, "core", "DONE", "Core properties ready."

    If FileExists(tempDir & "\docProps\app.xml") Then
        PulseDashboardState frm, "app", "Scanning application properties"
        ParseDocPropsFile tempDir & "\docProps\app.xml", report.DocPropsApp, "app"
    End If
    SetDashboardGroupState frm, "app", "DONE", "Application properties ready."

    If FileExists(tempDir & "\docProps\custom.xml") Then
        PulseDashboardState frm, "custom", "Scanning custom properties"
        ParseCustomDocProps tempDir & "\docProps\custom.xml", report.DocPropsCustom
    End If
    SetDashboardGroupState frm, "custom", "DONE", "Custom properties ready."

    ParseRelationshipsFile tempDir & "\_rels\.rels", report.Relationships
    ParseRelationshipsFile tempDir & "\word\_rels\document.xml.rels", report.Relationships
    ParseRelationshipsFile tempDir & "\customXml\_rels\item1.xml.rels", report.Relationships
    AddEmbeddedFileFindings tempDir, report.EmbeddedFiles
    SetDashboardGroupState frm, "package", "DONE", "Package summary ready."
    SetDashboardGroupState frm, "metadata", "WAIT", "Deep investigation available."
    SetDashboardGroupState frm, "crypto", "WAIT", "Deep investigation available."

    On Error Resume Next
    DeleteFolder tempDir
    On Error GoTo 0
End Sub

Private Function ExtractMetadataFindings(ByVal filePath As String, Optional ByVal frm As Object, _
                                         Optional ByVal groupKey As String = "metadata") As Collection
    Dim tempDir As String
    Dim xmlFiles As Collection
    Dim xmlFile As Variant
    Dim findings As Collection
    Dim partData As String

    Set findings = New Collection
    tempDir = CreateTempFolder()
    If Not UnzipDocx(filePath, tempDir) Then
        Set ExtractMetadataFindings = findings
        Exit Function
    End If

    Set xmlFiles = New Collection
    CollectXmlFiles tempDir, xmlFiles

    For Each xmlFile In xmlFiles
        partData = ReadTextFile(xmlFile)
        If Len(partData) > 0 Then
            PulseDashboardState frm, groupKey, "Investigating metadata"
            ScanXmlMetadata xmlFile, partData, findings
        End If
    Next xmlFile

    On Error Resume Next
    DeleteFolder tempDir
    On Error GoTo 0

    Set ExtractMetadataFindings = findings
End Function

Private Sub ExtractCryptoGuessFindings(ByVal filePath As String, ByRef cryptoFindings As Collection, _
                                       ByRef guessFindings As Collection, Optional ByVal frm As Object, _
                                       Optional ByVal groupKey As String = "crypto")
    Dim tempDir As String
    Dim xmlFiles As Collection
    Dim xmlFile As Variant
    Dim partData As String

    Set cryptoFindings = New Collection
    Set guessFindings = New Collection

    tempDir = CreateTempFolder()
    If Not UnzipDocx(filePath, tempDir) Then Exit Sub

    Set xmlFiles = New Collection
    CollectXmlFiles tempDir, xmlFiles

    For Each xmlFile In xmlFiles
        partData = ReadTextFile(xmlFile)
        If Len(partData) > 0 Then
            PulseDashboardState frm, groupKey, "Investigating crypto"
            ScanCryptoText xmlFile, partData, cryptoFindings
            ScanGuessing xmlFile, partData, guessFindings
        End If
    Next xmlFile

    On Error Resume Next
    DeleteFolder tempDir
    On Error GoTo 0
End Sub

Private Function ResolvePreferredDocxTarget(Optional ByVal fallbackPath As String = vbNullString, _
                                            Optional ByVal allowFilePicker As Boolean = True) As String
    Dim doc As Object
    Dim promptResult As VbMsgBoxResult

    Set doc = GetPreferredOpenDocument()
    If Not doc Is Nothing Then
        If Len(doc.Path) > 0 Then
            If IsSupportedDocxPath(doc.FullName) Then
                ResolvePreferredDocxTarget = doc.FullName
                Exit Function
            End If
        Else
            promptResult = MsgBox("The current document is open but has not been saved yet." & vbCrLf & vbCrLf & _
                                  "Click Yes to save it now, No to choose another file, or Cancel to stop.", _
                                  vbYesNoCancel + vbExclamation, "DOCX Metadata Audit")
            If promptResult = vbYes Then
                doc.Activate
                doc.Save
                If Len(doc.Path) > 0 And IsSupportedDocxPath(doc.FullName) Then
                    ResolvePreferredDocxTarget = doc.FullName
                    Exit Function
                End If
                MsgBox "Save the document as a .docx or .docm file before auditing it.", vbExclamation, "DOCX Metadata Audit"
            ElseIf promptResult = vbCancel Then
                Exit Function
            End If
        End If
    End If

    If IsSupportedDocxPath(fallbackPath) And FileExists(fallbackPath) Then
        ResolvePreferredDocxTarget = fallbackPath
        Exit Function
    End If

    If allowFilePicker Then
        ResolvePreferredDocxTarget = PickDocxFile()
    ElseIf Documents.Count = 0 Then
        MsgBox "No saved Word document is currently open. Use Audit Another Document to pick a file.", vbInformation, "MetaDataSuite"
    End If
End Function

Private Function GetPreferredOpenDocument() As Object
    Dim doc As Object

    On Error Resume Next
    Set doc = ActiveDocument
    On Error GoTo 0

    If Not doc Is Nothing Then
        If IsCandidateSourceDocument(doc, True) Then
            Set GetPreferredOpenDocument = doc
            Exit Function
        End If
    End If

    For Each doc In Documents
        If IsCandidateSourceDocument(doc, False) Then
            Set GetPreferredOpenDocument = doc
            Exit Function
        End If
    Next doc
End Function

Private Function PeekPreferredDocxTarget() As String
    Dim doc As Object

    Set doc = GetPreferredOpenDocument()
    If doc Is Nothing Then Exit Function
    If Len(doc.Path) = 0 Then Exit Function
    If IsSupportedDocxPath(doc.FullName) Then
        PeekPreferredDocxTarget = doc.FullName
    End If
End Function

Private Function IsCandidateSourceDocument(ByVal doc As Object, Optional ByVal allowHostDocument As Boolean = False) As Boolean
    If doc Is Nothing Then Exit Function
    If Not allowHostDocument Then
        If IsHostMacroDocument(doc) Then Exit Function
    End If
    If IsGeneratedMetadataDocument(doc) Then Exit Function

    If Len(doc.Path) = 0 Then
        IsCandidateSourceDocument = True
    Else
        IsCandidateSourceDocument = IsSupportedDocxPath(doc.FullName)
    End If
End Function

Private Sub ResetDashboardState()
    mDashboardHasQuickReport = False
    Set mDashboardMetadataFindings = New Collection
    Set mDashboardCryptoFindings = New Collection
    Set mDashboardGuessFindings = New Collection
    mDashboardMetadataInvestigated = False
    mDashboardCryptoInvestigated = False
    mDashboardIndicatorStep = 0
End Sub

Private Sub ConfigureDashboardDefaults(ByVal frm As Object)
    Const FORM_W As Single = 396
    On Error Resume Next
    ConfigureDashboardChrome frm
    ConfigureDashboardDetailTextBoxes frm
    frm.ScrollTop = 0
    frm.Width = FORM_W
    frm.lblTargetTitle.Visible = False
    frm.lblTargetTitle.Move 0, 0, 0, 0
    frm.lblTargetPath.Visible = False
    frm.lblTargetPath.Move 0, 0, 0, 0
    frm.lblOverallIndicator.Visible = False
    frm.lblOverallIndicator.Move 0, 0, 0, 0
    frm.lblOverallStatus.Visible = False
    frm.lblOverallStatus.Move 0, 0, 0, 0
    frm.lblHint.Visible = False
    frm.lblHint.Move 0, 0, 0, 0
    frm.Height = 476
    HideLegacyAuthorClearPanel frm
    frm.chkCoreDetails.Value = False
    frm.chkAppDetails.Value = False
    frm.chkCustomDetails.Value = False
    frm.chkPackageDetails.Value = False
    frm.chkMetadataDetails.Value = False
    frm.chkCryptoDetails.Value = False
    frm.chkShowSuiteMetadata.Value = False
    mDashboardShowSuiteMetadata = False
    ApplyMetadataDashboardVisibility frm
    LayoutMetadataDashboard frm
    On Error GoTo 0
End Sub

Private Sub UpdateDashboardDisplay(ByVal frm As Object)
    Dim visibleCore As Collection
    Dim visibleApp As Collection
    Dim visibleCustom As Collection
    Dim visiblePackageParts As Collection
    Dim visibleRelationships As Collection
    Dim visibleEmbeddedFiles As Collection
    Dim visibleMetadataFindings As Collection
    Dim visibleCryptoFindings As Collection
    Dim visibleGuessFindings As Collection

    On Error Resume Next

    If mDashboardHasQuickReport Then
        Set visibleCore = FilterDashboardMetadataItems(mDashboardQuickReport.DocPropsCore)
        Set visibleApp = FilterDashboardMetadataItems(mDashboardQuickReport.DocPropsApp)
        Set visibleCustom = FilterDashboardMetadataItems(mDashboardQuickReport.DocPropsCustom)
        Set visiblePackageParts = FilterDashboardMetadataItems(mDashboardQuickReport.PackageParts)
        Set visibleRelationships = FilterDashboardMetadataItems(mDashboardQuickReport.Relationships)
        Set visibleEmbeddedFiles = FilterDashboardMetadataItems(mDashboardQuickReport.EmbeddedFiles)

        frm.lblCoreSummary.Caption = "Quick check: " & CStr(visibleCore.Count) & " core properties found."
        frm.txtCoreDetails.Value = BuildDocumentPropertyRows(visibleCore, "No core properties found.", _
                                                             "docProps/core.xml", "core")
        frm.lblAppSummary.Caption = "Quick check: " & CStr(visibleApp.Count) & " application properties found."
        frm.txtAppDetails.Value = BuildDocumentPropertyRows(visibleApp, "No application properties found.", _
                                                            "docProps/app.xml", "application")
        frm.lblCustomSummary.Caption = "Quick check: " & CStr(visibleCustom.Count) & " custom properties found."
        frm.txtCustomDetails.Value = BuildDocumentPropertyRows(visibleCustom, "No custom properties found.", _
                                                               "docProps/custom.xml", "custom")
        frm.lblPackageSummary.Caption = "Quick check: " & CStr(visiblePackageParts.Count) & " package parts, " & _
                                        CStr(visibleRelationships.Count) & " relationships, " & _
                                        CStr(visibleEmbeddedFiles.Count) & " embedded files."
        frm.txtPackageDetails.Value = BuildPackageDetailsText(visibleRelationships, visibleEmbeddedFiles)
    Else
        frm.lblCoreSummary.Caption = "Quick check has not run yet."
        frm.txtCoreDetails.Value = "Use Refresh Quick Checks to scan the current document."
        frm.lblAppSummary.Caption = "Quick check has not run yet."
        frm.txtAppDetails.Value = "Use Refresh Quick Checks to scan the current document."
        frm.lblCustomSummary.Caption = "Quick check has not run yet."
        frm.txtCustomDetails.Value = "Use Refresh Quick Checks to scan the current document."
        frm.lblPackageSummary.Caption = "Quick check has not run yet."
        frm.txtPackageDetails.Value = "Use Refresh Quick Checks to scan the current document."
    End If

    If mDashboardMetadataInvestigated Then
        Set visibleMetadataFindings = FilterDashboardMetadataItems(mDashboardMetadataFindings)
        frm.lblMetadataSummary.Caption = "Deep check complete: " & CStr(visibleMetadataFindings.Count) & " metadata findings."
        frm.txtMetadataDetails.Value = BuildEditableMetadataRows(visibleMetadataFindings, "No metadata findings detected.", _
                                                                 DASHBOARD_ACTION_SAVE_OPEN, "package metadata findings", _
                                                                 "Open the source part before editing this payload.")
    Else
        frm.lblMetadataSummary.Caption = "Deep metadata investigation is available on demand."
        frm.txtMetadataDetails.Value = "Click Investigate to run the slower metadata-pattern scan for this document."
    End If

    If mDashboardCryptoInvestigated Then
        Set visibleCryptoFindings = FilterDashboardMetadataItems(mDashboardCryptoFindings)
        Set visibleGuessFindings = FilterDashboardMetadataItems(mDashboardGuessFindings)
        frm.lblCryptoSummary.Caption = "Deep check complete: " & CStr(visibleCryptoFindings.Count) & _
                                       " crypto findings, " & CStr(visibleGuessFindings.Count) & " guessing findings."
        frm.txtCryptoDetails.Value = BuildCryptoDetailsText(visibleCryptoFindings, visibleGuessFindings)
    Else
        frm.lblCryptoSummary.Caption = "Deep crypto and guessing investigation is available on demand."
        frm.txtCryptoDetails.Value = "Click Investigate to run the slower crypto, encoding, and guessing-pattern scan."
    End If

    ApplyMetadataDashboardVisibility frm
    frm.Repaint
    On Error GoTo 0
End Sub

Private Sub LayoutMetadataDashboard(ByVal frm As Object)
    Const FORM_W As Single = 396
    Const LEFT_COL As Single = 12
    Const DETAILS_LEFT As Single = 12
    Const DETAILS_WIDTH As Single = 372
    Const HEADER_HEIGHT As Single = 16
    Const GROUP_GAP As Single = 8
    Const DETAILS_GAP As Single = 3
    Const HEADER_TO_SUMMARY As Single = 14
    Const BOTTOM_PADDING As Single = 8
    Dim topPos As Single
    Dim contentBottom As Single
    Dim viewportHeight As Single

    If frm Is Nothing Then Exit Sub

    On Error Resume Next

    ConfigureDashboardChrome frm
    frm.Width = FORM_W
    frm.Height = 476
    frm.lblTitle.Move 12, 8, 260, 18
    frm.lblTargetTitle.Visible = False
    frm.lblTargetTitle.Move 0, 0, 0, 0
    frm.lblTargetPath.Visible = False
    frm.lblTargetPath.Move 0, 0, 0, 0
    frm.lblOverallIndicator.Visible = False
    frm.lblOverallIndicator.Move 0, 0, 0, 0
    frm.lblOverallStatus.Visible = False
    frm.lblOverallStatus.Move 0, 0, 0, 0
    frm.lblHint.Visible = False
    frm.lblHint.Move 0, 0, 0, 0
    frm.cmdRefreshCurrent.Move 12, 30, 116, 24
    frm.cmdAuditAnother.Move 136, 30, 120, 24
    frm.cmdOpenWordReport.Move 264, 30, 120, 24
    frm.cmdSafeEditCurrent.Move 12, 58, 116, 24
    frm.cmdGroupEditCurrent.Move 136, 58, 120, 24
    frm.cmdClose.Move 264, 58, 120, 24
    frm.cmdClearSharingProperties.Move 12, 86, 180, 24
    frm.cmdExpandAll.Move 204, 86, 180, 24
    frm.chkShowSuiteMetadata.Move 12, 114, 180, 18
    HideLegacyAuthorClearPanel frm

    topPos = 138
    LayoutMetadataDashboardGroup frm, "Core", topPos, False
    LayoutMetadataDashboardGroup frm, "App", topPos, False
    LayoutMetadataDashboardGroup frm, "Custom", topPos, False
    LayoutMetadataDashboardGroup frm, "Package", topPos, False
    LayoutMetadataDashboardGroup frm, "Metadata", topPos, True
    LayoutMetadataDashboardGroup frm, "Crypto", topPos, True

    contentBottom = topPos + BOTTOM_PADDING
    frm.ScrollHeight = contentBottom
    viewportHeight = frm.InsideHeight
    If viewportHeight <= 0 Then viewportHeight = frm.Height - 24
    If contentBottom <= viewportHeight + 10 Then
        frm.ScrollBars = fmScrollBarsNone
        frm.ScrollHeight = viewportHeight
    Else
        frm.ScrollBars = fmScrollBarsVertical
        frm.ScrollHeight = contentBottom
    End If
    frm.ScrollTop = 0

    On Error GoTo 0
End Sub

Public Sub ClearMetadataDashboardSharingProperties(ByVal frm As Object)
    Dim filePath As String
    Dim doc As Object
    Dim closeWhenDone As Boolean
    Dim previewText As String
    Dim promptText As String

    filePath = mDashboardTargetPath
    If Len(filePath) = 0 Then
        filePath = ResolvePreferredDocxTarget(vbNullString, False)
        If Len(filePath) = 0 Then Exit Sub
        mDashboardTargetPath = filePath
    End If

    Set doc = ResolveDocumentForMetadataEdit(filePath, closeWhenDone)
    If doc Is Nothing Then Exit Sub

    On Error GoTo ClearErr
    previewText = BuildSharingPropertiesPreview(doc)
    If InStr(1, previewText, "No populated sharing properties", vbTextCompare) > 0 Then
        MsgBox previewText, vbInformation, "Clear Sharing Properties"
        GoTo ClearExit
    End If

    promptText = previewText & vbCrLf & _
                 "This clears document info fields and Word personal-info metadata before sharing." & vbCrLf & _
                 "CleanupSuite-managed custom properties will be restored afterward." & vbCrLf & vbCrLf & _
                 "Continue?"
    If MsgBox(promptText, vbQuestion + vbYesNo + vbDefaultButton2, "Clear Sharing Properties") <> vbYes Then GoTo ClearExit

    ApplyClearSharingProperties doc
    doc.Save
    MsgBox "Sharing properties were cleared.", vbInformation, "Clear Sharing Properties"
    If Not frm Is Nothing Then RefreshMetadataDashboardQuickScan frm

ClearExit:
    If closeWhenDone And Not doc Is Nothing Then doc.Close SaveChanges:=False
    Set doc = Nothing
    Exit Sub

ClearErr:
    MsgBox "Unable to clear sharing properties: " & Err.Description, vbExclamation, "Clear Sharing Properties"
    Resume ClearExit
End Sub

Private Function BuildSharingPropertiesPreview(ByVal doc As Object) As String
    Dim fields As Variant
    Dim fieldName As Variant
    Dim valueText As String
    Dim info As String

    fields = SharingPropertyBuiltinFields()
    info = "Clear Sharing Properties preview:" & vbCrLf & vbCrLf

    For Each fieldName In fields
        valueText = GetBuiltInDocumentPropertyText(doc, CStr(fieldName))
        If Len(Trim$(valueText)) > 0 Then
            info = info & "- " & CStr(fieldName) & ": " & valueText & vbCrLf
        End If
    Next fieldName

    On Error Resume Next
    If doc.Revisions.Count > 0 Then info = info & "- Revisions: " & CStr(doc.Revisions.Count) & " revision author traces may be present." & vbCrLf
    On Error GoTo 0

    If info = "Clear Sharing Properties preview:" & vbCrLf & vbCrLf Then
        BuildSharingPropertiesPreview = "No populated sharing properties were found."
    Else
        BuildSharingPropertiesPreview = info
    End If
End Function

Private Sub ApplyClearSharingProperties(ByVal doc As Object)
    Dim protectedProps As Object

    If doc Is Nothing Then Exit Sub

    Set protectedProps = SnapshotCleanupSuiteManagedProperties(doc)
    On Error Resume Next
    doc.RemoveDocumentInformation wdRDIDocumentProperties
    doc.RemoveDocumentInformation wdRDIRemovePersonalInformation
    doc.RemovePersonalInformation = True
    On Error GoTo 0
    RestoreCleanupSuiteManagedProperties doc, protectedProps
End Sub

Private Function SharingPropertyBuiltinFields() As Variant
    SharingPropertyBuiltinFields = Array("Title", "Subject", "Author", "Last Author", "Company", "Manager", _
                                         "Category", "Keywords", "Comments", "Hyperlink Base", "Revision Number", _
                                         "Content Status")
End Function

Private Sub RunPackageMetadataGroupViewOnlyWorkflow(ByVal tempDir As String, ByVal filePath As String, ByVal reasonText As String)
    WritePackageMetadataEditGuide tempDir, filePath
    Shell "explorer.exe """ & tempDir & """", vbNormalFocus
    MsgBox reasonText & vbCrLf & vbCrLf & _
           "The extracted package is open for inspection only." & vbCrLf & _
           "Word keeps the current document locked while it is open, and this tool will not rebuild over a running macro host.", _
           vbInformation, "DOCX Group Metadata Editor"

    On Error Resume Next
    DeleteFolder tempDir
    On Error GoTo 0
End Sub

Private Sub LayoutMetadataDashboardGroup(ByVal frm As Object, ByVal prefix As String, ByRef topPos As Single, ByVal hasButton As Boolean)
    Const LEFT_COL As Single = 12
    Const CHECK_LEFT As Single = 166
    Const INDICATOR_LEFT As Single = 258
    Const BUTTON_LEFT As Single = 308
    Const DETAILS_LEFT As Single = 12
    Const DETAILS_WIDTH As Single = 372
    Const HEADER_HEIGHT As Single = 16
    Const GROUP_GAP As Single = 8
    Const DETAILS_GAP As Single = 4
    Const HEADER_TO_SUMMARY As Single = 14
    Dim txt As Object
    Dim chk As Object
    Dim summaryCtl As Object
    Dim summaryHeight As Single
    Dim detailHeight As Single

    Set txt = CallByName(frm, "txt" & prefix & "Details", VbGet)
    Set chk = CallByName(frm, "chk" & prefix & "Details", VbGet)
    Set summaryCtl = CallByName(frm, "lbl" & prefix & "Summary", VbGet)

    CallByName(frm, "lbl" & prefix & "Header", VbGet).Move LEFT_COL, topPos, 150, HEADER_HEIGHT
    chk.Move CHECK_LEFT, topPos - 2, 84, 18
    CallByName(frm, "lbl" & prefix & "Indicator", VbGet).Move INDICATOR_LEFT, topPos, 44, HEADER_HEIGHT

    If hasButton Then
        CallByName(frm, "cmd" & prefix & "Investigate", VbGet).Move BUTTON_LEFT, topPos - 4, 76, 22
    End If

    summaryHeight = MetadataWrappedLabelHeight(frm, summaryCtl, summaryCtl.Caption, DETAILS_WIDTH)
    summaryCtl.Move LEFT_COL, topPos + HEADER_TO_SUMMARY, DETAILS_WIDTH, summaryHeight

    If chk.Value Then
        detailHeight = MetadataWrappedTextBoxHeight(frm, txt, CStr(txt.Value), DETAILS_WIDTH)
        txt.Move DETAILS_LEFT, summaryCtl.Top + summaryCtl.Height + DETAILS_GAP, DETAILS_WIDTH, detailHeight
        txt.Visible = True
        If detailHeight <= 34 Then
            txt.ScrollBars = fmScrollBarsNone
        Else
            txt.ScrollBars = fmScrollBarsVertical
        End If
        topPos = txt.Top + txt.Height + GROUP_GAP
    Else
        txt.Visible = False
        txt.Move DETAILS_LEFT, summaryCtl.Top + summaryCtl.Height + DETAILS_GAP, DETAILS_WIDTH, 14
        txt.ScrollBars = fmScrollBarsVertical
        topPos = summaryCtl.Top + summaryCtl.Height + GROUP_GAP
    End If
End Sub

Private Function MetadataWrappedLabelHeight(ByVal frm As Object, ByVal sourceCtl As Object, ByVal detailText As String, ByVal detailWidth As Single) As Single
    Dim measureCtl As Object

    Set measureCtl = frm.lblHint
    With measureCtl
        .Caption = detailText
        .Font.Name = sourceCtl.Font.Name
        .Font.Size = sourceCtl.Font.Size
        .Font.Bold = sourceCtl.Font.Bold
        .WordWrap = True
        .Visible = False
        .Move -1000, -1000, detailWidth, 11
        .AutoSize = False
        .AutoSize = True
        MetadataWrappedLabelHeight = .Height
        .AutoSize = False
        .Move 0, 0, 0, 0
        .Caption = ""
    End With

    If MetadataWrappedLabelHeight < 16 Then MetadataWrappedLabelHeight = 16
End Function

Private Function MetadataWrappedTextBoxHeight(ByVal frm As Object, ByVal txt As Object, ByVal detailText As String, ByVal detailWidth As Single) As Single
    MetadataWrappedTextBoxHeight = MetadataWrappedLabelHeight(frm, txt, detailText, detailWidth - 8) + txt.Font.Size + 6
    If MetadataWrappedTextBoxHeight < 18 Then MetadataWrappedTextBoxHeight = 18
End Function

Private Sub SetDashboardOverallState(ByVal frm As Object, ByVal indicatorText As String, ByVal statusText As String)
    If frm Is Nothing Then Exit Sub
    On Error Resume Next
    frm.lblOverallIndicator.Caption = indicatorText
    frm.lblOverallStatus.Caption = statusText
    frm.Repaint
    DoEvents
    On Error GoTo 0
End Sub

Private Sub SetDashboardGroupState(ByVal frm As Object, ByVal groupKey As String, ByVal indicatorText As String, ByVal summaryText As String)
    If frm Is Nothing Then Exit Sub
    On Error Resume Next
    Select Case LCase$(groupKey)
        Case "core"
            frm.lblCoreIndicator.Caption = indicatorText
            If Len(summaryText) > 0 Then frm.lblCoreSummary.Caption = summaryText
        Case "app"
            frm.lblAppIndicator.Caption = indicatorText
            If Len(summaryText) > 0 Then frm.lblAppSummary.Caption = summaryText
        Case "custom"
            frm.lblCustomIndicator.Caption = indicatorText
            If Len(summaryText) > 0 Then frm.lblCustomSummary.Caption = summaryText
        Case "package"
            frm.lblPackageIndicator.Caption = indicatorText
            If Len(summaryText) > 0 Then frm.lblPackageSummary.Caption = summaryText
        Case "metadata"
            frm.lblMetadataIndicator.Caption = indicatorText
            If Len(summaryText) > 0 Then frm.lblMetadataSummary.Caption = summaryText
        Case "crypto"
            frm.lblCryptoIndicator.Caption = indicatorText
            If Len(summaryText) > 0 Then frm.lblCryptoSummary.Caption = summaryText
    End Select
    frm.Repaint
    DoEvents
    On Error GoTo 0
End Sub

Private Sub PulseDashboardState(ByVal frm As Object, ByVal groupKey As String, ByVal baseText As String)
    Dim suffix As String

    mDashboardIndicatorStep = mDashboardIndicatorStep + 1
    suffix = String$((mDashboardIndicatorStep Mod 3) + 1, ".")
    SetDashboardOverallState frm, "RUN" & suffix, baseText & suffix
    SetDashboardGroupState frm, groupKey, "RUN" & suffix, baseText & suffix
End Sub

Private Function CollectionToMultilineText(ByVal items As Collection, ByVal emptyText As String) As String
    Dim item As Variant
    Dim sb As String

    If items Is Nothing Then
        CollectionToMultilineText = emptyText
        Exit Function
    End If

    If items.Count = 0 Then
        CollectionToMultilineText = emptyText
        Exit Function
    End If

    For Each item In items
        sb = sb & CStr(item) & vbCrLf
    Next item

    If Right$(sb, 2) = vbCrLf Then sb = Left$(sb, Len(sb) - 2)
    CollectionToMultilineText = sb
End Function

Private Function FilterDashboardMetadataItems(ByVal items As Collection) As Collection
    Dim filtered As Collection
    Dim rawItem As Variant

    Set filtered = New Collection
    If items Is Nothing Then
        Set FilterDashboardMetadataItems = filtered
        Exit Function
    End If

    For Each rawItem In items
        If Not mDashboardShowSuiteMetadata Then
            If Not ShouldHideSuiteMetadataItem(CStr(rawItem)) Then filtered.Add CStr(rawItem)
        Else
            filtered.Add CStr(rawItem)
        End If
    Next rawItem

    Set FilterDashboardMetadataItems = filtered
End Function

Private Function ShouldHideSuiteMetadataItem(ByVal rawText As String) As Boolean
    Dim lowerText As String

    lowerText = LCase$(rawText)
    ShouldHideSuiteMetadataItem = _
        InStr(lowerText, "word\vbadata.xml") > 0 Or _
        InStr(lowerText, "word/vbadata.xml") > 0 Or _
        InStr(lowerText, "cleanupsuite") > 0 Or _
        InStr(lowerText, "metadata toolkit") > 0 Or _
        InStr(lowerText, "metadatatoolkitrole") > 0 Or _
        InStr(lowerText, "metadatasuite") > 0 Or _
        InStr(lowerText, "showcleanupsuitelauncher") > 0 Or _
        InStr(lowerText, "cleanupsuitereturntomainafterapply") > 0 Or _
        InStr(lowerText, "cleanupsuitereturntomainafterclose") > 0 Or _
        InStr(lowerText, LCase$("vbaProject.bin")) > 0 Or _
        InStr(lowerText, "vbaproject") > 0 Or _
        InStr(lowerText, "customui") > 0 Or _
        InStr(lowerText, "ribbon") > 0 Or _
        InStr(lowerText, "macros") > 0 Or _
        InStr(lowerText, "macro") > 0
End Function

Private Function ShouldSkipSuiteManagedMetadataLine(ByVal filePath As String, ByVal rawText As String) As Boolean
    Dim lowerPath As String
    Dim lowerText As String

    lowerPath = LCase$(Trim$(filePath))
    lowerText = LCase$(rawText)

    ShouldSkipSuiteManagedMetadataLine = _
        InStr(lowerPath, "word\vbadata.xml") > 0 Or _
        InStr(lowerPath, "word/vbadata.xml") > 0 Or _
        InStr(lowerText, "metadatatoolkitrole") > 0 Or _
        InStr(lowerText, "cleanupsuiteshowcompletionreviewafterapply") > 0 Or _
        InStr(lowerText, "cleanupsuitereturntomainafterapply") > 0 Or _
        InStr(lowerText, "cleanupsuitereturntomainafterclose") > 0 Or _
        InStr(lowerText, "cleanupsuiteautosave") > 0 Or _
        InStr(lowerText, "cleanupsuiteenableadvancedmetadatatools") > 0 Or _
        InStr(lowerText, "cleanupsuiteadvancedmetadatawarningshown") > 0 Or _
        InStr(lowerText, "project.modmetadatasuite.") > 0 Or _
        InStr(lowerText, "project.modcleanuplauncher.showcleanupsuitelauncher") > 0
End Function

Private Function BuildEditableMetadataRows(ByVal items As Collection, ByVal emptyText As String, ByVal primaryAction As String, _
                                           ByVal storageLocation As String, Optional ByVal capabilityNote As String = vbNullString, _
                                           Optional ByVal secondaryAction As String = vbNullString, _
                                           Optional ByVal isPossiblyEncrypted As Boolean = False) As String
    Dim rawItem As Variant
    Dim dashboardItem As MetadataDashboardItem
    Dim sb As String

    If items Is Nothing Then
        BuildEditableMetadataRows = emptyText
        Exit Function
    End If

    If items.Count = 0 Then
        BuildEditableMetadataRows = emptyText
        Exit Function
    End If

    For Each rawItem In items
        dashboardItem = CreateMetadataDashboardItem(CStr(rawItem), primaryAction, storageLocation, _
                                                    capabilityNote, secondaryAction, isPossiblyEncrypted)
        sb = sb & FormatMetadataDashboardItem(dashboardItem) & vbCrLf & vbCrLf
    Next rawItem

    If Right$(sb, 4) = vbCrLf & vbCrLf Then sb = Left$(sb, Len(sb) - 4)
    BuildEditableMetadataRows = sb
End Function

Private Function BuildDocumentPropertyRows(ByVal items As Collection, ByVal emptyText As String, _
                                           ByVal storageLocation As String, ByVal propertyArea As String) As String
    Dim rawItem As Variant
    Dim dashboardItem As MetadataDashboardItem
    Dim sb As String

    If items Is Nothing Then
        BuildDocumentPropertyRows = emptyText
        Exit Function
    End If

    If items.Count = 0 Then
        BuildDocumentPropertyRows = emptyText
        Exit Function
    End If

    For Each rawItem In items
        dashboardItem = CreateMetadataDashboardItem(CStr(rawItem), DASHBOARD_ACTION_SAVE_OPEN, _
                                                    storageLocation, vbNullString, vbNullString, False)
        dashboardItem.PrimaryAction = DashboardDocumentPropertyCapability(dashboardItem.DisplayName, _
                                                                          CStr(rawItem), propertyArea)
        sb = sb & FormatMetadataDashboardItem(dashboardItem) & vbCrLf & vbCrLf
    Next rawItem

    If Right$(sb, 4) = vbCrLf & vbCrLf Then sb = Left$(sb, Len(sb) - 4)
    BuildDocumentPropertyRows = sb
End Function

Private Function DashboardDocumentPropertyCapability(ByVal displayName As String, ByVal rawText As String, _
                                                      ByVal propertyArea As String) As String
    Dim normalizedName As String
    Dim propertyName As String

    propertyName = Trim$(displayName)
    If InStr(1, propertyName, " (pid=", vbTextCompare) > 0 Then
        propertyName = Left$(propertyName, InStr(1, propertyName, " (pid=", vbTextCompare) - 1)
    End If
    normalizedName = LCase$(propertyName)

    Select Case LCase$(Trim$(propertyArea))
        Case "core"
            Select Case normalizedName
                Case "title", "subject", "creator", "keywords", "description", "category", _
                     "lastmodifiedby", "contentstatus", "contenttype", "language"
                    DashboardDocumentPropertyCapability = DASHBOARD_ACTION_EDIT
                Case Else
                    DashboardDocumentPropertyCapability = DASHBOARD_ACTION_SAVE_OPEN
            End Select
        Case "application"
            Select Case normalizedName
                Case "company", "manager"
                    DashboardDocumentPropertyCapability = DASHBOARD_ACTION_EDIT
                Case Else
                    DashboardDocumentPropertyCapability = DASHBOARD_ACTION_SAVE_OPEN
            End Select
        Case "custom"
            If IsCleanupSuiteManagedProperty(propertyName) Then
                DashboardDocumentPropertyCapability = DASHBOARD_ACTION_SAVE_OPEN
            ElseIf InStr(1, rawText, "linked", vbTextCompare) > 0 Then
                DashboardDocumentPropertyCapability = DASHBOARD_ACTION_SAVE_OPEN
            Else
                DashboardDocumentPropertyCapability = DASHBOARD_ACTION_EDIT
            End If
        Case Else
            DashboardDocumentPropertyCapability = DASHBOARD_ACTION_SAVE_OPEN
    End Select
End Function

Private Function CreateMetadataDashboardItem(ByVal rawText As String, ByVal primaryAction As String, _
                                             ByVal storageLocation As String, ByVal capabilityNote As String, _
                                             ByVal secondaryAction As String, ByVal isPossiblyEncrypted As Boolean) As MetadataDashboardItem
    Dim item As MetadataDashboardItem
    Dim normalizedText As String
    Dim splitAt As Long
    Dim displayName As String
    Dim currentValue As String
    Dim colonAt As Long

    normalizedText = NormalizeAuditTempPathText(rawText)

    splitAt = InStr(normalizedText, " -> ")
    If splitAt > 0 Then
        displayName = Left$(normalizedText, splitAt - 1)
        currentValue = Mid$(normalizedText, splitAt + 4)
    Else
        splitAt = InStr(normalizedText, " = ")
        If splitAt > 0 Then
            displayName = Left$(normalizedText, splitAt - 1)
            currentValue = Mid$(normalizedText, splitAt + 3)
        Else
            displayName = normalizedText
            currentValue = normalizedText
        End If
    End If

    colonAt = InStr(displayName, ":")
    If colonAt > 0 Then displayName = Trim$(Mid$(displayName, colonAt + 1))
    displayName = Trim$(displayName)
    If Len(displayName) = 0 Then displayName = "metadata item"

    item.GroupKey = storageLocation
    item.DisplayName = displayName
    item.StorageLocation = storageLocation & "/" & displayName
    item.CurrentValue = VisualizeHiddenCharacters(Trim$(currentValue))
    item.PrimaryAction = primaryAction
    item.SecondaryAction = secondaryAction
    item.CapabilityNote = capabilityNote
    item.IsPossiblyEncrypted = isPossiblyEncrypted

    CreateMetadataDashboardItem = item
End Function

Private Function NormalizeAuditTempPathText(ByVal rawText As String) As String
    Dim normalized As String

    normalized = rawText
    normalized = Replace$(normalized, "/", "\")

    normalized = ReplaceAuditTempPathSegment(normalized)
    NormalizeAuditTempPathText = normalized
End Function

Private Function ReplaceAuditTempPathSegment(ByVal rawText As String) As String
    Dim lowerText As String
    Dim markerPos As Long
    Dim startPos As Long
    Dim scanPos As Long
    Dim pathEnd As Long
    Dim extractedPath As String
    Dim relativePath As String

    lowerText = LCase$(rawText)
    markerPos = InStr(lowerText, "\docxmetadataaudit_")
    If markerPos = 0 Then
        ReplaceAuditTempPathSegment = rawText
        Exit Function
    End If

    startPos = markerPos
    Do While startPos > 1
        If Mid$(rawText, startPos - 1, 1) = "\" Or Mid$(rawText, startPos - 1, 1) = ":" Then Exit Do
        startPos = startPos - 1
    Loop
    If Mid$(rawText, startPos, 1) = "\" And startPos > 3 Then startPos = startPos - 3

    pathEnd = Len(rawText) + 1
    For scanPos = markerPos To Len(rawText)
        Select Case Mid$(rawText, scanPos, 1)
            Case " ", ")", "]"
                pathEnd = scanPos
                Exit For
        End Select
    Next scanPos

    extractedPath = Mid$(rawText, startPos, pathEnd - startPos)
    relativePath = BuildPackageRelativePath(extractedPath)
    If Len(relativePath) = 0 Then
        ReplaceAuditTempPathSegment = rawText
    Else
        ReplaceAuditTempPathSegment = Left$(rawText, startPos - 1) & relativePath & Mid$(rawText, pathEnd)
    End If
End Function

Private Function BuildPackageRelativePath(ByVal extractedPath As String) As String
    Dim normalizedPath As String
    Dim markerPos As Long
    Dim slashPos As Long

    normalizedPath = Replace$(Trim$(extractedPath), "/", "\")
    markerPos = InStr(LCase$(normalizedPath), "\docxmetadataaudit_")
    If markerPos = 0 Then
        BuildPackageRelativePath = normalizedPath
        Exit Function
    End If

    slashPos = InStr(markerPos + Len("\docxmetadataaudit_"), normalizedPath, "\")
    If slashPos = 0 Then
        BuildPackageRelativePath = normalizedPath
    Else
        BuildPackageRelativePath = Mid$(normalizedPath, slashPos + 1)
    End If
End Function

Private Function FormatMetadataDashboardItem(ByRef item As MetadataDashboardItem) As String
    Dim lineText As String

    lineText = "[" & item.PrimaryAction & "] "
    If Len(item.SecondaryAction) > 0 Then lineText = lineText & "[" & item.SecondaryAction & "] "
    If item.IsPossiblyEncrypted Then lineText = lineText & "[possibly encrypted] "

    lineText = lineText & item.DisplayName & vbCrLf & _
               "Location: " & item.StorageLocation & vbCrLf & _
               "Value: " & item.CurrentValue
    If Len(item.CapabilityNote) > 0 Then
        lineText = lineText & vbCrLf & "Note: " & item.CapabilityNote
    End If

    FormatMetadataDashboardItem = lineText
End Function

Private Function VisualizeHiddenCharacters(ByVal text As String) As String
    Dim result As String

    result = text
    result = Replace$(result, vbCrLf, "[CRLF]")
    result = Replace$(result, vbCr, "[CR]")
    result = Replace$(result, vbLf, "[LF]")
    result = Replace$(result, vbTab, "[TAB]")
    result = Replace$(result, Chr$(0), "[NUL]")
    result = Replace$(result, Chr$(160), "[NBSP]")
    result = Replace$(result, ChrW$(8203), "[ZWSP]")
    result = Replace$(result, ChrW$(8204), "[ZWNJ]")
    result = Replace$(result, ChrW$(8205), "[ZWJ]")
    result = Replace$(result, ChrW$(65279), "[BOM]")

    VisualizeHiddenCharacters = result
End Function

Private Function BuildPackageDetailsText(ByVal relationships As Collection, ByVal embeddedFiles As Collection) As String
    Dim sb As String

    sb = "Relationships:" & vbCrLf & BuildEditableMetadataRows(relationships, "None found.", _
                                                               DASHBOARD_ACTION_SAVE_OPEN, "package relationships", _
                                                               "Relationship data usually needs package-level editing.") & vbCrLf & vbCrLf & _
         "Embedded files:" & vbCrLf & BuildEditableMetadataRows(embeddedFiles, "None found.", _
                                                                DASHBOARD_ACTION_SAVE_OPEN, "embedded package item", _
                                                                "Embedded data should be opened externally when possible.")
    BuildPackageDetailsText = sb
End Function

Private Function BuildCryptoDetailsText(ByVal cryptoFindings As Collection, ByVal guessFindings As Collection) As String
    BuildCryptoDetailsText = "Crypto / encoding findings:" & vbCrLf & _
                             BuildEditableMetadataRows(cryptoFindings, "None found.", _
                                                       DASHBOARD_ACTION_SAVE_OPEN, "crypto or encoded payload", _
                                                       "May be encrypted or encoded.", DASHBOARD_ACTION_DECRYPT, True) & vbCrLf & vbCrLf & _
                             "Guessing findings:" & vbCrLf & _
                             BuildEditableMetadataRows(guessFindings, "None found.", _
                                                       DASHBOARD_ACTION_SAVE_OPEN, "guessing finding", _
                                                       "Open the source payload before editing.")
End Function

Private Function BuildMetadataDashboardSmokeText(ByVal filePath As String) As String
    Dim sb As String

    sb = "Target: " & filePath & vbCrLf
    sb = sb & "QuickSummary" & vbCrLf
    sb = sb & "- Core=" & CStr(mDashboardQuickReport.DocPropsCore.Count) & vbCrLf
    sb = sb & "- App=" & CStr(mDashboardQuickReport.DocPropsApp.Count) & vbCrLf
    sb = sb & "- Custom=" & CStr(mDashboardQuickReport.DocPropsCustom.Count) & vbCrLf
    sb = sb & "- PackageParts=" & CStr(mDashboardQuickReport.PackageParts.Count) & vbCrLf
    sb = sb & "- EmbeddedFiles=" & CStr(mDashboardQuickReport.EmbeddedFiles.Count) & vbCrLf
    sb = sb & "- MetadataFindings=" & CStr(mDashboardQuickReport.XmlFindings.Count) & vbCrLf
    sb = sb & "- CryptoFindings=" & CStr(mDashboardQuickReport.CryptoFindings.Count) & vbCrLf
    sb = sb & "- GuessFindings=" & CStr(mDashboardQuickReport.GuessFindings.Count) & vbCrLf
    sb = sb & vbCrLf
    sb = sb & "[Core]" & vbCrLf & BuildDocumentPropertyRows(mDashboardQuickReport.DocPropsCore, "None found.", _
                                                            "docProps/core.xml", "core") & vbCrLf & vbCrLf
    sb = sb & "[Application]" & vbCrLf & BuildDocumentPropertyRows(mDashboardQuickReport.DocPropsApp, "None found.", _
                                                                   "docProps/app.xml", "application") & vbCrLf & vbCrLf
    sb = sb & "[Custom]" & vbCrLf & BuildDocumentPropertyRows(mDashboardQuickReport.DocPropsCustom, "None found.", _
                                                              "docProps/custom.xml", "custom") & vbCrLf & vbCrLf
    sb = sb & "[Package]" & vbCrLf & _
         BuildPackageDetailsText(mDashboardQuickReport.Relationships, _
                                 mDashboardQuickReport.EmbeddedFiles) & vbCrLf & vbCrLf
    sb = sb & "[MetadataQuick]" & vbCrLf & BuildEditableMetadataRows(mDashboardQuickReport.XmlFindings, "No metadata findings detected.", _
                                                                     DASHBOARD_ACTION_SAVE_OPEN, "quick metadata finding", _
                                                                     "Open the source part before editing this payload.") & vbCrLf & vbCrLf
    sb = sb & "[CryptoQuick]" & vbCrLf & _
         BuildCryptoDetailsText(mDashboardQuickReport.CryptoFindings, _
                                mDashboardQuickReport.GuessFindings)

    If mDashboardMetadataInvestigated Then
        sb = sb & vbCrLf & vbCrLf & "[MetadataDeep]" & vbCrLf & _
             BuildEditableMetadataRows(mDashboardMetadataFindings, "No metadata findings detected.", _
                                       DASHBOARD_ACTION_SAVE_OPEN, "package metadata finding", _
                                       "Open the source part before editing this payload.")
    End If

    If mDashboardCryptoInvestigated Then
        sb = sb & vbCrLf & vbCrLf & "[CryptoDeep]" & vbCrLf & _
             BuildCryptoDetailsText(mDashboardCryptoFindings, mDashboardGuessFindings)
    End If

    BuildMetadataDashboardSmokeText = sb
End Function

Private Function NormalizeAuthorMetadataRemovalLevel(ByVal removalLevel As String) As String
    Select Case LCase$(Trim$(removalLevel))
        Case AUTHOR_CLEAR_MINIMAL
            NormalizeAuthorMetadataRemovalLevel = AUTHOR_CLEAR_MINIMAL
        Case AUTHOR_CLEAR_NORMAL
            NormalizeAuthorMetadataRemovalLevel = AUTHOR_CLEAR_NORMAL
        Case AUTHOR_CLEAR_BROAD
            NormalizeAuthorMetadataRemovalLevel = AUTHOR_CLEAR_BROAD
    End Select
End Function

Private Function BuildAuthorMetadataRemovalPreview(ByVal filePath As String, ByVal removalLevel As String) As String
    Dim doc As Object
    Dim closeWhenDone As Boolean
    Dim fields As Variant
    Dim i As Long
    Dim valueText As String
    Dim sb As String
    Dim prop As Object
    Dim count As Long

    Set doc = ResolveDocumentForMetadataEdit(filePath, closeWhenDone, True)
    If doc Is Nothing Then
        BuildAuthorMetadataRemovalPreview = "No current document could be inspected."
        Exit Function
    End If

    fields = AuthorRemovalBuiltinFields(removalLevel)
    For i = LBound(fields) To UBound(fields)
        valueText = GetBuiltInDocumentPropertyText(doc, CStr(fields(i)))
        If Len(valueText) > 0 Then
            sb = sb & "[Will clear] Built-in property: " & CStr(fields(i)) & " | " & _
                 VisualizeHiddenCharacters(valueText) & vbCrLf
            count = count + 1
        End If
    Next i

    On Error Resume Next
    For Each prop In doc.CustomDocumentProperties
        If ShouldClearCustomAuthorProperty(prop.Name, removalLevel) Then
            valueText = CStr(prop.Value)
            If Err.Number <> 0 Then
                Err.Clear
                valueText = ""
            End If
            sb = sb & "[Will clear] Custom property: " & CStr(prop.Name) & " | " & _
                 VisualizeHiddenCharacters(valueText) & vbCrLf
            count = count + 1
        End If
    Next prop
    On Error GoTo 0

    If removalLevel = AUTHOR_CLEAR_BROAD Then
        sb = sb & "[Will enable] Word privacy flag: RemovePersonalInformation | Applies when saved" & vbCrLf
        count = count + 1
    End If

    If closeWhenDone Then
        On Error Resume Next
        doc.Close False
        On Error GoTo 0
    End If

    If count = 0 Then
        BuildAuthorMetadataRemovalPreview = "No matching non-empty author metadata was found."
    Else
        If Right$(sb, 2) = vbCrLf Then sb = Left$(sb, Len(sb) - 2)
        BuildAuthorMetadataRemovalPreview = sb
    End If
End Function

Private Sub ApplyAuthorMetadataRemoval(ByVal filePath As String, ByVal removalLevel As String)
    Dim doc As Object
    Dim closeWhenDone As Boolean
    Dim fields As Variant
    Dim i As Long

    Set doc = ResolveDocumentForMetadataEdit(filePath, closeWhenDone, False)
    If doc Is Nothing Then Exit Sub

    fields = AuthorRemovalBuiltinFields(removalLevel)
    For i = LBound(fields) To UBound(fields)
        ClearBuiltInDocumentProperty doc, CStr(fields(i))
    Next i

    ClearCustomAuthorProperties doc, removalLevel

    If removalLevel = AUTHOR_CLEAR_BROAD Then
        On Error Resume Next
        doc.RemovePersonalInformation = True
        On Error GoTo 0
    End If

    On Error Resume Next
    doc.Save
    If closeWhenDone Then doc.Close False
    On Error GoTo 0
End Sub

Private Function ResolveDocumentForMetadataEdit(ByVal filePath As String, ByRef closeWhenDone As Boolean, _
                                                Optional ByVal readOnly As Boolean = False) As Object
    Dim doc As Object
    Dim targetPath As String

    targetPath = LCase$(Trim$(filePath))
    closeWhenDone = False

    On Error Resume Next
    For Each doc In Application.Documents
        If LCase$(doc.FullName) = targetPath Then
            Set ResolveDocumentForMetadataEdit = doc
            Exit Function
        End If
    Next doc
    On Error GoTo 0

    On Error Resume Next
    Set doc = Application.Documents.Open(filePath, ReadOnly:=readOnly, AddToRecentFiles:=False)
    If Err.Number <> 0 Then
        Err.Clear
        Set doc = Nothing
    End If
    On Error GoTo 0

    If Not doc Is Nothing Then
        closeWhenDone = True
        Set ResolveDocumentForMetadataEdit = doc
    End If
End Function

Private Function AuthorRemovalBuiltinFields(ByVal removalLevel As String) As Variant
    Select Case removalLevel
        Case AUTHOR_CLEAR_MINIMAL
            AuthorRemovalBuiltinFields = Array("Author", "Last Author", "Company", "Manager")
        Case AUTHOR_CLEAR_NORMAL
            AuthorRemovalBuiltinFields = Array("Author", "Last Author", "Company", "Manager", _
                                               "Comments", "Keywords", "Category", "Subject")
        Case AUTHOR_CLEAR_BROAD
            AuthorRemovalBuiltinFields = Array("Author", "Last Author", "Company", "Manager", _
                                               "Comments", "Keywords", "Category", "Subject", _
                                               "Title", "Revision Number", "Hyperlink Base", "Content Status")
        Case Else
            AuthorRemovalBuiltinFields = Array("Author", "Last Author", "Company", "Manager")
    End Select
End Function

Private Function GetBuiltInDocumentPropertyText(ByVal doc As Object, ByVal propertyName As String) As String
    On Error Resume Next
    GetBuiltInDocumentPropertyText = CStr(doc.BuiltInDocumentProperties(propertyName).Value)
    If Err.Number <> 0 Then
        Err.Clear
        GetBuiltInDocumentPropertyText = vbNullString
    End If
    On Error GoTo 0
End Function

Private Sub ClearBuiltInDocumentProperty(ByVal doc As Object, ByVal propertyName As String)
    On Error Resume Next
    doc.BuiltInDocumentProperties(propertyName).Value = vbNullString
    Err.Clear
    On Error GoTo 0
End Sub

Private Function SnapshotCleanupSuiteManagedProperties(ByVal doc As Object) As Object
    Dim result As Object
    Dim prop As Object

    Set result = CreateObject("Scripting.Dictionary")
    If doc Is Nothing Then
        Set SnapshotCleanupSuiteManagedProperties = result
        Exit Function
    End If

    On Error Resume Next
    For Each prop In doc.CustomDocumentProperties
        If IsCleanupSuiteManagedProperty(CStr(prop.Name)) Then
            result(CStr(prop.Name)) = Array(prop.Type, prop.Value)
        End If
    Next prop
    On Error GoTo 0

    Set SnapshotCleanupSuiteManagedProperties = result
End Function

Private Sub RestoreCleanupSuiteManagedProperties(ByVal doc As Object, ByVal protectedProps As Object)
    Dim key As Variant
    Dim propertyData As Variant

    If doc Is Nothing Then Exit Sub
    If protectedProps Is Nothing Then Exit Sub

    On Error Resume Next
    For Each key In protectedProps.Keys
        propertyData = protectedProps(key)
        doc.CustomDocumentProperties(CStr(key)).Delete
        doc.CustomDocumentProperties.Add Name:=CStr(key), LinkToContent:=False, Type:=CLng(propertyData(0)), Value:=propertyData(1)
    Next key
    On Error GoTo 0
End Sub

Private Function IsCleanupSuiteManagedProperty(ByVal propertyName As String) As Boolean
    Select Case propertyName
        Case "CleanupSuiteAutoSave"
            IsCleanupSuiteManagedProperty = True
        Case "CleanupSuiteReturnToMainAfterApply"
            IsCleanupSuiteManagedProperty = True
        Case "CleanupSuiteReturnToMainAfterClose"
            IsCleanupSuiteManagedProperty = True
        Case "CleanupSuiteInProgress"
            IsCleanupSuiteManagedProperty = True
        Case "CleanupSuiteCapitalizationExceptions"
            IsCleanupSuiteManagedProperty = True
    End Select
End Function

Private Sub ClearCustomAuthorProperties(ByVal doc As Object, ByVal removalLevel As String)
    Dim prop As Object
    Dim namesToDelete As Collection
    Dim nameValue As Variant

    Set namesToDelete = New Collection

    On Error Resume Next
    For Each prop In doc.CustomDocumentProperties
        If ShouldClearCustomAuthorProperty(prop.Name, removalLevel) Then
            namesToDelete.Add CStr(prop.Name)
        End If
    Next prop

    For Each nameValue In namesToDelete
        doc.CustomDocumentProperties(CStr(nameValue)).Delete
    Next nameValue
    On Error GoTo 0
End Sub

Private Function ShouldClearCustomAuthorProperty(ByVal propertyName As String, ByVal removalLevel As String) As Boolean
    Dim lowerName As String

    If removalLevel = AUTHOR_CLEAR_BROAD Then
        ShouldClearCustomAuthorProperty = True
        Exit Function
    End If

    If removalLevel <> AUTHOR_CLEAR_NORMAL Then Exit Function

    lowerName = LCase$(propertyName)
    ShouldClearCustomAuthorProperty = _
        (InStr(lowerName, "author") > 0 Or _
         InStr(lowerName, "creator") > 0 Or _
         InStr(lowerName, "owner") > 0 Or _
         InStr(lowerName, "user") > 0 Or _
         InStr(lowerName, "email") > 0 Or _
         InStr(lowerName, "name") > 0 Or _
         InStr(lowerName, "company") > 0 Or _
         InStr(lowerName, "manager") > 0 Or _
         InStr(lowerName, "organization") > 0 Or _
         InStr(lowerName, "organisation") > 0 Or _
         InStr(lowerName, "reviewer") > 0)
End Function

Private Function IsSupportedDocxPath(ByVal filePath As String) As Boolean
    Dim lowerPath As String

    lowerPath = LCase$(Trim$(filePath))
    IsSupportedDocxPath = (Right$(lowerPath, 5) = ".docx" Or Right$(lowerPath, 5) = ".docm")
End Function

Private Function PickDocxFile() As String
    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.AllowMultiSelect = False
    fd.Title = "Select a DOCX file"
    fd.Filters.Clear
    fd.Filters.Add "Word Documents", "*.docx;*.docm", 1

    If fd.Show = -1 Then
        PickDocxFile = fd.SelectedItems(1)
    Else
        PickDocxFile = vbNullString
    End If
End Function

Private Function IsHostMacroDocument(ByVal doc As Object) As Boolean
    Dim hostPath As String

    On Error Resume Next
    hostPath = LCase$(ThisDocument.FullName)
    On Error GoTo 0

    If Len(hostPath) = 0 Then Exit Function

    On Error Resume Next
    IsHostMacroDocument = (LCase$(doc.FullName) = hostPath)
    On Error GoTo 0
End Function

Private Function IsRunningMacroHostPath(ByVal filePath As String) As Boolean
    Dim hostPath As String

    On Error Resume Next
    hostPath = ThisDocument.FullName
    On Error GoTo 0

    If Len(hostPath) = 0 Then Exit Function
    IsRunningMacroHostPath = (LCase$(Trim$(filePath)) = LCase$(Trim$(hostPath)))
End Function

Private Function IsOpenDocumentPath(ByVal filePath As String) As Boolean
    Dim doc As Object
    Dim targetPath As String

    targetPath = LCase$(Trim$(filePath))
    If Len(targetPath) = 0 Then Exit Function

    For Each doc In Documents
        On Error Resume Next
        If LCase$(Trim$(doc.FullName)) = targetPath Then
            IsOpenDocumentPath = True
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next doc
End Function

Private Function GetOutputPath(ByVal filePath As String) As String
    Dim fso As Object
    Dim basePath As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    basePath = fso.GetBaseName(filePath)
    GetOutputPath = fso.GetParentFolderName(filePath) & "\" & basePath & ".audit.txt"
End Function

Private Function CreateTempFolder() As String
    Dim fso As Object
    Dim tmp As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    tmp = Environ$("TEMP") & "\DocxMetadataAudit_" & Replace(Format(Now, "yyyymmddhhnnss"), ":", "")
    If Not fso.FolderExists(tmp) Then fso.CreateFolder tmp
    CreateTempFolder = tmp
End Function

Private Function UnzipDocx(ByVal docxPath As String, ByVal destDir As String) As Boolean
    Dim psCommand As String
    Dim shell As Object
    Dim result As Long
    Dim docxZipPath As String

    docxZipPath = destDir & "\source.zip"

    psCommand = "powershell -ExecutionPolicy Bypass -NoProfile -Command " & Chr(34) & _
               "$ErrorActionPreference='Stop'; " & _
               "Import-Module Microsoft.PowerShell.Archive -ErrorAction Stop; " & _
               "Copy-Item -LiteralPath '" & docxPath & "' -Destination '" & docxZipPath & "' -Force; " & _
               "Expand-Archive -LiteralPath '" & docxZipPath & "' -DestinationPath '" & destDir & "' -Force" & Chr(34)

    Set shell = CreateObject("WScript.Shell")
    result = shell.Run(psCommand, 0, True)
    UnzipDocx = (result = 0)
End Function

Private Function RepackDocx(ByVal sourceFolder As String, ByVal outputPath As String) As Boolean
    Dim psCommand As String
    Dim shell As Object
    Dim result As Long
    Dim tempZip As String

    tempZip = sourceFolder & "\repacked.zip"
    On Error Resume Next
    Kill tempZip
    On Error GoTo 0

    psCommand = "powershell -ExecutionPolicy Bypass -NoProfile -Command " & Chr(34) & _
               "$ErrorActionPreference='Stop'; " & _
               "Import-Module Microsoft.PowerShell.Archive -ErrorAction Stop; " & _
               "if (Test-Path '" & tempZip & "') { Remove-Item '" & tempZip & "' -Force }; " & _
               "Compress-Archive -Path '" & sourceFolder & "\\*' -DestinationPath '" & tempZip & "' -Force; " & _
               "Copy-Item '" & tempZip & "' -Destination '" & outputPath & "' -Force" & Chr(34)

    Set shell = CreateObject("WScript.Shell")
    result = shell.Run(psCommand, 0, True)
    RepackDocx = (result = 0)
End Function

Private Function ReadTextFile(ByVal path As String) As String
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(path, 1, False, -2)
    ReadTextFile = ts.ReadAll
    ts.Close
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal content As String)
    Dim fso As Object
    Dim ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(path, True, True)
    ts.Write content
    ts.Close
End Sub

Private Sub RenderAuditReportInWord(ByRef report As AuditReport, ByVal reportText As String)
    Dim doc As Object

    Set doc = Documents.Add
    MarkMetadataDocument doc, METADATA_ROLE_RESULTS

    AppendStyledParagraph doc, "Audit Results - " & GetFileNameFromPath(report.FilePath), "Title"
    AppendStyledParagraph doc, "Generated audit view. Review findings here first; export to text only if you need a shareable report file.", "Subtitle"
    AppendStyledParagraph doc, "Source document: " & report.FilePath, vbNullString
    AppendStyledParagraph doc, "Audit timestamp: " & Format$(mLastAuditTimestamp, "yyyy-mm-dd hh:nn:ss"), vbNullString
    AppendStyledParagraph doc, vbNullString, vbNullString
    AppendMacroButton doc, "RefreshLastAuditResult", "Refresh Audit"
    AppendMacroButton doc, "SaveLastAuditReportAsText", "Save Audit Report As Text"
    AppendMacroButton doc, "RunDocxMetadataAuditAnotherDocument", "Audit Another Document"
    AppendMacroButton doc, "SafeEditLastAuditedDocument", "Safe Edit Current Document"
    AppendMacroButton doc, "GroupEditLastAuditedDocument", "Group Edit Current Document"
    AppendStyledParagraph doc, vbNullString, vbNullString

    AppendStyledParagraph doc, "Summary", "Heading 1"
    AppendStyledParagraph doc, "Package parts: " & CStr(report.PackageParts.Count), vbNullString
    AppendStyledParagraph doc, "Core properties: " & CStr(report.DocPropsCore.Count), vbNullString
    AppendStyledParagraph doc, "Application properties: " & CStr(report.DocPropsApp.Count), vbNullString
    AppendStyledParagraph doc, "Custom properties: " & CStr(report.DocPropsCustom.Count), vbNullString
    AppendStyledParagraph doc, "Relationships: " & CStr(report.Relationships.Count), vbNullString
    AppendStyledParagraph doc, "Metadata findings: " & CStr(report.XmlFindings.Count), vbNullString
    AppendStyledParagraph doc, "Crypto findings: " & CStr(report.CryptoFindings.Count), vbNullString
    AppendStyledParagraph doc, "Guessing findings: " & CStr(report.GuessFindings.Count), vbNullString
    AppendStyledParagraph doc, "Embedded files: " & CStr(report.EmbeddedFiles.Count), vbNullString
    AppendStyledParagraph doc, vbNullString, vbNullString

    AppendStyledParagraph doc, "Status Legend", "Heading 1"
    AppendStyledParagraph doc, report.StatusLegend, vbNullString
    AppendStyledParagraph doc, vbNullString, vbNullString

    AppendCollectionSection doc, "Core Properties", report.DocPropsCore
    AppendCollectionSection doc, "Application Properties", report.DocPropsApp
    AppendCollectionSection doc, "Custom Properties", report.DocPropsCustom
    AppendCollectionSection doc, "Relationships", report.Relationships
    AppendCollectionSection doc, "Metadata Findings", report.XmlFindings
    AppendCollectionSection doc, "Crypto Findings", report.CryptoFindings
    AppendCollectionSection doc, "Guessing Findings", report.GuessFindings
    AppendCollectionSection doc, "Embedded Files", report.EmbeddedFiles

    doc.Variables.Add Name:="MetadataToolkitReportTextLength", Value:=CStr(Len(reportText))
End Sub

Private Sub AppendCollectionSection(ByVal doc As Object, ByVal headingText As String, ByVal items As Collection)
    Dim item As Variant

    AppendStyledParagraph doc, headingText, "Heading 1"
    If items.Count = 0 Then
        AppendStyledParagraph doc, "None found.", vbNullString
    Else
        For Each item In items
            AppendStyledParagraph doc, "- " & CStr(item), vbNullString
        Next item
    End If
    AppendStyledParagraph doc, vbNullString, vbNullString
End Sub

Private Sub AppendMacroButton(ByVal doc As Object, ByVal macroName As String, ByVal captionText As String)
    Dim buttonRange As Object
    Dim buttonField As Object

    Set buttonRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    Set buttonField = doc.Fields.Add(Range:=buttonRange, Type:=wdFieldMacroButton, Text:=macroName & " " & captionText)
    buttonField.Result.Font.Bold = True
    buttonField.Result.Font.Color = wdColorBlue
    buttonField.Result.Underline = wdUnderlineSingle
    buttonField.Result.InsertParagraphAfter
End Sub

Private Sub AppendStyledParagraph(ByVal doc As Object, ByVal paragraphText As String, ByVal styleName As String)
    Dim textRange As Object

    Set textRange = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
    textRange.InsertAfter paragraphText
    textRange.InsertParagraphAfter
    textRange.SetRange textRange.Start, textRange.Start + Len(paragraphText)

    If Len(styleName) > 0 Then
        On Error Resume Next
        textRange.Style = styleName
        On Error GoTo 0
    End If
End Sub

Private Sub MarkMetadataDocument(ByVal doc As Object, ByVal roleName As String)
    On Error Resume Next
    doc.Variables(METADATA_ROLE_VARIABLE).Delete
    doc.Variables.Add Name:=METADATA_ROLE_VARIABLE, Value:=roleName
    doc.BuiltInDocumentProperties("Title").Value = "Metadata Toolkit " & roleName
    On Error GoTo 0
End Sub

Private Function IsGeneratedMetadataDocument(ByVal doc As Object) As Boolean
    Dim roleName As String

    On Error Resume Next
    roleName = CStr(doc.Variables(METADATA_ROLE_VARIABLE).Value)
    On Error GoTo 0

    IsGeneratedMetadataDocument = (roleName = METADATA_ROLE_RESULTS)
End Function

Private Function GetFileNameFromPath(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    GetFileNameFromPath = fso.GetFileName(filePath)
End Function

Private Function FileExists(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    FileExists = fso.FileExists(path)
End Function

Private Function DeleteFolder(ByVal path As String) As Boolean
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.DeleteFolder path, True
    DeleteFolder = Not fso.FolderExists(path)
    On Error GoTo 0
End Function

Private Sub ParseDocPropsFile(ByVal path As String, ByRef coll As Collection, ByVal context As String)
    Dim xmlDoc As Object
    Dim nodes As Object
    Dim node As Object
    Dim textValue As String
    Dim keyName As String

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load path
    If xmlDoc.parseError.ErrorCode <> 0 Then Exit Sub

    Set nodes = xmlDoc.SelectNodes("//*")
    For Each node In nodes
        If NodeHasElementChildren(node) Then GoTo NextDocPropNode
        textValue = Trim$(node.Text)
        keyName = node.BaseName
        If Len(textValue) > 0 Then
            If Len(keyName) > 0 Then
                coll.Add context & ": " & keyName & " = " & textValue
            Else
                coll.Add context & ": " & textValue
            End If
        End If
NextDocPropNode:
    Next node
End Sub

Private Function NodeHasElementChildren(ByVal node As Object) As Boolean
    Dim childNode As Object

    If node Is Nothing Then Exit Function

    On Error Resume Next
    For Each childNode In node.ChildNodes
        If childNode.NodeType = 1 Then
            NodeHasElementChildren = True
            Exit Function
        End If
    Next childNode
    On Error GoTo 0
End Function

Private Sub ParseCustomDocProps(ByVal path As String, ByRef coll As Collection)
    Dim xmlDoc As Object
    Dim nodes As Object
    Dim node As Object
    Dim valueText As String
    Dim nameAttr As String
    Dim pidAttr As String

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load path
    If xmlDoc.parseError.ErrorCode <> 0 Then Exit Sub

    Set nodes = xmlDoc.SelectNodes("//*[local-name()='property']")
    For Each node In nodes
        nameAttr = node.getAttribute("name")
        pidAttr = node.getAttribute("pid")
        valueText = Trim$(node.Text)
        If Len(valueText) > 0 Then
            If Len(nameAttr) > 0 Then
                coll.Add "custom: " & nameAttr & " (pid=" & pidAttr & ") = " & valueText
            Else
                coll.Add "custom: " & valueText
            End If
        End If
    Next node
End Sub

Private Sub ParseRelationshipsFile(ByVal path As String, ByRef coll As Collection)
    Dim xmlDoc As Object
    Dim nodes As Object
    Dim node As Object
    Dim relText As String

    If Not FileExists(path) Then Exit Sub

    Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
    xmlDoc.async = False
    xmlDoc.Load path
    If xmlDoc.parseError.ErrorCode <> 0 Then Exit Sub

    Set nodes = xmlDoc.SelectNodes("//*[local-name()='Relationship']")
    For Each node In nodes
        relText = node.getAttribute("Id") & " -> " & node.getAttribute("Target") & " (" & node.getAttribute("Type") & ")"
        coll.Add relText
    Next node
End Sub

Private Sub CollectXmlFiles(ByVal rootFolder As String, ByRef outColl As Collection)
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(rootFolder)

    For Each file In folder.Files
        If LCase$(fso.GetExtensionName(file.Path)) = "xml" Then
            outColl.Add file.Path
        End If
    Next file

    For Each folder In folder.SubFolders
        CollectXmlFiles folder.Path, outColl
    Next folder
End Sub

Private Sub ScanXmlMetadata(ByVal filePath As String, ByVal xmlText As String, ByRef outColl As Collection)
    Dim i As Long
    Dim lineText As String
    Dim lines() As String
    Dim regex As Object
    Dim matches As Object
    Dim m As Object
    Dim lineNo As Long
    Dim foundText As String

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = META_WORDS & "|" & GUID_PATTERN & "|" & URL_PATTERN & "|" & DATE_PATTERN

    lines = Split(xmlText, vbNewLine)
    For i = LBound(lines) To UBound(lines)
        lineText = Trim$(lines(i))
        If Len(lineText) = 0 Then GoTo NextLine
        If ShouldSkipSuiteManagedMetadataLine(filePath, lineText) Then GoTo NextLine
        lineNo = i + 1

        If LooksLikeMetadataText(lineText) Or regex.Test(lineText) Then
            Set matches = regex.Execute(lineText)
            If matches.Count > 0 Then
                For Each m In matches
                    foundText = m.value
                    If Len(foundText) > 0 Then
                        outColl.Add "Part: " & filePath & " (line " & CStr(lineNo) & ") -> " & foundText
                    End If
                Next m
            Else
                outColl.Add "Part: " & filePath & " (line " & CStr(lineNo) & ") -> " & lineText
            End If
        End If
    NextLine:
    Next i
End Sub

Private Function LooksLikeMetadataText(ByVal text As String) As Boolean
    Dim lowerText As String
    lowerText = LCase$(text)
    LooksLikeMetadataText = _
        InStr(lowerText, "author") > 0 Or _
        InStr(lowerText, "creator") > 0 Or _
        InStr(lowerText, "modified") > 0 Or _
        InStr(lowerText, "created") > 0 Or _
        InStr(lowerText, "revision") > 0 Or _
        InStr(lowerText, "template") > 0 Or _
        InStr(lowerText, "comment") > 0 Or _
        InStr(lowerText, "company") > 0 Or _
        InStr(lowerText, "title") > 0 Or _
        InStr(lowerText, "subject") > 0 Or _
        InStr(lowerText, "keyword") > 0 Or _
        InStr(lowerText, "user") > 0 Or _
        InStr(lowerText, "owner") > 0 Or _
        InStr(lowerText, "guid") > 0 Or _
        InStr(lowerText, "uuid") > 0 Or _
        InStr(lowerText, "url") > 0 Or _
        InStr(lowerText, "path") > 0 Or _
        InStr(lowerText, "mailto") > 0 Or _
        InStr(lowerText, "rsid") > 0
End Function

Private Sub ScanCryptoText(ByVal filePath As String, ByVal text As String, ByRef outColl As Collection)
    Dim regex As Object
    Dim matches As Object
    Dim m As Object
    Dim candidate As String

    If ShouldSkipSuiteManagedMetadataLine(filePath, text) Then Exit Sub

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = HASH_PATTERN_LIST
    Set matches = regex.Execute(text)
    For Each m In matches
        candidate = m.value
        outColl.Add "CRY | Hash-like pattern in " & filePath & ": " & candidate
    Next m

    ' Additional crypto / encoding indicators
    If InStr(text, "Salted__") > 0 Then
        outColl.Add "CRY | OpenSSL Salt marker in " & filePath
    End If
    If InStr(text, "PK") > 0 And InStr(text, Chr$(3)) > 0 And InStr(text, Chr$(4)) > 0 Then
        outColl.Add "CRY | ZIP/container signature in " & filePath
    End If
    If RegexContains(text, GUID_PATTERN) Then
        outColl.Add "CRY | GUID-like identifier in " & filePath
    End If
    If RegexContains(text, URL_PATTERN) Then
        outColl.Add "REL | URL-like reference in " & filePath
    End If
    If LooksLikeBase64(text) Then
        outColl.Add "ENC | Base64-like segment in " & filePath
    End If
    If RegexContains(text, HEX_PATTERN) And Len(text) > 64 Then
        outColl.Add "ENC | Large hex-like segment in " & filePath
    End If
End Sub

Private Function LooksLikeBase64(ByVal text As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = BASE64_PATTERN
    regex.Global = False
    regex.IgnoreCase = True
    LooksLikeBase64 = regex.Test(Replace$(text, vbCrLf, ""))
End Function

Private Function RegexContains(ByVal text As String, ByVal pattern As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.IgnoreCase = True
    regex.Pattern = pattern
    RegexContains = regex.Test(text)
End Function

Private Sub ScanGuessing(ByVal filePath As String, ByVal text As String, ByRef outColl As Collection)
    Dim fieldWords As Variant
    Dim i As Long
    Dim lowText As String
    If ShouldSkipSuiteManagedMetadataLine(filePath, text) Then Exit Sub
    lowText = LCase$(text)
    fieldWords = Array("author", "creator", "document", "subject", "title", "comments", "template", "manager", "company", "owner", "password", "secret", _
                      "key", "token", "uuid", "guid", "revision", "last", "modified", "created", "user", "email", "path", "url")
    For i = LBound(fieldWords) To UBound(fieldWords)
        If InStr(lowText, fieldWords(i)) > 0 Then
            outColl.Add "GES | Guessing hint in " & filePath & " -> " & CStr(fieldWords(i))
        End If
    Next i
End Sub

Private Sub AddEmbeddedFileFindings(ByVal rootFolder As String, ByRef outColl As Collection)
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim subFolder As Object
    Dim extensionName As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(rootFolder)

    For Each file In folder.Files
        extensionName = LCase$(fso.GetExtensionName(file.Path))
        Select Case extensionName
            Case "png", "jpg", "jpeg", "gif", "bmp", "pdf"
                outColl.Add file.Path
        End Select
    Next file

    For Each subFolder In folder.SubFolders
        AddEmbeddedFileFindings subFolder.Path, outColl
    Next subFolder
End Sub

Private Function BuildReportText(ByRef report As AuditReport) As String
    Dim sb As String
    Dim item As Variant

    sb = "DOCX Metadata Audit Report" & vbCrLf & _
         "==========================" & vbCrLf & _
         "File: " & report.FilePath & vbCrLf & _
         "Extracted root: " & report.ExtractedRoot & vbCrLf & _
         "Is ZIP/Open XML: " & IIf(report.IsZip, "Yes", "No") & vbCrLf & vbCrLf & _
         "Status legend:" & vbCrLf & report.StatusLegend & vbCrLf & vbCrLf & _
         "Summary:" & vbCrLf & _
         "- Package parts: " & CStr(report.PackageParts.Count) & vbCrLf & _
         "- Core props: " & CStr(report.DocPropsCore.Count) & vbCrLf & _
         "- App props: " & CStr(report.DocPropsApp.Count) & vbCrLf & _
         "- Custom props: " & CStr(report.DocPropsCustom.Count) & vbCrLf & _
         "- Relationships: " & CStr(report.Relationships.Count) & vbCrLf & _
         "- XML findings: " & CStr(report.XmlFindings.Count) & vbCrLf & _
         "- Crypto findings: " & CStr(report.CryptoFindings.Count) & vbCrLf & _
         "- Guessing findings: " & CStr(report.GuessFindings.Count) & vbCrLf & _
         "- Embedded files: " & CStr(report.EmbeddedFiles.Count) & vbCrLf & vbCrLf

    sb = sb & "Document properties (core):" & vbCrLf
    If report.DocPropsCore.Count = 0 Then
        sb = sb & "- None found" & vbCrLf
    Else
        For Each item In report.DocPropsCore
            sb = sb & "- " & item & vbCrLf
        Next item
    End If
    sb = sb & vbCrLf

    sb = sb & "Document properties (app):" & vbCrLf
    If report.DocPropsApp.Count = 0 Then
        sb = sb & "- None found" & vbCrLf
    Else
        For Each item In report.DocPropsApp
            sb = sb & "- " & item & vbCrLf
        Next item
    End If
    sb = sb & vbCrLf

    sb = sb & "Document properties (custom):" & vbCrLf
    If report.DocPropsCustom.Count = 0 Then
        sb = sb & "- None found" & vbCrLf
    Else
        For Each item In report.DocPropsCustom
            sb = sb & "- " & item & vbCrLf
        Next item
    End If
    sb = sb & vbCrLf

    sb = sb & "Relationships:" & vbCrLf
    If report.Relationships.Count = 0 Then
        sb = sb & "- None found" & vbCrLf
    Else
        For Each item In report.Relationships
            sb = sb & "- " & item & vbCrLf
        Next item
    End If
    sb = sb & vbCrLf

    sb = sb & "XML metadata findings:" & vbCrLf
    If report.XmlFindings.Count = 0 Then
        sb = sb & "- None found" & vbCrLf
    Else
        For Each item In report.XmlFindings
            sb = sb & "- " & item & vbCrLf
        Next item
    End If
    sb = sb & vbCrLf

    sb = sb & "Cryptography / encoding findings:" & vbCrLf
    If report.CryptoFindings.Count = 0 Then
        sb = sb & "- None found" & vbCrLf
    Else
        For Each item In report.CryptoFindings
            sb = sb & "- " & item & vbCrLf
        Next item
    End If
    sb = sb & vbCrLf

    sb = sb & "Guessing findings:" & vbCrLf
    If report.GuessFindings.Count = 0 Then
        sb = sb & "- None found" & vbCrLf
    Else
        For Each item In report.GuessFindings
            sb = sb & "- " & item & vbCrLf
        Next item
    End If
    sb = sb & vbCrLf

    sb = sb & "Embedded / nested files detected:" & vbCrLf
    If report.EmbeddedFiles.Count = 0 Then
        sb = sb & "- None found" & vbCrLf
    Else
        For Each item In report.EmbeddedFiles
            sb = sb & "- " & item & vbCrLf
        Next item
    End If

    BuildReportText = sb
End Function

Private Function IIf(ByVal condition As Boolean, ByVal trueVal As Variant, ByVal falseVal As Variant) As Variant
    If condition Then
        IIf = trueVal
    Else
        IIf = falseVal
    End If
End Function

Private Function GetAllFiles(ByVal rootFolder As String) As Collection
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim result As Collection
    Set result = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(rootFolder)

    For Each file In folder.Files
        result.Add file.Path
    Next file

    For Each folder In folder.SubFolders
        Dim nested As Collection
        Set nested = GetAllFiles(folder.Path)
        Dim v As Variant
        For Each v In nested
            result.Add v
        Next v
    Next folder

    Set GetAllFiles = result
End Function
