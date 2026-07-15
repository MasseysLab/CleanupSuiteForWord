Option Explicit

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ReturnToMainAfterToolClose Me, Cancel, CloseMode
End Sub

Private Sub UserForm_Initialize()
    InitializeMetadataDashboard Me
End Sub

Private Sub chkCoreDetails_Click()
    ApplyMetadataDashboardVisibility Me
End Sub

Private Sub chkAppDetails_Click()
    ApplyMetadataDashboardVisibility Me
End Sub

Private Sub chkCustomDetails_Click()
    ApplyMetadataDashboardVisibility Me
End Sub

Private Sub chkPackageDetails_Click()
    ApplyMetadataDashboardVisibility Me
End Sub

Private Sub chkMetadataDetails_Click()
    ApplyMetadataDashboardVisibility Me
End Sub

Private Sub chkCryptoDetails_Click()
    ApplyMetadataDashboardVisibility Me
End Sub

Private Sub chkShowSuiteMetadata_Click()
    SetMetadataDashboardShowSuiteMetadata Me, chkShowSuiteMetadata.Value
End Sub

Private Sub cmdRefreshCurrent_Click()
    RefreshMetadataDashboardQuickScan Me
End Sub

Private Sub cmdAuditAnother_Click()
    PickMetadataDashboardTarget Me
End Sub

Private Sub cmdOpenWordReport_Click()
    OpenMetadataDashboardWordReport Me
End Sub

Private Sub cmdExpandAll_Click()
    ExpandMetadataDashboardAllGroups Me
End Sub

Private Sub cmdSafeEditCurrent_Click()
    SafeEditMetadataDashboardTarget Me
End Sub

Private Sub cmdGroupEditCurrent_Click()
    GroupEditMetadataDashboardTarget Me
End Sub

Private Sub cmdClearSharingProperties_Click()
    ClearMetadataDashboardSharingProperties Me
End Sub

Private Sub cmdMetadataInvestigate_Click()
    InvestigateMetadataDashboardGroup Me, "metadata"
End Sub

Private Sub cmdCryptoInvestigate_Click()
    InvestigateMetadataDashboardGroup Me, "crypto"
End Sub

Private Sub cmdClose_Click()
    CloseMetadataDashboard Me
End Sub
