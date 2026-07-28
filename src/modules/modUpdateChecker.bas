Option Explicit

Private Const UPDATE_SETTINGS_APP As String = "CleanupSuiteForWord"
Private Const UPDATE_SETTINGS_SECTION As String = "GlobalDefaults"
Private Const UPDATE_ENABLED_SETTING As String = "UpdateChecks"
Private Const UPDATE_LAST_CHECK_SETTING As String = "UpdateLastCheck"
Private Const UPDATE_INTERVAL_DAYS As Long = 7
Private Const UPDATE_MANIFEST_URL As String = "https://raw.githubusercontent.com/MasseysLab/CleanupSuiteForWord/main/release/latest-release.ini"
Private Const UPDATE_ALLOWED_URL_PREFIX As String = "https://github.com/MasseysLab/CleanupSuiteForWord/"

Private mUpdateCheckScheduled As Boolean

Public Function GetCleanupSuiteUpdateChecksEnabled() As Boolean
    Dim settingText As String
    On Error GoTo UseDefault
    settingText = GetSetting(UPDATE_SETTINGS_APP, UPDATE_SETTINGS_SECTION, UPDATE_ENABLED_SETTING, "True")
    GetCleanupSuiteUpdateChecksEnabled = (StrComp(Trim$(settingText), "True", vbTextCompare) = 0)
    Exit Function
UseDefault:
    GetCleanupSuiteUpdateChecksEnabled = True
End Function

Public Sub SetCleanupSuiteUpdateChecksEnabled(ByVal enabled As Boolean)
    On Error Resume Next
    SaveSetting UPDATE_SETTINGS_APP, UPDATE_SETTINGS_SECTION, UPDATE_ENABLED_SETTING, IIf(enabled, "True", "False")
    On Error GoTo 0
End Sub

Public Sub ExplainCleanupSuiteUpdateChecksDisabled()
    MsgBox "Periodic update checks are now off." & vbCrLf & vbCrLf & _
           "To turn them back on, open CleanupSuite and select:" & vbCrLf & _
           "Check periodically for updates." & vbCrLf & vbCrLf & _
           "You can still choose Check for Updates at any time.", _
           vbInformation, "CleanupSuite Updates"
End Sub

Public Sub ScheduleCleanupSuiteUpdateCheck()
    On Error GoTo ScheduleDone
    If mUpdateCheckScheduled Then Exit Sub
    If Not GetCleanupSuiteUpdateChecksEnabled() Then Exit Sub
    If Not CleanupSuiteUpdateCheckIsDue() Then Exit Sub
    mUpdateCheckScheduled = True
    Application.OnTime When:=Now + TimeValue("00:00:08"), _
                       Name:="CheckCleanupSuiteUpdateIfDue", _
                       Tolerance:=30
ScheduleDone:
End Sub

Public Sub CheckCleanupSuiteUpdateIfDue()
    mUpdateCheckScheduled = False
    If Not GetCleanupSuiteUpdateChecksEnabled() Then Exit Sub
    If Not CleanupSuiteUpdateCheckIsDue() Then Exit Sub
    CheckForCleanupSuiteUpdates False
End Sub

Public Sub CheckForCleanupSuiteUpdates(Optional ByVal manualCheck As Boolean = True)
    Dim manifestText As String
    Dim remoteVersion As String
    Dim setupUrl As String
    Dim releaseUrl As String
    Dim messageText As String

    On Error GoTo CheckFailed
    SaveCleanupSuiteUpdateAttemptTime
    manifestText = DownloadCleanupSuiteUpdateManifest()
    remoteVersion = ManifestValue(manifestText, "version")
    setupUrl = ManifestValue(manifestText, "setup_url")
    releaseUrl = ManifestValue(manifestText, "release_url")

    If Len(remoteVersion) = 0 Then Err.Raise vbObjectError + 1700, "CheckForCleanupSuiteUpdates", "The update manifest did not include a version."
    If CompareCleanupSuiteVersions(remoteVersion, SUITE_VERSION) > 0 Then
        messageText = "CleanupSuite " & remoteVersion & " is available." & vbCrLf & vbCrLf & _
                      "Installed version: " & SUITE_VERSION & vbCrLf & vbCrLf & _
                      "Open the newest per-user installer now?"
        If MsgBox(messageText, vbInformation + vbYesNo, "CleanupSuite Update Available") = vbYes Then
            If CleanupSuiteUpdateUrlIsAllowed(setupUrl) Then
                OpenCleanupSuiteUpdateUrl setupUrl
            ElseIf CleanupSuiteUpdateUrlIsAllowed(releaseUrl) Then
                OpenCleanupSuiteUpdateUrl releaseUrl
            Else
                Err.Raise vbObjectError + 1701, "CheckForCleanupSuiteUpdates", "The update link was not from the CleanupSuite repository."
            End If
        End If
    ElseIf manualCheck Then
        MsgBox "CleanupSuite is current on the stable release channel." & vbCrLf & vbCrLf & _
               "Installed version: " & SUITE_VERSION, _
               vbInformation, "CleanupSuite Updates"
    End If
    Exit Sub

CheckFailed:
    If manualCheck Then
        MsgBox "CleanupSuite could not check for updates right now." & vbCrLf & vbCrLf & _
               "Your installed version was not changed." & vbCrLf & _
               "Try Check for Updates again later.", vbExclamation, "CleanupSuite Updates"
    End If
End Sub

Private Function CleanupSuiteUpdateCheckIsDue() As Boolean
    Dim savedValue As String
    Dim lastCheckValue As Double
    On Error GoTo CheckNow
    savedValue = GetSetting(UPDATE_SETTINGS_APP, UPDATE_SETTINGS_SECTION, UPDATE_LAST_CHECK_SETTING, "")
    If Len(Trim$(savedValue)) = 0 Then GoTo CheckNow
    lastCheckValue = CDbl(savedValue)
    CleanupSuiteUpdateCheckIsDue = (DateDiff("d", CDate(lastCheckValue), Now) >= UPDATE_INTERVAL_DAYS)
    Exit Function
CheckNow:
    CleanupSuiteUpdateCheckIsDue = True
End Function

Private Sub SaveCleanupSuiteUpdateAttemptTime()
    On Error Resume Next
    SaveSetting UPDATE_SETTINGS_APP, UPDATE_SETTINGS_SECTION, UPDATE_LAST_CHECK_SETTING, CStr(CDbl(Now))
    On Error GoTo 0
End Sub

Private Function DownloadCleanupSuiteUpdateManifest() As String
    Dim request As Object
    Set request = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    request.setTimeouts 3000, 3000, 5000, 5000
    request.Open "GET", UPDATE_MANIFEST_URL, False
    request.setRequestHeader "User-Agent", "CleanupSuiteForWord/" & SUITE_VERSION
    request.Send
    If CLng(request.Status) <> 200 Then
        Err.Raise vbObjectError + 1702, "DownloadCleanupSuiteUpdateManifest", "HTTP status " & CStr(request.Status)
    End If
    DownloadCleanupSuiteUpdateManifest = CStr(request.responseText)
End Function

Private Function ManifestValue(ByVal manifestText As String, ByVal keyName As String) As String
    Dim lines As Variant
    Dim lineValue As Variant
    Dim equalsAt As Long
    Dim currentKey As String

    lines = Split(Replace$(manifestText, vbCr, vbNullString), vbLf)
    For Each lineValue In lines
        If Len(Trim$(CStr(lineValue))) > 0 Then
            If Left$(Trim$(CStr(lineValue)), 1) <> "#" Then
                equalsAt = InStr(CStr(lineValue), "=")
                If equalsAt > 0 Then
                    currentKey = Trim$(Left$(CStr(lineValue), equalsAt - 1))
                    If StrComp(currentKey, keyName, vbTextCompare) = 0 Then
                        ManifestValue = Trim$(Mid$(CStr(lineValue), equalsAt + 1))
                        Exit Function
                    End If
                End If
            End If
        End If
    Next lineValue
End Function

Private Function CleanupSuiteUpdateUrlIsAllowed(ByVal urlValue As String) As Boolean
    CleanupSuiteUpdateUrlIsAllowed = (StrComp(Left$(urlValue, Len(UPDATE_ALLOWED_URL_PREFIX)), _
                                               UPDATE_ALLOWED_URL_PREFIX, vbTextCompare) = 0)
End Function

Private Sub OpenCleanupSuiteUpdateUrl(ByVal urlValue As String)
    Dim shellObject As Object
    Set shellObject = CreateObject("WScript.Shell")
    shellObject.Run Chr$(34) & urlValue & Chr$(34), 1, False
End Sub

Private Function CompareCleanupSuiteVersions(ByVal leftVersion As String, ByVal rightVersion As String) As Long
    Dim leftCore As Variant
    Dim rightCore As Variant
    Dim i As Long
    Dim leftValue As Long
    Dim rightValue As Long
    Dim leftRank As Long
    Dim rightRank As Long

    leftCore = Split(CleanupVersionCore(leftVersion), ".")
    rightCore = Split(CleanupVersionCore(rightVersion), ".")
    For i = 0 To 2
        leftValue = CleanupVersionNumberAt(leftCore, i)
        rightValue = CleanupVersionNumberAt(rightCore, i)
        If leftValue > rightValue Then CompareCleanupSuiteVersions = 1: Exit Function
        If leftValue < rightValue Then CompareCleanupSuiteVersions = -1: Exit Function
    Next i

    leftRank = CleanupPrereleaseRank(leftVersion)
    rightRank = CleanupPrereleaseRank(rightVersion)
    If leftRank > rightRank Then CompareCleanupSuiteVersions = 1
    If leftRank < rightRank Then CompareCleanupSuiteVersions = -1
End Function

Private Function CleanupVersionCore(ByVal versionText As String) As String
    Dim normalized As String
    normalized = LCase$(Trim$(versionText))
    If Left$(normalized, 1) = "v" Then normalized = Mid$(normalized, 2)
    If InStr(normalized, "-") > 0 Then normalized = Left$(normalized, InStr(normalized, "-") - 1)
    CleanupVersionCore = normalized
End Function

Private Function CleanupVersionNumberAt(ByVal versionParts As Variant, ByVal indexValue As Long) As Long
    On Error GoTo MissingValue
    If indexValue <= UBound(versionParts) Then CleanupVersionNumberAt = CLng(Val(CStr(versionParts(indexValue))))
    Exit Function
MissingValue:
    CleanupVersionNumberAt = 0
End Function

Private Function CleanupPrereleaseRank(ByVal versionText As String) As Long
    Dim normalized As String
    normalized = LCase$(Trim$(versionText))
    If InStr(normalized, "-") = 0 Then CleanupPrereleaseRank = 40: Exit Function
    If InStr(normalized, "rc") > 0 Then CleanupPrereleaseRank = 30: Exit Function
    If InStr(normalized, "beta") > 0 Then CleanupPrereleaseRank = 20: Exit Function
    If InStr(normalized, "alpha") > 0 Then CleanupPrereleaseRank = 10: Exit Function
    CleanupPrereleaseRank = 1
End Function
