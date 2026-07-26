Option Explicit

' Contract 1.0 bridge. The engine analyzes immutable text; Word/VBA retains
' Preview, revalidation, Apply, and Undo ownership.

#If VBA7 Then
Private Declare PtrSafe Function CoCreateGuid Lib "ole32" (ByRef pguid As HYBRID_GUID) As Long
Private Declare PtrSafe Function StringFromGUID2 Lib "ole32" (ByRef rguid As HYBRID_GUID, ByVal lpsz As LongPtr, ByVal cchMax As Long) As Long
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As HYBRID_SYSTEMTIME)
Private Declare PtrSafe Function CreateProcessW Lib "kernel32" (ByVal lpApplicationName As LongPtr, ByVal lpCommandLine As LongPtr, ByVal lpProcessAttributes As LongPtr, ByVal lpThreadAttributes As LongPtr, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As LongPtr, ByVal lpCurrentDirectory As LongPtr, ByRef lpStartupInfo As HYBRID_STARTUPINFO, ByRef lpProcessInformation As HYBRID_PROCESS_INFORMATION) As Long
Private Declare PtrSafe Function CreatePipe Lib "kernel32" (ByRef hReadPipe As LongPtr, ByRef hWritePipe As LongPtr, ByRef lpPipeAttributes As HYBRID_SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare PtrSafe Function SetHandleInformation Lib "kernel32" (ByVal hObject As LongPtr, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As LongPtr, ByRef lpBuffer As Any, ByVal nBufferSize As Long, ByRef lpBytesRead As Long, ByRef lpTotalBytesAvail As Long, ByRef lpBytesLeftThisMessage As Long) As Long
Private Declare PtrSafe Function ReadFile Lib "kernel32" (ByVal hFile As LongPtr, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As LongPtr) As Long
Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As LongPtr, ByVal dwMilliseconds As Long) As Long
Private Declare PtrSafe Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByRef lpExitCode As Long) As Long
Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As LongPtr, ByVal uExitCode As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cbMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
Private Declare PtrSafe Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cbMultiByte As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long) As Long
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt" (ByRef phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptGetProperty Lib "bcrypt" (ByVal hObject As LongPtr, ByVal pszProperty As LongPtr, ByRef pbOutput As Any, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptCreateHash Lib "bcrypt" (ByVal hAlgorithm As LongPtr, ByRef phHash As LongPtr, ByRef pbHashObject As Any, ByVal cbHashObject As Long, ByVal pbSecret As LongPtr, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptHashData Lib "bcrypt" (ByVal hHash As LongPtr, ByRef pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptFinishHash Lib "bcrypt" (ByVal hHash As LongPtr, ByRef pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptDestroyHash Lib "bcrypt" (ByVal hHash As LongPtr) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
#Else
Private Declare Function CoCreateGuid Lib "ole32" (ByRef pguid As HYBRID_GUID) As Long
Private Declare Function StringFromGUID2 Lib "ole32" (ByRef rguid As HYBRID_GUID, ByVal lpsz As Long, ByVal cchMax As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As HYBRID_SYSTEMTIME)
Private Declare Function CreateProcessW Lib "kernel32" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, ByRef lpStartupInfo As HYBRID_STARTUPINFO, ByRef lpProcessInformation As HYBRID_PROCESS_INFORMATION) As Long
Private Declare Function CreatePipe Lib "kernel32" (ByRef hReadPipe As Long, ByRef hWritePipe As Long, ByRef lpPipeAttributes As HYBRID_SECURITY_ATTRIBUTES, ByVal nSize As Long) As Long
Private Declare Function SetHandleInformation Lib "kernel32" (ByVal hObject As Long, ByVal dwMask As Long, ByVal dwFlags As Long) As Long
Private Declare Function PeekNamedPipe Lib "kernel32" (ByVal hNamedPipe As Long, ByRef lpBuffer As Any, ByVal nBufferSize As Long, ByRef lpBytesRead As Long, ByRef lpTotalBytesAvail As Long, ByRef lpBytesLeftThisMessage As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByRef lpMultiByteStr As Any, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (ByRef phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptGetProperty Lib "bcrypt" (ByVal hObject As Long, ByVal pszProperty As Long, ByRef pbOutput As Any, ByVal cbOutput As Long, ByRef pcbResult As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCreateHash Lib "bcrypt" (ByVal hAlgorithm As Long, ByRef phHash As Long, ByRef pbHashObject As Any, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptHashData Lib "bcrypt" (ByVal hHash As Long, ByRef pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptFinishHash Lib "bcrypt" (ByVal hHash As Long, ByRef pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyHash Lib "bcrypt" (ByVal hHash As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long
#End If

Private Type HYBRID_GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
Private Type HYBRID_SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As LongPtr
    bInheritHandle As Long
End Type
#Else
Private Type HYBRID_SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
#End If

#If VBA7 Then
Private Type HYBRID_STARTUPINFO
    cb As Long
    lpReserved As LongPtr
    lpDesktop As LongPtr
    lpTitle As LongPtr
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As LongPtr
    hStdInput As LongPtr
    hStdOutput As LongPtr
    hStdError As LongPtr
End Type

Private Type HYBRID_PROCESS_INFORMATION
    hProcess As LongPtr
    hThread As LongPtr
    dwProcessId As Long
    dwThreadId As Long
End Type
#Else
Private Type HYBRID_STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type HYBRID_PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
#End If

Private Type HYBRID_SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Const HYBRID_CONTRACT_VERSION As String = "1.0"
Private Const HYBRID_PROTOCOL_VERSION As String = "1.0"
Private Const HYBRID_ENGINE_ID As String = "com.masseyslab.cleanupsuite.engine"
Private Const HYBRID_ENGINE_VERSION As String = "0.2.0"
Private Const HYBRID_UNICODE_TOOL_ID As String = "invisible-unicode-cleaner"
Private Const HYBRID_UNICODE_TOOL_VERSION As String = "1.0.0"
Private Const HYBRID_UNICODE_MODE As String = "selected-characters"
Private Const HYBRID_CREATE_NO_WINDOW As Long = &H8000000
Private Const HYBRID_STARTF_USESTDHANDLES As Long = &H100
Private Const HYBRID_HANDLE_FLAG_INHERIT As Long = 1
Private Const HYBRID_WAIT_OBJECT_0 As Long = 0
Private Const HYBRID_WAIT_TIMEOUT As Long = &H102
Private Const HYBRID_CP_UTF8 As Long = 65001
Private Const HYBRID_WC_ERR_INVALID_CHARS As Long = &H80
Private Const HYBRID_MB_ERR_INVALID_CHARS As Long = &H8
Private Const HYBRID_MAX_RESULT_BYTES As Long = 134217728
Private Const HYBRID_ANALYSIS_TIMEOUT_SECONDS As Long = 90

#If VBA7 Then
Private mHybridShaAlgorithm As LongPtr
#Else
Private mHybridShaAlgorithm As Long
#End If
Private mHybridShaObjectLength As Long
Private mHybridShaDigestLength As Long

Public Function HybridBridgeSelfTest(ByRef details As String) As Boolean
    Dim parsed As Object
    Dim nested As Object
    Dim items As Collection
    Dim bytes() As Byte
    Dim byteCount As Long
    Dim uuid As String
    Dim stepName As String
    Dim rejectedUnpairedSurrogate As Boolean

    On Error GoTo SelfTestFailed
    stepName = "strict JSON"
    Set parsed = HybridJsonParseObject("{""name"":""CleanupSuite"",""active"":true,""nested"":{""count"":2},""items"":[1,""\u2011""]}")
    HybridRequireExactKeys parsed, "name|active|nested|items"
    HybridRequireString parsed, "name", "CleanupSuite"
    If Not HybridRequireBoolean(parsed, "active") Then Err.Raise 5, "HybridBridgeSelfTest", "JSON Boolean parsing failed."
    Set nested = HybridRequireObject(parsed, "nested")
    If HybridRequireLong(nested, "count") <> 2 Then Err.Raise 5, "HybridBridgeSelfTest", "JSON integer parsing failed."
    Set items = HybridRequireCollection(parsed, "items")
    If items.Count <> 2 Or CStr(items(2)) <> ChrW$(&H2011) Then Err.Raise 5, "HybridBridgeSelfTest", "JSON array or Unicode parsing failed."
    Set parsed = HybridJsonParseObject("{""emoji"":""\uD83D\uDE00""}")
    If CStr(parsed("emoji")) <> ChrW$(CLng(&HD83D)) & ChrW$(CLng(&HDE00)) Then Err.Raise 5, "HybridBridgeSelfTest", "Escaped supplementary Unicode parsing failed."
    On Error Resume Next
    Err.Clear
    Set parsed = HybridJsonParseObject("{""invalid"":""\uD83D""}")
    rejectedUnpairedSurrogate = (Err.Number <> 0)
    Err.Clear
    On Error GoTo SelfTestFailed
    If Not rejectedUnpairedSurrogate Then Err.Raise 5, "HybridBridgeSelfTest", "An unpaired JSON surrogate was accepted."

    stepName = "SHA-256"
    If HybridSha256Utf8("abc") <> "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad" Then Err.Raise 5, "HybridBridgeSelfTest", "SHA-256 validation failed."
    stepName = "supplementary UTF-8"
    bytes = HybridUtf8Bytes("A" & ChrW$(CLng(&HD83D)) & ChrW$(CLng(&HDE00)) & "B", byteCount)
    If byteCount <> 6 Then Err.Raise 5, "HybridBridgeSelfTest", "Supplementary Unicode UTF-8 encoding failed."
    stepName = "UUID"
    uuid = HybridNewUuid()
    If Len(uuid) <> 36 Or LCase$(uuid) <> uuid Then Err.Raise 5, "HybridBridgeSelfTest", "UUID generation failed."

    details = "Strict JSON, UTF-8, SHA-256, supplementary Unicode, and UUID checks passed."
    HybridBridgeSelfTest = True
    Exit Function

SelfTestFailed:
    details = stepName & ": " & Err.Description
End Function

Public Function HybridBridgeEngineSelfTest(ByRef details As String) As Boolean
    Dim testDocument As Document
    Dim targetRange As Range
    Dim analysis As Object
    Dim failureMessage As String
    Dim candidates As Collection
    Dim analysisSucceeded As Boolean

    On Error GoTo SelfTestFailed
    Set testDocument = Documents.Add
    testDocument.Content.Text = _
        "A" & ChrW$(&HA0) & _
        "B" & ChrW$(&H200B) & _
        "C" & ChrW$(&H200C) & _
        "D" & ChrW$(&H200D) & _
        "E" & ChrW$(CLng(&HFEFF)) & _
        "F" & ChrW$(&HAD) & _
        "G" & ChrW$(&H2011) & "H"
    Set targetRange = testDocument.Content.Duplicate

    analysisSucceeded = HybridAnalyzeInvisibleUnicode( _
            targetRange, False, _
            True, True, True, True, True, True, True, _
            analysis, failureMessage)
    If Not analysisSucceeded Then
        Err.Raise 5, "HybridBridgeEngineSelfTest", failureMessage
    End If
    Set candidates = analysis("candidates")
    If candidates.Count <> 7 Then Err.Raise 5, "HybridBridgeEngineSelfTest", "The engine did not return all seven deterministic Unicode candidates."
    If Not HybridRevalidateInvisibleUnicode(analysis, targetRange, failureMessage) Then Err.Raise 5, "HybridBridgeEngineSelfTest", failureMessage

    testDocument.Range(0, 0).InsertBefore "changed"
    Set targetRange = testDocument.Range(CLng(analysis("scopeStart")), CLng(analysis("scopeEnd")))
    If HybridRevalidateInvisibleUnicode(analysis, targetRange, failureMessage) Then Err.Raise 5, "HybridBridgeEngineSelfTest", "Stale analysis was not rejected."

    details = "Word-to-engine analysis returned seven candidates and rejected a stale scope before Apply."
    testDocument.Close wdDoNotSaveChanges
    Set testDocument = Nothing
    HybridBridgeEngineSelfTest = True
    Exit Function

SelfTestFailed:
    details = Err.Description
    On Error Resume Next
    If Not testDocument Is Nothing Then testDocument.Close wdDoNotSaveChanges
    On Error GoTo 0
End Function

Public Function HybridBridgeApplySelfTest(ByRef details As String) As Boolean
    Dim testDocument As Document
    Dim targetRange As Range
    Dim analysis As Object
    Dim failureMessage As String
    Dim candidates As Collection
    Dim candidate As Object
    Dim candidateRange As Range
    Dim originalText As String
    Dim expectedText As String
    Dim index As Long
    Dim undoRecord As UndoRecord
    Dim analysisSucceeded As Boolean

    On Error GoTo SelfTestFailed
    Set testDocument = Documents.Add
    testDocument.Content.Text = _
        "A" & ChrW$(&HA0) & _
        "B" & ChrW$(&H200B) & _
        "C" & ChrW$(&H200C) & _
        "D" & ChrW$(&H200D) & _
        "E" & ChrW$(CLng(&HFEFF)) & _
        "F" & ChrW$(&HAD) & _
        "G" & ChrW$(&H2011) & "H"
    Set targetRange = testDocument.Content.Duplicate
    originalText = targetRange.Text
    expectedText = "A BCDEF-G-H" & vbCr

    analysisSucceeded = HybridAnalyzeInvisibleUnicode( _
        targetRange, False, _
        True, True, True, True, True, True, True, _
        analysis, failureMessage)
    If Not analysisSucceeded Then Err.Raise 5, "HybridBridgeApplySelfTest", failureMessage
    If Not HybridRevalidateInvisibleUnicode(analysis, targetRange, failureMessage) Then Err.Raise 5, "HybridBridgeApplySelfTest", failureMessage

    Set candidates = analysis("candidates")
    Set undoRecord = Application.UndoRecord
    undoRecord.StartCustomRecord "CleanupSuite Hybrid Apply Self-Test"
    For index = candidates.Count To 1 Step -1
        Set candidate = candidates(index)
        Set candidateRange = HybridCandidateAbsoluteRange(analysis, candidate, testDocument)
        candidateRange.Text = HybridCandidateReplacement(candidate)
    Next index
    undoRecord.EndCustomRecord
    If testDocument.Content.Text <> expectedText Then Err.Raise 5, "HybridBridgeApplySelfTest", "The reverse-order Apply result is incorrect."

    testDocument.Undo
    If testDocument.Content.Text <> originalText Then Err.Raise 5, "HybridBridgeApplySelfTest", "One Word Undo did not restore the original Unicode text."

    details = "Seven candidates applied bottom-up and one Word Undo restored the exact original text."
    testDocument.Close wdDoNotSaveChanges
    Set testDocument = Nothing
    HybridBridgeApplySelfTest = True
    Exit Function

SelfTestFailed:
    details = Err.Description
    On Error Resume Next
    undoRecord.EndCustomRecord
    If Not testDocument Is Nothing Then testDocument.Close wdDoNotSaveChanges
    On Error GoTo 0
End Function

Public Function HybridBridgePerformanceSelfTest(ByRef details As String) As Boolean
    Dim testDocument As Document
    Dim targetRange As Range
    Dim analysis As Object
    Dim failureMessage As String
    Dim candidates As Collection
    Dim snapshotText As String
    Dim blockText As String
    Dim index As Long
    Dim startedAt As Single
    Dim elapsed As Double
    Dim analysisSucceeded As Boolean

    On Error GoTo SelfTestFailed
    blockText = String$(500, "x") & ChrW$(&HA0)
    For index = 1 To 400
        snapshotText = snapshotText & blockText
    Next index
    Set testDocument = Documents.Add
    testDocument.Content.Text = snapshotText
    Set targetRange = testDocument.Content.Duplicate

    startedAt = Timer
    analysisSucceeded = HybridAnalyzeInvisibleUnicode( _
        targetRange, False, _
        True, False, False, False, False, False, False, _
        analysis, failureMessage)
    If Not analysisSucceeded Then
        Err.Raise 5, "HybridBridgePerformanceSelfTest", failureMessage
    End If
    elapsed = Timer - startedAt
    If elapsed < 0 Then elapsed = elapsed + 86400#
    Set candidates = analysis("candidates")
    If candidates.Count <> 400 Then Err.Raise 5, "HybridBridgePerformanceSelfTest", "The performance fixture candidate count is incorrect."
    If Not HybridRevalidateInvisibleUnicode(analysis, targetRange, failureMessage) Then Err.Raise 5, "HybridBridgePerformanceSelfTest", failureMessage
    If elapsed > 30# Then Err.Raise 5, "HybridBridgePerformanceSelfTest", "The 200,000-character bridge regression exceeded 30 seconds."

    details = "200,000 characters and 400 candidates analyzed and revalidated in " & Format$(elapsed, "0.00") & " seconds."
    testDocument.Close wdDoNotSaveChanges
    Set testDocument = Nothing
    HybridBridgePerformanceSelfTest = True
    Exit Function

SelfTestFailed:
    details = Err.Description
    On Error Resume Next
    If Not testDocument Is Nothing Then testDocument.Close wdDoNotSaveChanges
    On Error GoTo 0
End Function

Public Function HybridAnalyzeInvisibleUnicode( _
        ByVal targetRange As Range, _
        ByVal selectionOnly As Boolean, _
        ByVal nonBreakingSpace As Boolean, _
        ByVal zeroWidthSpace As Boolean, _
        ByVal zeroWidthNonJoiner As Boolean, _
        ByVal zeroWidthJoiner As Boolean, _
        ByVal byteOrderMark As Boolean, _
        ByVal softHyphen As Boolean, _
        ByVal nonBreakingHyphen As Boolean, _
        ByRef analysis As Object, _
        ByRef failureMessage As String) As Boolean
    Dim enginePath As String
    Dim expectedEngineHash As String
    Dim jobId As String
    Dim jobDirectory As String
    Dim snapshotText As String
    Dim snapshotBytes() As Byte
    Dim snapshotByteCount As Long
    Dim snapshotHash As String
    Dim requestText As String
    Dim exitCode As Long
    Dim resultText As String
    Dim parsed As Object

    On Error GoTo AnalysisFailed
    failureMessage = vbNullString
    If targetRange Is Nothing Then Err.Raise 5, "HybridAnalyzeInvisibleUnicode", "The analysis range is unavailable."
    If targetRange.StoryType <> wdMainTextStory Then Err.Raise 5, "HybridAnalyzeInvisibleUnicode", "The pilot currently supports the main document story only."
    If Not (nonBreakingSpace Or zeroWidthSpace Or zeroWidthNonJoiner Or zeroWidthJoiner Or byteOrderMark Or softHyphen Or nonBreakingHyphen) Then Err.Raise 5, "HybridAnalyzeInvisibleUnicode", "No Unicode character type is selected."
    If Not HybridResolveTrustedDevelopmentEngine(enginePath, expectedEngineHash, failureMessage) Then Exit Function

    jobId = HybridNewUuid()
    jobDirectory = HybridCreateEmptyJobDirectory(jobId)
    exitCode = HybridRunEngine(enginePath, "--prepare-job", jobDirectory, 10)
    If exitCode <> 0 Then Err.Raise 5, "HybridAnalyzeInvisibleUnicode", "CleanupSuite could not secure the analysis workspace."

    snapshotText = targetRange.Text
    snapshotBytes = HybridUtf8Bytes(snapshotText, snapshotByteCount)
    snapshotHash = HybridSha256Bytes(snapshotBytes, snapshotByteCount)
    HybridWriteBytesAtomic jobDirectory & "\document.utf8.txt", snapshotBytes, snapshotByteCount

    requestText = HybridBuildUnicodeRequest( _
        jobId, targetRange, selectionOnly, snapshotHash, _
        snapshotByteCount, Len(snapshotText), _
        nonBreakingSpace, zeroWidthSpace, zeroWidthNonJoiner, _
        zeroWidthJoiner, byteOrderMark, softHyphen, nonBreakingHyphen)
    HybridWriteTextUtf8Atomic jobDirectory & "\request.json", requestText

    exitCode = HybridRunEngine(enginePath, "--analyze", jobDirectory, HYBRID_ANALYSIS_TIMEOUT_SECONDS)
    If exitCode <> 0 Then Err.Raise 5, "HybridAnalyzeInvisibleUnicode", "The Unicode analysis did not complete safely."
    resultText = HybridReadUtf8File(jobDirectory & "\result.json", HYBRID_MAX_RESULT_BYTES)
    Set parsed = HybridJsonParseObject(resultText)
    HybridValidateUnicodeResponse parsed, jobId, snapshotHash

    If HybridSha256Utf8(targetRange.Text) <> snapshotHash Then Err.Raise 5, "HybridAnalyzeInvisibleUnicode", "The document changed during analysis. Run Preview again."

    Set analysis = CreateObject("Scripting.Dictionary")
    analysis.CompareMode = vbBinaryCompare
    analysis.Add "enginePath", enginePath
    analysis.Add "engineSha256", expectedEngineHash
    analysis.Add "snapshotSha256", snapshotHash
    analysis.Add "scopeStart", targetRange.Start
    analysis.Add "scopeEnd", targetRange.End
    analysis.Add "selectionOnly", selectionOnly
    analysis.Add "candidates", parsed("candidates")
    analysis.Add "summary", parsed("summary")
    analysis.Add "jobId", jobId

    HybridDeleteJobDirectory jobDirectory
    HybridAnalyzeInvisibleUnicode = True
    Exit Function

AnalysisFailed:
    failureMessage = Err.Description
    On Error Resume Next
    If Len(jobDirectory) > 0 Then HybridDeleteJobDirectory jobDirectory
    On Error GoTo 0
End Function

Public Function HybridRevalidateInvisibleUnicode( _
        ByVal analysis As Object, _
        ByVal currentRange As Range, _
        ByRef failureMessage As String) As Boolean
    Dim currentText As String
    Dim expectedHash As String
    Dim candidates As Collection
    Dim candidate As Object
    Dim fingerprint As Object
    Dim location As Object
    Dim relativeStart As Long
    Dim relativeEnd As Long
    Dim prefixLength As Long
    Dim suffixLength As Long
    Dim exactText As String
    Dim prefixText As String
    Dim suffixText As String
    Dim paragraphText As String
    Dim paragraphKey As String
    Dim paragraphHashes As Object

    On Error GoTo RevalidationFailed
    failureMessage = vbNullString
    If analysis Is Nothing Or currentRange Is Nothing Then Err.Raise 5, "HybridRevalidateInvisibleUnicode", "The saved Preview analysis is unavailable."
    If currentRange.Start <> CLng(analysis("scopeStart")) Or currentRange.End <> CLng(analysis("scopeEnd")) Then Err.Raise 5, "HybridRevalidateInvisibleUnicode", "The analyzed range moved or changed. Run Preview again."

    currentText = currentRange.Text
    expectedHash = CStr(analysis("snapshotSha256"))
    If HybridSha256Utf8(currentText) <> expectedHash Then Err.Raise 5, "HybridRevalidateInvisibleUnicode", "The document changed after Preview. No changes were applied; run Preview again."

    Set candidates = analysis("candidates")
    Set paragraphHashes = CreateObject("Scripting.Dictionary")
    paragraphHashes.CompareMode = vbBinaryCompare
    For Each candidate In candidates
        Set location = candidate("location")
        Set fingerprint = candidate("fingerprint")
        relativeStart = CLng(location("startUtf16"))
        relativeEnd = CLng(location("endUtf16"))
        If relativeStart < 0 Or relativeEnd <= relativeStart Or relativeEnd > Len(currentText) Then Err.Raise 5, "HybridRevalidateInvisibleUnicode", "A candidate range is no longer valid."

        exactText = Mid$(currentText, relativeStart + 1, relativeEnd - relativeStart)
        If HybridSha256Utf8(exactText) <> CStr(fingerprint("exactTextSha256")) Then Err.Raise 5, "HybridRevalidateInvisibleUnicode", "A candidate changed after Preview. No changes were applied."

        prefixLength = CLng(fingerprint("prefixLengthUtf16"))
        suffixLength = CLng(fingerprint("suffixLengthUtf16"))
        If prefixLength > relativeStart Or relativeEnd + suffixLength > Len(currentText) Then Err.Raise 5, "HybridRevalidateInvisibleUnicode", "A candidate context is no longer valid."
        prefixText = Mid$(currentText, relativeStart - prefixLength + 1, prefixLength)
        suffixText = Mid$(currentText, relativeEnd + 1, suffixLength)
        If HybridSha256Utf8(prefixText) <> CStr(fingerprint("prefixSha256")) Then Err.Raise 5, "HybridRevalidateInvisibleUnicode", "Text before a candidate changed after Preview."
        If HybridSha256Utf8(suffixText) <> CStr(fingerprint("suffixSha256")) Then Err.Raise 5, "HybridRevalidateInvisibleUnicode", "Text after a candidate changed after Preview."

        paragraphText = HybridSnapshotParagraph(currentText, relativeStart, relativeEnd, paragraphKey)
        If Not paragraphHashes.Exists(paragraphKey) Then paragraphHashes.Add paragraphKey, HybridSha256Utf8(paragraphText)
        If CStr(paragraphHashes(paragraphKey)) <> CStr(fingerprint("paragraphSha256")) Then Err.Raise 5, "HybridRevalidateInvisibleUnicode", "A candidate paragraph changed after Preview."
    Next candidate

    HybridRevalidateInvisibleUnicode = True
    Exit Function

RevalidationFailed:
    failureMessage = Err.Description
End Function

Public Function HybridCandidateAbsoluteRange( _
        ByVal analysis As Object, _
        ByVal candidate As Object, _
        ByVal document As Document) As Range
    Dim location As Object
    Set location = candidate("location")
    Set HybridCandidateAbsoluteRange = document.Range( _
        CLng(analysis("scopeStart")) + CLng(location("startUtf16")), _
        CLng(analysis("scopeStart")) + CLng(location("endUtf16")))
End Function

Public Function HybridCandidateReplacement(ByVal candidate As Object) As String
    Dim operation As Object
    Dim parameters As Object
    Set operation = candidate("operation")
    Set parameters = operation("parameters")
    HybridCandidateReplacement = CStr(parameters("replacementText"))
End Function

Private Function HybridBuildUnicodeRequest( _
        ByVal jobId As String, _
        ByVal targetRange As Range, _
        ByVal selectionOnly As Boolean, _
        ByVal snapshotHash As String, _
        ByVal snapshotByteCount As Long, _
        ByVal snapshotUtf16Length As Long, _
        ByVal nonBreakingSpace As Boolean, _
        ByVal zeroWidthSpace As Boolean, _
        ByVal zeroWidthNonJoiner As Boolean, _
        ByVal zeroWidthJoiner As Boolean, _
        ByVal byteOrderMark As Boolean, _
        ByVal softHyphen As Boolean, _
        ByVal nonBreakingHyphen As Boolean) As String
    Dim sessionId As String
    Dim documentSessionId As String
    Dim requestText As String

    sessionId = HybridNewUuid()
    documentSessionId = HybridNewUuid()
    requestText = _
        "{" & _
        """contractVersion"":""1.0""," & _
        """messageType"":""analysis-request""," & _
        """jobId"":" & HybridJsonEscape(jobId) & "," & _
        """createdUtc"":" & HybridJsonEscape(HybridUtcTimestamp()) & "," & _
        """client"":{" & _
            """suiteVersion"":" & HybridJsonEscape(SUITE_VERSION) & "," & _
            """protocolVersion"":""1.0""," & _
            """processId"":" & CStr(GetCurrentProcessId()) & "," & _
            """sessionId"":" & HybridJsonEscape(sessionId) & "},"
    requestText = requestText & _
        """tool"":{" & _
            """id"":""invisible-unicode-cleaner""," & _
            """definitionVersion"":""1.0.0""," & _
            """analysisMode"":""selected-characters""},"
    requestText = requestText & _
        """scope"":{" & _
            """documentSessionId"":" & HybridJsonEscape(documentSessionId) & "," & _
            """storyType"":""main-text""," & _
            """startUtf16"":" & CStr(targetRange.Start) & "," & _
            """endUtf16"":" & CStr(targetRange.End) & "," & _
            """selectionOnly"":" & HybridJsonBoolean(selectionOnly) & "},"
    requestText = requestText & _
        """snapshot"":{" & _
            """snapshotId"":" & HybridJsonEscape(snapshotHash) & "," & _
            """contentFile"":""document.utf8.txt""," & _
            """encoding"":""utf-8""," & _
            """byteLength"":" & CStr(snapshotByteCount) & "," & _
            """utf16Length"":" & CStr(snapshotUtf16Length) & "," & _
            """sha256"":" & HybridJsonEscape(snapshotHash) & "," & _
            """lineEndingPolicy"":""preserve-word-story-text""},"
    requestText = requestText & _
        """options"":{" & _
            """nonBreakingSpace"":" & HybridJsonBoolean(nonBreakingSpace) & "," & _
            """zeroWidthSpace"":" & HybridJsonBoolean(zeroWidthSpace) & "," & _
            """zeroWidthNonJoiner"":" & HybridJsonBoolean(zeroWidthNonJoiner) & "," & _
            """zeroWidthJoiner"":" & HybridJsonBoolean(zeroWidthJoiner) & "," & _
            """byteOrderMark"":" & HybridJsonBoolean(byteOrderMark) & "," & _
            """softHyphen"":" & HybridJsonBoolean(softHyphen) & "," & _
            """nonBreakingHyphen"":" & HybridJsonBoolean(nonBreakingHyphen) & "},"
    requestText = requestText & _
        """requestedCapabilities"":[""analysis.invisible-unicode"",""fingerprint.sha256-utf8-exact""]," & _
        """privacy"":{""documentPathIncluded"":false,""documentNameIncluded"":false}" & _
        "}"
    HybridBuildUnicodeRequest = requestText
End Function

Private Sub HybridValidateUnicodeResponse( _
        ByVal root As Object, _
        ByVal expectedJobId As String, _
        ByVal expectedSnapshotHash As String)
    Dim engine As Object
    Dim echo As Object
    Dim summary As Object
    Dim diagnostics As Object
    Dim candidates As Collection
    Dim candidate As Object
    Dim ordinal As Long

    HybridRequireExactKeys root, "contractVersion|messageType|jobId|completedUtc|engine|echo|status|candidates|summary|diagnostics"
    HybridRequireString root, "contractVersion", HYBRID_CONTRACT_VERSION
    HybridRequireString root, "messageType", "analysis-response"
    HybridRequireString root, "jobId", expectedJobId
    HybridRequireString root, "status", "completed"

    Set engine = HybridRequireObject(root, "engine")
    HybridRequireExactKeys engine, "id|version|protocolVersion"
    HybridRequireString engine, "id", HYBRID_ENGINE_ID
    HybridRequireString engine, "version", HYBRID_ENGINE_VERSION
    HybridRequireString engine, "protocolVersion", HYBRID_PROTOCOL_VERSION

    Set echo = HybridRequireObject(root, "echo")
    HybridRequireExactKeys echo, "snapshotId|toolId|toolDefinitionVersion"
    HybridRequireString echo, "snapshotId", expectedSnapshotHash
    HybridRequireString echo, "toolId", HYBRID_UNICODE_TOOL_ID
    HybridRequireString echo, "toolDefinitionVersion", HYBRID_UNICODE_TOOL_VERSION

    Set candidates = HybridRequireCollection(root, "candidates")
    ordinal = 0
    For Each candidate In candidates
        ordinal = ordinal + 1
        HybridValidateUnicodeCandidate candidate, ordinal, expectedSnapshotHash
    Next candidate

    Set summary = HybridRequireObject(root, "summary")
    HybridRequireExactKeys summary, "total|applicable|reviewOnly|protected|skipped"
    If HybridRequireLong(summary, "total") <> candidates.Count Then Err.Raise 5, "HybridValidateUnicodeResponse", "The engine summary does not match its candidates."
    If HybridRequireLong(summary, "applicable") <> candidates.Count Then Err.Raise 5, "HybridValidateUnicodeResponse", "The engine returned a non-applicable Unicode candidate."
    If HybridRequireLong(summary, "reviewOnly") <> 0 Or HybridRequireLong(summary, "protected") <> 0 Or HybridRequireLong(summary, "skipped") <> 0 Then Err.Raise 5, "HybridValidateUnicodeResponse", "The Unicode pilot returned an unsupported candidate state."

    Set diagnostics = HybridRequireObject(root, "diagnostics")
    HybridRequireExactKeys diagnostics, "durationMs|warningCodes|logContainsDocumentContent"
    If HybridRequireLong(diagnostics, "durationMs") < 0 Then Err.Raise 5, "HybridValidateUnicodeResponse", "The engine duration is invalid."
    Dim warningCodes As Collection
    Set warningCodes = HybridRequireCollection(diagnostics, "warningCodes")
    If warningCodes.Count <> 0 Then Err.Raise 5, "HybridValidateUnicodeResponse", "The engine returned an unsupported warning."
    If HybridRequireBoolean(diagnostics, "logContainsDocumentContent") Then Err.Raise 5, "HybridValidateUnicodeResponse", "The engine reported unsafe diagnostic logging."
End Sub

Private Sub HybridValidateUnicodeCandidate( _
        ByVal candidate As Object, _
        ByVal ordinal As Long, _
        ByVal expectedSnapshotHash As String)
    Dim location As Object
    Dim fingerprint As Object
    Dim operation As Object
    Dim parameters As Object
    Dim revalidation As Object
    Dim confidence As Object
    Dim display As Object
    Dim reasonCode As String
    Dim replacementText As String

    HybridRequireExactKeys candidate, "candidateId|state|reasonCode|location|fingerprint|operation|revalidation|confidence|display"
    HybridRequireString candidate, "candidateId", "c" & Format$(ordinal, "000000")
    HybridRequireString candidate, "state", "applicable"
    reasonCode = HybridRequireString(candidate, "reasonCode")

    Set location = HybridRequireObject(candidate, "location")
    HybridRequireExactKeys location, "storyType|startUtf16|endUtf16|offsetBasis"
    HybridRequireString location, "storyType", "main-text"
    If HybridRequireLong(location, "startUtf16") < 0 Then Err.Raise 5, "HybridValidateUnicodeCandidate", "A candidate start is invalid."
    If HybridRequireLong(location, "endUtf16") - HybridRequireLong(location, "startUtf16") <> 1 Then Err.Raise 5, "HybridValidateUnicodeCandidate", "A Unicode candidate must cover one UTF-16 code unit."
    HybridRequireString location, "offsetBasis", "snapshot-relative-utf16-code-units"

    Set fingerprint = HybridRequireObject(candidate, "fingerprint")
    HybridRequireExactKeys fingerprint, "algorithm|snapshotId|exactTextSha256|prefixLengthUtf16|prefixSha256|suffixLengthUtf16|suffixSha256|paragraphSha256"
    HybridRequireString fingerprint, "algorithm", "sha256-utf8-exact"
    HybridRequireString fingerprint, "snapshotId", expectedSnapshotHash
    HybridRequireSha256 fingerprint, "exactTextSha256"
    If HybridRequireLong(fingerprint, "prefixLengthUtf16") < 0 Or HybridRequireLong(fingerprint, "prefixLengthUtf16") > 32 Then Err.Raise 5, "HybridValidateUnicodeCandidate", "A candidate prefix length is invalid."
    HybridRequireSha256 fingerprint, "prefixSha256"
    If HybridRequireLong(fingerprint, "suffixLengthUtf16") < 0 Or HybridRequireLong(fingerprint, "suffixLengthUtf16") > 32 Then Err.Raise 5, "HybridValidateUnicodeCandidate", "A candidate suffix length is invalid."
    HybridRequireSha256 fingerprint, "suffixSha256"
    HybridRequireSha256 fingerprint, "paragraphSha256"

    Set operation = HybridRequireObject(candidate, "operation")
    HybridRequireExactKeys operation, "type|safetyClass|parameters"
    HybridRequireString operation, "type", "replaceText"
    HybridRequireString operation, "safetyClass", "textual"
    Set parameters = HybridRequireObject(operation, "parameters")
    HybridRequireExactKeys parameters, "replacementText"
    replacementText = HybridRequireString(parameters, "replacementText")
    HybridValidateUnicodeReasonAndReplacement reasonCode, replacementText

    Set revalidation = HybridRequireObject(candidate, "revalidation")
    HybridRequireExactKeys revalidation, "requireWholeScopeSnapshot|requireExactRange|requireContext|requireStructure|onMismatch|allowRelocation"
    If Not HybridRequireBoolean(revalidation, "requireWholeScopeSnapshot") Then Err.Raise 5, "HybridValidateUnicodeCandidate", "Whole-scope revalidation is required."
    If Not HybridRequireBoolean(revalidation, "requireExactRange") Then Err.Raise 5, "HybridValidateUnicodeCandidate", "Exact-range revalidation is required."
    If Not HybridRequireBoolean(revalidation, "requireContext") Then Err.Raise 5, "HybridValidateUnicodeCandidate", "Context revalidation is required."
    If HybridRequireBoolean(revalidation, "requireStructure") Then Err.Raise 5, "HybridValidateUnicodeCandidate", "The Unicode pilot must not require structural mutation."
    HybridRequireString revalidation, "onMismatch", "abort-apply"
    If HybridRequireBoolean(revalidation, "allowRelocation") Then Err.Raise 5, "HybridValidateUnicodeCandidate", "Candidate relocation is forbidden."

    Set confidence = HybridRequireObject(candidate, "confidence")
    HybridRequireExactKeys confidence, "tier|score"
    HybridRequireString confidence, "tier", "deterministic"
    If CDbl(confidence("score")) <> 1# Then Err.Raise 5, "HybridValidateUnicodeCandidate", "A Unicode candidate must be deterministic."

    Set display = HybridRequireObject(candidate, "display")
    HybridRequireExactKeys display, "label|canHighlight"
    If Len(HybridRequireString(display, "label")) = 0 Then Err.Raise 5, "HybridValidateUnicodeCandidate", "A candidate label is missing."
    If Not HybridRequireBoolean(display, "canHighlight") Then Err.Raise 5, "HybridValidateUnicodeCandidate", "A Unicode candidate must support Preview highlighting."
End Sub

Private Sub HybridValidateUnicodeReasonAndReplacement(ByVal reasonCode As String, ByVal replacementText As String)
    Select Case reasonCode
        Case "unicode.non-breaking-space"
            If replacementText <> " " Then Err.Raise 5, "HybridValidateUnicodeReasonAndReplacement", "The non-breaking-space replacement is invalid."
        Case "unicode.zero-width-space", "unicode.zero-width-nonjoiner", "unicode.zero-width-joiner", "unicode.byte-order-mark"
            If replacementText <> vbNullString Then Err.Raise 5, "HybridValidateUnicodeReasonAndReplacement", "A zero-width replacement is invalid."
        Case "unicode.soft-hyphen", "unicode.non-breaking-hyphen"
            If replacementText <> "-" Then Err.Raise 5, "HybridValidateUnicodeReasonAndReplacement", "A hyphen replacement is invalid."
        Case Else
            Err.Raise 5, "HybridValidateUnicodeReasonAndReplacement", "The engine returned an unknown Unicode reason."
    End Select
End Sub

Private Function HybridResolveTrustedDevelopmentEngine( _
        ByRef enginePath As String, _
        ByRef expectedHash As String, _
        ByRef failureMessage As String) As Boolean
    enginePath = Trim$(Environ$("CLEANUPSUITE_HYBRID_DEV_ENGINE"))
    expectedHash = LCase$(Trim$(Environ$("CLEANUPSUITE_HYBRID_DEV_SHA256")))
    If Len(enginePath) = 0 Or Len(expectedHash) = 0 Then
        failureMessage = "The matched CleanupSuite analysis engine is not installed. Use Setup Repair after the hybrid installer is available."
        Exit Function
    End If
    If Mid$(enginePath, 2, 2) <> ":\" Or InStr(enginePath, "..") > 0 Or InStr(enginePath, """") > 0 Or Len(Dir$(enginePath, vbNormal)) = 0 Then
        failureMessage = "The configured development engine path is invalid."
        Exit Function
    End If
    If LCase$(Right$(enginePath, 4)) <> ".exe" Or Not HybridIsLowercaseSha256(expectedHash) Then
        failureMessage = "The configured development engine identity is invalid."
        Exit Function
    End If
    If HybridSha256File(enginePath) <> expectedHash Then
        failureMessage = "The analysis engine hash does not match the trusted development value."
        Exit Function
    End If
    If Not HybridValidateCapabilityHandshake(enginePath, failureMessage) Then Exit Function
    HybridResolveTrustedDevelopmentEngine = True
End Function

Private Function HybridValidateCapabilityHandshake( _
        ByVal enginePath As String, _
        ByRef failureMessage As String) As Boolean
    Dim capabilityText As String
    Dim root As Object
    Dim engine As Object
    Dim protocolRange As Object
    Dim supportedTools As Collection
    Dim supportedOperations As Collection
    Dim security As Object
    Dim transport As Object
    Dim distribution As Object
    Dim tool As Object
    Dim foundTool As Boolean

    On Error GoTo CapabilityFailed
    capabilityText = HybridCaptureEngineOutput(enginePath, "--capabilities", 10)
    Set root = HybridJsonParseObject(capabilityText)
    HybridRequireExactKeys root, "contractVersion|messageType|engine|protocolRange|supportedTools|supportedOperations|security|transport|distribution"
    HybridRequireString root, "contractVersion", HYBRID_CONTRACT_VERSION
    HybridRequireString root, "messageType", "engine-capabilities"

    Set engine = HybridRequireObject(root, "engine")
    HybridRequireExactKeys engine, "id|version"
    HybridRequireString engine, "id", HYBRID_ENGINE_ID
    HybridRequireString engine, "version", HYBRID_ENGINE_VERSION

    Set protocolRange = HybridRequireObject(root, "protocolRange")
    HybridRequireExactKeys protocolRange, "minimum|maximum"
    HybridRequireString protocolRange, "minimum", HYBRID_PROTOCOL_VERSION
    HybridRequireString protocolRange, "maximum", HYBRID_PROTOCOL_VERSION

    Set supportedTools = HybridRequireCollection(root, "supportedTools")
    For Each tool In supportedTools
        HybridRequireExactKeys tool, "id|definitionVersions|analysisModes"
        If HybridRequireString(tool, "id") = HYBRID_UNICODE_TOOL_ID Then
            If Not HybridCollectionContainsString(HybridRequireCollection(tool, "definitionVersions"), HYBRID_UNICODE_TOOL_VERSION) Then Err.Raise 5, "HybridValidateCapabilityHandshake", "The engine does not support the Unicode tool definition."
            If Not HybridCollectionContainsString(HybridRequireCollection(tool, "analysisModes"), HYBRID_UNICODE_MODE) Then Err.Raise 5, "HybridValidateCapabilityHandshake", "The engine does not support the Unicode analysis mode."
            foundTool = True
        End If
    Next tool
    If Not foundTool Then Err.Raise 5, "HybridValidateCapabilityHandshake", "The engine does not advertise the Unicode pilot."

    Set supportedOperations = HybridRequireCollection(root, "supportedOperations")
    If Not HybridCollectionContainsString(supportedOperations, "replaceText") Then Err.Raise 5, "HybridValidateCapabilityHandshake", "The engine does not support textual replacement candidates."

    Set security = HybridRequireObject(root, "security")
    HybridRequireExactKeys security, "editsWordDocuments|requiresNetwork|opensListeningEndpoint|runsAsService|requiresElevation|logsDocumentContent"
    If HybridRequireBoolean(security, "editsWordDocuments") Or HybridRequireBoolean(security, "requiresNetwork") Or HybridRequireBoolean(security, "opensListeningEndpoint") Or HybridRequireBoolean(security, "runsAsService") Or HybridRequireBoolean(security, "requiresElevation") Or HybridRequireBoolean(security, "logsDocumentContent") Then Err.Raise 5, "HybridValidateCapabilityHandshake", "The engine advertises an unsafe capability."

    Set transport = HybridRequireObject(root, "transport")
    HybridRequireExactKeys transport, "mode|requestFile|resultFile"
    HybridRequireString transport, "mode", "owner-only-job-directory"
    HybridRequireString transport, "requestFile", "request.json"
    HybridRequireString transport, "resultFile", "result.json"

    Set distribution = HybridRequireObject(root, "distribution")
    HybridRequireExactKeys distribution, "publisher|authenticodeRequiredForOfficialBeta"
    HybridRequireString distribution, "publisher", "MasseysLab"
    If Not HybridRequireBoolean(distribution, "authenticodeRequiredForOfficialBeta") Then Err.Raise 5, "HybridValidateCapabilityHandshake", "The engine does not require signed official distribution."

    HybridValidateCapabilityHandshake = True
    Exit Function

CapabilityFailed:
    failureMessage = "The analysis engine capability handshake failed: " & Err.Description
End Function

Private Function HybridCaptureEngineOutput( _
        ByVal enginePath As String, _
        ByVal commandName As String, _
        ByVal timeoutSeconds As Long) As String
    Dim securityAttributes As HYBRID_SECURITY_ATTRIBUTES
    Dim startup As HYBRID_STARTUPINFO
    Dim processInfo As HYBRID_PROCESS_INFORMATION
    Dim readPipe As LongPtr
    Dim writePipe As LongPtr
    Dim commandLine As String
    Dim waitResult As Long
    Dim exitCode As Long
    Dim startedAt As Single
    Dim elapsed As Double
    Dim available As Long
    Dim bytesRead As Long
    Dim buffer() As Byte
    Dim outputBytes() As Byte
    Dim outputCount As Long

    On Error GoTo CaptureFailed
    ReDim buffer(0 To 4095)
    securityAttributes.nLength = LenB(securityAttributes)
    securityAttributes.bInheritHandle = 1
    If CreatePipe(readPipe, writePipe, securityAttributes, 0) = 0 Then Err.Raise 5, "HybridCaptureEngineOutput", "The capability output pipe could not be created."
    If SetHandleInformation(readPipe, HYBRID_HANDLE_FLAG_INHERIT, 0) = 0 Then Err.Raise 5, "HybridCaptureEngineOutput", "The capability output pipe could not be secured."

    startup.cb = LenB(startup)
    startup.dwFlags = HYBRID_STARTF_USESTDHANDLES
    startup.hStdOutput = writePipe
    startup.hStdError = writePipe
    startup.hStdInput = 0
    commandLine = """" & enginePath & """ " & commandName
    If CreateProcessW(StrPtr(enginePath), StrPtr(commandLine), 0, 0, 1, HYBRID_CREATE_NO_WINDOW, 0, 0, startup, processInfo) = 0 Then Err.Raise 5, "HybridCaptureEngineOutput", "The capability command could not be started."
    CloseHandle writePipe
    writePipe = 0

    startedAt = Timer
    Do
        HybridDrainPipe readPipe, outputBytes, outputCount, buffer, available, bytesRead
        waitResult = WaitForSingleObject(processInfo.hProcess, 25)
        If waitResult = HYBRID_WAIT_OBJECT_0 Then Exit Do
        If waitResult <> HYBRID_WAIT_TIMEOUT Then Err.Raise 5, "HybridCaptureEngineOutput", "The capability command wait failed."
        DoEvents
        elapsed = Timer - startedAt
        If elapsed < 0 Then elapsed = elapsed + 86400#
        If elapsed >= timeoutSeconds Then
            TerminateProcess processInfo.hProcess, 124
            WaitForSingleObject processInfo.hProcess, 5000
            Err.Raise 5, "HybridCaptureEngineOutput", "The capability command timed out."
        End If
    Loop
    HybridDrainPipe readPipe, outputBytes, outputCount, buffer, available, bytesRead
    If GetExitCodeProcess(processInfo.hProcess, exitCode) = 0 Or exitCode <> 0 Then Err.Raise 5, "HybridCaptureEngineOutput", "The capability command failed."
    If outputCount < 1 Or outputCount > 1048576 Then Err.Raise 5, "HybridCaptureEngineOutput", "The capability response size is invalid."
    HybridCaptureEngineOutput = HybridUtf8Decode(outputBytes, outputCount)

CaptureExit:
    On Error Resume Next
    If writePipe <> 0 Then CloseHandle writePipe
    If readPipe <> 0 Then CloseHandle readPipe
    If processInfo.hThread <> 0 Then CloseHandle processInfo.hThread
    If processInfo.hProcess <> 0 Then CloseHandle processInfo.hProcess
    On Error GoTo 0
    Exit Function

CaptureFailed:
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureDescription As String
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    If writePipe <> 0 Then CloseHandle writePipe
    If readPipe <> 0 Then CloseHandle readPipe
    If processInfo.hThread <> 0 Then CloseHandle processInfo.hThread
    If processInfo.hProcess <> 0 Then CloseHandle processInfo.hProcess
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Function

Private Sub HybridDrainPipe( _
        ByVal readPipe As LongPtr, _
        ByRef outputBytes() As Byte, _
        ByRef outputCount As Long, _
        ByRef buffer() As Byte, _
        ByRef available As Long, _
        ByRef bytesRead As Long)
    Dim oldCount As Long
    Dim i As Long
    Do
        available = 0
        If PeekNamedPipe(readPipe, ByVal 0&, 0, bytesRead, available, bytesRead) = 0 Then Exit Do
        If available <= 0 Then Exit Do
        If ReadFile(readPipe, buffer(0), IIf(available > 4096, 4096, available), bytesRead, 0) = 0 Then Err.Raise 5, "HybridDrainPipe", "The capability response could not be read."
        If outputCount + bytesRead > 1048576 Then Err.Raise 5, "HybridDrainPipe", "The capability response exceeded its size limit."
        oldCount = outputCount
        outputCount = outputCount + bytesRead
        If oldCount = 0 Then
            ReDim outputBytes(0 To outputCount - 1)
        Else
            ReDim Preserve outputBytes(0 To outputCount - 1)
        End If
        For i = 0 To bytesRead - 1
            outputBytes(oldCount + i) = buffer(i)
        Next i
    Loop
End Sub

Private Function HybridCollectionContainsString( _
        ByVal values As Collection, _
        ByVal expected As String) As Boolean
    Dim value As Variant
    For Each value In values
        If VarType(value) = vbString Then
            If CStr(value) = expected Then
                HybridCollectionContainsString = True
                Exit Function
            End If
        End If
    Next value
End Function

Private Function HybridCreateEmptyJobDirectory(ByVal jobId As String) As String
    Dim localRoot As String
    Dim publisherRoot As String
    Dim suiteRoot As String
    Dim jobsRoot As String
    Dim jobDirectory As String

    localRoot = Environ$("LOCALAPPDATA")
    If Len(localRoot) = 0 Then Err.Raise 5, "HybridCreateEmptyJobDirectory", "Local application data is unavailable."
    publisherRoot = localRoot & "\MasseysLab"
    suiteRoot = publisherRoot & "\CleanupSuite"
    jobsRoot = suiteRoot & "\Jobs"
    HybridEnsureDirectory publisherRoot
    HybridEnsureDirectory suiteRoot
    HybridEnsureDirectory jobsRoot
    jobDirectory = jobsRoot & "\" & jobId
    If Len(Dir$(jobDirectory, vbDirectory)) > 0 Then Err.Raise 5, "HybridCreateEmptyJobDirectory", "The analysis job already exists."
    MkDir jobDirectory
    HybridCreateEmptyJobDirectory = jobDirectory
End Function

Private Sub HybridEnsureDirectory(ByVal path As String)
    If Len(Dir$(path, vbDirectory)) = 0 Then MkDir path
    If (GetAttr(path) And vbDirectory) = 0 Then Err.Raise 5, "HybridEnsureDirectory", "An analysis workspace path is not a directory."
End Sub

Private Function HybridRunEngine( _
        ByVal enginePath As String, _
        ByVal commandName As String, _
        ByVal jobDirectory As String, _
        ByVal timeoutSeconds As Long) As Long
    Dim startup As HYBRID_STARTUPINFO
    Dim processInfo As HYBRID_PROCESS_INFORMATION
    Dim commandLine As String
    Dim waitResult As Long
    Dim exitCode As Long
    Dim startedAt As Single
    Dim elapsed As Double

    On Error GoTo RunEngineFailed
    commandLine = """" & enginePath & """ " & commandName & " """ & jobDirectory & """"
    startup.cb = LenB(startup)
    If CreateProcessW(StrPtr(enginePath), StrPtr(commandLine), 0, 0, 0, HYBRID_CREATE_NO_WINDOW, 0, 0, startup, processInfo) = 0 Then Err.Raise 5, "HybridRunEngine", "The analysis engine could not be started."
    startedAt = Timer
    Do
        waitResult = WaitForSingleObject(processInfo.hProcess, 50)
        If waitResult = HYBRID_WAIT_OBJECT_0 Then Exit Do
        If waitResult <> HYBRID_WAIT_TIMEOUT Then Err.Raise 5, "HybridRunEngine", "The analysis engine wait failed."
        DoEvents
        elapsed = Timer - startedAt
        If elapsed < 0 Then elapsed = elapsed + 86400#
        If elapsed >= timeoutSeconds Then
            TerminateProcess processInfo.hProcess, 124
            WaitForSingleObject processInfo.hProcess, 5000
            Err.Raise 5, "HybridRunEngine", "The analysis engine timed out."
        End If
    Loop
    If GetExitCodeProcess(processInfo.hProcess, exitCode) = 0 Then Err.Raise 5, "HybridRunEngine", "The analysis engine exit status is unavailable."
    HybridRunEngine = exitCode

RunEngineExit:
    On Error Resume Next
    If processInfo.hThread <> 0 Then CloseHandle processInfo.hThread
    If processInfo.hProcess <> 0 Then CloseHandle processInfo.hProcess
    On Error GoTo 0
    Exit Function

RunEngineFailed:
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureDescription As String
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    If processInfo.hThread <> 0 Then CloseHandle processInfo.hThread
    If processInfo.hProcess <> 0 Then CloseHandle processInfo.hProcess
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Function

Private Function HybridNewUuid() As String
    Dim value As HYBRID_GUID
    Dim buffer As String
    If CoCreateGuid(value) <> 0 Then Err.Raise 5, "HybridNewUuid", "A secure job identifier could not be created."
    buffer = String$(39, vbNullChar)
    If StringFromGUID2(value, StrPtr(buffer), 39) <= 0 Then Err.Raise 5, "HybridNewUuid", "A job identifier could not be formatted."
    HybridNewUuid = LCase$(Mid$(buffer, 2, 36))
End Function

Private Function HybridJsonBoolean(ByVal value As Boolean) As String
    If value Then HybridJsonBoolean = "true" Else HybridJsonBoolean = "false"
End Function

Private Function HybridUtcTimestamp() As String
    Dim value As HYBRID_SYSTEMTIME
    GetSystemTime value
    HybridUtcTimestamp = _
        Format$(value.wYear, "0000") & "-" & _
        Format$(value.wMonth, "00") & "-" & _
        Format$(value.wDay, "00") & "T" & _
        Format$(value.wHour, "00") & ":" & _
        Format$(value.wMinute, "00") & ":" & _
        Format$(value.wSecond, "00") & "." & _
        Format$(value.wMilliseconds, "000") & "Z"
End Function

Private Function HybridUtf8Bytes(ByVal text As String, ByRef byteCount As Long) As Byte()
    Dim bytes() As Byte
    If Len(text) = 0 Then
        ReDim bytes(0 To 0)
        byteCount = 0
        HybridUtf8Bytes = bytes
        Exit Function
    End If
    byteCount = WideCharToMultiByte(HYBRID_CP_UTF8, HYBRID_WC_ERR_INVALID_CHARS, StrPtr(text), Len(text), ByVal 0&, 0, 0, 0)
    If byteCount <= 0 Then Err.Raise 5, "HybridUtf8Bytes", "The Word snapshot contains invalid Unicode."
    ReDim bytes(0 To byteCount - 1)
    If WideCharToMultiByte(HYBRID_CP_UTF8, HYBRID_WC_ERR_INVALID_CHARS, StrPtr(text), Len(text), bytes(0), byteCount, 0, 0) <> byteCount Then Err.Raise 5, "HybridUtf8Bytes", "The Word snapshot could not be encoded exactly."
    HybridUtf8Bytes = bytes
End Function

Private Function HybridUtf8Decode(ByRef bytes() As Byte, ByVal byteCount As Long) As String
    Dim charCount As Long
    Dim result As String
    If byteCount = 0 Then Exit Function
    charCount = MultiByteToWideChar(HYBRID_CP_UTF8, HYBRID_MB_ERR_INVALID_CHARS, bytes(0), byteCount, 0, 0)
    If charCount <= 0 Then Err.Raise 5, "HybridUtf8Decode", "The engine result is not valid UTF-8."
    result = String$(charCount, vbNullChar)
    If MultiByteToWideChar(HYBRID_CP_UTF8, HYBRID_MB_ERR_INVALID_CHARS, bytes(0), byteCount, StrPtr(result), charCount) <> charCount Then Err.Raise 5, "HybridUtf8Decode", "The engine result could not be decoded exactly."
    HybridUtf8Decode = result
End Function

Private Sub HybridWriteTextUtf8Atomic(ByVal path As String, ByVal text As String)
    Dim bytes() As Byte
    Dim byteCount As Long
    bytes = HybridUtf8Bytes(text, byteCount)
    HybridWriteBytesAtomic path, bytes, byteCount
End Sub

Private Sub HybridWriteBytesAtomic(ByVal path As String, ByRef bytes() As Byte, ByVal byteCount As Long)
    Dim temporaryPath As String
    Dim fileNumber As Integer
    On Error GoTo WriteFailed
    temporaryPath = path & ".tmp"
    If Len(Dir$(path, vbNormal)) > 0 Or Len(Dir$(temporaryPath, vbNormal)) > 0 Then Err.Raise 5, "HybridWriteBytesAtomic", "A protocol output already exists."
    fileNumber = FreeFile
    Open temporaryPath For Binary Access Write Lock Read Write As #fileNumber
    If byteCount > 0 Then Put #fileNumber, , bytes
    Close #fileNumber
    fileNumber = 0
    Name temporaryPath As path
    Exit Sub

WriteFailed:
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureDescription As String
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Sub

Private Function HybridReadUtf8File(ByVal path As String, ByVal maximumBytes As Long) As String
    Dim fileNumber As Integer
    Dim byteCount As Long
    Dim bytes() As Byte

    If Len(Dir$(path, vbNormal)) = 0 Then Err.Raise 5, "HybridReadUtf8File", "The engine did not create a result."
    byteCount = FileLen(path)
    If byteCount < 1 Or byteCount > maximumBytes Then Err.Raise 5, "HybridReadUtf8File", "The engine result has an invalid size."
    ReDim bytes(0 To byteCount - 1)
    fileNumber = FreeFile
    Open path For Binary Access Read Lock Write As #fileNumber
    Get #fileNumber, , bytes
    Close #fileNumber
    fileNumber = 0
    If byteCount >= 3 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then Err.Raise 5, "HybridReadUtf8File", "The engine result contains a forbidden UTF-8 byte-order mark."
    End If
    HybridReadUtf8File = HybridUtf8Decode(bytes, byteCount)
End Function

Private Function HybridSha256Utf8(ByVal text As String) As String
    Dim bytes() As Byte
    Dim byteCount As Long
    bytes = HybridUtf8Bytes(text, byteCount)
    HybridSha256Utf8 = HybridSha256Bytes(bytes, byteCount)
End Function

Private Function HybridSha256File(ByVal path As String) As String
    Dim bytes() As Byte
    Dim byteCount As Long
    Dim fileNumber As Integer
    On Error GoTo FileHashFailed
    byteCount = FileLen(path)
    If byteCount < 0 Then Err.Raise 5, "HybridSha256File", "The engine file size is invalid."
    If byteCount = 0 Then
        ReDim bytes(0 To 0)
    Else
        ReDim bytes(0 To byteCount - 1)
        fileNumber = FreeFile
        Open path For Binary Access Read Lock Write As #fileNumber
        Get #fileNumber, , bytes
        Close #fileNumber
        fileNumber = 0
    End If
    HybridSha256File = HybridSha256Bytes(bytes, byteCount)
    Exit Function

FileHashFailed:
    Dim failureNumber As Long
    Dim failureSource As String
    Dim failureDescription As String
    failureNumber = Err.Number
    failureSource = Err.Source
    failureDescription = Err.Description
    On Error Resume Next
    If fileNumber <> 0 Then Close #fileNumber
    On Error GoTo 0
    Err.Raise failureNumber, failureSource, failureDescription
End Function

Private Function HybridSha256Bytes(ByRef bytes() As Byte, ByVal byteCount As Long) As String
    Dim hashHandle As LongPtr
    Dim hashObject() As Byte
    Dim digest() As Byte
    Dim i As Long
    Dim result As String

    HybridEnsureSha256Provider
    ReDim hashObject(0 To mHybridShaObjectLength - 1)
    ReDim digest(0 To mHybridShaDigestLength - 1)
    If BCryptCreateHash(mHybridShaAlgorithm, hashHandle, hashObject(0), mHybridShaObjectLength, 0, 0, 0) <> 0 Then GoTo HashFailed
    If byteCount > 0 Then
        If BCryptHashData(hashHandle, bytes(0), byteCount, 0) <> 0 Then GoTo HashFailed
    End If
    If BCryptFinishHash(hashHandle, digest(0), mHybridShaDigestLength, 0) <> 0 Then GoTo HashFailed
    For i = 0 To mHybridShaDigestLength - 1
        result = result & LCase$(Right$("0" & Hex$(digest(i)), 2))
    Next i
    HybridSha256Bytes = result

HashExit:
    On Error Resume Next
    If hashHandle <> 0 Then BCryptDestroyHash hashHandle
    On Error GoTo 0
    Exit Function

HashFailed:
    On Error Resume Next
    If hashHandle <> 0 Then BCryptDestroyHash hashHandle
    On Error GoTo 0
    Err.Raise 5, "HybridSha256Bytes", "SHA-256 hashing failed."
End Function

Private Sub HybridEnsureSha256Provider()
    Dim returnedLength As Long
    If mHybridShaAlgorithm <> 0 Then Exit Sub
    If BCryptOpenAlgorithmProvider(mHybridShaAlgorithm, StrPtr("SHA256"), 0, 0) <> 0 Then Err.Raise 5, "HybridEnsureSha256Provider", "SHA-256 is unavailable."
    If BCryptGetProperty(mHybridShaAlgorithm, StrPtr("ObjectLength"), mHybridShaObjectLength, 4, returnedLength, 0) <> 0 Then GoTo ProviderFailed
    If BCryptGetProperty(mHybridShaAlgorithm, StrPtr("HashDigestLength"), mHybridShaDigestLength, 4, returnedLength, 0) <> 0 Then GoTo ProviderFailed
    If mHybridShaObjectLength < 1 Or mHybridShaDigestLength <> 32 Then GoTo ProviderFailed
    Exit Sub

ProviderFailed:
    On Error Resume Next
    If mHybridShaAlgorithm <> 0 Then BCryptCloseAlgorithmProvider mHybridShaAlgorithm, 0
    mHybridShaAlgorithm = 0
    mHybridShaObjectLength = 0
    mHybridShaDigestLength = 0
    On Error GoTo 0
    Err.Raise 5, "HybridEnsureSha256Provider", "SHA-256 initialization failed."
End Sub

Private Function HybridSnapshotParagraph( _
        ByVal snapshotText As String, _
        ByVal relativeStart As Long, _
        ByVal relativeEnd As Long, _
        Optional ByRef paragraphKey As String) As String
    Dim paragraphStart As Long
    Dim paragraphMark As Long
    paragraphStart = InStrRev(Left$(snapshotText, relativeStart), vbCr, -1, vbBinaryCompare)
    If paragraphStart = 0 Then paragraphStart = 1 Else paragraphStart = paragraphStart + 1
    paragraphMark = InStr(relativeEnd + 1, snapshotText, vbCr, vbBinaryCompare)
    If paragraphMark = 0 Then
        HybridSnapshotParagraph = Mid$(snapshotText, paragraphStart)
        paragraphKey = CStr(paragraphStart) & "|" & CStr(Len(snapshotText))
    Else
        HybridSnapshotParagraph = Mid$(snapshotText, paragraphStart, paragraphMark - paragraphStart + 1)
        paragraphKey = CStr(paragraphStart) & "|" & CStr(paragraphMark)
    End If
End Function

Private Sub HybridDeleteJobDirectory(ByVal jobDirectory As String)
    Dim jobsRoot As String
    Dim jobName As String
    Dim fileName As Variant
    jobsRoot = Environ$("LOCALAPPDATA") & "\MasseysLab\CleanupSuite\Jobs\"
    If LCase$(Left$(jobDirectory, Len(jobsRoot))) <> LCase$(jobsRoot) Then Exit Sub
    jobName = Mid$(jobDirectory, Len(jobsRoot) + 1)
    If Len(jobName) <> 36 Or InStr(jobName, "\") > 0 Or InStr(jobName, "/") > 0 Then Exit Sub
    For Each fileName In Array("request.json", "document.utf8.txt", "structure.json", "result.json", "cancel.request", "engine.log", "request.json.tmp", "document.utf8.txt.tmp", "result.json.tmp")
        If Len(Dir$(jobDirectory & "\" & CStr(fileName), vbNormal)) > 0 Then Kill jobDirectory & "\" & CStr(fileName)
    Next fileName
    If Len(Dir$(jobDirectory, vbDirectory)) > 0 Then RmDir jobDirectory
End Sub

Private Sub HybridRequireExactKeys(ByVal value As Object, ByVal keyList As String)
    Dim allowed As Object
    Dim key As Variant
    Dim expected As Variant
    Set allowed = CreateObject("Scripting.Dictionary")
    allowed.CompareMode = vbBinaryCompare
    For Each expected In Split(keyList, "|")
        allowed.Add CStr(expected), True
    Next expected
    If value.Count <> allowed.Count Then Err.Raise 5, "HybridRequireExactKeys", "A protocol object contains missing or unknown fields."
    For Each key In value.Keys
        If Not allowed.Exists(CStr(key)) Then Err.Raise 5, "HybridRequireExactKeys", "A protocol object contains an unknown field."
    Next key
End Sub

Private Function HybridRequireObject(ByVal value As Object, ByVal key As String) As Object
    If Not value.Exists(key) Or Not IsObject(value(key)) Or TypeName(value(key)) <> "Dictionary" Then Err.Raise 5, "HybridRequireObject", "A required protocol object is missing."
    Set HybridRequireObject = value(key)
End Function

Private Function HybridRequireCollection(ByVal value As Object, ByVal key As String) As Collection
    If Not value.Exists(key) Or Not IsObject(value(key)) Or TypeName(value(key)) <> "Collection" Then Err.Raise 5, "HybridRequireCollection", "A required protocol array is missing."
    Set HybridRequireCollection = value(key)
End Function

Private Function HybridRequireString(ByVal value As Object, ByVal key As String, Optional ByVal expected As String = vbNullString) As String
    If Not value.Exists(key) Or VarType(value(key)) <> vbString Then Err.Raise 5, "HybridRequireString", "A required protocol string is missing."
    HybridRequireString = CStr(value(key))
    If Len(expected) > 0 And HybridRequireString <> expected Then Err.Raise 5, "HybridRequireString", "A protocol string does not match the expected value."
End Function

Private Function HybridRequireLong(ByVal value As Object, ByVal key As String) As Long
    If Not value.Exists(key) Or (VarType(value(key)) <> vbInteger And VarType(value(key)) <> vbLong And VarType(value(key)) <> vbDouble) Then Err.Raise 5, "HybridRequireLong", "A required protocol integer is missing."
    If CDbl(value(key)) <> Fix(CDbl(value(key))) Then Err.Raise 5, "HybridRequireLong", "A protocol integer contains a fraction."
    HybridRequireLong = CLng(value(key))
End Function

Private Function HybridRequireBoolean(ByVal value As Object, ByVal key As String) As Boolean
    If Not value.Exists(key) Or VarType(value(key)) <> vbBoolean Then Err.Raise 5, "HybridRequireBoolean", "A required protocol Boolean is missing."
    HybridRequireBoolean = CBool(value(key))
End Function

Private Sub HybridRequireSha256(ByVal value As Object, ByVal key As String)
    If Not HybridIsLowercaseSha256(HybridRequireString(value, key)) Then Err.Raise 5, "HybridRequireSha256", "A protocol SHA-256 value is invalid."
End Sub

Private Function HybridIsLowercaseSha256(ByVal value As String) As Boolean
    Dim i As Long
    Dim ch As String
    If Len(value) <> 64 Then Exit Function
    For i = 1 To 64
        ch = Mid$(value, i, 1)
        If InStr(1, "0123456789abcdef", ch, vbBinaryCompare) = 0 Then Exit Function
    Next i
    HybridIsLowercaseSha256 = True
End Function
