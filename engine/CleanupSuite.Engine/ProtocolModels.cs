using System.Text.Json;
using System.Text.Json.Serialization;

namespace MasseysLab.CleanupSuite.Engine;

[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class AnalysisRequest
{
    public required string ContractVersion { get; init; }
    public required string MessageType { get; init; }
    public required string JobId { get; init; }
    public required DateTimeOffset CreatedUtc { get; init; }
    public required ClientRecord Client { get; init; }
    public required ToolRecord Tool { get; init; }
    public required ScopeRecord Scope { get; init; }
    public required SnapshotRecord Snapshot { get; init; }
    public required Dictionary<string, JsonElement> Options { get; init; }
    public required List<string> RequestedCapabilities { get; init; }
    public required PrivacyRecord Privacy { get; init; }
}

[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class ClientRecord
{
    public required string SuiteVersion { get; init; }
    public required string ProtocolVersion { get; init; }
    public required int ProcessId { get; init; }
    public required string SessionId { get; init; }
}

[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class ToolRecord
{
    public required string Id { get; init; }
    public required string DefinitionVersion { get; init; }
    public required string AnalysisMode { get; init; }
}

[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class ScopeRecord
{
    public required string DocumentSessionId { get; init; }
    public required string StoryType { get; init; }
    public required int StartUtf16 { get; init; }
    public required int EndUtf16 { get; init; }
    public required bool SelectionOnly { get; init; }
}

[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class SnapshotRecord
{
    public required string SnapshotId { get; init; }
    public required string ContentFile { get; init; }
    public required string Encoding { get; init; }
    public required long ByteLength { get; init; }
    public required int Utf16Length { get; init; }
    public required string Sha256 { get; init; }
    public required string LineEndingPolicy { get; init; }
    public string? StructureFile { get; init; }
    public string? StructureSha256 { get; init; }
}

[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class PrivacyRecord
{
    public required bool DocumentPathIncluded { get; init; }
    public required bool DocumentNameIncluded { get; init; }
}

public sealed class AnalysisResponse
{
    public required string ContractVersion { get; init; }
    public required string MessageType { get; init; }
    public required string JobId { get; init; }
    public required DateTimeOffset CompletedUtc { get; init; }
    public required EngineRecord Engine { get; init; }
    public required EchoRecord Echo { get; init; }
    public required string Status { get; init; }
    public required List<CandidateRecord> Candidates { get; init; }
    public required SummaryRecord Summary { get; init; }
    public required DiagnosticsRecord Diagnostics { get; init; }
    public ErrorRecord? Error { get; init; }
}

public sealed class EngineRecord
{
    public required string Id { get; init; }
    public required string Version { get; init; }
    public required string ProtocolVersion { get; init; }
}

public sealed class EchoRecord
{
    public required string SnapshotId { get; init; }
    public required string ToolId { get; init; }
    public required string ToolDefinitionVersion { get; init; }
}

public sealed class CandidateRecord
{
    public required string CandidateId { get; init; }
    public required string State { get; init; }
    public required string ReasonCode { get; init; }
    public required LocationRecord Location { get; init; }
    public required FingerprintRecord Fingerprint { get; init; }
    public required OperationRecord Operation { get; init; }
    public required RevalidationRecord Revalidation { get; init; }
    public required ConfidenceRecord Confidence { get; init; }
    public required DisplayRecord Display { get; init; }
}

public sealed class LocationRecord
{
    public required string StoryType { get; init; }
    public required int StartUtf16 { get; init; }
    public required int EndUtf16 { get; init; }
    public required string OffsetBasis { get; init; }
}

public sealed class FingerprintRecord
{
    public required string Algorithm { get; init; }
    public required string SnapshotId { get; init; }
    public required string ExactTextSha256 { get; init; }
    public required int PrefixLengthUtf16 { get; init; }
    public required string PrefixSha256 { get; init; }
    public required int SuffixLengthUtf16 { get; init; }
    public required string SuffixSha256 { get; init; }
    public string? ParagraphSha256 { get; init; }
    public string? StructureSha256 { get; init; }
}

public sealed class OperationRecord
{
    public required string Type { get; init; }
    public required string SafetyClass { get; init; }
    public required Dictionary<string, object?> Parameters { get; init; }
}

public sealed class RevalidationRecord
{
    public bool RequireWholeScopeSnapshot { get; init; } = true;
    public bool RequireExactRange { get; init; } = true;
    public bool RequireContext { get; init; } = true;
    public bool RequireStructure { get; init; }
    public string OnMismatch { get; init; } = "abort-apply";
    public bool AllowRelocation { get; init; }
}

public sealed class ConfidenceRecord
{
    public required string Tier { get; init; }
    public required decimal Score { get; init; }
}

public sealed class DisplayRecord
{
    public required string Label { get; init; }
    public required bool CanHighlight { get; init; }
}

public sealed class SummaryRecord
{
    public required int Total { get; init; }
    public required int Applicable { get; init; }
    public required int ReviewOnly { get; init; }
    public required int Protected { get; init; }
    public required int Skipped { get; init; }
}

public sealed class DiagnosticsRecord
{
    public required long DurationMs { get; init; }
    public required List<string> WarningCodes { get; init; }
    public required bool LogContainsDocumentContent { get; init; }
}

public sealed class ErrorRecord
{
    public required string Code { get; init; }
    public required string Message { get; init; }
}
