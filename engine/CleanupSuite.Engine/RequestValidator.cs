using System.Text.Json;
using System.Text;

namespace MasseysLab.CleanupSuite.Engine;

public sealed record ValidatedJob(
    string JobDirectory,
    AnalysisRequest Request,
    byte[] SnapshotBytes,
    string SnapshotText);

public static class RequestValidator
{
    private static readonly HashSet<string> SupportedCapabilities =
    [
        "analysis.contract-fixture",
        "analysis.invisible-unicode",
        "fingerprint.sha256-utf8-exact"
    ];

    public static ValidatedJob Load(string suppliedJobDirectory)
    {
        string jobDirectory = JobPathPolicy.ValidateJobDirectory(
            suppliedJobDirectory);
        string requestPath = JobPathPolicy.FixedFile(
            jobDirectory,
            ContractConstants.RequestFileName);
        string contentPath = JobPathPolicy.FixedFile(
            jobDirectory,
            ContractConstants.ContentFileName);

        byte[] requestBytes = ReadBoundedFile(
            requestPath,
            ContractConstants.MaximumRequestBytes,
            "The request is missing or exceeds its size limit.");
        AnalysisRequest request;
        try
        {
            request = JsonSupport.DeserializeStrict<AnalysisRequest>(
                requestBytes);
        }
        catch (Exception exception)
            when (exception is JsonException
                or DecoderFallbackException
                or InvalidDataException)
        {
            throw InvalidRequest("The request JSON is invalid.");
        }

        ValidateEnvelope(request, jobDirectory);

        byte[] snapshotBytes = ReadBoundedFile(
            contentPath,
            ContractConstants.MaximumSnapshotBytes,
            "The snapshot is missing or exceeds its size limit.");
        JsonSupport.RejectUtf8Bom(snapshotBytes);

        string snapshotText;
        try
        {
            snapshotText = JsonSupport.StrictUtf8.GetString(snapshotBytes);
            // Re-encoding catches malformed surrogate data introduced through
            // escaped JSON options or future in-memory snapshot sources.
            _ = JsonSupport.StrictUtf8.GetBytes(snapshotText);
        }
        catch (DecoderFallbackException)
        {
            throw new EngineContractException(
                "invalid-request",
                "invalid-snapshot-text",
                "The snapshot is not strict UTF-8 text.",
                ContractConstants.ExitInvalidRequest);
        }

        ValidateSnapshot(request, snapshotBytes, snapshotText);
        ValidateToolOptions(request);
        ValidateOptionalStructure(jobDirectory, request.Snapshot);

        return new ValidatedJob(
            jobDirectory,
            request,
            snapshotBytes,
            snapshotText);
    }

    private static void ValidateEnvelope(
        AnalysisRequest request,
        string jobDirectory)
    {
        Require(
            request.ContractVersion == ContractConstants.ContractVersion,
            "incompatible",
            "unsupported-protocol",
            "The contract version is not supported.",
            ContractConstants.ExitIncompatible);
        Require(
            request.MessageType == "analysis-request",
            "invalid-request",
            "invalid-request",
            "The message type is invalid.",
            ContractConstants.ExitInvalidRequest);
        RequireLowercaseUuid(request.JobId, "jobId");
        RequireLowercaseUuid(request.Client.SessionId, "client.sessionId");
        RequireLowercaseUuid(
            request.Scope.DocumentSessionId,
            "scope.documentSessionId");
        Require(
            request.JobId == Path.GetFileName(jobDirectory),
            "security-error",
            "security-policy-violation",
            "The request job ID does not match its directory.",
            ContractConstants.ExitSecurityError);
        Require(
            request.Client.ProtocolVersion == ContractConstants.ProtocolVersion,
            "incompatible",
            "unsupported-protocol",
            "The protocol version is not supported.",
            ContractConstants.ExitIncompatible);
        Require(
            request.Client.ProcessId > 0,
            "invalid-request",
            "invalid-request",
            "The client process ID is invalid.",
            ContractConstants.ExitInvalidRequest);
        ValidateToolEnvelope(request);
        Require(
            !request.Privacy.DocumentPathIncluded
                && !request.Privacy.DocumentNameIncluded,
            "security-error",
            "security-policy-violation",
            "Document identity must not be included in an analysis request.",
            ContractConstants.ExitSecurityError);
        Require(
            request.RequestedCapabilities.All(
                capability => SupportedCapabilities.Contains(capability)),
            "incompatible",
            "unsupported-tool",
            "A requested engine capability is not supported.",
            ContractConstants.ExitIncompatible);
    }

    private static void ValidateToolEnvelope(AnalysisRequest request)
    {
        string requiredCapability;
        if (request.Tool.Id == ContractConstants.FixtureToolId)
        {
            Require(
                request.Tool.DefinitionVersion
                    == ContractConstants.FixtureToolDefinitionVersion
                    && request.Tool.AnalysisMode
                    == ContractConstants.FixtureAnalysisMode,
                "incompatible",
                "unsupported-tool-version",
                "The requested contract fixture version or mode is not supported.",
                ContractConstants.ExitIncompatible);
            requiredCapability = "analysis.contract-fixture";
        }
        else if (request.Tool.Id == ContractConstants.UnicodeToolId)
        {
            Require(
                request.Tool.DefinitionVersion
                    == ContractConstants.UnicodeToolDefinitionVersion
                    && request.Tool.AnalysisMode
                    == ContractConstants.UnicodeAnalysisMode,
                "incompatible",
                "unsupported-tool-version",
                "The requested Unicode tool version or mode is not supported.",
                ContractConstants.ExitIncompatible);
            requiredCapability = "analysis.invisible-unicode";
        }
        else
        {
            throw new EngineContractException(
                "incompatible",
                "unsupported-tool",
                "The requested tool is not supported.",
                ContractConstants.ExitIncompatible);
        }

        Require(
            request.RequestedCapabilities.Contains(
                requiredCapability,
                StringComparer.Ordinal)
                && request.RequestedCapabilities.Contains(
                    "fingerprint.sha256-utf8-exact",
                    StringComparer.Ordinal),
            "incompatible",
            "unsupported-tool",
            "Required engine capabilities were not requested.",
            ContractConstants.ExitIncompatible);
    }

    private static void ValidateToolOptions(AnalysisRequest request)
    {
        if (request.Tool.Id == ContractConstants.FixtureToolId)
        {
            ValidateFixtureOptions(request.Options);
            return;
        }

        if (request.Tool.Id == ContractConstants.UnicodeToolId)
        {
            _ = InvisibleUnicodeOptions.From(request.Options);
            return;
        }

        throw new EngineContractException(
            "incompatible",
            "unsupported-tool",
            "The requested tool is not supported.",
            ContractConstants.ExitIncompatible);
    }

    private static void ValidateSnapshot(
        AnalysisRequest request,
        byte[] snapshotBytes,
        string snapshotText)
    {
        SnapshotRecord snapshot = request.Snapshot;
        Require(
            snapshot.ContentFile == ContractConstants.ContentFileName,
            "security-error",
            "security-policy-violation",
            "The snapshot filename is not allowed.",
            ContractConstants.ExitSecurityError);
        Require(
            snapshot.Encoding == "utf-8"
                && snapshot.LineEndingPolicy
                    == "preserve-word-story-text",
            "invalid-request",
            "invalid-request",
            "The snapshot encoding policy is invalid.",
            ContractConstants.ExitInvalidRequest);
        Require(
            snapshot.ByteLength == snapshotBytes.LongLength,
            "invalid-request",
            "invalid-request",
            "The snapshot byte length does not match.",
            ContractConstants.ExitInvalidRequest);
        Require(
            snapshot.Utf16Length == snapshotText.Length,
            "invalid-request",
            "invalid-request",
            "The snapshot UTF-16 length does not match.",
            ContractConstants.ExitInvalidRequest);
        Require(
            request.Scope.StartUtf16 >= 0
                && request.Scope.EndUtf16 >= request.Scope.StartUtf16
                && request.Scope.EndUtf16 - request.Scope.StartUtf16
                    == snapshot.Utf16Length,
            "invalid-request",
            "invalid-request",
            "The scope does not match the snapshot length.",
            ContractConstants.ExitInvalidRequest);
        string actualHash = Hashing.Sha256Hex(snapshotBytes);
        Require(
            Hashing.IsLowercaseSha256(snapshot.Sha256)
                && snapshot.Sha256 == actualHash
                && snapshot.SnapshotId == actualHash,
            "invalid-request",
            "invalid-request",
            "The snapshot fingerprint does not match its content.",
            ContractConstants.ExitInvalidRequest);
    }

    private static void ValidateFixtureOptions(
        Dictionary<string, JsonElement> options)
    {
        if (options.Count != 2
            || !options.TryGetValue("from", out JsonElement from)
            || !options.TryGetValue("to", out JsonElement to)
            || from.ValueKind != JsonValueKind.String
            || to.ValueKind != JsonValueKind.String)
        {
            throw InvalidRequest(
                "The contract fixture requires only string from/to options.");
        }

        string source = from.GetString() ?? string.Empty;
        string replacement = to.GetString() ?? string.Empty;
        Require(
            source.Length > 0,
            "invalid-request",
            "invalid-request",
            "The contract fixture source text must not be empty.",
            ContractConstants.ExitInvalidRequest);
        try
        {
            _ = JsonSupport.StrictUtf8.GetBytes(source);
            _ = JsonSupport.StrictUtf8.GetBytes(replacement);
        }
        catch (EncoderFallbackException)
        {
            throw InvalidRequest(
                "The contract fixture options contain malformed Unicode.");
        }
    }

    private static void ValidateOptionalStructure(
        string jobDirectory,
        SnapshotRecord snapshot)
    {
        bool hasFile = snapshot.StructureFile is not null;
        bool hasHash = snapshot.StructureSha256 is not null;
        Require(
            hasFile == hasHash,
            "invalid-request",
            "invalid-request",
            "The structure filename and hash must be supplied together.",
            ContractConstants.ExitInvalidRequest);
        if (!hasFile)
        {
            return;
        }

        Require(
            snapshot.StructureFile == ContractConstants.StructureFileName,
            "security-error",
            "security-policy-violation",
            "The structure filename is not allowed.",
            ContractConstants.ExitSecurityError);
        string structurePath = JobPathPolicy.FixedFile(
            jobDirectory,
            ContractConstants.StructureFileName);
        byte[] structureBytes = ReadBoundedFile(
            structurePath,
            67_108_864,
            "The structure snapshot is missing or exceeds its size limit.");
        Require(
            Hashing.IsLowercaseSha256(snapshot.StructureSha256!)
                && Hashing.Sha256Hex(structureBytes)
                    == snapshot.StructureSha256,
            "invalid-request",
            "invalid-request",
            "The structure snapshot fingerprint does not match.",
            ContractConstants.ExitInvalidRequest);
    }

    private static byte[] ReadBoundedFile(
        string path,
        long maximumBytes,
        string safeMessage)
    {
        if (!File.Exists(path))
        {
            throw InvalidRequest(safeMessage);
        }

        FileInfo info = new(path);
        if (info.Length > maximumBytes)
        {
            throw InvalidRequest(safeMessage);
        }

        return File.ReadAllBytes(path);
    }

    private static void RequireLowercaseUuid(string value, string fieldName)
    {
        Require(
            Guid.TryParseExact(value, "D", out Guid parsed)
                && parsed.ToString("D") == value,
            "invalid-request",
            "invalid-request",
            $"The {fieldName} value is not a lowercase UUID.",
            ContractConstants.ExitInvalidRequest);
    }

    private static void Require(
        bool condition,
        string status,
        string errorCode,
        string safeMessage,
        int exitCode)
    {
        if (!condition)
        {
            throw new EngineContractException(
                status,
                errorCode,
                safeMessage,
                exitCode);
        }
    }

    private static EngineContractException InvalidRequest(string message) =>
        new(
            "invalid-request",
            "invalid-request",
            message,
            ContractConstants.ExitInvalidRequest);
}
