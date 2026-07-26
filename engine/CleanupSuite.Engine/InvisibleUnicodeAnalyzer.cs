using System.Diagnostics;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace MasseysLab.CleanupSuite.Engine;

[JsonUnmappedMemberHandling(JsonUnmappedMemberHandling.Disallow)]
public sealed class InvisibleUnicodeOptions
{
    public required bool NonBreakingSpace { get; init; }
    public required bool ZeroWidthSpace { get; init; }
    public required bool ZeroWidthNonJoiner { get; init; }
    public required bool ZeroWidthJoiner { get; init; }
    public required bool ByteOrderMark { get; init; }
    public required bool SoftHyphen { get; init; }
    public required bool NonBreakingHyphen { get; init; }

    public bool AnySelected =>
        NonBreakingSpace
        || ZeroWidthSpace
        || ZeroWidthNonJoiner
        || ZeroWidthJoiner
        || ByteOrderMark
        || SoftHyphen
        || NonBreakingHyphen;

    public static InvisibleUnicodeOptions From(
        Dictionary<string, JsonElement> options)
    {
        try
        {
            byte[] json = JsonSerializer.SerializeToUtf8Bytes(
                options,
                JsonSupport.WriteOptions);
            InvisibleUnicodeOptions parsed =
                JsonSupport.DeserializeStrict<InvisibleUnicodeOptions>(json);
            if (!parsed.AnySelected)
            {
                throw InvalidOptions(
                    "At least one Unicode character type must be selected.");
            }

            return parsed;
        }
        catch (EngineContractException)
        {
            throw;
        }
        catch (Exception exception)
            when (exception is JsonException
                or InvalidDataException)
        {
            throw InvalidOptions(
                "The Unicode cleaner options are invalid.");
        }
    }

    private static EngineContractException InvalidOptions(string message) =>
        new(
            "invalid-request",
            "invalid-request",
            message,
            ContractConstants.ExitInvalidRequest);
}

public static class InvisibleUnicodeAnalyzer
{
    private sealed record Mapping(
        char Character,
        string Replacement,
        string ReasonCode,
        string Label);

    public static AnalysisResponse Analyze(ValidatedJob job)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        InvisibleUnicodeOptions options =
            InvisibleUnicodeOptions.From(job.Request.Options);
        Dictionary<char, Mapping> mappings = SelectedMappings(options);
        List<CandidateRecord> candidates = [];

        for (int index = 0; index < job.SnapshotText.Length; index++)
        {
            ThrowIfCancelled(job.JobDirectory);
            if (!mappings.TryGetValue(
                    job.SnapshotText[index],
                    out Mapping? mapping))
            {
                continue;
            }

            if (candidates.Count >= ContractConstants.MaximumCandidates)
            {
                throw new EngineContractException(
                    "analysis-error",
                    "analysis-failed",
                    "The candidate limit was exceeded.",
                    ContractConstants.ExitAnalysisError);
            }

            candidates.Add(TextCandidateFactory.CreateReplacement(
                candidates.Count + 1,
                job,
                index,
                index + 1,
                mapping.Replacement,
                mapping.ReasonCode,
                mapping.Label));
        }

        stopwatch.Stop();
        return CompletedResponse(
            job.Request,
            candidates,
            stopwatch.ElapsedMilliseconds);
    }

    private static Dictionary<char, Mapping> SelectedMappings(
        InvisibleUnicodeOptions options)
    {
        Dictionary<char, Mapping> mappings = [];
        AddIf(
            options.NonBreakingSpace,
            '\u00A0',
            " ",
            "unicode.non-breaking-space",
            "Non-breaking space (U+00A0)",
            mappings);
        AddIf(
            options.ZeroWidthSpace,
            '\u200B',
            "",
            "unicode.zero-width-space",
            "Zero-width space (U+200B)",
            mappings);
        AddIf(
            options.ZeroWidthNonJoiner,
            '\u200C',
            "",
            "unicode.zero-width-nonjoiner",
            "Zero-width non-joiner (U+200C)",
            mappings);
        AddIf(
            options.ZeroWidthJoiner,
            '\u200D',
            "",
            "unicode.zero-width-joiner",
            "Zero-width joiner (U+200D)",
            mappings);
        AddIf(
            options.ByteOrderMark,
            '\uFEFF',
            "",
            "unicode.byte-order-mark",
            "Byte order mark (U+FEFF)",
            mappings);
        AddIf(
            options.SoftHyphen,
            '\u00AD',
            "-",
            "unicode.soft-hyphen",
            "Soft hyphen (U+00AD)",
            mappings);
        AddIf(
            options.NonBreakingHyphen,
            '\u2011',
            "-",
            "unicode.non-breaking-hyphen",
            "Non-breaking hyphen (U+2011)",
            mappings);
        return mappings;
    }

    private static void AddIf(
        bool selected,
        char character,
        string replacement,
        string reasonCode,
        string label,
        Dictionary<char, Mapping> mappings)
    {
        if (selected)
        {
            mappings.Add(
                character,
                new Mapping(
                    character,
                    replacement,
                    reasonCode,
                    label));
        }
    }

    private static void ThrowIfCancelled(string jobDirectory)
    {
        string cancelPath = JobPathPolicy.FixedFile(
            jobDirectory,
            ContractConstants.CancelFileName);
        if (File.Exists(cancelPath))
        {
            throw new EngineContractException(
                "cancelled",
                "cancelled",
                "The analysis was cancelled.",
                ContractConstants.ExitCancelled);
        }
    }

    private static AnalysisResponse CompletedResponse(
        AnalysisRequest request,
        List<CandidateRecord> candidates,
        long durationMs) =>
        new()
        {
            ContractVersion = ContractConstants.ContractVersion,
            MessageType = "analysis-response",
            JobId = request.JobId,
            CompletedUtc = DateTimeOffset.UtcNow,
            Engine = ResponseFactory.EngineIdentity(),
            Echo = ResponseFactory.Echo(request),
            Status = "completed",
            Candidates = candidates,
            Summary = new SummaryRecord
            {
                Total = candidates.Count,
                Applicable = candidates.Count,
                ReviewOnly = 0,
                Protected = 0,
                Skipped = 0
            },
            Diagnostics = new DiagnosticsRecord
            {
                DurationMs = durationMs,
                WarningCodes = [],
                LogContainsDocumentContent = false
            }
        };
}
