using System.Diagnostics;
using System.Text.Json;

namespace MasseysLab.CleanupSuite.Engine;

public static class ContractFixtureAnalyzer
{
    private const int ContextLimitUtf16 = 32;

    public static AnalysisResponse Analyze(ValidatedJob job)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        string source = job.Request.Options["from"].GetString()!;
        string replacement = job.Request.Options["to"].GetString()!;
        List<CandidateRecord> candidates = [];
        int searchStart = 0;

        while (searchStart <= job.SnapshotText.Length - source.Length)
        {
            ThrowIfCancelled(job.JobDirectory);
            int found = job.SnapshotText.IndexOf(
                source,
                searchStart,
                StringComparison.Ordinal);
            if (found < 0)
            {
                break;
            }

            if (candidates.Count >= ContractConstants.MaximumCandidates)
            {
                throw new EngineContractException(
                    "analysis-error",
                    "analysis-failed",
                    "The candidate limit was exceeded.",
                    ContractConstants.ExitAnalysisError);
            }

            candidates.Add(CreateCandidate(
                candidates.Count + 1,
                job,
                found,
                found + source.Length,
                replacement));
            searchStart = found + source.Length;
        }

        stopwatch.Stop();
        return CompletedResponse(
            job.Request,
            candidates,
            stopwatch.ElapsedMilliseconds);
    }

    private static CandidateRecord CreateCandidate(
        int ordinal,
        ValidatedJob job,
        int start,
        int end,
        string replacement)
    {
        string text = job.SnapshotText;
        int prefixStart = AlignedPrefixStart(text, start);
        int suffixEnd = AlignedSuffixEnd(text, end);
        string exact = text[start..end];
        string prefix = text[prefixStart..start];
        string suffix = text[end..suffixEnd];
        string paragraph = ParagraphContaining(text, start, end);

        return new CandidateRecord
        {
            CandidateId = $"c{ordinal:000000}",
            State = "applicable",
            ReasonCode = "contract-fixture.literal-match",
            Location = new LocationRecord
            {
                StoryType = job.Request.Scope.StoryType,
                StartUtf16 = start,
                EndUtf16 = end,
                OffsetBasis = "snapshot-relative-utf16-code-units"
            },
            Fingerprint = new FingerprintRecord
            {
                Algorithm = "sha256-utf8-exact",
                SnapshotId = job.Request.Snapshot.SnapshotId,
                ExactTextSha256 = Hashing.Sha256Utf8(exact),
                PrefixLengthUtf16 = prefix.Length,
                PrefixSha256 = Hashing.Sha256Utf8(prefix),
                SuffixLengthUtf16 = suffix.Length,
                SuffixSha256 = Hashing.Sha256Utf8(suffix),
                ParagraphSha256 = Hashing.Sha256Utf8(paragraph)
            },
            Operation = new OperationRecord
            {
                Type = "replaceText",
                SafetyClass = "textual",
                Parameters = new Dictionary<string, object?>
                {
                    ["replacementText"] = replacement
                }
            },
            Revalidation = new RevalidationRecord
            {
                RequireWholeScopeSnapshot = true,
                RequireExactRange = true,
                RequireContext = true,
                RequireStructure = false,
                OnMismatch = "abort-apply",
                AllowRelocation = false
            },
            Confidence = new ConfidenceRecord
            {
                Tier = "deterministic",
                Score = 1.0m
            },
            Display = new DisplayRecord
            {
                Label = "Contract fixture literal replacement",
                CanHighlight = true
            }
        };
    }

    private static int AlignedPrefixStart(string text, int candidateStart)
    {
        int start = Math.Max(0, candidateStart - ContextLimitUtf16);
        if (start > 0
            && start < candidateStart
            && char.IsLowSurrogate(text[start])
            && char.IsHighSurrogate(text[start - 1]))
        {
            start++;
        }

        return start;
    }

    private static int AlignedSuffixEnd(string text, int candidateEnd)
    {
        int end = Math.Min(
            text.Length,
            candidateEnd + ContextLimitUtf16);
        if (end > candidateEnd
            && end < text.Length
            && char.IsLowSurrogate(text[end])
            && char.IsHighSurrogate(text[end - 1]))
        {
            end--;
        }

        return end;
    }

    private static string ParagraphContaining(
        string text,
        int candidateStart,
        int candidateEnd)
    {
        int paragraphStart = candidateStart == 0
            ? 0
            : text.LastIndexOf('\r', candidateStart - 1) + 1;
        int paragraphMark = text.IndexOf('\r', candidateEnd);
        int paragraphEnd = paragraphMark < 0
            ? text.Length
            : paragraphMark + 1;
        return text[paragraphStart..paragraphEnd];
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
