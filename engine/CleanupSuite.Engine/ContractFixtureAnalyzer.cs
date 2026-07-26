using System.Diagnostics;
using System.Text.Json;

namespace MasseysLab.CleanupSuite.Engine;

public static class ContractFixtureAnalyzer
{
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

            candidates.Add(TextCandidateFactory.CreateReplacement(
                candidates.Count + 1,
                job,
                found,
                found + source.Length,
                replacement,
                "contract-fixture.literal-match",
                "Contract fixture literal replacement"));
            searchStart = found + source.Length;
        }

        stopwatch.Stop();
        return CompletedResponse(
            job.Request,
            candidates,
            stopwatch.ElapsedMilliseconds);
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
