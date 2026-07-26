namespace MasseysLab.CleanupSuite.Engine;

public static class ResponseFactory
{
    private const string UnknownSnapshot =
        "0000000000000000000000000000000000000000000000000000000000000000";

    public static EngineRecord EngineIdentity() =>
        new()
        {
            Id = ContractConstants.EngineId,
            Version = ContractConstants.EngineVersion,
            ProtocolVersion = ContractConstants.ProtocolVersion
        };

    public static EchoRecord Echo(AnalysisRequest request) =>
        new()
        {
            SnapshotId = request.Snapshot.SnapshotId,
            ToolId = request.Tool.Id,
            ToolDefinitionVersion = request.Tool.DefinitionVersion
        };

    public static AnalysisResponse Error(
        string jobId,
        AnalysisRequest? request,
        EngineContractException exception,
        long durationMs) =>
        new()
        {
            ContractVersion = ContractConstants.ContractVersion,
            MessageType = "analysis-response",
            JobId = jobId,
            CompletedUtc = DateTimeOffset.UtcNow,
            Engine = EngineIdentity(),
            Echo = request is null
                ? new EchoRecord
                {
                    SnapshotId = UnknownSnapshot,
                    ToolId = "unknown-tool",
                    ToolDefinitionVersion = "0.0.0"
                }
                : Echo(request),
            Status = exception.Status,
            Candidates = [],
            Summary = new SummaryRecord
            {
                Total = 0,
                Applicable = 0,
                ReviewOnly = 0,
                Protected = 0,
                Skipped = 0
            },
            Diagnostics = new DiagnosticsRecord
            {
                DurationMs = durationMs,
                WarningCodes = [],
                LogContainsDocumentContent = false
            },
            Error = new ErrorRecord
            {
                Code = exception.ErrorCode,
                Message = exception.SafeMessage
            }
        };
}
