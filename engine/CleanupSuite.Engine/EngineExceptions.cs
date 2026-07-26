namespace MasseysLab.CleanupSuite.Engine;

public sealed class EngineContractException : Exception
{
    public EngineContractException(
        string status,
        string errorCode,
        string safeMessage,
        int exitCode)
        : base(safeMessage)
    {
        Status = status;
        ErrorCode = errorCode;
        SafeMessage = safeMessage;
        ExitCode = exitCode;
    }

    public string Status { get; }
    public string ErrorCode { get; }
    public string SafeMessage { get; }
    public int ExitCode { get; }
}
