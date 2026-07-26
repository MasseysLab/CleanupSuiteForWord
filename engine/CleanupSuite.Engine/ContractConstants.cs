namespace MasseysLab.CleanupSuite.Engine;

public static class ContractConstants
{
    public const string ContractVersion = "1.0";
    public const string EngineId = "com.masseyslab.cleanupsuite.engine";
    public const string EngineVersion = "0.2.0";
    public const string ProtocolVersion = "1.0";
    public const string FixtureToolId = "contract-fixture";
    public const string FixtureToolDefinitionVersion = "1.0.0";
    public const string FixtureAnalysisMode = "replace-literal";
    public const string UnicodeToolId = "invisible-unicode-cleaner";
    public const string UnicodeToolDefinitionVersion = "1.0.0";
    public const string UnicodeAnalysisMode = "selected-characters";

    public const string RequestFileName = "request.json";
    public const string ContentFileName = "document.utf8.txt";
    public const string StructureFileName = "structure.json";
    public const string ResultFileName = "result.json";
    public const string CancelFileName = "cancel.request";
    public const string LogFileName = "engine.log";
    public const string AtomicSuffix = ".tmp";

    public const long MaximumRequestBytes = 1_048_576;
    public const long MaximumSnapshotBytes = 268_435_456;
    public const long MaximumResultBytes = 134_217_728;
    public const int MaximumCandidates = 1_000_000;

    public const int ExitCompleted = 0;
    public const int ExitInvalidRequest = 10;
    public const int ExitIncompatible = 11;
    public const int ExitCancelled = 12;
    public const int ExitAnalysisError = 13;
    public const int ExitSecurityError = 14;
    public const int ExitInternalError = 15;
}
