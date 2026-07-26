namespace MasseysLab.CleanupSuite.Engine;

public static class Capabilities
{
    public static object Create() =>
        new
        {
            contractVersion = ContractConstants.ContractVersion,
            messageType = "engine-capabilities",
            engine = new
            {
                id = ContractConstants.EngineId,
                version = ContractConstants.EngineVersion
            },
            protocolRange = new
            {
                minimum = ContractConstants.ProtocolVersion,
                maximum = ContractConstants.ProtocolVersion
            },
            supportedTools = new[]
            {
                new
                {
                    id = ContractConstants.FixtureToolId,
                    definitionVersions = new[]
                    {
                        ContractConstants.FixtureToolDefinitionVersion
                    },
                    analysisModes = new[]
                    {
                        ContractConstants.FixtureAnalysisMode
                    }
                },
                new
                {
                    id = ContractConstants.UnicodeToolId,
                    definitionVersions = new[]
                    {
                        ContractConstants.UnicodeToolDefinitionVersion
                    },
                    analysisModes = new[]
                    {
                        ContractConstants.UnicodeAnalysisMode
                    }
                }
            },
            supportedOperations = new[]
            {
                "reportOnly",
                "replaceText"
            },
            security = new
            {
                editsWordDocuments = false,
                requiresNetwork = false,
                opensListeningEndpoint = false,
                runsAsService = false,
                requiresElevation = false,
                logsDocumentContent = false
            },
            transport = new
            {
                mode = "owner-only-job-directory",
                requestFile = ContractConstants.RequestFileName,
                resultFile = ContractConstants.ResultFileName
            },
            distribution = new
            {
                publisher = "MasseysLab",
                authenticodeRequiredForOfficialBeta = true
            }
        };
}
