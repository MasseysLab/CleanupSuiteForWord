namespace MasseysLab.CleanupSuite.Engine;

public static class SanitizedLog
{
    public static void TryAppend(
        string jobDirectory,
        string eventCode,
        int candidateCount = 0)
    {
        try
        {
            string logPath = JobPathPolicy.FixedFile(
                jobDirectory,
                ContractConstants.LogFileName);
            string line = string.Create(
                System.Globalization.CultureInfo.InvariantCulture,
                $"{DateTimeOffset.UtcNow:O}|{eventCode}|candidates={candidateCount}\n");
            File.AppendAllText(logPath, line, JsonSupport.StrictUtf8);
        }
        catch
        {
            // Diagnostics must never make analysis less safe or expose content.
        }
    }
}
