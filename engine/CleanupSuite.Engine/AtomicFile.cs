namespace MasseysLab.CleanupSuite.Engine;

public static class AtomicFile
{
    public static void WriteNew(string destinationPath, ReadOnlySpan<byte> content)
    {
        string temporaryPath = destinationPath + ContractConstants.AtomicSuffix;
        EnsureOrdinaryOrMissing(temporaryPath);
        EnsureOrdinaryOrMissing(destinationPath);

        using (FileStream stream = new(
            temporaryPath,
            FileMode.CreateNew,
            FileAccess.Write,
            FileShare.None,
            bufferSize: 16_384,
            options: FileOptions.WriteThrough))
        {
            stream.Write(content);
            stream.Flush(flushToDisk: true);
        }

        File.Move(temporaryPath, destinationPath);
    }

    private static void EnsureOrdinaryOrMissing(string path)
    {
        if (!File.Exists(path))
        {
            return;
        }

        FileAttributes attributes = File.GetAttributes(path);
        if ((attributes & FileAttributes.ReparsePoint) != 0)
        {
            throw new EngineContractException(
                "security-error",
                "security-policy-violation",
                "A protocol output path is a reparse point.",
                ContractConstants.ExitSecurityError);
        }

        throw new EngineContractException(
            "security-error",
            "security-policy-violation",
            "A protocol output file already exists.",
            ContractConstants.ExitSecurityError);
    }
}
