namespace MasseysLab.CleanupSuite.Engine;

public static class JobPathPolicy
{
    public static string ExpectedJobsRoot
    {
        get
        {
            string? testRoot = Environment.GetEnvironmentVariable(
                "CLEANUPSUITE_HYBRID_TEST_JOBS_ROOT");
            if (!string.IsNullOrWhiteSpace(testRoot))
            {
                return Path.GetFullPath(testRoot);
            }

            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MasseysLab",
                "CleanupSuite",
                "Jobs");
        }
    }

    public static string ValidateJobDirectory(string suppliedPath)
    {
        string job = ValidateUnhardenedJobDirectory(suppliedPath);
        JobAccessPolicy.ValidateOwnerOnly(job);
        return job;
    }

    public static string ValidateUnhardenedJobDirectory(string suppliedPath)
    {
        if (string.IsNullOrWhiteSpace(suppliedPath))
        {
            throw SecurityError("The job directory was not supplied.");
        }

        string root = Path.GetFullPath(ExpectedJobsRoot)
            .TrimEnd(Path.DirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        string job = Path.GetFullPath(suppliedPath)
            .TrimEnd(Path.DirectorySeparatorChar);
        string parent = Path.GetDirectoryName(job) ?? string.Empty;

        if (!string.Equals(
                parent.TrimEnd(Path.DirectorySeparatorChar),
                root.TrimEnd(Path.DirectorySeparatorChar),
                StringComparison.OrdinalIgnoreCase))
        {
            throw SecurityError("The job directory is outside the approved root.");
        }

        string jobName = Path.GetFileName(job);
        if (!Guid.TryParseExact(jobName, "D", out Guid parsed)
            || !string.Equals(
                parsed.ToString("D"),
                jobName,
                StringComparison.Ordinal))
        {
            throw SecurityError("The job directory name is not a lowercase UUID.");
        }

        if (!Directory.Exists(job))
        {
            throw SecurityError("The job directory does not exist.");
        }

        RejectReparsePoint(root.TrimEnd(Path.DirectorySeparatorChar));
        RejectReparsePoint(job);
        return job;
    }

    public static string FixedFile(string jobDirectory, string fileName)
    {
        if (fileName != Path.GetFileName(fileName)
            || fileName.Contains("..", StringComparison.Ordinal))
        {
            throw SecurityError("A protocol filename is not fixed and relative.");
        }

        string combined = Path.GetFullPath(
            Path.Combine(jobDirectory, fileName));
        string expectedParent = Path.GetFullPath(jobDirectory)
            .TrimEnd(Path.DirectorySeparatorChar);
        string actualParent = (Path.GetDirectoryName(combined) ?? string.Empty)
            .TrimEnd(Path.DirectorySeparatorChar);
        if (!string.Equals(
                expectedParent,
                actualParent,
                StringComparison.OrdinalIgnoreCase))
        {
            throw SecurityError("A protocol file escaped its job directory.");
        }

        if (File.Exists(combined))
        {
            RejectReparsePoint(combined);
        }

        return combined;
    }

    private static void RejectReparsePoint(string path)
    {
        FileAttributes attributes = File.GetAttributes(path);
        if ((attributes & FileAttributes.ReparsePoint) != 0)
        {
            throw SecurityError("A protocol path is a reparse point.");
        }
    }

    private static EngineContractException SecurityError(string message) =>
        new(
            "security-error",
            "security-policy-violation",
            message,
            ContractConstants.ExitSecurityError);
}
