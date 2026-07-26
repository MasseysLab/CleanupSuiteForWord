using System.Text;
using System.Text.Json;
using System.Diagnostics;
using MasseysLab.CleanupSuite.Engine;

return ContractTestRunner.Run(args);

internal static class ContractTestRunner
{
    private static int _passed;
    private static string? _resultOutputPath;

    public static int Run(string[] args)
    {
        if (args.Length == 2 && args[0] == "--result-output")
        {
            _resultOutputPath = Path.GetFullPath(args[1]);
        }
        else if (args.Length != 0)
        {
            Console.Error.WriteLine(
                "Usage: CleanupSuite.Engine.ContractTests [--result-output <path>]");
            return 2;
        }

        Environment.SetEnvironmentVariable(
            "DOTNET_CLI_TELEMETRY_OPTOUT",
            "1");
        RunTest("Capabilities are safe and contract compatible", TestCapabilities);
        RunTest("Valid analysis returns deterministic candidates", TestValidAnalysis);
        RunTest("Cancellation returns no candidates", TestCancellation);
        RunTest("Hash mismatch returns no candidates", TestHashMismatch);
        RunTest("Malformed UTF-8 returns no candidates", TestMalformedUtf8);
        RunTest("Version mismatch returns no candidates", TestVersionMismatch);
        RunTest("Unknown request fields are rejected", TestUnknownField);
        RunTest("Outside-root jobs are rejected", TestOutsideRoot);
        RunTest("Inherited job permissions are rejected", TestInheritedPermissions);
        RunTest("Existing result files are never overwritten", TestExistingResult);
        RunTest("UTF-16 context never splits a surrogate pair", TestSurrogateContext);
        RunTest("Large snapshots stay on the fast deterministic path", TestLargeSnapshot);
        Console.WriteLine($"PASS|Hybrid Engine Contract|{_passed} tests passed.");
        return 0;
    }

    private static void RunTest(string name, Action test)
    {
        try
        {
            test();
            _passed++;
            Console.WriteLine($"PASS|{name}");
        }
        catch (Exception exception)
        {
            Console.Error.WriteLine(
                $"FAIL|{name}|{exception.GetType().Name}: {exception.Message}");
            Environment.Exit(1);
        }
    }

    private static void TestCapabilities()
    {
        TextWriter original = Console.Out;
        using StringWriter capture = new();
        Console.SetOut(capture);
        int exitCode;
        try
        {
            exitCode = EngineApp.Run(["--capabilities"]);
        }
        finally
        {
            Console.SetOut(original);
        }

        AssertEqual(0, exitCode, "capability exit code");
        using JsonDocument document = JsonDocument.Parse(capture.ToString());
        JsonElement root = document.RootElement;
        AssertEqual("1.0", root.GetProperty("contractVersion").GetString(), "contract");
        JsonElement security = root.GetProperty("security");
        AssertFalse(security.GetProperty("editsWordDocuments").GetBoolean(), "Word edits");
        AssertFalse(security.GetProperty("requiresNetwork").GetBoolean(), "network");
        AssertFalse(security.GetProperty("runsAsService").GetBoolean(), "service");
        AssertFalse(security.GetProperty("requiresElevation").GetBoolean(), "elevation");
        AssertFalse(security.GetProperty("logsDocumentContent").GetBoolean(), "content logging");
    }

    private static void TestValidAnalysis()
    {
        using JobFixture fixture = JobFixture.Create(
            "alpha 😀 alpha\r",
            "alpha",
            "omega");
        int exitCode = RunWithCapturedError(
            ["--analyze", fixture.DirectoryPath]);
        AssertEqual(0, exitCode, "analysis exit code");
        using JsonDocument result = fixture.ReadResult();
        JsonElement root = result.RootElement;
        AssertEqual("completed", root.GetProperty("status").GetString(), "status");
        JsonElement candidates = root.GetProperty("candidates");
        AssertEqual(2, candidates.GetArrayLength(), "candidate count");
        AssertEqual(0, candidates[0].GetProperty("location").GetProperty("startUtf16").GetInt32(), "first start");
        AssertEqual(9, candidates[1].GetProperty("location").GetProperty("startUtf16").GetInt32(), "second start");
        AssertStrictRevalidation(candidates[0]);
        AssertStrictRevalidation(candidates[1]);
        AssertFalse(File.Exists(fixture.ResultPath + ".tmp"), "temporary result");
        if (_resultOutputPath is not null)
        {
            string? outputDirectory = Path.GetDirectoryName(_resultOutputPath);
            if (!string.IsNullOrEmpty(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            File.Copy(
                fixture.ResultPath,
                _resultOutputPath,
                overwrite: true);
        }

        string log = File.ReadAllText(fixture.LogPath, JsonSupport.StrictUtf8);
        AssertFalse(log.Contains("alpha", StringComparison.Ordinal), "source leaked to log");
        AssertFalse(log.Contains(fixture.DirectoryPath, StringComparison.OrdinalIgnoreCase), "path leaked to log");
    }

    private static void TestCancellation()
    {
        using JobFixture fixture = JobFixture.Create("alpha\r", "alpha", "omega");
        File.WriteAllBytes(fixture.CancelPath, []);
        int exitCode = RunWithCapturedError(
            ["--analyze", fixture.DirectoryPath]);
        AssertEqual(ContractConstants.ExitCancelled, exitCode, "cancel exit code");
        using JsonDocument result = fixture.ReadResult();
        AssertEqual("cancelled", result.RootElement.GetProperty("status").GetString(), "cancel status");
        AssertEqual(0, result.RootElement.GetProperty("candidates").GetArrayLength(), "cancel candidates");
    }

    private static void TestHashMismatch()
    {
        using JobFixture fixture = JobFixture.Create("alpha\r", "alpha", "omega");
        fixture.ReplaceRequestText(
            fixture.RequestText.Replace(
                fixture.SnapshotSha256,
                new string('a', 64),
                StringComparison.Ordinal));
        int exitCode = RunWithCapturedError(
            ["--analyze", fixture.DirectoryPath]);
        AssertEqual(ContractConstants.ExitInvalidRequest, exitCode, "hash exit code");
        using JsonDocument result = fixture.ReadResult();
        AssertEqual(0, result.RootElement.GetProperty("candidates").GetArrayLength(), "hash candidates");
    }

    private static void TestUnknownField()
    {
        using JobFixture fixture = JobFixture.Create("alpha\r", "alpha", "omega");
        fixture.ReplaceRequestText(
            fixture.RequestText.TrimEnd('}', '\r', '\n')
                + ",\"documentPath\":\"C:\\\\private.docx\"}");
        int exitCode = RunWithCapturedError(
            ["--analyze", fixture.DirectoryPath]);
        AssertEqual(ContractConstants.ExitInvalidRequest, exitCode, "unknown field exit");
        using JsonDocument result = fixture.ReadResult();
        AssertEqual(0, result.RootElement.GetProperty("candidates").GetArrayLength(), "unknown field candidates");
    }

    private static void TestMalformedUtf8()
    {
        using JobFixture fixture = JobFixture.Create("alpha\r", "alpha", "omega");
        File.WriteAllBytes(fixture.ContentPath, [0xFF]);
        int exitCode = RunWithCapturedError(
            ["--analyze", fixture.DirectoryPath]);
        AssertEqual(
            ContractConstants.ExitInvalidRequest,
            exitCode,
            "malformed UTF-8 exit");
        using JsonDocument result = fixture.ReadResult();
        AssertEqual(
            "invalid-snapshot-text",
            result.RootElement.GetProperty("error").GetProperty("code").GetString(),
            "malformed UTF-8 code");
        AssertEqual(
            0,
            result.RootElement.GetProperty("candidates").GetArrayLength(),
            "malformed UTF-8 candidates");
    }

    private static void TestVersionMismatch()
    {
        using JobFixture fixture = JobFixture.Create("alpha\r", "alpha", "omega");
        fixture.ReplaceRequestText(
            fixture.RequestText.Replace(
                "\"protocolVersion\": \"1.0\"",
                "\"protocolVersion\": \"2.0\"",
                StringComparison.Ordinal));
        int exitCode = RunWithCapturedError(
            ["--analyze", fixture.DirectoryPath]);
        AssertEqual(
            ContractConstants.ExitIncompatible,
            exitCode,
            "version mismatch exit");
        using JsonDocument result = fixture.ReadResult();
        AssertEqual(
            "incompatible",
            result.RootElement.GetProperty("status").GetString(),
            "version mismatch status");
        AssertEqual(
            0,
            result.RootElement.GetProperty("candidates").GetArrayLength(),
            "version mismatch candidates");
    }

    private static void TestOutsideRoot()
    {
        string outside = Path.Combine(
            Path.GetTempPath(),
            Guid.NewGuid().ToString("D"));
        Directory.CreateDirectory(outside);
        try
        {
            int exitCode = RunWithCapturedError(["--analyze", outside]);
            AssertEqual(ContractConstants.ExitSecurityError, exitCode, "outside-root exit");
            AssertFalse(File.Exists(Path.Combine(outside, "result.json")), "outside-root result");
        }
        finally
        {
            Directory.Delete(outside, recursive: true);
        }
    }

    private static void TestExistingResult()
    {
        using JobFixture fixture = JobFixture.Create("alpha\r", "alpha", "omega");
        byte[] sentinel = JsonSupport.StrictUtf8.GetBytes("{\"sentinel\":true}");
        File.WriteAllBytes(fixture.ResultPath, sentinel);
        int exitCode = RunWithCapturedError(
            ["--analyze", fixture.DirectoryPath]);
        AssertEqual(ContractConstants.ExitSecurityError, exitCode, "existing result exit");
        AssertSequenceEqual(sentinel, File.ReadAllBytes(fixture.ResultPath), "existing result bytes");
    }

    private static void TestInheritedPermissions()
    {
        string jobId = Guid.NewGuid().ToString("D");
        Directory.CreateDirectory(JobPathPolicy.ExpectedJobsRoot);
        string directory = Path.Combine(
            JobPathPolicy.ExpectedJobsRoot,
            jobId);
        Directory.CreateDirectory(directory);
        try
        {
            int exitCode = RunWithCapturedError(["--analyze", directory]);
            AssertEqual(
                ContractConstants.ExitSecurityError,
                exitCode,
                "inherited permission exit");
            AssertFalse(
                File.Exists(Path.Combine(directory, "result.json")),
                "inherited permission result");
        }
        finally
        {
            Directory.Delete(directory, recursive: true);
        }
    }

    private static void TestSurrogateContext()
    {
        string text = "😀" + new string('A', 31) + "token\r";
        using JobFixture fixture = JobFixture.Create(text, "token", "replacement");
        int exitCode = EngineApp.Run(["--analyze", fixture.DirectoryPath]);
        AssertEqual(0, exitCode, "surrogate context exit");
        using JsonDocument result = fixture.ReadResult();
        JsonElement fingerprint = result.RootElement
            .GetProperty("candidates")[0]
            .GetProperty("fingerprint");
        AssertEqual(31, fingerprint.GetProperty("prefixLengthUtf16").GetInt32(), "aligned prefix length");
    }

    private static void TestLargeSnapshot()
    {
        const int repeats = 1_024;
        string block = new('x', 2_040);
        StringBuilder builder = new(repeats * (block.Length + 5) + 1);
        for (int index = 0; index < repeats; index++)
        {
            builder.Append(block);
            builder.Append("token");
        }

        builder.Append('\r');
        using JobFixture fixture = JobFixture.Create(
            builder.ToString(),
            "token",
            "replacement");
        Stopwatch stopwatch = Stopwatch.StartNew();
        int exitCode = EngineApp.Run(["--analyze", fixture.DirectoryPath]);
        stopwatch.Stop();
        AssertEqual(0, exitCode, "large snapshot exit");
        using JsonDocument result = fixture.ReadResult();
        AssertEqual(
            repeats,
            result.RootElement.GetProperty("candidates").GetArrayLength(),
            "large snapshot candidate count");
        AssertTrue(
            stopwatch.Elapsed < TimeSpan.FromSeconds(15),
            "large snapshot completed within the regression ceiling");
    }

    private static void AssertStrictRevalidation(JsonElement candidate)
    {
        JsonElement revalidation = candidate.GetProperty("revalidation");
        AssertTrue(revalidation.GetProperty("requireWholeScopeSnapshot").GetBoolean(), "whole scope");
        AssertTrue(revalidation.GetProperty("requireExactRange").GetBoolean(), "exact range");
        AssertTrue(revalidation.GetProperty("requireContext").GetBoolean(), "context");
        AssertFalse(revalidation.GetProperty("allowRelocation").GetBoolean(), "relocation");
        AssertEqual("abort-apply", revalidation.GetProperty("onMismatch").GetString(), "mismatch");
    }

    private static int RunWithCapturedError(string[] args)
    {
        TextWriter original = Console.Error;
        using StringWriter capture = new();
        Console.SetError(capture);
        try
        {
            return EngineApp.Run(args);
        }
        finally
        {
            Console.SetError(original);
        }
    }

    private static void AssertTrue(bool value, string label)
    {
        if (!value)
        {
            throw new InvalidOperationException($"Expected true: {label}.");
        }
    }

    private static void AssertFalse(bool value, string label) =>
        AssertTrue(!value, label);

    private static void AssertEqual<T>(T expected, T actual, string label)
    {
        if (!EqualityComparer<T>.Default.Equals(expected, actual))
        {
            throw new InvalidOperationException(
                $"{label}: expected {expected}, received {actual}.");
        }
    }

    private static void AssertSequenceEqual(
        byte[] expected,
        byte[] actual,
        string label)
    {
        if (!expected.AsSpan().SequenceEqual(actual))
        {
            throw new InvalidOperationException($"{label} changed.");
        }
    }
}

internal sealed class JobFixture : IDisposable
{
    private JobFixture(
        string directoryPath,
        string requestText,
        string snapshotSha256)
    {
        DirectoryPath = directoryPath;
        RequestText = requestText;
        SnapshotSha256 = snapshotSha256;
    }

    public string DirectoryPath { get; }
    public string RequestText { get; private set; }
    public string SnapshotSha256 { get; }
    public string RequestPath => Path.Combine(DirectoryPath, "request.json");
    public string ContentPath => Path.Combine(
        DirectoryPath,
        "document.utf8.txt");
    public string ResultPath => Path.Combine(DirectoryPath, "result.json");
    public string CancelPath => Path.Combine(DirectoryPath, "cancel.request");
    public string LogPath => Path.Combine(DirectoryPath, "engine.log");

    public static JobFixture Create(
        string snapshotText,
        string source,
        string replacement)
    {
        string jobId = Guid.NewGuid().ToString("D");
        string root = JobPathPolicy.ExpectedJobsRoot;
        Directory.CreateDirectory(root);
        string directory = Path.Combine(root, jobId);
        Directory.CreateDirectory(directory);
        JobAccessPolicy.HardenForCurrentUser(directory);
        byte[] snapshotBytes = JsonSupport.StrictUtf8.GetBytes(snapshotText);
        string snapshotHash = Hashing.Sha256Hex(snapshotBytes);
        File.WriteAllBytes(
            Path.Combine(directory, "document.utf8.txt"),
            snapshotBytes);

        object request = new
        {
            contractVersion = "1.0",
            messageType = "analysis-request",
            jobId,
            createdUtc = DateTimeOffset.UtcNow,
            client = new
            {
                suiteVersion = "0.9.5-beta",
                protocolVersion = "1.0",
                processId = Environment.ProcessId,
                sessionId = Guid.NewGuid().ToString("D")
            },
            tool = new
            {
                id = "contract-fixture",
                definitionVersion = "1.0.0",
                analysisMode = "replace-literal"
            },
            scope = new
            {
                documentSessionId = Guid.NewGuid().ToString("D"),
                storyType = "main-text",
                startUtf16 = 0,
                endUtf16 = snapshotText.Length,
                selectionOnly = false
            },
            snapshot = new
            {
                snapshotId = snapshotHash,
                contentFile = "document.utf8.txt",
                encoding = "utf-8",
                byteLength = snapshotBytes.LongLength,
                utf16Length = snapshotText.Length,
                sha256 = snapshotHash,
                lineEndingPolicy = "preserve-word-story-text"
            },
            options = new
            {
                from = source,
                to = replacement
            },
            requestedCapabilities = new[]
            {
                "analysis.contract-fixture",
                "fingerprint.sha256-utf8-exact"
            },
            privacy = new
            {
                documentPathIncluded = false,
                documentNameIncluded = false
            }
        };
        string requestText = JsonSerializer.Serialize(
            request,
            JsonSupport.WriteOptions);
        File.WriteAllText(
            Path.Combine(directory, "request.json"),
            requestText,
            JsonSupport.StrictUtf8);
        return new JobFixture(directory, requestText, snapshotHash);
    }

    public void ReplaceRequestText(string text)
    {
        RequestText = text;
        File.WriteAllText(RequestPath, text, JsonSupport.StrictUtf8);
    }

    public JsonDocument ReadResult() =>
        JsonDocument.Parse(File.ReadAllBytes(ResultPath));

    public void Dispose()
    {
        string fullPath = Path.GetFullPath(DirectoryPath);
        string expectedRoot = Path.GetFullPath(JobPathPolicy.ExpectedJobsRoot)
            .TrimEnd(Path.DirectorySeparatorChar)
            + Path.DirectorySeparatorChar;
        if (!fullPath.StartsWith(
                expectedRoot,
                StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                "Refusing to delete a test job outside the approved root.");
        }

        if (Directory.Exists(fullPath))
        {
            Directory.Delete(fullPath, recursive: true);
        }
    }
}
