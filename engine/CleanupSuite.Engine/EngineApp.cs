using System.Diagnostics;
using System.Text;
using System.Text.Json;

namespace MasseysLab.CleanupSuite.Engine;

public static class EngineApp
{
    public static int Run(string[] args)
    {
        Console.OutputEncoding = new UTF8Encoding(false);
        if (args.Length == 1 && args[0] == "--capabilities")
        {
            Console.WriteLine(
                JsonSerializer.Serialize(
                    Capabilities.Create(),
                    JsonSupport.WriteOptions));
            return ContractConstants.ExitCompleted;
        }

        if (args.Length == 1 && args[0] == "--version")
        {
            Console.WriteLine(ContractConstants.EngineVersion);
            return ContractConstants.ExitCompleted;
        }

        if (args.Length == 2 && args[0] == "--analyze")
        {
            return Analyze(args[1]);
        }

        Console.Error.WriteLine(
            "Usage: CleanupSuite.Engine --capabilities | --version | --analyze <job-directory>");
        return ContractConstants.ExitInvalidRequest;
    }

    private static int Analyze(string suppliedJobDirectory)
    {
        Stopwatch stopwatch = Stopwatch.StartNew();
        string? jobDirectory = null;
        AnalysisRequest? request = null;
        string jobId = "00000000-0000-4000-8000-000000000000";

        try
        {
            jobDirectory = JobPathPolicy.ValidateJobDirectory(
                suppliedJobDirectory);
            jobId = Path.GetFileName(jobDirectory);
            ValidatedJob job = RequestValidator.Load(jobDirectory);
            request = job.Request;
            AnalysisResponse response = ContractFixtureAnalyzer.Analyze(job);
            WriteResult(jobDirectory, response);
            SanitizedLog.TryAppend(
                jobDirectory,
                "analysis-completed",
                response.Candidates.Count);
            return ContractConstants.ExitCompleted;
        }
        catch (EngineContractException exception)
        {
            if (jobDirectory is not null)
            {
                AnalysisResponse errorResponse = ResponseFactory.Error(
                    jobId,
                    request,
                    exception,
                    stopwatch.ElapsedMilliseconds);
                TryWriteErrorResult(jobDirectory, errorResponse);
                SanitizedLog.TryAppend(
                    jobDirectory,
                    $"analysis-{exception.ErrorCode}");
            }

            Console.Error.WriteLine(exception.SafeMessage);
            return exception.ExitCode;
        }
        catch
        {
            EngineContractException exception = new(
                "internal-error",
                "internal-error",
                "The analysis engine encountered an internal error.",
                ContractConstants.ExitInternalError);
            if (jobDirectory is not null)
            {
                AnalysisResponse errorResponse = ResponseFactory.Error(
                    jobId,
                    request,
                    exception,
                    stopwatch.ElapsedMilliseconds);
                TryWriteErrorResult(jobDirectory, errorResponse);
                SanitizedLog.TryAppend(
                    jobDirectory,
                    "analysis-internal-error");
            }

            Console.Error.WriteLine(exception.SafeMessage);
            return exception.ExitCode;
        }
    }

    private static void WriteResult(
        string jobDirectory,
        AnalysisResponse response)
    {
        byte[] bytes = JsonSupport.SerializeUtf8(response);
        if (bytes.LongLength > ContractConstants.MaximumResultBytes)
        {
            throw new EngineContractException(
                "analysis-error",
                "analysis-failed",
                "The result exceeds its size limit.",
                ContractConstants.ExitAnalysisError);
        }

        string resultPath = JobPathPolicy.FixedFile(
            jobDirectory,
            ContractConstants.ResultFileName);
        AtomicFile.WriteNew(resultPath, bytes);
    }

    private static void TryWriteErrorResult(
        string jobDirectory,
        AnalysisResponse response)
    {
        try
        {
            WriteResult(jobDirectory, response);
        }
        catch
        {
            // A failed or unsafe result path must never be overwritten.
        }
    }
}
