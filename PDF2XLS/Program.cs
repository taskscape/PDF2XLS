using System.Diagnostics;
using System.Reflection;
using System.Text.Json.Nodes;
using Azure;
using Google.Apis.Sheets.v4;
using Microsoft.Extensions.Configuration;
using PDF2XLS.Helpers;
using Polly;
using Polly.Retry;
using Serilog;

namespace PDF2XLS;

class Program
{
    private static string PreferredApi { get; set; } = string.Empty;
    private static string ResponseSchema { get; set; } = string.Empty;
    private static Dictionary<string, string> Mappings { get; set; } = new();
    private static string SeqAddress { get; set; } = string.Empty;
    private static string SeqAppName { get; set; } = string.Empty;
    private static string SeqApiKey { get; set; } = string.Empty;
    private static string ExeDirectory { get; set; } = string.Empty;
    private static bool UploadPDFStatus { get; set; }
    private static string PDF2URLPath { get; set; } = string.Empty;
    private static Guid RunID { get; set; }
    private static string RunTime { get; set; } = string.Empty;

    static async Task Main(string[] args)
    {
        try
        {
            string? exePath = Process.GetCurrentProcess().MainModule?.FileName;
            ExeDirectory = Path.GetDirectoryName(exePath)!;

            IConfiguration config = new ConfigurationBuilder()
                .SetBasePath(ExeDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            // Load embedded JSON schema (used by OpenAIResponses to inject into the prompt).
            Assembly assembly = Assembly.GetExecutingAssembly();
            await using Stream schemaStream =
                assembly.GetManifestResourceStream("PDF2XLS.schema.json")!;
            using StreamReader schemaReader = new(schemaStream);
            ResponseSchema = await schemaReader.ReadToEndAsync();

            PreferredApi    = config["PreferredAPI"] ?? string.Empty;
            SeqAddress      = config["Seq:ServerAddress"] ?? string.Empty;
            SeqAppName      = config["Seq:AppName"] ?? string.Empty;
            SeqApiKey       = config["Seq:ApiKey"] ?? string.Empty;
            UploadPDFStatus = bool.Parse(config["UploadPDF:Enabled"] ?? "false");
            PDF2URLPath     = config["UploadPDF:PDF2URLPath"] ?? string.Empty;
            Mappings        = config.GetSection("GoogleSheets:Mappings")
                                    .Get<Dictionary<string, string>>() ?? new Dictionary<string, string>();
            RunID   = Guid.NewGuid();
            RunTime = DateTime.UtcNow.ToString("yyyyMMdd HHmmss");

            ConfigureLogger();

            Log.Information("Starting PDF2XLS application, Run ID: {RunID}", RunID);
            Log.Information(
                "Invocation — Exe: {Exe} | PreferredAPI: {Api} | UploadPDF: {UploadEnabled} | ArgCount: {ArgCount} | Args: {Args}",
                exePath,
                PreferredApi,
                UploadPDFStatus,
                args.Length,
                args.Select(a => $"\"{a}\"").ToArray());

            if (args.Length < 1)
            {
                Console.WriteLine("Usage: PDF2XLS <file.pdf> [file2.pdf ...] | <folder path>");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                return;
            }

            // ── Configuration validation ─────────────────────────────────────
            List<string> configErrors = ConfigurationValidator.Validate(config, PreferredApi, UploadPDFStatus);
            if (configErrors.Count > 0)
            {
                foreach (string error in configErrors)
                {
                    Console.WriteLine($"Configuration error: {error}");
                    Log.Error("Configuration error: {Error}", error);
                }
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                return;
            }

            // ── Spreadsheet name verification ───────────────────────────────
            GSheets sheetsValidator = new(config, string.Empty);
            SheetsService sheetsService = sheetsValidator.CreateSheetsService();
            if (!await sheetsValidator.VerifySpreadsheetName(sheetsService))
            {
                Console.WriteLine("Spreadsheet name mismatch — application cannot proceed. Check 'GoogleSheets:ExpectedSpreadsheetName' in appsettings.json.");
                Log.Error("Application aborted due to spreadsheet name mismatch.");
                return;
            }

            if (!await sheetsValidator.VerifySheetName(sheetsService))
            {
                Console.WriteLine("Google Sheets tab mismatch — application cannot proceed. Check 'GoogleSheets:SheetName' in appsettings.json.");
                Log.Error("Application aborted due to Google Sheets tab mismatch.");
                return;
            }

            AzureDocumentIntelligenceQuotaTracker? azureQuotaTracker = null;
            if (PreferredApi == "AzureDocumentIntelligence")
            {
                azureQuotaTracker = new AzureDocumentIntelligenceQuotaTracker(config, ExeDirectory);
                if (azureQuotaTracker.IsQuotaLimitReached())
                {
                    azureQuotaTracker.LogQuotaLimitReached();
                    return;
                }
            }

            // ── Input: accept file(s) or folder ─────────────────────────────
            InputPathResult inputResult = InputPathResolver.Resolve(args);
            if (!inputResult.IsSuccess)
            {
                Console.WriteLine(inputResult.ErrorMessage);
                if (inputResult.FailureKind == InputPathFailureKind.EmptyDirectory)
                    Log.Warning("{Error}", inputResult.ErrorMessage);
                else
                    Log.Error("Invalid input path: {Error}", inputResult.ErrorMessage);
                return;
            }

            if (inputResult.IsDirectory)
            {
                Log.Information(
                    "Processing {Count} PDF files from folder: {Folder}",
                    inputResult.Files.Count,
                    inputResult.InputPath);
            }
            else if (inputResult.IsMultipleInputs)
            {
                Log.Information(
                    "Processing {Count} PDF files from command line",
                    inputResult.Files.Count);
            }

            int processedCount = 0;
            foreach (string inputFilePath in inputResult.Files)
            {
                processedCount++;

                if (azureQuotaTracker?.IsQuotaLimitReached() == true)
                {
                    azureQuotaTracker.LogQuotaLimitReached();
                    break;
                }

                RunID = Guid.NewGuid();
                try
                {
                    bool shouldContinue = await ProcessFileAsync(inputFilePath, config, azureQuotaTracker);
                    if (!shouldContinue)
                    {
                        int remaining = inputResult.Files.Count - processedCount;
                        Log.Warning(
                            "Run stopped early after a blocking condition. {Remaining} of {Total} file(s) were not processed and will be retried automatically on the next scheduled run.",
                            remaining, inputResult.Files.Count);
                        break;
                    }
                }
                finally
                {
                    await FlushLogsAsync();
                }
            }
        }
        catch (Exception ex)
        {
            Environment.ExitCode = 1;
            Log.Fatal(ex, "An unhandled exception occurred in Main ({ExceptionType}: {ExceptionMessage}).", ex.GetType().Name, ex.Message);
        }
        finally
        {
            Log.Information("PDF2XLS application exiting with code {ExitCode}.", Environment.ExitCode);
            await Log.CloseAndFlushAsync();
        }
    }

    // ── Helpers ──────────────────────────────────────────────────────────────

    private static void ConfigureLogger()
    {
        Directory.CreateDirectory(Path.Combine(ExeDirectory, "logs"));

        LoggerConfiguration loggerConfig = new LoggerConfiguration()
            .Enrich.WithProperty("Application", SeqAppName)
            .MinimumLevel.Debug()
            .WriteTo.File(
                path: Path.Combine(ExeDirectory, "logs", "log-.txt"),
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 365,
                outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}"
            );

        if (!string.IsNullOrWhiteSpace(SeqAddress))
        {
            loggerConfig = loggerConfig.WriteTo.Seq(SeqAddress, apiKey: SeqApiKey);
        }

        Log.Logger = loggerConfig.CreateLogger();
    }

    private static async Task FlushLogsAsync()
    {
        await Log.CloseAndFlushAsync();
        ConfigureLogger();
    }

    private static async Task<bool> ProcessFileAsync(
        string inputFilePath,
        IConfiguration config,
        AzureDocumentIntelligenceQuotaTracker? azureQuotaTracker)
    {
        var stopwatch = System.Diagnostics.Stopwatch.StartNew();
        long fileSizeKb = new FileInfo(inputFilePath).Length / 1024;

        Log.Information(
            "──────────────────────────────────────────────────────────────────────");
        Log.Information(
            "Processing started. RunID: {RunID} | API: {Api} | File: {File} ({Size} KB)",
            RunID, PreferredApi, Path.GetFileName(inputFilePath), fileSizeKb);

        try
        {
            JsonNode? root = null;
            string? documentLink = string.Empty;

            switch (PreferredApi)
            {
                // ── NuDelta ──────────────────────────────────────────────────
                case "NuDelta":
                {
                    NuDeltaProcessor nuDeltaProcessor = new(config);

                    AsyncRetryPolicy retryPolicy = Policy
                        .Handle<Exception>(ShouldRetryWorkflowException)
                        .WaitAndRetryAsync(
                            retryCount: 5,
                            sleepDurationProvider: _ => TimeSpan.FromSeconds(1),
                            onRetry: (ex, ts, count, _) =>
                            {
                                Console.WriteLine(
                                    $"Retry {count} after {ts.TotalSeconds}s due to: {ex.Message}");
                                Log.Warning(ex,
                                    "Retry {Count} after {Delay}s. File: {file}",
                                    count, ts.TotalSeconds, inputFilePath);
                            });

                    await retryPolicy.ExecuteAsync(async () =>
                    {
                        string? r = await nuDeltaProcessor.ProcessPdfAsync(inputFilePath);
                        root = r != null ? JsonNode.Parse(r) : null;
                        EnsureInvoiceIssueData(root, "NuDelta");

                        if (UploadPDFStatus)
                        {
                            documentLink = await RunPDF2URL(PDF2URLPath, inputFilePath);
                            if (!StringHelper.IsValidHttpUrl(documentLink))
                                throw new Exception("Document failed to upload");
                            Log.Information("PDF uploaded. DocumentLink: {Link}. File: {file}", documentLink, inputFilePath);
                        }
                    });

                    break;
                }

                // ── OpenAI Responses API ──────────────────────────────────────
                case "OpenAIResponses":
                {
                    OpenAIResponsesProcessor openAIProcessor = new(config, ResponseSchema);

                    AsyncRetryPolicy openAIRetry = Policy
                        .Handle<Exception>(ShouldRetryWorkflowException)
                        .WaitAndRetryAsync(
                            retryCount: 3,
                            sleepDurationProvider: attempt =>
                                TimeSpan.FromSeconds(Math.Pow(2, attempt)),
                            onRetry: (ex, ts, count, _) =>
                            {
                                Console.WriteLine(
                                    $"Retry {count} after {ts.TotalSeconds}s due to: {ex.Message}");
                                Log.Warning(ex,
                                    "OpenAI Responses retry {Count} after {Delay}s. File: {file}",
                                    count, ts.TotalSeconds, inputFilePath);
                            });

                    await openAIRetry.ExecuteAsync(async () =>
                    {
                        string? r = await openAIProcessor.ProcessPdfAsync(inputFilePath);
                        root = r != null ? JsonNode.Parse(r) : null;
                        EnsureInvoiceIssueData(root, "OpenAI Responses");

                        if (UploadPDFStatus)
                        {
                            documentLink = await RunPDF2URL(PDF2URLPath, inputFilePath);
                            if (!StringHelper.IsValidHttpUrl(documentLink))
                                throw new Exception("Document failed to upload");
                            Log.Information("PDF uploaded. DocumentLink: {Link}. File: {file}", documentLink, inputFilePath);
                        }
                    });

                    break;
                }

                // ── Azure Document Intelligence ───────────────────────────────
                case "AzureDocumentIntelligence":
                {
                    AzureDocumentIntelligenceProcessor azDIProcessor = new(config, azureQuotaTracker);

                    // Retry ONCE when Azure Document Intelligence returns no usable invoice
                    // data (empty/unparseable result or not invoice-like). If it still has no
                    // data afterwards, the SkippedDocumentException propagates and the file is
                    // marked with the .skp extension.
                    AsyncRetryPolicy azDINoDataRetry = Policy
                        .Handle<SkippedDocumentException>()
                        .WaitAndRetryAsync(
                            retryCount: 1,
                            sleepDurationProvider: _ => TimeSpan.FromSeconds(2),
                            onRetry: (ex, ts, count, _) =>
                            {
                                Console.WriteLine(
                                    $"Retry {count} (no usable data) after {ts.TotalSeconds}s: {ex.Message}");
                                Log.Warning(ex,
                                    "Azure Document Intelligence returned no usable invoice data. Retry {Count} of 1 after {Delay}s. Reason: {Reason}. File: {file}",
                                    count, ts.TotalSeconds, ex.Message, inputFilePath);
                            });

                    // Retry TWICE when Azure Document Intelligence is not responding
                    // (transient/network errors, HTTP 5xx, etc.).
                    AsyncRetryPolicy azDINotRespondingRetry = Policy
                        .Handle<Exception>(ShouldRetryAzureDocumentIntelligenceException)
                        .WaitAndRetryAsync(
                            retryCount: 2,
                            sleepDurationProvider: attempt =>
                                TimeSpan.FromSeconds(Math.Pow(2, attempt)),
                            onRetry: (ex, ts, count, _) =>
                            {
                                Console.WriteLine(
                                    $"Retry {count} (not responding) after {ts.TotalSeconds}s: {ex.Message}");
                                Log.Warning(ex,
                                    "Azure Document Intelligence not responding. Retry {Count} of 2 after {Delay}s. Reason: {Reason}. File: {file}",
                                    count, ts.TotalSeconds, ex.Message, inputFilePath);
                            });

                    // Inner policy (no-data, retry once) runs closest to the call; the outer
                    // policy (not-responding, retry twice) wraps it.
                    var azDIPolicy = Policy.WrapAsync(azDINotRespondingRetry, azDINoDataRetry);

                    await azDIPolicy.ExecuteAsync(async () =>
                    {
                        string? r = await azDIProcessor.ProcessPdfAsync(inputFilePath);

                        root = r != null ? JsonNode.Parse(r) : null;
                        if (root == null)
                        {
                            throw new SkippedDocumentException(
                                "Azure Document Intelligence returned no usable result (empty or unparseable response).");
                        }

                        EnsureInvoiceIssueData(root, "Azure Document Intelligence");
                    });

                    if (UploadPDFStatus)
                    {
                        documentLink = await RunPDF2URL(PDF2URLPath, inputFilePath);
                        if (!StringHelper.IsValidHttpUrl(documentLink))
                            throw new Exception("Document failed to upload");
                        Log.Information("PDF uploaded. DocumentLink: {Link}. File: {file}", documentLink, inputFilePath);
                    }

                    break;
                }
            }

            // ── Parse JSON result and write to Google Sheets ──────────────────
            GSheets sheets = new(config, inputFilePath);
            SheetsService sheetsService = sheets.CreateSheetsService();

            Dictionary<string, string?> data = InvoiceDataMapper.Map(root, RunID, documentLink);

            Log.Information(
                "Extracted data — Invoice: {InvoiceNumber} | Date: {IssueDate} | Seller: {Seller} | Total: {Total} {Currency} | File: {File}",
                data.GetValueOrDefault("InvoiceNumber"),
                data.GetValueOrDefault("IssueDate"),
                data.GetValueOrDefault("SellerName"),
                data.GetValueOrDefault("TotalAmount"),
                data.GetValueOrDefault("Currency"),
                Path.GetFileName(inputFilePath));

            bool sheetsSuccess = await sheets.AppendRowWithBatchUpdate(sheetsService, data, Mappings);

            if (!sheetsSuccess)
            {
                Environment.ExitCode = 1;
                string? skippedPath = TryMarkFileAsSkipped(inputFilePath);
                Log.Warning(
                    "Google Sheets had no data to write (check field mappings). File marked as skipped to avoid reprocessing. File: {file} | SkippedFile: {SkippedFile}",
                    inputFilePath, skippedPath);
            }
            else
            {
                string bakPath = Path.Combine(
                    Path.GetDirectoryName(inputFilePath)!,
                    $"{RunTime} {RunID} {Path.GetFileName(inputFilePath)}.bak");
                File.Move(inputFilePath, bakPath);
                Log.Information("File archived as: {BakFile}", Path.GetFileName(bakPath));
            }

            stopwatch.Stop();
            Log.Information(
                "Processing complete. RunID: {RunID} | Elapsed: {Elapsed:F1}s | File: {File}",
                RunID, stopwatch.Elapsed.TotalSeconds, Path.GetFileName(inputFilePath));

            return true;
        }
        catch (AzureDocumentIntelligenceQuotaReachedException e)
        {
            stopwatch.Stop();
            Log.Information(e,
                "Processing stopped. RunID: {RunID} | Elapsed: {Elapsed:F1}s | File: {file}",
                RunID, stopwatch.Elapsed.TotalSeconds, inputFilePath);
            return false;
        }
        catch (AzureDocumentIntelligenceQuotaAccountingException e)
        {
            stopwatch.Stop();
            Console.WriteLine("There was an error while updating the Azure Document Intelligence quota counter. Processing stopped.");
            Log.Error(e,
                "Processing stopped because the Azure Document Intelligence quota counter could not be updated. RunID: {RunID} | Elapsed: {Elapsed:F1}s | File: {file}",
                RunID, stopwatch.Elapsed.TotalSeconds, inputFilePath);
            return false;
        }
        catch (GoogleSheetsConfigurationException e)
        {
            stopwatch.Stop();
            Console.WriteLine("Google Sheets is misconfigured or access was denied. The whole run has been aborted; no files were skipped. Fix the configuration before the next run.");
            Environment.ExitCode = 1;
            Log.Error(e,
                "RUN ABORTED — Google Sheets is misconfigured or permission was denied ({Reason}). This is a configuration problem that affects every file, so the entire run is being stopped without processing any further documents. The current file is left UNCHANGED (NOT marked as skipped) and will be retried automatically on the next scheduled run once the configuration is corrected. Review 'GoogleSheets:SpreadsheetId', 'GoogleSheets:SheetName', 'GoogleSheets:ServiceAccountFile', and the service account's share/permission on the spreadsheet. RunID: {RunID} | Elapsed: {Elapsed:F1}s | File: {file}",
                e.Message, RunID, stopwatch.Elapsed.TotalSeconds, inputFilePath);
            return false;
        }
        catch (GoogleSheetsCommunicationException e)
        {
            stopwatch.Stop();
            Environment.ExitCode = 1;
            string? skippedPath = TryMarkFileAsSkipped(inputFilePath);
            Console.WriteLine("Google Sheets could not be communicated. The file will be marked as skipped.");
            Log.Error(e,
                "Processing FAILED — Google Sheets could not be communicated after retries. File marked as skipped to avoid reprocessing. RunID: {RunID} | Elapsed: {Elapsed:F1}s | File: {file} | SkippedFile: {SkippedFile}",
                RunID, stopwatch.Elapsed.TotalSeconds, inputFilePath, skippedPath);
            return true;
        }
        catch (SkippedDocumentException e)
        {
            stopwatch.Stop();
            string? skippedPath = TryMarkFileAsSkipped(inputFilePath);
            Console.WriteLine(skippedPath == null
                ? "The file could not be processed as an invoice, but could not be marked as skipped."
                : "The file could not be processed as an invoice and was marked as skipped.");
            Log.Warning(e,
                "Processing skipped. RunID: {RunID} | Elapsed: {Elapsed:F1}s | File: {file} | SkippedFile: {SkippedFile}",
                RunID, stopwatch.Elapsed.TotalSeconds, inputFilePath, skippedPath);
            return true;
        }
        catch (Exception e)
        {
            stopwatch.Stop();
            Environment.ExitCode = 1;
            string? skippedPath = TryMarkFileAsSkipped(inputFilePath);
            Console.WriteLine("There was an error while processing the file. The file will be marked as skipped.");
            Log.Error(e,
                "Processing FAILED ({ExceptionType}: {ExceptionMessage}). File marked as skipped to avoid reprocessing. RunID: {RunID} | Elapsed: {Elapsed:F1}s | File: {file} | SkippedFile: {SkippedFile}",
                e.GetType().Name, e.Message, RunID, stopwatch.Elapsed.TotalSeconds, inputFilePath, skippedPath);
            return true;
        }
    }

    private static bool ShouldRetryAzureDocumentIntelligenceException(Exception ex)
    {
        if (ex is SkippedDocumentException)
            return false;

        if (ex is AzureDocumentIntelligenceQuotaException)
            return false;

        if (ex is OperationCanceledException)
            return false;

        if (ex is RequestFailedException { Status: >= 400 and < 500 })
            return false;

        return true;
    }

    private static bool ShouldRetryWorkflowException(Exception ex) =>
        ex is not OperationCanceledException and not SkippedDocumentException;

    private static void EnsureInvoiceIssueData(JsonNode? root, string source)
    {
        if (root == null)
            throw new InvalidOperationException($"{source} returned no parseable result");

        string issue = root["data"]?["issue"]?.ToString() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(issue))
        {
            throw new SkippedDocumentException(
                $"{source} response is missing issue data; document does not look like an invoice.");
        }
    }

    private static string? TryMarkFileAsSkipped(string inputFilePath)
    {
        try
        {
            string skippedPath = GetAvailableSiblingPath($"{inputFilePath}.skp");
            File.Move(inputFilePath, skippedPath);
            return skippedPath;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to mark file as skipped. File: {file}", inputFilePath);
            return null;
        }
    }

    private static string GetAvailableSiblingPath(string preferredPath)
    {
        if (!File.Exists(preferredPath))
            return preferredPath;

        string directory = Path.GetDirectoryName(preferredPath) ?? string.Empty;
        string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(preferredPath);
        string extension = Path.GetExtension(preferredPath);

        for (int i = 1; i <= 999; i++)
        {
            string candidate = Path.Combine(directory, $"{fileNameWithoutExtension} ({i}){extension}");
            if (!File.Exists(candidate))
                return candidate;
        }

        return Path.Combine(directory, $"{fileNameWithoutExtension} {Guid.NewGuid():N}{extension}");
    }

    private static async Task<string> RunPDF2URL(string exePath, string filePath)
    {
        ProcessStartInfo startInfo = new()
        {
            FileName = exePath,
            Arguments = $"\"{filePath}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        var uploadStopwatch = System.Diagnostics.Stopwatch.StartNew();
        Log.Information("Starting PDF upload via PDF2URL: {Exe} \"{file}\"", exePath, filePath);

        using var process = Process.Start(startInfo);
        if (process == null)
        {
            Log.Warning("PDF2URL process could not be started ({Exe}). File: {file}", exePath, filePath);
            return string.Empty;
        }

        // Read stdout/stderr asynchronously to avoid pipe-buffer deadlock with WaitForExit.
        Task<string> readOutputTask = process.StandardOutput.ReadToEndAsync();
        Task<string> readErrorTask = process.StandardError.ReadToEndAsync();

        bool exited = process.WaitForExit(300_000); // 5-minute timeout
        if (!exited)
        {
            try { process.Kill(); } catch { /* best effort */ }
            uploadStopwatch.Stop();
            Log.Warning(
                "PDF2URL process timed out after {Elapsed:F1}s (5-minute limit) and was killed. File: {file}",
                uploadStopwatch.Elapsed.TotalSeconds, filePath);
            return string.Empty;
        }

        string output = await readOutputTask;
        string error = await readErrorTask;
        uploadStopwatch.Stop();

        if (process.ExitCode != 0)
        {
            Log.Warning(
                "PDF2URL exited with code {Code} after {Elapsed:F1}s. Stdout: {Output} | Stderr: {Error}. File: {file}",
                process.ExitCode, uploadStopwatch.Elapsed.TotalSeconds, output.TrimEnd(), error.TrimEnd(), filePath);
            return string.Empty;
        }

        Log.Information(
            "PDF upload finished in {Elapsed:F1}s. File: {file}",
            uploadStopwatch.Elapsed.TotalSeconds, filePath);

        return output.TrimEnd();
    }
}
