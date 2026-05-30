using System.Diagnostics;
using System.Reflection;
using System.Text.Json.Nodes;
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

            foreach (string inputFilePath in inputResult.Files)
            {
                RunID = Guid.NewGuid();
                try
                {
                    await ProcessFileAsync(inputFilePath, config);
                }
                finally
                {
                    await FlushLogsAsync();
                }
            }
        }
        catch (Exception ex)
        {
            Log.Fatal(ex, "An unhandled exception occurred in Main.");
        }
        finally
        {
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

    private static async Task ProcessFileAsync(string inputFilePath, IConfiguration config)
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
                        .Handle<Exception>(ex => ex is not OperationCanceledException)
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
                        if (root?["data"]?["issue"] == null ||
                            string.IsNullOrEmpty(root["data"]!["issue"]?.ToString()))
                        {
                            throw new InvalidOperationException(
                                "JSON response is missing issue data");
                        }

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
                        .Handle<Exception>(ex => ex is not OperationCanceledException)
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
                        if (root?["data"]?["issue"] == null ||
                            string.IsNullOrEmpty(root["data"]!["issue"]?.ToString()))
                        {
                            throw new InvalidOperationException(
                                "JSON response is missing issue data");
                        }

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
                    AzureDocumentIntelligenceProcessor azDIProcessor = new(config);

                    AsyncRetryPolicy azDIRetry = Policy
                        .Handle<Exception>(ex => ex is not OperationCanceledException)
                        .WaitAndRetryAsync(
                            retryCount: 3,
                            sleepDurationProvider: attempt =>
                                TimeSpan.FromSeconds(Math.Pow(2, attempt)),
                            onRetry: (ex, ts, count, _) =>
                            {
                                Console.WriteLine(
                                    $"Retry {count} after {ts.TotalSeconds}s due to: {ex.Message}");
                                Log.Warning(ex,
                                    "Azure Document Intelligence retry {Count} after {Delay}s. File: {file}",
                                    count, ts.TotalSeconds, inputFilePath);
                            });

                    await azDIRetry.ExecuteAsync(async () =>
                    {
                        string? r = await azDIProcessor.ProcessPdfAsync(inputFilePath);

                        root = r != null ? JsonNode.Parse(r) : null;
                        if (root == null)
                        {
                            throw new InvalidOperationException(
                                "Azure Document Intelligence returned no result");
                        }

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
                Log.Warning("Google Sheets write failed — file will NOT be archived. File: {file}", inputFilePath);
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
        }
        catch (Exception e)
        {
            stopwatch.Stop();
            Console.WriteLine("There was an error while processing the file. Please try again.");
            Log.Error(e,
                "Processing FAILED. RunID: {RunID} | Elapsed: {Elapsed:F1}s | File: {file}",
                RunID, stopwatch.Elapsed.TotalSeconds, inputFilePath);
        }
    }

    private static async Task<string> RunPDF2URL(string exePath, string filePath)
    {
        ProcessStartInfo startInfo = new()
        {
            FileName = exePath,
            Arguments = $"\"{filePath}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            CreateNoWindow = true
        };

        using var process = Process.Start(startInfo);
        if (process == null) return string.Empty;

        // Read stdout asynchronously to avoid pipe-buffer deadlock with WaitForExit.
        Task<string> readTask = process.StandardOutput.ReadToEndAsync();

        bool exited = process.WaitForExit(300_000); // 5-minute timeout
        if (!exited)
        {
            try { process.Kill(); } catch { /* best effort */ }
            Log.Warning("PDF2URL process timed out after 5 minutes and was killed. File: {file}", filePath);
            return string.Empty;
        }

        string output = await readTask;

        if (process.ExitCode != 0)
        {
            Log.Warning("PDF2URL exited with code {Code}. Output: {Output}. File: {file}",
                process.ExitCode, output.TrimEnd(), filePath);
            return string.Empty;
        }

        return output.TrimEnd();
    }
}