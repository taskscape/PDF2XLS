using System.ClientModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Text.Json.Nodes;
using Google.Apis.Sheets.v4;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using OpenAI;
using OpenAI.Assistants;
using OpenAI.Files;
using Polly;
using Polly.Retry;
using Serilog;
using JsonSerializer = System.Text.Json.JsonSerializer;

namespace PDF2XLS;

class Program
{
    private static string Username { get; set; }
    private static string Password { get; set; }
    private static string PreferredApi { get; set; }
    private static string OpenAiApiKey { get; set; }
    private static string ResponseSchema { get; set; }
    private static string NuDeltaBaseUrl = "https://www.nudelta.pl/api/v1";
    private static bool DeleteAfter { get; set; }
    private static Dictionary<string, string> Mappings { get; set; }
    private static string SeqAddress { get; set; }
    private static string SeqAppName { get; set; }
    private static bool UploadPDFStatus { get; set; }
    private static string PDF2URLPath { get; set; }

    [Experimental("OPENAI001")]
    static async Task Main(string[] args)
    {
        try
        {
            string? exePath = Process.GetCurrentProcess().MainModule?.FileName;
            string realExeDirectory = Path.GetDirectoryName(exePath);
            IConfiguration config = new ConfigurationBuilder()
                .SetBasePath(realExeDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();
            
            Assembly assembly = Assembly.GetExecutingAssembly();
            string resourceName = "PDF2XLS.schema.json";
            string fileContents;
            await using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new(stream))
            {
                fileContents = await reader.ReadToEndAsync();
            }
            
            Username = config["NuDeltaCredentials:Username"] ?? string.Empty;
            Password = config["NuDeltaCredentials:Password"] ?? string.Empty;
            PreferredApi = config["PreferredAPI"] ?? string.Empty;
            OpenAiApiKey = config["OpenAI_APIKey"] ?? string.Empty;
            ResponseSchema = fileContents;
            DeleteAfter = bool.Parse(config["DeleteFileAfterProcessing"]);
            SeqAddress = config["Seq:ServerAddress"] ?? string.Empty;
            SeqAppName = config["Seq:AppName"] ?? string.Empty;
            UploadPDFStatus = bool.Parse(config["UploadPDF:Enabled"]);
            PDF2URLPath = config["UploadPDF:PDF2URLPath"] ?? string.Empty;
            Mappings = config.GetSection("GoogleSheets:Mappings")
                .Get<Dictionary<string, string>>() ?? new Dictionary<string, string>();

            Log.Logger = new LoggerConfiguration()
                .Enrich.WithProperty("Application", SeqAppName)
                .MinimumLevel.Debug()
                .WriteTo.File(
                    path: $"{realExeDirectory}/logs/log-.txt",
                    rollingInterval: RollingInterval.Day,
                    retainedFileCountLimit: 7,
                    outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}"
                )
                .WriteTo.Seq(SeqAddress)
                .CreateLogger();

            Log.Information("Starting PDF2XLS application...");
            
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: PDF2XLS <input file path>");
                Console.WriteLine("Press any key to exit...");
                Console.ReadKey();
                return;
            }

            string inputFilePath = args[0];
            if (!File.Exists(inputFilePath))
            {
                Console.WriteLine($"File {inputFilePath} does not exist");
                Log.Error("File {InputFilePath} does not exist", inputFilePath);
                return;
            }

            if (!string.Equals(Path.GetExtension(inputFilePath), ".pdf", StringComparison.OrdinalIgnoreCase))
            {
                Console.WriteLine($"File {inputFilePath} is not a PDF file");
                Log.Error("File {InputFilePath} is not a PDF file", inputFilePath);
                return;
            }

            try
            {
                AsyncRetryPolicy? retryPolicy = Policy
                    .Handle<Exception>()
                    .WaitAndRetryAsync(
                        retryCount: 5,
                        sleepDurationProvider: attempt => TimeSpan.FromSeconds(1),
                        onRetry: (exception, timeSpan, retryCount, context) =>
                        {
                            Console.WriteLine($"Retry {retryCount} after {timeSpan.TotalSeconds}s due to: {exception.Message}");
                            Log.Warning(exception,
                                "Retry {RetryCount} after {TimeSpanSeconds}s due to exception",
                                retryCount, timeSpan.TotalSeconds);
                        }
                    );

                JsonNode root = null;
                await retryPolicy.ExecuteAsync(async () =>
                {
                    string response = await GetJsonResponse(inputFilePath);
                    root = JsonNode.Parse(response);
                    if (root?["data"]?["issue"] == null || string.IsNullOrEmpty(root["data"]["issue"].ToString()))
                    {
                        throw new InvalidOperationException("JSON response is missing or empty issue data");
                    }
                });
                JsonNode? dataNode = root?["data"];

                // Extract top-level fields
                string invNumber = GetValFromNode(dataNode?["invn"]);
                string refNumber = GetValFromNode(dataNode?["reference"]);
                string issueDateString = GetValFromNode(dataNode?["issue"]);
                DateTime.TryParse(issueDateString, out DateTime issueDate);
                issueDateString = issueDate.ToString("yyyy-MM-dd");
                string saleDateString = GetValFromNode(dataNode?["sale"]);
                DateTime.TryParse(saleDateString, out DateTime saleDate);
                saleDateString = saleDate.ToString("yyyy-MM-dd");
                string paymentMethod = GetValFromNode(dataNode?["payment"]);
                string maturity = GetValFromNode(dataNode?["maturity"]);
                string currency = GetValFromNode(dataNode?["currency"]);
                string totalAmount = GetValFromNode(dataNode?["total"]);
                string paidAmount = GetValFromNode(dataNode?["paid"]);
                string leftToPay = GetValFromNode(dataNode?["left"]);
                string iban = GetValFromNode(dataNode?["iban"]);

                // Seller info
                JsonNode? seller = dataNode?["seller"];
                string sellerNip = GetValFromNode(seller?["nip"]);
                string sellerName = GetValFromNode(seller?["name"]);
                string sellerCity = GetValFromNode(seller?["city"]);
                string sellerStreet = GetValFromNode(seller?["street"]);
                string sellerZip = GetValFromNode(seller?["zipcode"]);

                // Buyer info
                JsonNode? buyer = dataNode?["buyer"];
                string buyerNip = GetValFromNode(buyer?["nip"]);
                string buyerName = GetValFromNode(buyer?["name"]);
                string buyerCity = GetValFromNode(buyer?["city"]);
                string buyerStreet = GetValFromNode(buyer?["street"]);
                string buyerZip = GetValFromNode(buyer?["zipcode"]);

                // Table rows
                JsonNode? tablesNode = dataNode?["tables"];
                JsonArray totals = tablesNode?["total"]?.AsArray() ?? [];
               
                JsonNode? totalNode = totals.Count > 0 ? totals[0] : null;
                string totalNet = GetValFromNode(totalNode?["valNetto"]);
                string totalVat = GetValFromNode(totalNode?["valVat"]);
                string totalGross = GetValFromNode(totalNode?["valBrutto"]);

                string documentLink = string.Empty;
                if (UploadPDFStatus)
                {
                    documentLink = RunPDF2URL(PDF2URLPath, inputFilePath);
                }

                GSheets sheets = new GSheets(config);
                SheetsService sheetsService = sheets.CreateSheetsService();
                Dictionary<string, string> data = new Dictionary<string, string>
                {
                    { "InvoiceNumber", invNumber },
                    { "ReferenceNumber", refNumber },
                    { "IssueDate", issueDateString },
                    { "SaleDate", saleDateString },
                    { "PaymentMethod", paymentMethod },
                    { "Maturity", maturity },
                    { "Currency", currency },
                    { "TotalAmount", totalAmount },
                    { "PaidAmount", paidAmount },
                    { "AmountLeftToPay", leftToPay },
                    { "IBAN", iban },
                    { "SellerNIP", sellerNip },
                    { "SellerName", sellerName },
                    { "SellerCity", sellerCity },
                    { "SellerStreet", sellerStreet },
                    { "SellerZIP", sellerZip },
                    { "BuyerNIP", buyerNip },
                    { "BuyerName", buyerName },
                    { "BuyerCity", buyerCity },
                    { "BuyerStreet", buyerStreet },
                    { "BuyerZIP", buyerZip },
                    { "DocumentLink", documentLink },
                    { "TotalNet", totalNet },
                    { "TotalVat", totalVat },
                    { "TotalGross", totalGross }
                };
                
                sheets.AddRow(sheetsService, data, Mappings);
                
                if (DeleteAfter)
                {
                    File.Delete(inputFilePath);
                }
                else
                {
                    File.Move(inputFilePath, Path.Combine(
                        Path.GetDirectoryName(inputFilePath),
                        $"{DateTime.UtcNow:yyyyMMdd HHmm}_{Path.GetFileName(inputFilePath)}.bak"));
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("There was an error while parsing the response, please try again.");
                Log.Error(e, "Error while parsing the response in Main method.");
                throw;
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

    /// <summary>
    /// Returns value of the node as a string
    /// </summary>
    private static string GetValFromNode(JsonNode node)
    {
        switch (node)
        {
            case null:
                return "";
            case JsonValue:
                return node.ToString();
        }

        JsonNode? ansNode = node["ans"];
        if (ansNode?["val"] != null)
        {
            return ansNode["val"]?.ToString() ?? "";
        }

        return node.ToString();
    }

    [Experimental("OPENAI001")]
    private static async Task<string> GetJsonResponse(string inputFilePath)
    {
        try
        {
            string response = PreferredApi switch
            {
                "NuDelta" => await UploadPdfToNuDelta(NuDeltaBaseUrl, Username, Password, inputFilePath),
                "OpenAI"  => await UploadPdfToChatGpt(inputFilePath, OpenAiApiKey, ResponseSchema),
                _         => await UploadPdfToChatGpt(inputFilePath, OpenAiApiKey, ResponseSchema)
            };

            Log.Information("Received JSON response from {PreferredApi}", PreferredApi);
            return response;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error communicating with {PreferredApi}: " + ex.Message);
            Log.Error(ex, "Error communicating with {PreferredApi}", PreferredApi);
            return string.Empty;
        }
    }

    private static async Task<string> UploadPdfToNuDelta(string baseUrl, string username, string password, string inputFilePath)
    {
        string documentId = await UploadDocumentAsync(baseUrl, username, password, inputFilePath);

        if (string.IsNullOrEmpty(documentId))
        {
            Console.WriteLine("Document upload failed. No Document ID received.");
            Log.Error("Document upload failed. No Document ID received from NuDelta.");
            return string.Empty;
        }

        Console.WriteLine($"File uploaded successfully. File ID: {documentId}");
        Log.Information("File uploaded successfully to NuDelta. Document ID: {DocumentId}", documentId);
        return await GetProcessedResultAsync(baseUrl, username, password, documentId);
    }

    /// <summary>
    /// Uploads a file to the NuDelta API. Returns the generated document ID on success.
    /// </summary>
    private static async Task<string> UploadDocumentAsync(string baseUrl, string username, string password, string filePath)
    {
        try
        {
            using HttpClient client = new();
            string authToken = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{username}:{password}"));
            client.DefaultRequestHeaders.Authorization = 
                new AuthenticationHeaderValue("Basic", authToken);

            using MultipartFormDataContent multipartContent = new();
            byte[] fileBytes = await File.ReadAllBytesAsync(filePath);
            ByteArrayContent fileContent = new(fileBytes);
            fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            multipartContent.Add(fileContent, "file", Path.GetFileName(filePath));
            string uploadUrl = $"{baseUrl}/documents";
            HttpResponseMessage response = await client.PostAsync(uploadUrl, multipartContent);
            response.EnsureSuccessStatusCode();

            string responseBody = await response.Content.ReadAsStringAsync();
            JObject jsonObj = JObject.Parse(responseBody);
            string docId = jsonObj["document_id"]?.Value<string>();

            Log.Information("NuDelta UploadDocumentAsync success. Document ID: {DocId}", docId);
            return docId;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"UploadDocumentAsync error: {ex.Message}");
            Log.Error(ex, "UploadDocumentAsync error");
            return null;
        }
    }

    /// <summary>
    /// Fetches the processed JSON result from the NuDelta API for the given document ID.
    /// </summary>
    private static async Task<string> GetProcessedResultAsync(string baseUrl, string username, string password, string documentId)
    {
        AsyncRetryPolicy<string>? retryPolicy = Policy<string>
            .HandleResult(resultJson =>
            {
                JsonNode? root = JsonNode.Parse(resultJson);
                switch (root)
                {
                    case JsonObject jsonObject when jsonObject.ContainsKey("state"):
                    {
                        string state = GetValFromNode(root["state"]!);
                        return !string.Equals(state, "done", StringComparison.OrdinalIgnoreCase);
                    }
                    default:
                        return true;
                }
            })
            .WaitAndRetryAsync(
                retryCount: 5,
                sleepDurationProvider: retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                onRetry: (outcome, timespan, retryAttempt, context) =>
                {
                    Log.Warning("Result not ready. Retry attempt {RetryAttempt} after {TimespanSeconds} seconds.", retryAttempt, timespan.TotalSeconds);
                });

        try
        {
            using HttpClient client = new();
            string authToken = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{username}:{password}"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authToken);

            string resultUrl = $"{baseUrl}/documents/{documentId}?compact-response=true";

            Console.WriteLine("Waiting for result");
            Log.Information("Waiting for result from NuDelta: Document ID = {DocumentId}", documentId);

            string resultJson = await retryPolicy.ExecuteAsync(async () =>
            {
                HttpResponseMessage response = await client.GetAsync(resultUrl);
                response.EnsureSuccessStatusCode();

                return await response.Content.ReadAsStringAsync();
            });

            Log.Information("Received final JSON result from NuDelta for Document ID: {DocumentId}", documentId);
            return resultJson;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"GetProcessedResultAsync error: {ex.Message}");
            Log.Error(ex, "GetProcessedResultAsync error");
            return null;
        }
    }

    [Experimental("OPENAI001")]
    private static async Task<string> UploadPdfToChatGpt(string filePath, string apiKey, string schema)
    {
        OpenAIClient client = new(apiKey);

        OpenAIFileClient? fileClient = client.GetOpenAIFileClient();
        FileUploadPurpose uploadPurpose = FileUploadPurpose.Assistants;
        string fileId;

        await using (FileStream fileStream = File.OpenRead(filePath))
        {
            ClientResult<OpenAIFile>? uploadResult = await fileClient.UploadFileAsync(fileStream, Path.GetFileName(filePath), uploadPurpose);
            fileId = uploadResult.Value.Id;
            Console.WriteLine($"File uploaded successfully. File ID: {fileId}");
            Log.Information("File uploaded to OpenAI. File ID: {FileId}", fileId);
        }
        Console.WriteLine("Waiting for result");
        Log.Information("Waiting for ChatGPT to process file {FileId}", fileId);

        AssistantClient? assistantClient = client.GetAssistantClient();

        ClientResult<Assistant>? assistant = await assistantClient.CreateAssistantAsync("gpt-4o-mini", new AssistantCreationOptions
        {
            Instructions = $"You are supposed to analyze the PDFs given to you and always respond with ONLY a valid json object (without markdown codeblocks) filled in with information from the PDF, validated by a schema. Remove quotation marks in names. If information is missing in the PDF, leave the string empty. The schema: {schema}",
            Tools = { new FileSearchToolDefinition() }
        });

        ClientResult<AssistantThread>? thread = await assistantClient.CreateThreadAsync();

        using HttpClient httpClient = new();
        httpClient.DefaultRequestHeaders.Add("Authorization", $"Bearer {apiKey}");
        httpClient.DefaultRequestHeaders.Add("OpenAI-Beta", "assistants=v2");

        PromptRequestModel requestData = new()
        { 
            Role = "user", 
            Content = "Please analyze the file",
            Attachments =
            [
                new AttachmentModel
                {
                    FileId = fileId,
                    Tools = [
                        new ToolModel
                        {
                            Type = "file_search"
                        }
                    ]
                }
            ]
        };
        
        string endpoint = $"https://api.openai.com/v1/threads/{thread.Value.Id}/messages";
        string json = JsonSerializer.Serialize(requestData);

        StringContent requestContent = new(json, Encoding.UTF8, "application/json");

        HttpResponseMessage response = await httpClient.PostAsync(endpoint, requestContent);
        response.EnsureSuccessStatusCode();

        ClientResult<ThreadRun>? run = await assistantClient.CreateRunAsync(thread.Value.Id, assistant.Value.Id);
        List<ThreadMessage> messages = await GetMessagesWithRetryAsync(thread.Value.Id, run.Value.Id, assistantClient);
        string actualAnswer = messages.FirstOrDefault(message => message.Role == MessageRole.Assistant).Content[0].Text ?? "Please try again.";

        // Cleanup
        await fileClient.DeleteFileAsync(fileId);
        await assistantClient.DeleteThreadAsync(thread.Value.Id);
        await assistantClient.DeleteAssistantAsync(assistant.Value.Id);

        Log.Information("Received final JSON response from ChatGPT");
        return actualAnswer;
    }

    [Experimental("OPENAI001")]
    private static async Task<List<ThreadMessage>> GetMessagesWithRetryAsync(string threadId, string runId, AssistantClient assistantClient)
    {
        AsyncRetryPolicy? retryPolicy = Policy
            .Handle<Exception>()
            .WaitAndRetryAsync(
                retryCount: 10,
                sleepDurationProvider: retryAttempt => TimeSpan.FromSeconds(Math.Pow(2, retryAttempt)),
                onRetry: (exception, timeSpan, retryCount, context) =>
                {
                    Log.Warning(exception, "Retry {RetryCount} after {TimespanSeconds} seconds while fetching messages", 
                        retryCount, timeSpan.TotalSeconds);
                });

        // Retrieve the run asynchronously
        Task<ClientResult<ThreadRun>>? retrievedRun = assistantClient.GetRunAsync(threadId, runId);
        await retryPolicy.ExecuteAsync(async () =>
        {
            if (retrievedRun == null || !retrievedRun.IsCompleted)
            {
                throw new Exception("Run is not completed yet.");
            }
        });

        // Retrieve messages asynchronously
        List<ThreadMessage> messages = [];
        await retryPolicy.ExecuteAsync(async () =>
        {
            List<ThreadMessage> fetchedMessages = await assistantClient.GetMessagesAsync(threadId).ToListAsync();
            messages = fetchedMessages.ToList();

            if (messages.Count < 2 || messages[0].Content.Count == 0)
            {
                throw new Exception("Messages are incomplete.");
            }
        });

        return messages;
    }
    
    private static string RunPDF2URL(string exePath, string filePath)
    {
        var startInfo = new ProcessStartInfo
        {
            FileName = exePath,
            Arguments = $"\"{filePath}\"",
            UseShellExecute = false,
            RedirectStandardOutput = true,
            CreateNoWindow = true
        };

        using var process = Process.Start(startInfo);
        string output = process?.StandardOutput.ReadToEnd() ?? string.Empty;
    
        process?.WaitForExit();
        return output;
    }
}