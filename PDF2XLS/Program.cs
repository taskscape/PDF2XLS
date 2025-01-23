using System.ClientModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Net.Http.Headers;
using System.Reflection;
using System.Text;
using System.Text.Json.Nodes;
using ClosedXML.Excel;
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

    [Experimental("OPENAI001")]
    static async Task Main(string[] args)
    {
        string? exePath = Process.GetCurrentProcess().MainModule?.FileName;
        string realExeDirectory = Path.GetDirectoryName(exePath);
        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Debug()
            .WriteTo.File(
                path: $"{realExeDirectory}/logs/log-.txt",
                rollingInterval: RollingInterval.Day,
                retainedFileCountLimit: 7,
                outputTemplate: "[{Timestamp:yyyy-MM-dd HH:mm:ss} {Level:u3}] {Message:lj}{NewLine}{Exception}"
            )
            .CreateLogger();

        Log.Information("Starting PDF2XLS application...");

        try
        {
            
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
            
            Username = config["NuDeltaCredentials:Username"] ?? "";
            Password = config["NuDeltaCredentials:Password"] ?? "";
            PreferredApi = config["PreferredAPI"] ?? "";
            OpenAiApiKey = config["OpenAI_APIKey"] ?? "";
            ResponseSchema = fileContents;
            DeleteAfter = bool.Parse(config["DeleteFileAfterProcessing"]);
            if (args.Length < 1)
            {
                Console.WriteLine("Usage: PDF2XLS <input file path> [output directory]");
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

            string outputDir = args.Length >= 2 ? args[1] : Environment.CurrentDirectory;

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
                JsonArray rows = tablesNode?["rows"]?.AsArray() ?? [];
                JsonArray totals = tablesNode?["total"]?.AsArray() ?? [];

                string fileNameNoExt = Path.GetFileNameWithoutExtension(inputFilePath);
                string outputPath = Path.Combine(outputDir, fileNameNoExt + ".xlsx");

                using XLWorkbook wb = new();
                IXLWorksheet ws = wb.Worksheets.Add("Faktura");

                // Styles
                IXLStyle? headerStyle = wb.Style;
                headerStyle.Font.Bold = true;
                headerStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                headerStyle.Fill.BackgroundColor = XLColor.LightGray;
                headerStyle.Border.OutsideBorder = XLBorderStyleValues.Thin;

                // Invoice Header
                ws.Cell("A1").Value = "Faktura";
                ws.Range("A1:B1").Merge().Style.Font.Bold = true;
                ws.Row(1).Height = 20;

                ws.Cell("A3").Value = "Numer faktury:";
                ws.Cell("B3").Value = invNumber;
                ws.Cell("A4").Value = "Data wystawienia:";
                ws.Cell("B4").Value = issueDateString;
                ws.Cell("A5").Value = "Data sprzedaży:";
                ws.Cell("B5").Value = saleDateString;
                ws.Cell("A6").Value = "Termin zapłaty:";
                ws.Cell("B6").Value = maturity;
                ws.Cell("A7").Value = "Forma zapłaty:";
                ws.Cell("B7").Value = paymentMethod;
                ws.Cell("A8").Value = "Waluta:";
                ws.Cell("B8").Value = currency;

                // Seller & Buyer
                ws.Cell("D3").Value = "Sprzedawca";
                ws.Cell("D3").Style.Font.Bold = true;
                ws.Cell("D4").Value = $"NIP: {sellerNip}";
                ws.Cell("D5").Value = sellerName;
                ws.Cell("D6").Value = sellerStreet;
                ws.Cell("D7").Value = $"{sellerZip} {sellerCity}";

                ws.Cell("F3").Value = "Kupujący";
                ws.Cell("F3").Style.Font.Bold = true;
                ws.Cell("F4").Value = $"NIP: {buyerNip}";
                ws.Cell("F5").Value = buyerName;
                ws.Cell("F6").Value = buyerStreet;
                ws.Cell("F7").Value = $"{buyerZip} {buyerCity}";

                // Line Items Table
                int startRow = 10;
                ws.Cell(startRow, 1).Value = "lp";
                ws.Cell(startRow, 2).Value = "Nazwa towaru lub usługi";
                ws.Cell(startRow, 3).Value = "Ilość";
                ws.Cell(startRow, 4).Value = "Jm";
                ws.Cell(startRow, 5).Value = "Cena netto";
                ws.Cell(startRow, 6).Value = "Stawka VAT %";
                ws.Cell(startRow, 7).Value = "Wartość Netto";
                ws.Cell(startRow, 8).Value = "Wartość VAT";
                ws.Cell(startRow, 9).Value = "Wartość Brutto";

                ws.Range(startRow, 1, startRow, 9).Style = headerStyle;

                int currentRow = startRow + 1;
                int lastItemRow = currentRow;
                foreach (JsonNode? r in rows)
                {
                    lastItemRow = currentRow;
                    string noVal = GetValFromNode(r?["no"]);
                    string nameVal = GetValFromNode(r?["name"]);
                    string amountVal = GetValFromNode(r?["amount"]);
                    string unitVal = GetValFromNode(r?["unit"]);
                    string priceNettoVal = GetValFromNode(r?["priceNetto"]);
                    string vatVal = GetValFromNode(r?["vat"]);
                    string valNettoVal = GetValFromNode(r?["valNetto"]);
                    string valVatVal = GetValFromNode(r?["valVat"]);
                    string valBruttoVal = GetValFromNode(r?["valBrutto"]);

                    ws.Cell(currentRow, 1).Value = noVal;
                    ws.Cell(currentRow, 2).Value = nameVal;
                    ws.Cell(currentRow, 3).Value = amountVal;
                    ws.Cell(currentRow, 4).Value = unitVal;
                    ws.Cell(currentRow, 5).Value = priceNettoVal;
                    ws.Cell(currentRow, 6).Value = vatVal;
                    ws.Cell(currentRow, 7).Value = valNettoVal;
                    ws.Cell(currentRow, 8).Value = valVatVal;
                    ws.Cell(currentRow, 9).Value = valBruttoVal;

                    currentRow++;
                }

                // Totals Section
                currentRow += 1;
                JsonNode? totalNode = totals.Count > 0 ? totals[0] : null;
                string totalNetto = GetValFromNode(totalNode?["valNetto"]);
                string totalVat = GetValFromNode(totalNode?["valVat"]);
                string totalBrutto = GetValFromNode(totalNode?["valBrutto"]);

                ws.Cell(currentRow, 8).Value = "Netto Razem:";
                ws.Cell(currentRow, 9).Value = totalNetto;
                ws.Cell(currentRow, 1).Value = "IBAN:";
                ws.Cell(currentRow, 2).Value = iban;
                currentRow++;
                ws.Cell(currentRow, 8).Value = "VAT Razem:";
                ws.Cell(currentRow, 9).Value = totalVat;
                ws.Cell(currentRow, 1).Value = "Zaliczka otrzymana:";
                ws.Cell(currentRow, 2).Value = paidAmount;
                currentRow++;
                ws.Cell(currentRow, 8).Value = "Brutto Razem:";
                ws.Cell(currentRow, 9).Value = totalBrutto;
                ws.Cell(currentRow, 1).Value = "Do zapłaty:";
                ws.Cell(currentRow, 2).Value = leftToPay;

                // Formatting
                ws.Columns().AdjustToContents();
                ws.Range("A3:A9").Style.Font.Bold = true;
                ws.Range("A3:A9").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                ws.Range("D3").Style.Font.Bold = true;
                ws.Range("F3").Style.Font.Bold = true;
                ws.Range(startRow, 1, lastItemRow, 9).Style.Border.OutsideBorder = XLBorderStyleValues.Thin;
                ws.Range(startRow, 1, lastItemRow, 9).Style.Border.InsideBorder = XLBorderStyleValues.Hair;

                wb.SaveAs(outputPath);

                Console.WriteLine($"Excel file saved to: {outputPath}");
                Log.Information("Excel file saved successfully to {OutputPath}", outputPath);

                GSheets sheets = new GSheets(config);
                SheetsService sheetsService = sheets.CreateSheetsService();
                sheets.AddRow(sheetsService, issueDateString, refNumber, sellerName, invNumber, totalAmount, leftToPay, currency);
                
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
}