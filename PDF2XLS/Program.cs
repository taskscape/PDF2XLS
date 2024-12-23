using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;
using ClosedXML.Excel;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using Polly;
using Polly.Retry;

namespace PDF2XLS;

class Program
{
    static async Task Main(string[] args)
    {
        IConfiguration config = new ConfigurationBuilder()
            .SetBasePath(Environment.CurrentDirectory)
            .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
            .Build();
        
        const string baseUrl = "https://www.nudelta.pl/api/v1";
        string username = config["NuDeltaCredentials:Username"] ?? "";
        string password = config["NuDeltaCredentials:Password"] ?? "";

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
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
            return;
        }

        if (!string.Equals(Path.GetExtension(inputFilePath), ".pdf", StringComparison.OrdinalIgnoreCase))
        {
            Console.WriteLine($"File {inputFilePath} is not a PDF file");
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
            return;
        }
            
        string outputDir;
        if (args.Length >= 2)
        {
            outputDir = args[1];
        }
        else
        {
            outputDir = Environment.CurrentDirectory;
        }

        string documentId = await UploadDocumentAsync(baseUrl, username, password, inputFilePath);

        if (string.IsNullOrEmpty(documentId))
        {
            Console.WriteLine("Document upload failed. No Document ID received.");
            return;
        }
            
        Console.WriteLine($"Successfully uploaded. Document ID: {documentId}");
        string processedJson = await GetProcessedResultAsync(baseUrl, username, password, documentId);
        JsonNode root = JsonNode.Parse(processedJson);
        JsonNode? dataNode = root?["data"];

        // Extract top-level fields
        string invNumber = GetValFromNode(dataNode?["invn"]);
        string issueDate = GetValFromNode(dataNode?["issue"]);
        string saleDate = GetValFromNode(dataNode?["sale"]);
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
        ws.Cell("B4").Value = issueDate;
        ws.Cell("A5").Value = "Data sprzedaży:";
        ws.Cell("B5").Value = saleDate;
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

            return docId;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"UploadDocumentAsync error: {ex.Message}");
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
                    //Console.WriteLine($"Retry {retryAttempt}. Waiting {timespan.TotalSeconds} seconds before next attempt...");
                });

        try
        {
            using HttpClient client = new();
            string authToken = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{username}:{password}"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic", authToken);

            string resultUrl = $"{baseUrl}/documents/{documentId}";
            
            Console.WriteLine("Waiting for result");
            string resultJson = await retryPolicy.ExecuteAsync(async () =>
            {
                HttpResponseMessage response = await client.GetAsync(resultUrl);
                response.EnsureSuccessStatusCode();
                
                return await response.Content.ReadAsStringAsync();
            });

            return resultJson;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"GetProcessedResultAsync error: {ex.Message}");
            return null;
        }
    }
}