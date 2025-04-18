using System.Net.Http.Headers;
using System.Text.Json;
using Microsoft.AspNetCore.WebUtilities;
using Microsoft.Extensions.Configuration;
using Polly;
using Polly.Retry;
using Serilog;

namespace PDF2XLS;

public class LLMWhisperer
{
    private IConfiguration Config { get; }
    private readonly Uri _baseUrl;
    private readonly string _apiKey;
    private readonly Guid _runId;
    private readonly string _runTime;
    
    public LLMWhisperer(IConfiguration config, Guid runId, string runTime)
    {
        Config = config;
        _baseUrl = new Uri(Config["Whisperer:BaseUrl"] ?? string.Empty);
        _apiKey = Config["Whisperer:ApiKey"] ?? string.Empty;
        _runId = runId;
        _runTime = runTime;
    }
    
    public async Task ProcessPdfWorkflow(string pdfFilePath)
    {
        string whisperHash = await UploadPdfToWhisperer(pdfFilePath);
        if (string.IsNullOrEmpty(whisperHash))
        {
            Log.Error("PDF upload to Whisperer failed. Aborting workflow. File: {file}", pdfFilePath);
            return;
        }
        
        Log.Information("PDF upload to Whisperer successful. File: {file}", pdfFilePath);
        
        AsyncRetryPolicy<string> retryPolicy = Policy<string>
            .HandleResult(status => status != "processed")
            .WaitAndRetryAsync(
                retryCount: 10,
                sleepDurationProvider: retryAttempt => TimeSpan.FromSeconds(5),
                onRetry: (outcome, timespan, retryCount, context) =>
                {
                    Log.Information("Attempt {retryCount}: status is '{status}'. Waiting {retryInterval} seconds before next check. File: {file}", retryCount, outcome.Result, timespan.TotalSeconds, pdfFilePath);
                });


        string finalStatus = await retryPolicy.ExecuteAsync(async () => await GetWhispererProcessingStatus(whisperHash));
        
        Log.Information("Final status: {status}. File: {file}", finalStatus, pdfFilePath);
        
        if (finalStatus == "processed")
        {
            string pdfText = await GetPdfTextFromWhisperer(whisperHash);

            if (!string.IsNullOrEmpty(pdfText))
            {
                string txtFilePath = Path.Combine(
                    Path.GetDirectoryName(pdfFilePath),
                    $"{_runTime}_{_runId}_{Path.GetFileName(pdfFilePath)}.txt");
                await File.WriteAllTextAsync(txtFilePath, pdfText);
                Log.Information("PDF text extracted and saved to {filePath}. File {file}", txtFilePath, pdfFilePath);
            }
            else
            {
                Log.Error("Failed to extract PDF text. File: {file}", pdfFilePath);
            }
        }
        else
        {
            Log.Error("PDF processing did not complete successfully within the allowed retries. File: {file}", pdfFilePath);
        }
    }

    private async Task<string> GetPdfTextFromWhisperer(string whisperHash)
    {
        Uri endpoint = new("api/v2/whisper-retrieve", UriKind.Relative);
        Uri requestUrl = new(_baseUrl, endpoint);

        using HttpClient client = new();
        client.DefaultRequestHeaders.Add("unstract-key", _apiKey);
        Dictionary<string, string?> parameters = new()
        {
            ["whisper_hash"] = whisperHash,
            ["text_only"] = "true"
        };
        Uri fullRequest = new(QueryHelpers.AddQueryString(requestUrl.ToString(), parameters));
        HttpResponseMessage response = await client.GetAsync(fullRequest);

        if (response.IsSuccessStatusCode)
        {
            return await response.Content.ReadAsStringAsync();
        }
        
        return string.Empty;
    }

    private async Task<string> GetWhispererProcessingStatus(string whisperHash)
    {
        Uri endpoint = new("api/v2/whisper-status", UriKind.Relative);
        Uri requestUrl = new(_baseUrl, endpoint);

        using HttpClient client = new();
        client.DefaultRequestHeaders.Add("unstract-key", _apiKey);
        Uri fullRequest = new(QueryHelpers.AddQueryString(requestUrl.ToString(), "whisper_hash", whisperHash));
        HttpResponseMessage response = await client.GetAsync(fullRequest);

        if (!response.IsSuccessStatusCode) return string.Empty;
        string jsonResponse = await response.Content.ReadAsStringAsync();

        try
        {
            using JsonDocument document = JsonDocument.Parse(jsonResponse);
            JsonElement root = document.RootElement;
            if (root.TryGetProperty("status", out JsonElement statusElement))
            {
                return statusElement.ToString();
            }
        }
        catch (JsonException)
        {
            return string.Empty;
        }

        return string.Empty;
    }

    private async Task<string> UploadPdfToWhisperer(string pdfFilePath)
    {
        Uri endpoint = new("api/v2/whisper", UriKind.Relative);
        Uri requestUrl = new(_baseUrl, endpoint);

        byte[] pdfBytes = await File.ReadAllBytesAsync(pdfFilePath);

        using HttpClient client = new();
        client.DefaultRequestHeaders.Add("unstract-key", _apiKey);

        using ByteArrayContent content = new(pdfBytes);
        content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");

        HttpResponseMessage response = await client.PostAsync(requestUrl, content);

        if (!response.IsSuccessStatusCode) return string.Empty;
        string jsonResponse = await response.Content.ReadAsStringAsync();

        try
        {
            using JsonDocument document = JsonDocument.Parse(jsonResponse);
            JsonElement root = document.RootElement;
            if (root.TryGetProperty("whisper_hash", out JsonElement whisperHashElement))
            {
                return whisperHashElement.ToString();
            }
        }
        catch (JsonException)
        {
            return string.Empty;
        }

        return string.Empty;
    }
}