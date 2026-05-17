using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Polly;
using Polly.Retry;
using Serilog;

namespace PDF2XLS;

/// <summary>
/// Processes PDF invoices using the OpenAI Responses API with native PDF reading (GPT-4o vision pipeline).
/// Replaces the deprecated Assistants API approach with a single, stateless API call.
/// </summary>
public class OpenAIResponsesProcessor
{
    private const string ResponsesEndpoint = "https://api.openai.com/v1/responses";

    private readonly string _apiKey;
    private readonly string _model;
    private readonly string _prompt;

    public OpenAIResponsesProcessor(IConfiguration config, string responseSchema)
    {
        _apiKey = config["OpenAI:OpenAI_APIKey"] ?? string.Empty;
        _model = config["OpenAI:OpenAI_Model"] ?? string.Empty;
        _prompt = config["OpenAI:Prompt"]?.Replace("{schema}", responseSchema) ?? string.Empty;
    }

    /// <summary>
    /// Sends the PDF file directly to the OpenAI Responses API as a base64-encoded input_file.
    /// Returns the raw JSON string produced by the model.
    /// </summary>
    public async Task<string?> ProcessPdfAsync(string filePath)
    {
        byte[] pdfBytes = await File.ReadAllBytesAsync(filePath);
        string base64Pdf = Convert.ToBase64String(pdfBytes);
        string fileName = Path.GetFileName(filePath).ToLowerInvariant();

        Log.Information("Sending PDF to OpenAI Responses API. File: {file}", filePath);

        string? result = await CallResponsesApiAsync(base64Pdf, fileName, filePath);

        Log.Information("Received result from OpenAI Responses API. File: {file}", filePath);
        return result;
    }

    private async Task<string?> CallResponsesApiAsync(string base64Pdf, string fileName, string filePath)
    {
        // Retry on transient server errors and rate-limit responses.
        AsyncRetryPolicy<HttpResponseMessage> retryPolicy = Policy
            .HandleResult<HttpResponseMessage>(r =>
                (int)r.StatusCode >= 500 || r.StatusCode == System.Net.HttpStatusCode.TooManyRequests)
            .WaitAndRetryAsync(
                retryCount: 3,
                sleepDurationProvider: attempt => TimeSpan.FromSeconds(Math.Pow(2, attempt)),
                onRetry: (result, ts, attempt, _) =>
                    Log.Warning("OpenAI Responses API retry {Attempt} after {Delay}s (HTTP {Status}). File: {file}",
                        attempt, ts.TotalSeconds, (int)result.Result.StatusCode, filePath));

        object requestBody = new
        {
            model = _model,
            input = new[]
            {
                new
                {
                    role = "user",
                    content = new object[]
                    {
                        new { type = "input_text", text = _prompt },
                        new { type = "input_file", filename = fileName, file_data = base64Pdf }
                    }
                }
            },
            text = new
            {
                format = new { type = "json_object" }
            }
        };

        string requestJson = JsonSerializer.Serialize(requestBody);

        using HttpClient http = new();
        http.DefaultRequestHeaders.Authorization =
            new AuthenticationHeaderValue("Bearer", _apiKey);
        http.Timeout = TimeSpan.FromMinutes(5);

        using StringContent httpContent = new(requestJson, Encoding.UTF8, "application/json");

        HttpResponseMessage response = await retryPolicy.ExecuteAsync(
            () => http.PostAsync(ResponsesEndpoint, httpContent));

        response.EnsureSuccessStatusCode();

        string responseBody = await response.Content.ReadAsStringAsync();
        return ExtractTextFromResponse(responseBody, filePath);
    }

    /// <summary>
    /// Extracts the model's text output from the Responses API envelope:
    /// response.output[0].content[0].text
    /// </summary>
    private static string? ExtractTextFromResponse(string responseBody, string filePath)
    {
        try
        {
            using JsonDocument doc = JsonDocument.Parse(responseBody);
            JsonElement root = doc.RootElement;

            if (!root.TryGetProperty("output", out JsonElement output) ||
                output.ValueKind != JsonValueKind.Array ||
                output.GetArrayLength() == 0)
            {
                Log.Error("OpenAI Responses API: 'output' array missing or empty. File: {file}", filePath);
                return null;
            }

            JsonElement firstOutput = output[0];
            if (!firstOutput.TryGetProperty("content", out JsonElement content) ||
                content.ValueKind != JsonValueKind.Array ||
                content.GetArrayLength() == 0)
            {
                Log.Error("OpenAI Responses API: 'content' array missing or empty in output[0]. File: {file}", filePath);
                return null;
            }

            JsonElement textPart = content[0];
            if (textPart.TryGetProperty("text", out JsonElement text))
                return text.GetString();

            Log.Error("OpenAI Responses API: 'text' field not found in output[0].content[0]. File: {file}", filePath);
            return null;
        }
        catch (JsonException ex)
        {
            Log.Error(ex, "Failed to parse OpenAI Responses API response. File: {file}", filePath);
            return null;
        }
    }
}
