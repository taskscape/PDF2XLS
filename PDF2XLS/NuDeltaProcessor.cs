using System.Net.Http.Headers;
using System.Text;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json.Linq;
using Polly;
using Polly.Retry;
using Serilog;

namespace PDF2XLS;

public class NuDeltaProcessor
{
    private const string BaseUrl = "https://www.nudelta.pl/api/v1";
    private readonly string _username;
    private readonly string _password;

    public NuDeltaProcessor(IConfiguration config)
    {
        _username = config["NuDeltaCredentials:Username"] ?? string.Empty;
        _password = config["NuDeltaCredentials:Password"] ?? string.Empty;
    }

    public async Task<string?> ProcessPdfAsync(string filePath)
    {
        string? documentId = await UploadDocumentAsync(filePath);
        if (string.IsNullOrEmpty(documentId))
        {
            Log.Error("NuDelta document upload failed — no Document ID received. File: {file}", filePath);
            return null;
        }

        Log.Information("File uploaded to NuDelta. Document ID: {DocumentId}. File: {file}", documentId, filePath);
        return await GetProcessedResultAsync(documentId, filePath);
    }

    private async Task<string?> UploadDocumentAsync(string filePath)
    {
        try
        {
            using HttpClient client = new();
            client.Timeout = TimeSpan.FromMinutes(5);
            string authToken = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{_username}:{_password}"));
            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Basic", authToken);

            using MultipartFormDataContent multipartContent = new();
            byte[] fileBytes = await File.ReadAllBytesAsync(filePath);
            ByteArrayContent fileContent = new(fileBytes);
            fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
            multipartContent.Add(fileContent, "file", Path.GetFileName(filePath));

            HttpResponseMessage response = await client.PostAsync($"{BaseUrl}/documents", multipartContent);
            response.EnsureSuccessStatusCode();

            string responseBody = await response.Content.ReadAsStringAsync();
            JObject jsonObj = JObject.Parse(responseBody);
            string? docId = jsonObj["document_id"]?.Value<string>();

            Log.Information("NuDelta upload success. Document ID: {DocId}. File: {file}", docId, filePath);
            return docId;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "NuDelta UploadDocumentAsync error. File: {file}", filePath);
            return null;
        }
    }

    private async Task<string?> GetProcessedResultAsync(string documentId, string filePath)
    {
        AsyncRetryPolicy<string> retryPolicy = Policy<string>
            .HandleResult(resultJson =>
            {
                try
                {
                    JsonNode? root = JsonNode.Parse(resultJson);
                    if (root is not JsonObject jsonObject || !jsonObject.ContainsKey("state"))
                        return true;

                    string? state = ExtractStateValue(root["state"]!);
                    return !string.Equals(state, "done", StringComparison.OrdinalIgnoreCase);
                }
                catch
                {
                    return true;
                }
            })
            .WaitAndRetryAsync(
                retryCount: 5,
                sleepDurationProvider: attempt => TimeSpan.FromSeconds(Math.Pow(2, attempt)),
                onRetry: (_, ts, attempt, _) =>
                    Log.Warning("NuDelta result not ready. Retry {Attempt} after {Delay}s. File: {file}",
                        attempt, ts.TotalSeconds, filePath));

        try
        {
            using HttpClient client = new();
            client.Timeout = TimeSpan.FromMinutes(5);
            string authToken = Convert.ToBase64String(Encoding.ASCII.GetBytes($"{_username}:{_password}"));
            client.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Basic", authToken);

            string resultUrl = $"{BaseUrl}/documents/{documentId}?compact-response=true";
            Log.Information("Waiting for NuDelta result. Document ID: {DocumentId}. File: {file}", documentId, filePath);

            string? resultJson = await retryPolicy.ExecuteAsync(async () =>
            {
                HttpResponseMessage response = await client.GetAsync(resultUrl);
                response.EnsureSuccessStatusCode();
                return await response.Content.ReadAsStringAsync();
            });

            Log.Information("Received NuDelta result for Document ID: {DocumentId}. File: {file}", documentId, filePath);
            return resultJson;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "NuDelta GetProcessedResultAsync error. File: {file}", filePath);
            return null;
        }
    }

    /// <summary>
    /// Extracts the state string value from a NuDelta node, handling the wrapped
    /// {"ans": {"val": "..."}} format that NuDelta uses for some fields.
    /// </summary>
    private static string? ExtractStateValue(JsonNode node)
    {
        if (node is JsonValue)
            return node.ToString();

        JsonNode? ansNode = node["ans"];
        if (ansNode?["val"] != null)
            return ansNode["val"]?.ToString();

        return node.ToString();
    }
}
