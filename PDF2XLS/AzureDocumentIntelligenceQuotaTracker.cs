using System.Text.Json;
using Microsoft.Extensions.Configuration;
using Serilog;

namespace PDF2XLS;

public sealed class AzureDocumentIntelligenceQuotaTracker
{
    private readonly object _sync = new();
    private readonly string _counterFilePath;
    private readonly JsonSerializerOptions _jsonOptions = new() { WriteIndented = true };

    public int MonthlyPageLimit { get; }
    public bool IsEnabled => MonthlyPageLimit > 0;

    public AzureDocumentIntelligenceQuotaTracker(IConfiguration config, string baseDirectory)
    {
        MonthlyPageLimit = int.TryParse(config["AzureDocumentIntelligence:MonthlyPageLimit"], out int limit)
            ? limit
            : 0;

        string configuredCounterPath = config["AzureDocumentIntelligence:MonthlyQuotaCounterPath"] ?? string.Empty;
        _counterFilePath = ResolveCounterFilePath(configuredCounterPath, baseDirectory);
    }

    public bool IsQuotaLimitReached()
    {
        if (!IsEnabled)
            return false;

        QuotaCounterState state = ReadStateForCurrentMonth();
        return state.PagesSubmitted >= MonthlyPageLimit;
    }

    public void LogQuotaLimitReached()
    {
        if (!IsEnabled)
            return;

        QuotaCounterState state = ReadStateForCurrentMonth();
        Log.Information(
            "Azure Document Intelligence monthly page quota limit has been achieved. Month: {Month}, PagesSubmitted: {PagesSubmitted}, Limit: {Limit}. No further Azure processing will be attempted.",
            state.Month,
            state.PagesSubmitted,
            MonthlyPageLimit);
    }

    public void EnsureCanSubmit(string filePath, int documentPageCount)
    {
        if (!IsEnabled)
            return;

        QuotaCounterState state = ReadStateForCurrentMonth();
        if (state.PagesSubmitted >= MonthlyPageLimit)
        {
            LogQuotaLimitReached();
            throw new AzureDocumentIntelligenceQuotaReachedException(
                "Azure Document Intelligence monthly page quota limit has been achieved.");
        }

        int projectedPages = state.PagesSubmitted + documentPageCount;
        if (projectedPages > MonthlyPageLimit)
        {
            Log.Information(
                "Azure Document Intelligence monthly page quota would be exceeded. Month: {Month}, PagesSubmitted: {PagesSubmitted}, DocumentPages: {DocumentPages}, Limit: {Limit}, File: {File}. No further Azure processing will be attempted.",
                state.Month,
                state.PagesSubmitted,
                documentPageCount,
                MonthlyPageLimit,
                filePath);
            throw new AzureDocumentIntelligenceQuotaReachedException(
                "Azure Document Intelligence monthly page quota would be exceeded by this document.");
        }
    }

    public void RecordSuccessfulSubmission(string filePath, int documentPageCount)
    {
        if (!IsEnabled)
            return;

        lock (_sync)
        {
            QuotaCounterState state = ReadStateForCurrentMonth();
            state.PageLimit = MonthlyPageLimit;
            state.PagesSubmitted += documentPageCount;
            state.LastSubmittedFile = filePath;
            state.LastSubmittedPageCount = documentPageCount;
            state.UpdatedUtc = DateTimeOffset.UtcNow;

            WriteState(state);

            Log.Information(
                "Azure Document Intelligence quota counter updated. Month: {Month}, AddedPages: {AddedPages}, PagesSubmitted: {PagesSubmitted}, Limit: {Limit}, CounterFile: {CounterFile}, File: {File}",
                state.Month,
                documentPageCount,
                state.PagesSubmitted,
                MonthlyPageLimit,
                _counterFilePath,
                filePath);

            if (state.PagesSubmitted >= MonthlyPageLimit)
            {
                Log.Information(
                    "Azure Document Intelligence monthly page quota limit has been achieved. Month: {Month}, PagesSubmitted: {PagesSubmitted}, Limit: {Limit}. No further Azure processing will be attempted.",
                    state.Month,
                    state.PagesSubmitted,
                    MonthlyPageLimit);
            }
        }
    }

    private QuotaCounterState ReadStateForCurrentMonth()
    {
        lock (_sync)
        {
            string currentMonth = GetCurrentMonth();
            if (!File.Exists(_counterFilePath))
            {
                return NewState(currentMonth);
            }

            try
            {
                string json = File.ReadAllText(_counterFilePath);
                QuotaCounterState? state = JsonSerializer.Deserialize<QuotaCounterState>(json, _jsonOptions);
                if (state == null || string.IsNullOrWhiteSpace(state.Month))
                {
                    throw new JsonException("Counter file did not contain a valid quota state.");
                }

                if (!string.Equals(state.Month, currentMonth, StringComparison.Ordinal))
                {
                    return NewState(currentMonth);
                }

                state.PageLimit = MonthlyPageLimit;
                return state;
            }
            catch (Exception ex)
            {
                throw new AzureDocumentIntelligenceQuotaAccountingException(
                    $"Could not read Azure Document Intelligence quota counter file: {_counterFilePath}",
                    ex);
            }
        }
    }

    private void WriteState(QuotaCounterState state)
    {
        try
        {
            string? directory = Path.GetDirectoryName(_counterFilePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string tempPath = $"{_counterFilePath}.{Guid.NewGuid():N}.tmp";
            File.WriteAllText(tempPath, JsonSerializer.Serialize(state, _jsonOptions));
            File.Move(tempPath, _counterFilePath, true);
        }
        catch (Exception ex)
        {
            throw new AzureDocumentIntelligenceQuotaAccountingException(
                $"Could not write Azure Document Intelligence quota counter file: {_counterFilePath}",
                ex);
        }
    }

    private QuotaCounterState NewState(string month) => new()
    {
        Month = month,
        PageLimit = MonthlyPageLimit,
        PagesSubmitted = 0,
        UpdatedUtc = DateTimeOffset.UtcNow
    };

    private static string ResolveCounterFilePath(string configuredCounterPath, string baseDirectory)
    {
        if (string.IsNullOrWhiteSpace(configuredCounterPath))
            return Path.Combine(baseDirectory, "azure-document-intelligence-quota.json");

        return Path.IsPathRooted(configuredCounterPath)
            ? configuredCounterPath
            : Path.Combine(baseDirectory, configuredCounterPath);
    }

    private static string GetCurrentMonth() => DateTimeOffset.UtcNow.ToString("yyyy-MM");

    private sealed class QuotaCounterState
    {
        public string Month { get; set; } = string.Empty;
        public int PageLimit { get; set; }
        public int PagesSubmitted { get; set; }
        public string? LastSubmittedFile { get; set; }
        public int LastSubmittedPageCount { get; set; }
        public DateTimeOffset UpdatedUtc { get; set; }
    }
}

public abstract class AzureDocumentIntelligenceQuotaException : Exception
{
    protected AzureDocumentIntelligenceQuotaException(string message)
        : base(message)
    {
    }

    protected AzureDocumentIntelligenceQuotaException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}

public sealed class AzureDocumentIntelligenceQuotaReachedException : AzureDocumentIntelligenceQuotaException
{
    public AzureDocumentIntelligenceQuotaReachedException(string message)
        : base(message)
    {
    }
}

public sealed class AzureDocumentIntelligenceQuotaAccountingException : AzureDocumentIntelligenceQuotaException
{
    public AzureDocumentIntelligenceQuotaAccountingException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
