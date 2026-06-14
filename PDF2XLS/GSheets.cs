using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Configuration;
using Polly;
using Polly.Retry;
using Serilog;
using System.Globalization;

namespace PDF2XLS;

public class GSheets
{
    // Retry twice on transient Google API failures (5xx, 429) and network errors
    // ("Google Sheets cannot be communicated"). Configuration/permission errors
    // (4xx) are NOT retried here and are surfaced as GoogleSheetsConfigurationException.
    // OperationCanceledException (our 5-min timeout) is intentionally excluded.
    private static readonly AsyncRetryPolicy SheetsRetryPolicy = Policy
        .Handle<Google.GoogleApiException>(ex => (int)ex.HttpStatusCode >= 500 || (int)ex.HttpStatusCode == 429)
        .Or<HttpRequestException>()
        .Or<IOException>()
        .WaitAndRetryAsync(
            retryCount: 2,
            sleepDurationProvider: attempt => TimeSpan.FromSeconds(Math.Pow(2, attempt)),
            onRetry: (ex, ts, attempt, _) =>
                Log.Warning(ex, "Google Sheets could not be communicated (transient error). Retry {Attempt} of 2 after {Delay:F1}s. Reason: {Reason}",
                    attempt, ts.TotalSeconds, ex.Message));

    private IConfiguration Config { get; }
    private readonly string _serviceAccountFile;
    private readonly string _spreadsheetId;
    private readonly string _expectedSpreadsheetName;
    private readonly string _sheetName;
    private readonly string _applicationName;
    private readonly string _inputFilePath;

    public GSheets(IConfiguration config, string inputFilePath)
    {
        Config = config;
        _serviceAccountFile = Config["GoogleSheets:ServiceAccountFile"] ?? string.Empty;
        _spreadsheetId = Config["GoogleSheets:SpreadsheetId"] ?? string.Empty;
        _expectedSpreadsheetName = Config["GoogleSheets:ExpectedSpreadsheetName"] ?? string.Empty;
        _sheetName = Config["GoogleSheets:SheetName"] ?? string.Empty;
        _applicationName = Config["GoogleSheets:ApplicationName"] ?? string.Empty;
        _inputFilePath = inputFilePath;
    }

    private static int? GetColumnIndex(string columnLetter)
    {
        if (string.IsNullOrEmpty(columnLetter))
            return null;

        int index = 0;
        foreach (char c in columnLetter.ToUpperInvariant())
        {
            if (c is < 'A' or > 'Z')
                return null;
            index = index * 26 + (c - 'A' + 1);
        }
        return index - 1;
    }

    private async Task<int?> GetSheetIdAsync(SheetsService sheetsService, CancellationToken cancellationToken)
    {
        Spreadsheet? spreadsheet = await sheetsService.Spreadsheets.Get(_spreadsheetId).ExecuteAsync(cancellationToken);
        Sheet? sheet = FindConfiguredSheet(spreadsheet);
        return sheet?.Properties.SheetId;
    }

    private Sheet? FindConfiguredSheet(Spreadsheet spreadsheet) =>
        spreadsheet.Sheets.FirstOrDefault(s =>
            string.Equals(s.Properties.Title, _sheetName, StringComparison.Ordinal));

    public async Task<bool> VerifySpreadsheetName(SheetsService sheetsService)
    {
        if (string.IsNullOrEmpty(_expectedSpreadsheetName))
            return true;

        try
        {
            using CancellationTokenSource cts = new(TimeSpan.FromMinutes(5));

            Spreadsheet? spreadsheet = await SheetsRetryPolicy.ExecuteAsync(
                ct => sheetsService.Spreadsheets.Get(_spreadsheetId).ExecuteAsync(ct),
                cts.Token);

            string actualName = spreadsheet.Properties.Title;

            if (!string.Equals(actualName, _expectedSpreadsheetName, StringComparison.OrdinalIgnoreCase))
            {
                Log.Error(
                    "Spreadsheet name mismatch. Expected: '{Expected}', Actual: '{Actual}'. SpreadsheetId: {Id}",
                    _expectedSpreadsheetName, actualName, _spreadsheetId);
                return false;
            }

            Log.Information("Spreadsheet name verified: '{Name}'", actualName);
            return true;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to verify spreadsheet name after retries. SpreadsheetId: {Id}", _spreadsheetId);
            return false;
        }
    }

    public async Task<bool> VerifySheetName(SheetsService sheetsService)
    {
        try
        {
            using CancellationTokenSource cts = new(TimeSpan.FromMinutes(5));

            Spreadsheet? spreadsheet = await SheetsRetryPolicy.ExecuteAsync(
                ct => sheetsService.Spreadsheets.Get(_spreadsheetId).ExecuteAsync(ct),
                cts.Token);

            Sheet? exactSheet = FindConfiguredSheet(spreadsheet);
            if (exactSheet != null)
            {
                Log.Information("Spreadsheet sheet verified: '{SheetName}'", _sheetName);
                return true;
            }

            Sheet? caseInsensitiveMatch = spreadsheet.Sheets.FirstOrDefault(s =>
                string.Equals(s.Properties.Title, _sheetName, StringComparison.OrdinalIgnoreCase));

            if (caseInsensitiveMatch != null)
            {
                Log.Error(
                    "Spreadsheet sheet name mismatch. Expected exact tab name: '{Expected}', found different casing: '{Actual}'. SpreadsheetId: {Id}",
                    _sheetName,
                    caseInsensitiveMatch.Properties.Title,
                    _spreadsheetId);
                return false;
            }

            string availableSheets = string.Join(", ", spreadsheet.Sheets.Select(s => $"'{s.Properties.Title}'"));
            Log.Error(
                "Spreadsheet sheet not found. Expected exact tab name: '{Expected}'. Available tabs: {Available}. SpreadsheetId: {Id}",
                _sheetName,
                availableSheets,
                _spreadsheetId);
            return false;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Failed to verify spreadsheet sheet after retries. SpreadsheetId: {Id}, SheetName: {SheetName}",
                _spreadsheetId, _sheetName);
            return false;
        }
    }

    public SheetsService CreateSheetsService()
    {
        GoogleCredential? credential = GoogleCredential
            .FromFile(_serviceAccountFile)
            .CreateScoped(SheetsService.Scope.Spreadsheets, SheetsService.Scope.Drive);

        return new SheetsService(new Google.Apis.Services.BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = _applicationName
        });
    }
    
    private (ExtendedValue Value, CellFormat? Format) GetExtendedValueAndFormat(string? value)
    {
        if (TryParseFlexibleDecimal(value, out decimal decValue))
        {
            decValue = decimal.Round(decValue, 2, MidpointRounding.AwayFromZero);
            
            double dbl = (double)decValue;

            return (
                new ExtendedValue { NumberValue = dbl },
                new CellFormat
                {
                    NumberFormat = new NumberFormat
                    {
                        Type    = "NUMBER",
                        Pattern = "#,##0.00"
                    }
                }
            );
        }
        
        if (DateTime.TryParse(value,
                CultureInfo.CurrentCulture,
                DateTimeStyles.None,
                out DateTime dateValue))
        {
            double serial = (dateValue - new DateTime(1899, 12, 30)).TotalDays;
            return (
                new ExtendedValue { NumberValue = serial },
                new CellFormat
                {
                    NumberFormat = new NumberFormat
                    {
                        Type    = "DATE",
                        Pattern = "yyyy‑MM‑dd"
                    }
                }
            );
        }
        
        return (new ExtendedValue { StringValue = value }, null);
    }

    public async Task<bool> AppendRowWithBatchUpdate(
        SheetsService sheetsService,
        Dictionary<string, string?> data,
        Dictionary<string, string> columnMappings)
    {
        try
        {
            using CancellationTokenSource cts = new(TimeSpan.FromMinutes(5));
            CancellationToken ct = cts.Token;

            int? sheetId = await SheetsRetryPolicy.ExecuteAsync(
                token => GetSheetIdAsync(sheetsService, token),
                ct);

            if (sheetId == null)
            {
                // The configured sheet/tab does not exist — this is a configuration problem,
                // not something a retry can fix. Surface it so the run can be terminated.
                throw new GoogleSheetsConfigurationException(
                    $"Sheet '{_sheetName}' not found in spreadsheet '{_spreadsheetId}'. Check 'GoogleSheets:SheetName'.");
            }

            string range = $"{_sheetName}!A:Z";
            ValueRange getResponse = await SheetsRetryPolicy.ExecuteAsync(
                token => sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range).ExecuteAsync(token),
                ct);

            int lastNonEmptyRowIndex = -1;
            if (getResponse.Values != null)
            {
                for (int i = 0; i < getResponse.Values.Count; i++)
                {
                    IList<object>? row = getResponse.Values[i];
                    if (row.Any(cell => !string.IsNullOrWhiteSpace(cell.ToString())))
                    {
                        lastNonEmptyRowIndex = i;
                    }
                }
            }

            int nextRow = lastNonEmptyRowIndex + 2;

            List<Request> requests = [];

            foreach (KeyValuePair<string, string> mapping in columnMappings)
            {
                if (string.IsNullOrEmpty(mapping.Value) ||
                    !data.TryGetValue(mapping.Key, out string? value) ||
                    GetColumnIndex(mapping.Value) is not { } columnIndex) continue;

                (ExtendedValue extendedValue, CellFormat? cellFormat) = GetExtendedValueAndFormat(value);

                CellData cellData = new()
                {
                    UserEnteredValue = extendedValue
                };

                if (cellFormat != null)
                {
                    cellData.UserEnteredFormat = cellFormat;
                }

                Request updateCellRequest = new()
                {
                    UpdateCells = new UpdateCellsRequest
                    {
                        Start = new GridCoordinate
                        {
                            SheetId = sheetId.Value,
                            RowIndex = nextRow - 1,
                            ColumnIndex = columnIndex
                        },
                        Rows = new List<RowData>
                        {
                            new() { Values = new List<CellData> { cellData } }
                        },
                        Fields = "userEnteredValue,userEnteredFormat.numberFormat"
                    }
                };

                requests.Add(updateCellRequest);
            }

            if (requests.Any())
            {
                BatchUpdateSpreadsheetRequest batchUpdateRequest = new() { Requests = requests };
                await SheetsRetryPolicy.ExecuteAsync(
                    token => sheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, _spreadsheetId).ExecuteAsync(token),
                    ct);

                Log.Information("Batch update executed successfully. Data appended in row {row}. File: {file}", nextRow, _inputFilePath);
                return true;
            }
            else
            {
                Log.Warning("No data to update for the specified mappings. File: {file}", _inputFilePath);
                return false;
            }
        }
        catch (GoogleSheetsConfigurationException)
        {
            // Already a classified configuration/permission failure — propagate to terminate the run.
            throw;
        }
        catch (Google.GoogleApiException ex) when (IsConfigurationOrPermissionError(ex))
        {
            Log.Error(ex,
                "Google Sheets is misconfigured or access was denied (HTTP {Status}). File: {file}",
                (int)ex.HttpStatusCode, _inputFilePath);
            throw new GoogleSheetsConfigurationException(
                $"Google Sheets is misconfigured or access was denied (HTTP {(int)ex.HttpStatusCode}).", ex);
        }
        catch (Exception ex)
        {
            // Anything left here (5xx/429 after retries, network/IO errors, request timeout)
            // means Google Sheets could not be communicated.
            Log.Error(ex,
                "Google Sheets could not be communicated after retries: {message}. File: {file}",
                ex.Message, _inputFilePath);
            throw new GoogleSheetsCommunicationException(
                "Google Sheets could not be communicated after retries.", ex);
        }
    }

    private static bool IsConfigurationOrPermissionError(Google.GoogleApiException ex)
    {
        int status = (int)ex.HttpStatusCode;
        return status is 400 or 401 or 403 or 404;
    }

    private static bool TryParseFlexibleDecimal(string? raw, out decimal result)
    {
        result = 0;
        if (string.IsNullOrWhiteSpace(raw))
            return false;

        string s = raw.Trim();
        
        if (s.Contains(',') && s.Contains('.'))
        {
            s = s.LastIndexOf(',') > s.LastIndexOf('.') ? s.Replace(".", "").Replace(",", ".") : s.Replace(",", "");
        }
        else if (s.Contains(','))
        {
            int commas = s.Count(c => c == ',');
            s = (commas == 1) ? s.Replace(",", ".") : s.Replace(",", "");
        }
        
        return decimal.TryParse(
            s,
            NumberStyles.AllowDecimalPoint
            | NumberStyles.AllowThousands
            | NumberStyles.AllowLeadingSign,
            CultureInfo.InvariantCulture,
            out result
        );
    }
}

/// <summary>
/// Raised when Google Sheets is misconfigured or access is denied (e.g. wrong sheet/tab,
/// missing share permission, invalid credentials). The run should terminate; the file is NOT skipped.
/// </summary>
public sealed class GoogleSheetsConfigurationException : Exception
{
    public GoogleSheetsConfigurationException(string message)
        : base(message)
    {
    }

    public GoogleSheetsConfigurationException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}

/// <summary>
/// Raised when Google Sheets could not be communicated after the configured retries
/// (transient 5xx/429, network/IO errors, or request timeout).
/// </summary>
public sealed class GoogleSheetsCommunicationException : Exception
{
    public GoogleSheetsCommunicationException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
