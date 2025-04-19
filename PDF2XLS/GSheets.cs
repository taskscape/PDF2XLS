using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Configuration;
using Serilog;
using System.Globalization;

namespace PDF2XLS;

public class GSheets
{
    private IConfiguration Config { get; }
    private readonly string _serviceAccountFile;
    private readonly string _spreadsheetId;
    private readonly string _sheetName;
    private readonly string _applicationName;
    private readonly string _inputFilePath;

    public GSheets(IConfiguration config, string inputFilePath)
    {
        Config = config;
        _serviceAccountFile = Config["GoogleSheets:ServiceAccountFile"] ?? string.Empty;
        _spreadsheetId = Config["GoogleSheets:SpreadsheetId"] ?? string.Empty;
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

    private int? GetSheetId(SheetsService sheetsService)
    {
        Spreadsheet? spreadsheet = sheetsService.Spreadsheets.Get(_spreadsheetId).Execute();
        Sheet? sheet = spreadsheet.Sheets.FirstOrDefault(s =>
            s.Properties.Title.Equals(_sheetName, StringComparison.OrdinalIgnoreCase));
        return sheet?.Properties.SheetId;
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

    public void AppendRowWithBatchUpdate(
        SheetsService sheetsService,
        Dictionary<string, string?> data,
        Dictionary<string, string> columnMappings)
    {
        try
        {
            int? sheetId = GetSheetId(sheetsService);
            if (sheetId == null)
            {
                Log.Error("Sheet {sheetName} not found in spreadsheet {spreadsheetId}", _sheetName, _spreadsheetId);
                return;
            }
            
            string range = $"{_sheetName}!A:Z";
            SpreadsheetsResource.ValuesResource.GetRequest? getRequest = sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
            ValueRange getResponse = getRequest.Execute();
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
                sheetsService.Spreadsheets.BatchUpdate(batchUpdateRequest, _spreadsheetId).Execute();

                Log.Information("Batch update executed successfully. Data appended in row {row}. File: {file}", nextRow, _inputFilePath);
            }
            else
            {
                Log.Warning("No data to update for the specified mappings. File: {file}", _inputFilePath);
            }
        }
        catch (Exception ex)
        {
            Log.Error("An error occurred during batch update: {message}. File: {file}", ex.Message, _inputFilePath);
        }
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