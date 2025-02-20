using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Configuration;
using Serilog;

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
            if (c < 'A' || c > 'Z')
                return null;
            index = index * 26 + (c - 'A' + 1);
        }
        return index - 1;
    }
    
    private int? GetSheetId(SheetsService sheetsService)
    {
        var spreadsheet = sheetsService.Spreadsheets.Get(_spreadsheetId).Execute();
        var sheet = spreadsheet.Sheets.FirstOrDefault(s =>
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
    
    public void AppendRowWithBatchUpdate(
        SheetsService sheetsService,
        Dictionary<string, string> data,
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
            
            string range = $"{_sheetName}!A:A";
            SpreadsheetsResource.ValuesResource.GetRequest? getRequest = sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);
            ValueRange? getResponse = getRequest.Execute();
            int nextRow = (getResponse.Values?.Count ?? 0) + 1;
            
            List<Request> requests = [];

            foreach (KeyValuePair<string, string> mapping in columnMappings)
            {
                if (!string.IsNullOrEmpty(mapping.Value) &&
                    data.TryGetValue(mapping.Key, out string value) &&
                    GetColumnIndex(mapping.Value) is { } columnIndex)
                {
                    CellData cellData = new CellData
                    {
                        UserEnteredValue = new ExtendedValue { StringValue = value }
                    };
                    
                    Request updateCellRequest = new Request
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
                                new()
                                {
                                    Values = new List<CellData> { cellData }
                                }
                            },
                            Fields = "userEnteredValue"
                        }
                    };

                    requests.Add(updateCellRequest);
                }
            }

            if (requests.Any())
            {
                BatchUpdateSpreadsheetRequest batchUpdateRequest = new BatchUpdateSpreadsheetRequest { Requests = requests };
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
}