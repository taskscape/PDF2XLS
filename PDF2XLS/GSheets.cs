using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Microsoft.Extensions.Configuration;

namespace PDF2XLS;

public class GSheets
{
    private IConfiguration Config { get; }
    private readonly string _serviceAccountFile;
    private readonly string _spreadsheetId;
    private readonly string _sheetName;
    private readonly string _applicationName;
    private static readonly string[] Scopes =
    {
        SheetsService.Scope.Spreadsheets,
        SheetsService.Scope.Drive
    };

    public GSheets(IConfiguration config)
    {
        Config = config;
        _serviceAccountFile = Config["GoogleSheets:ServiceAccountFile"] ?? string.Empty;
        _spreadsheetId = Config["GoogleSheets:SpreadsheetId"] ?? string.Empty;
        _sheetName = Config["GoogleSheets:SheetName"] ?? string.Empty;
        _applicationName = Config["GoogleSheets:ApplicationName"] ?? string.Empty;
    }

    public SheetsService CreateSheetsService()
    {
        var credential = GoogleCredential.FromFile(_serviceAccountFile).CreateScoped(Scopes);
        return new SheetsService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credential,
            ApplicationName = _applicationName
        });
    }

    public void AddRow(
        SheetsService sheetsService,
        Dictionary<string, string> data,
        Dictionary<string, string> columnMappings
    )
    {
        try
        {
            int maxColumns = columnMappings
                .Values
                .Where(column => !string.IsNullOrEmpty(column))
                .Select(column => GetColumnIndex(column) ?? 0)
                .DefaultIfEmpty(0)
                .Max() + 1;
            
            List<object> row = [..new object[maxColumns]];
            
            foreach (KeyValuePair<string, string> mapping in columnMappings)
            {
                if (!string.IsNullOrEmpty(mapping.Value) && data.TryGetValue(mapping.Key, out string value))
                {
                    int? columnIndex = GetColumnIndex(mapping.Value);
                    if (columnIndex.HasValue)
                    {
                        row[columnIndex.Value] = value;
                    }
                }
            }
            
            ValueRange valueRange = new ValueRange
            {
                Values = new List<IList<object>> { row }
            };

            var request = sheetsService.Spreadsheets.Values.Append(
                valueRange,
                _spreadsheetId,
                _sheetName
            );

            request.ValueInputOption = SpreadsheetsResource.ValuesResource
                .AppendRequest
                .ValueInputOptionEnum.USERENTERED;

            request.InsertDataOption = SpreadsheetsResource.ValuesResource
                .AppendRequest
                .InsertDataOptionEnum.INSERTROWS;

            request.Execute();

            Console.WriteLine("Data appended to Google Sheet successfully!");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    private static int? GetColumnIndex(string columnLetter)
    {
        if (string.IsNullOrEmpty(columnLetter))
            return null;

        int index = 0;
        foreach (char c in columnLetter.ToUpper())
        {
            index = index * 26 + (c - 'A') + 1;
        }
        return index - 1;
    }
}