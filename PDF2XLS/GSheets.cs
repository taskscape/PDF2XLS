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

    public void AddRow(SheetsService sheetsService, string invoiceDate, string reference, string sellerName, string invoiceNumber, string priceInPln, string priceToPay, string originalCurrency)
    {
        try
        {
            List<object> row =
            [
                invoiceDate, sellerName, invoiceNumber, reference, "", "", priceInPln, priceToPay, originalCurrency, "", "", ""
            ];
            
            ValueRange valueRange = new ValueRange
            {
                Values = new List<IList<object>> { row }
            };
            
            SpreadsheetsResource.ValuesResource.AppendRequest request =
                sheetsService.Spreadsheets.Values.Append(
                    valueRange,
                    _spreadsheetId,
                    $"{_sheetName}!A:A"
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
            Console.WriteLine($"An error occured: {ex.Message}");
        }
    }
}