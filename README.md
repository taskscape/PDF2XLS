# PDF2XLS

This program lets you easily convert pdf invoices into excel spreadsheets by utilising NuDelta or OpenAI's API.
It also appends a row in a specified Google Spreadsheet.
The APIs are not perfect, always double check the output and fix errors manually.

# Prerequisites:
- NuDelta Invoice account
- OpenAI Token with balance
- Google Service Account file in json format

# Installation

Download the latest release and unpack it to your destination folder.
Inside `appsettings.json`, there are field which you need to fill in:

- `Username`: Your NuDelta Invoice login username.
- `Password`: Your NuDelta Invoice login Password.
- `PreferredAPI`: The API that you want to use. It can either be `NuDelta` or `OpenAI`.
- `OpenAI_APIKey`: Your OpenAI API token.
- `ServiceAccountFile`: Path to your Google Service Account json file.
- `SpreadsheetId`: ID of your Google Spreadsheet. You can find it by going into your spreadsheet in a browser and copying it from the URL (it comes after /spreadsheets/d/).
- `SheetName`: Name of your spreadsheet sheet.
- `ApplicationName`: Name of your Google Sheets API Service Account (Not email).
- `DeleteFileAfterProcessing`: Set to `true` if you want the processed file to be deleted, or `false` it you want it to be backed up.
- `Mappings`: Which columns in the Google Spreadsheet should have what information.

These fields are necessary for API communication.

# Usage

This is a CLI application, which you can use in two ways:
- Running from the console by specifying arguments:
    - First argument is the path to your pdf invoice you want to convert
    - Second [optional] is the output directory where you want your spreadsheet to go.
    If not specified, it is set to wherever this program is run from.

- Drag and Drop:
    - You can drag your pdf file onto the app's executable and the spreadsheet will be created in the same location the pdf was in.

# Logging

When ran, the program will create a subfolder named `logs` in it's location, in which it will produce and store logs for up to 7 different days.

# Important notes

When using OpenAI API, the results will be inconsistent and may require multiple tries to successfully process the file.
Please be patient as it is based on an AI's response.