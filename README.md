# PDF2XLS

This program lets you easily convert pdf invoices into rows in Google Spreadsheets by utilising NuDelta or OpenAI's API.
It has the ability to upload your file to a service and provide a link to it in the spreadsheet. It also generates a text file with extracted text using LLMWhisperer's API.
The APIs are not perfect, always double check the output and fix errors manually.

# Prerequisites:
- NuDelta Invoice account
- OpenAI Token with balance
- Google Service Account file in json format
- PDF2URL program if wanted
- LLMWhisperer account

# Installation

Download the latest release and unpack it to your destination folder.
Inside `appsettings.json`, there are field which you need to fill in:

- `Username`: Your NuDelta Invoice login username.
- `Password`: Your NuDelta Invoice login Password.
- `PreferredAPI`: The API that you want to use. It can either be `NuDelta` or `OpenAI`.
- `OpenAI_APIKey`: Your OpenAI API token.
- `GoogleSheets`:
    - `ServiceAccountFile`: Path to your Google Service Account json file.
    - `SpreadsheetId`: ID of your Google Spreadsheet. You can find it by going into your spreadsheet in a browser and copying it from the URL (it comes after /spreadsheets/d/).
    - `SheetName`: Name of your spreadsheet sheet.
    - `ApplicationName`: Name of your Google Sheets API Service Account (Not email).
    - `Mappings`: Which columns in the Google Spreadsheet should have what information. `DocumentLink` refers to the url where your file is uploaded.
- `DeleteFileAfterProcessing`: Set to `true` if you want the processed file to be deleted, or `false` it you want it to be backed up.
- `Seq`:
    - `ServerAddress`: Your Seq server's address.
    - `AppName`: The application name by which you can filter in Seq.
- `UploadPDF`:
    - `Enabled`: Set to `true` if you want your file to be uploaded and accessible through `DocumentLink` mapping, or `false` if you don't want that.
    - `PDF2URLPath`: Path to your PDF2URL executable, which is responsible for uploading the file and returning the url to it.
- `Whisperer`:
    - `BaseUrl`: Base url of the LLMWhisperer's API.
    - `ApiKey`: Your LLMWhisperer API key. Your key must match the correct base url for either Europe or US.

# Usage

This is a CLI application, which you can use in two ways:
- Running from the console by specifying an argument:
    - The path to your pdf invoice you want to convert

- Drag and Drop:
    - You can drag your pdf file onto the app's executable.

# Logging

When ran, the program will create a subfolder named `logs` in it's location, in which it will produce and store logs for up to 7 different days.
It also logs to a Seq server.

# Important notes

When using OpenAI API, the results will be inconsistent and may require multiple tries to successfully process the file. LLMWhisperer is also utilised to extract the text from your PDF and that is used as input when all retries fail.
Please be patient as it is based on an AI's response.