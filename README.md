# PDF2XLS

This program lets you easily convert pdf invoices into excel spreadsheets by utilising NuDelta's API.
The API is not perfect, always double check the output and fix errors manually.

# Prerequisites:
- NuDelta Invoice account

# Configuration

Inside `appsettings.json`, there are field which you need to fill in:

- `Username`: Your NuDelta Invoice login username
- `Password`: Your NuDelta Invoice login Password

These are necessary for API communication.

# Usage

This is a CLI application, which you can use in two ways:
- Running from the console by specifying arguments:
    - First argument is the path to your pdf invoice you want to convert
    - Second [optional] is the output directory where you want your spreadsheet to go.
    If not specified, it is set to wherever this program is run from.

- Drag and Drop:
    - You can drag your pdf file onto the app's executable and the spreadsheet will be created in the same location the pdf was in.