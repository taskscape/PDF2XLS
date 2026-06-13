# PDF2XLS

PDF2XLS converts PDF invoices into rows in a Google Spreadsheet. It supports three independent extraction workflows that you can switch between via configuration. Optionally it can upload the original PDF to a public location and write the link to the spreadsheet.

The APIs are not perfect — always double-check the output and fix errors manually.

## Available workflows

You choose the workflow with the `PreferredAPI` setting in [appsettings.json](PDF2XLS/appsettings.json). Valid values:

| `PreferredAPI` value | What it does | External services required |
|---|---|---|
| `NuDelta` | Sends the PDF to the NuDelta Invoice service and polls until extraction is done. | NuDelta Invoice account |
| `OpenAIResponses` | Sends the PDF as an `input_file` to the OpenAI Responses API with a structured prompt + JSON schema. | OpenAI account with API key |
| `AzureDocumentIntelligence` | Sends the PDF to the Azure Document Intelligence `prebuilt-invoice` model and maps the structured fields to the internal schema. | Azure subscription with a Document Intelligence resource |

All three workflows share the **Google Sheets**, **Seq**, and (optional) **PDF upload** configuration described below.

The application validates that the configuration fields required by the selected workflow are present at startup and exits with a clear error message if anything is missing.

---

# Prerequisites

Regardless of which workflow you use, you always need:

- A **Google Service Account** with access to your target spreadsheet (JSON key file).
- A Google Spreadsheet with column headers matching the `Mappings` values.
- *(Optional)* A **Seq** server if you want centralised logs.
- *(Optional)* The **PDF2URL** helper executable if you want to upload PDFs and store a link in the sheet.

Then, depending on the workflow:

- **NuDelta workflow** → NuDelta Invoice account.
- **OpenAIResponses workflow** → OpenAI account with a funded API key.
- **AzureDocumentIntelligence workflow** → Azure subscription with a Document Intelligence (Form Recognizer) resource.

---

# Account setup — how to obtain the parameters

## 1. Google Sheets (required for every workflow)

1. Go to <https://console.cloud.google.com/> and create (or select) a project.
2. Open **APIs & Services → Library** and enable **Google Sheets API**.
3. Open **APIs & Services → Credentials → Create Credentials → Service Account**. Give it a name (this name becomes `GoogleSheets:ApplicationName`).
4. Open the new service account, go to the **Keys** tab, **Add Key → Create new key → JSON**. Save the downloaded file somewhere safe — the path goes into `GoogleSheets:ServiceAccountFile`.
5. Open the JSON file and copy the `client_email` value (looks like `name@project.iam.gserviceaccount.com`). In your Google Spreadsheet, click **Share** and grant **Editor** access to that email.
6. From the spreadsheet URL `https://docs.google.com/spreadsheets/d/<SPREADSHEET_ID>/edit#gid=0`, copy `<SPREADSHEET_ID>` into `GoogleSheets:SpreadsheetId`.
7. Set `GoogleSheets:SheetName` to the tab name (e.g. `Sheet1`).
8. Fill in `GoogleSheets:Mappings` — each field's value is the **column letter** in the sheet where that field should be written (e.g. `"InvoiceNumber": "A"`). Leave blank to skip a field.

## 2. NuDelta workflow

### NuDelta Invoice account
1. Sign up at the NuDelta Invoice portal and confirm you have an active subscription.
2. Put your portal login into `NuDeltaCredentials:Username` and `NuDeltaCredentials:Password`. These are used with HTTP Basic auth against the NuDelta API.

## 3. OpenAIResponses workflow

1. Create an account at <https://platform.openai.com/>.
2. Add credit / payment method under **Billing**. The Responses API with `input_file` requires a funded account.
3. Go to **API keys → Create new secret key**. Copy the value (`sk-…`) into `OpenAI:OpenAI_APIKey`. You won't be able to see it again — store it safely.
4. Choose a model that supports the `input_file` content type and set it in `OpenAI:OpenAI_Model`. Recommended: `gpt-4o-mini` (good cost/quality trade-off). Other supported options at the time of writing: `gpt-4o`, `gpt-4.1`, `gpt-4.1-mini`.
5. Leave `OpenAI:Prompt` at the default unless you know what you're doing. The placeholder `{schema}` is substituted at runtime with the JSON schema the program expects — do **not** remove it.

> Cost note: each invoice costs a few cents on `gpt-4o-mini`. Larger / multi-page PDFs cost more because the file is sent as base64.

## 4. AzureDocumentIntelligence workflow

1. In the [Azure Portal](https://portal.azure.com/), click **Create a resource → AI + Machine Learning → Document Intelligence** (previously called *Form Recognizer*).
2. Choose a subscription, a resource group, a region (e.g. `westeurope`), and a pricing tier. **F0** (free) lets you process 500 pages/month; **S0** (standard) is pay-as-you-go.
3. After deployment, open the resource and go to **Keys and Endpoint** in the left menu.
4. Copy **Endpoint** (looks like `https://<resource>.cognitiveservices.azure.com/`) into `AzureDocumentIntelligence:Endpoint`.
5. Copy **KEY 1** into `AzureDocumentIntelligence:ApiKey`.
6. This workflow uses the **`prebuilt-invoice`** model, which is built into the service — you do not need to train anything. It is available in all regions that support Document Intelligence v4.
7. Set `AzureDocumentIntelligence:MonthlyPageLimit` to `500` for the F0 free-tier guard. Set it to `0` to disable the internal guard. The app stores the month counter in `AzureDocumentIntelligence:MonthlyQuotaCounterPath`, or next to the executable as `azure-document-intelligence-quota.json` when that value is empty.

> Note: Document Intelligence is deterministic and returns structured fields directly. It does not use OpenAI.

## 5. Optional: PDF upload to a public URL

If you want a `DocumentLink` column in the sheet pointing to the uploaded PDF:
1. Obtain or build a small uploader CLI (the `PDF2URL` helper) that takes a file path as an argument and prints the public URL to stdout.
2. Set `UploadPDF:Enabled` to `"true"` and `UploadPDF:PDF2URLPath` to the executable path.
3. Add `DocumentLink` to `GoogleSheets:Mappings` with the column letter to write the URL into.

Set `UploadPDF:Enabled` to `"false"` to skip uploading.

## 6. Optional: Seq logging

1. Run a [Seq](https://datalust.co/seq) server (locally or remote).
2. Put its URL into `Seq:ServerAddress` (e.g. `http://localhost:5341/`).
3. Create an API key in Seq and set it in `Seq:ApiKey`.
4. `Seq:AppName` is the value of the `Application` property used to filter events in Seq.

If you don't have a Seq server, leave `Seq:ServerAddress` empty — the file logs under `logs/` still work.

---

# Installation

## Download a release

1. Open the [GitHub Releases](https://github.com/taskscape/PDF2XLS/releases) page and download the latest `PDF2XLS-<version>-win-x64.zip` asset.
2. Unpack the archive into a folder. It contains:
   - `PDF2XLS.exe` — self-contained, single-file executable (64-bit Windows)
   - `appsettings.json` — configuration file (kept alongside the executable so you can edit it without rebuilding)
3. Open `appsettings.json` and fill in:
    - **Always:** `PreferredAPI`, `GoogleSheets.*`.
    - **Workflow-specific** (only for the workflow you chose — see above).
    - **Optional:** `UploadPDF.*`, `Seq.*`.

## Full configuration reference

| Section | Key | Required for | Description |
|---|---|---|---|
| (root) | `PreferredAPI` | All | `NuDelta`, `OpenAIResponses`, or `AzureDocumentIntelligence`. |
| `NuDeltaCredentials` | `Username`, `Password` | NuDelta | NuDelta Invoice portal login (HTTP Basic auth). |
| `OpenAI` | `OpenAI_APIKey`, `OpenAI_Model`, `Prompt` | OpenAIResponses | OpenAI key, model id, and prompt containing `{schema}`. |
| `AzureDocumentIntelligence` | `Endpoint`, `ApiKey`, `MonthlyPageLimit`, `MonthlyQuotaCounterPath` | AzureDocumentIntelligence | Azure Document Intelligence endpoint/key and optional internal monthly page guard. `MonthlyPageLimit` is a non-negative integer; `0` disables the guard. `MonthlyQuotaCounterPath` may be blank to use `{app}/azure-document-intelligence-quota.json`. |
| `GoogleSheets` | `ServiceAccountFile`, `SpreadsheetId`, `ExpectedSpreadsheetName`, `SheetName`, `ApplicationName`, `Mappings` | All | Google Sheets target. `ExpectedSpreadsheetName` is verified against the live spreadsheet title at startup; leave blank to skip the check. `Mappings` values are spreadsheet **column letters**. |
| `UploadPDF` | `Enabled`, `PDF2URLPath` | Optional | Enable to upload the PDF and store the link in `DocumentLink`. |
| `Seq` | `ServerAddress`, `ApiKey`, `AppName` | Optional | Centralised Seq logging. `ApiKey` is required when `ServerAddress` is set. |

---

# Usage

This is a CLI application. Pass one or more PDF file paths, or a folder path, as arguments. You can also drag-and-drop a PDF (or folder) onto the executable.

```
PDF2XLS.exe <file.pdf> [file2.pdf ...] | <folder path>
```

## Processing a single file

```
PDF2XLS.exe "C:\invoices\invoice-001.pdf"
```

The selected workflow processes the file, parses the response into the internal schema, and appends a row to the configured Google Sheet.

## Processing multiple files

Pass two or more PDF paths on the command line:

```
PDF2XLS.exe "C:\invoices\invoice-001.pdf" "C:\invoices\invoice-002.pdf"
```

- Files are processed **in the order given** on the command line.
- Duplicate paths are skipped automatically.

## Processing a folder

Pass a single folder path to process all PDF files inside it:

```
PDF2XLS.exe "C:\invoices\2026-05\"
```

- The app scans for **all `.pdf` files** in the folder (top-level only — subfolders are not scanned).
- Files are processed **in alphabetical order**.

## Batch behaviour (all input modes)

When more than one file is processed in a single application run (multiple CLI paths or a folder):

- Each file gets its own unique **RunID** (GUID), but they all share the same **RunTime** timestamp.
- Each processed file is appended as a separate row in the Google Sheet.
- Each processed file is renamed to `.bak` after it succeeds.
- If a file fails, it is **left untouched** and processing continues with the next file.
- If a folder contains no PDF files, the app exits with a warning message.

# Logging

On startup the program creates a `logs/` subfolder next to the executable and writes a daily rolling log file (365 days retained, up to 365 files).

If `Seq:ServerAddress` is set, events are also forwarded to Seq using the configured `Seq:ApiKey`. After each PDF file finishes processing (success or failure), buffered log events are flushed to both the local file and Seq before the next file starts.

| Destination | When active | Notes |
|---|---|---|
| Local file | Always | `{exe}/logs/log-YYYYMMDD.txt` |
| Seq | When `Seq:ServerAddress` is configured | Requires `Seq:ApiKey` |

## Processed file naming

When the Google Sheets write succeeds, the original PDF is renamed in-place using the following convention:

```
{RunTime} {RunID} {OriginalFileName}.bak
```

| Part | Format | Details |
|---|---|---|
| `RunTime` | `yyyyMMdd HHmmss` | UTC timestamp of application start (UTC). Set once per batch run; all files in the same run share this value. |
| `RunID` | GUID | A unique ID generated for **each file**. Also written to the Google Sheet (column mapped to `RunID`). |
| `OriginalFileName` | — | The original filename, unchanged, including its extension. |
| `.bak` | — | Fixed suffix appended after the original extension. |

All three parts are separated by a **single space**. There are no underscores in the filename.

**Example:**

```
20260517 013221 48070c04-bce5-4205-8ee7-ff506c8f2533 invoice-001.pdf.bak
```

> **Why spaces?** Spaces make the three logical parts visually distinct without requiring a special delimiter character. The GUID already contains hyphens internally, so using an underscore as a separator would be ambiguous when splitting the name programmatically. Spaces allow a simple `Split(' ', 3)` to recover `RunTime`, `RunID`, and `OriginalFileName` unambiguously.

The file is left untouched if processing fails or the Google Sheets write does not succeed.

# Notes per workflow

- **NuDelta** — polls for the result with exponential backoff (up to 5 attempts, 1-second base delay). The outer operation also retries on exception with a 1-second delay (up to 5 attempts).
- **OpenAIResponses** — the inner HTTP client retries on HTTP 429 / 5xx with exponential backoff (3 attempts). The outer operation also retries on exception with exponential backoff (3 attempts).
- **AzureDocumentIntelligence** — checks the configured internal monthly page guard before every Azure submission attempt, then retries transient extraction failures with exponential backoff (3 retries, 2 s → 4 s → 8 s). `OperationCanceledException`, quota guard stops, and Azure `RequestFailedException` 4xx responses are not retried by the outer policy. Azure SDK retries are disabled for this client so every retryable submission attempt passes through the quota guard. PDF upload happens after the Azure retry block, so an upload failure does not resubmit the document to Azure in the same run.

---

# Resilience — timeouts and retries

The application is designed so that a network outage or a service disruption during any stage of processing always leaves the source file untouched. The next run will pick it up and try again.

## File-safety guarantee

A file is only renamed/deleted **after all of the following have succeeded**:
1. PDF extraction (NuDelta / OpenAI / Azure DI)
2. PDF upload to public URL *(if `UploadPDF:Enabled` is `"true"`)*
3. Google Sheets row write

If any step fails — even after all retries are exhausted — the file stays in place.

For the AzureDocumentIntelligence workflow, the internal monthly quota counter is updated immediately after Azure accepts a document submission, before polling, PDF upload, or Google Sheets writes. This prevents post-submission failures from causing uncounted repeat submissions.

## Timeouts

| Layer | Timeout | Behaviour on expiry |
|---|---|---|
| **Azure DI polling** | 5 minutes | `OperationCanceledException` thrown; not retried by the outer policy |
| **NuDelta HTTP client** (upload + poll requests) | 5 minutes per request | `TaskCanceledException` thrown; propagates to outer retry policy |
| **OpenAI HTTP client** | 5 minutes per request | `TaskCanceledException` thrown; propagates to outer retry policy |
| **Google Sheets API calls** (all three calls per write) | 5 minutes per call | `OperationCanceledException` thrown; propagates as a write failure — file stays |
| **PDF2URL process** | 5 minutes | Process is killed; empty URL returned → write not attempted → file stays |

## Retry policies

| Workflow / stage | Retries | Delay strategy | What triggers a retry |
|---|---|---|---|
| **NuDelta outer** (whole operation) | up to 5 | 1 s fixed | Any exception except `OperationCanceledException` |
| **NuDelta inner** (result polling) | up to 5 | Exponential (`2^n` s) | Document state is not `done` |
| **OpenAI outer** (whole operation) | up to 3 | Exponential (`2^n` s) | Any exception except `OperationCanceledException` |
| **OpenAI inner** (HTTP call) | up to 3 | Exponential (`2^n` s) | HTTP 5xx or HTTP 429 |
| **Azure DI** (extraction only) | up to 3 retries | Exponential (`2^n` s) | Transient exceptions except `OperationCanceledException`, Azure quota guard stops, and Azure `RequestFailedException` 4xx responses |
| **Google Sheets** (each API call) | up to 3 | Exponential (`2^n` s) | HTTP 5xx, HTTP 429, `HttpRequestException`, `IOException` |

`OperationCanceledException` is never retried in any policy — it signals an intentional 5-minute timeout and should propagate immediately so the file is left untouched for the next run.

## Azure monthly quota guard

When `AzureDocumentIntelligence:MonthlyPageLimit` is greater than `0`, the app keeps a UTC-month counter in JSON. Before each Azure submission attempt, including retry attempts, it counts the local PDF pages and stops the batch without calling Azure if the monthly limit has already been reached or the next document would exceed it. The source PDF is left untouched and the log records that the internally configured monthly page quota has been achieved or would be exceeded.

## Spreadsheet name verification

Before processing any files the application fetches the spreadsheet title and compares it to `GoogleSheets:ExpectedSpreadsheetName`. If the names do not match, or if the API call fails after 3 retries, the application logs the error and exits without processing any files. This prevents writing data to the wrong spreadsheet when `SpreadsheetId` is misconfigured.

---

# Building and releasing

## Build locally

Requires the [.NET 10 SDK](https://dotnet.microsoft.com/download).

```powershell
dotnet publish PDF2XLS/PDF2XLS.csproj `
  -c Release `
  -r win-x64 `
  --self-contained true `
  -p:PublishSingleFile=true `
  -p:PublishReadyToRun=true `
  -p:EnableCompressionInSingleFile=true `
  -o publish
```

The output directory contains `PDF2XLS.exe` and `appsettings.json`.

## GitHub release pipeline

Pushing a version tag triggers the [Release workflow](.github/workflows/release.yml), which:

1. Builds a self-contained, single-file `win-x64` executable on `windows-latest`
2. Packages `PDF2XLS.exe` and `appsettings.json` into `PDF2XLS-<tag>-win-x64.zip`
3. Creates a GitHub release with auto-generated release notes and attaches the zip

To publish a release:

```bash
git tag v1.0.0
git push origin v1.0.0
```
