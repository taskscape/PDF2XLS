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

> Note: Document Intelligence is deterministic and returns structured fields directly, so there is no retry loop or LLM fallback. It does not use OpenAI.

## 5. Optional: PDF upload to a public URL

If you want a `DocumentLink` column in the sheet pointing to the uploaded PDF:
1. Obtain or build a small uploader CLI (the `PDF2URL` helper) that takes a file path as an argument and prints the public URL to stdout.
2. Set `UploadPDF:Enabled` to `"true"` and `UploadPDF:PDF2URLPath` to the executable path.
3. Add `DocumentLink` to `GoogleSheets:Mappings` with the column letter to write the URL into.

Set `UploadPDF:Enabled` to `"false"` to skip uploading.

## 6. Optional: Seq logging

1. Run a [Seq](https://datalust.co/seq) server (locally or remote).
2. Put its URL into `Seq:ServerAddress` (e.g. `http://localhost:5341/`).
3. `Seq:AppName` is the value of the `App` property used to filter events in Seq.

If you don't have a Seq server, leave the address empty — the file logs under `logs/` still work.

---

# Installation

1. Download the latest release and unpack it into a folder.
2. Open [appsettings.json](PDF2XLS/appsettings.json) and fill in:
    - **Always:** `PreferredAPI`, `GoogleSheets.*`, `DeleteFileAfterProcessing`.
    - **Workflow-specific** (only for the workflow you chose — see above).
    - **Optional:** `UploadPDF.*`, `Seq.*`.

## Full configuration reference

| Section | Key | Required for | Description |
|---|---|---|---|
| (root) | `PreferredAPI` | All | `NuDelta`, `OpenAIResponses`, or `AzureDocumentIntelligence`. |
| (root) | `DeleteFileAfterProcessing` | All | `"true"` deletes the processed PDF; `"false"` keeps it. |
| `NuDeltaCredentials` | `Username`, `Password` | NuDelta | NuDelta portal credentials (HTTP Basic auth). |
| `OpenAI` | `OpenAI_APIKey`, `OpenAI_Model`, `Prompt` | OpenAIResponses | OpenAI key, model id, and prompt containing `{schema}`. |
| `AzureDocumentIntelligence` | `Endpoint`, `ApiKey` | AzureDocumentIntelligence | Azure Document Intelligence endpoint and key. |
| `GoogleSheets` | `ServiceAccountFile`, `SpreadsheetId`, `SheetName`, `ApplicationName`, `Mappings` | All | Google Sheets target. `Mappings` values are spreadsheet **column letters**. |
| `UploadPDF` | `Enabled`, `PDF2URLPath` | Optional | Enable to upload the PDF and store the link in `DocumentLink`. |
| `Seq` | `ServerAddress`, `AppName` | Optional | Centralised Seq logging. |

---

# Usage

This is a CLI application. Pass the path to the PDF invoice as the first argument, or drag-and-drop the PDF onto the executable.

```
PDF2XLS.exe "C:\invoices\invoice-001.pdf"
```

The selected workflow processes the file, parses the response into the internal schema, and appends a row to the configured Google Sheet.

# Logging

On startup the program creates a `logs/` subfolder next to the executable and writes a daily rolling log file (365 days retained). If `Seq:ServerAddress` is set, it also forwards events to Seq.

## Processed file naming

When `DeleteFileAfterProcessing` is `"false"` and the Google Sheets write succeeds, the original PDF is renamed in-place using the following convention:

```
{RunTime}_{RunID}_{OriginalFileName}.bak
```

| Part | Format | Details |
|---|---|---|
| `RunTime` | `yyyyMMdd HHmmss` | UTC timestamp of application start. Set once for the whole batch run. |
| `RunID` | GUID | A new unique ID generated per file. Also written to the Google Sheet (column mapped to `RunID`). |
| `OriginalFileName` | — | The original filename, unchanged. |
| `.bak` | — | Fixed suffix appended after the original extension. |

**Example:**

```
20260517 013221_48070c04-bce5-4205-8ee7-ff506c8f2533_invoice-001.pdf.bak
```

The file is left untouched if processing fails or the Google Sheets write does not succeed.

# Notes per workflow

- **NuDelta** — retries up to 5 times with a 1-second delay while the document status is `processing`. If all retries fail, the file is skipped and the error is logged.
- **OpenAIResponses** — retries on HTTP 429 / 5xx with exponential backoff (3 attempts). Output quality depends on the model; results can vary between runs.
- **AzureDocumentIntelligence** — single deterministic call (the Azure SDK handles transient retries internally). Output is the most consistent of the three because fields come from a trained invoice model rather than a generative LLM.