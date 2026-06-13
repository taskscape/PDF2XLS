using Azure;
using Azure.AI.DocumentIntelligence;
using Microsoft.Extensions.Configuration;
using Serilog;
using System.Globalization;
using System.Text.Json;

namespace PDF2XLS;

/// <summary>
/// Processes PDF invoices using the Azure Document Intelligence prebuilt-invoice model.
/// Handles OCR and structured field extraction without requiring a separate LLM call.
/// Returns a JSON string in the application's internal schema format.
/// </summary>
public class AzureDocumentIntelligenceProcessor
{
    private readonly string _endpoint;
    private readonly string _apiKey;
    private readonly AzureDocumentIntelligenceQuotaTracker? _quotaTracker;

    public AzureDocumentIntelligenceProcessor(
        IConfiguration config,
        AzureDocumentIntelligenceQuotaTracker? quotaTracker = null)
    {
        _endpoint = config["AzureDocumentIntelligence:Endpoint"] ?? string.Empty;
        _apiKey = config["AzureDocumentIntelligence:ApiKey"] ?? string.Empty;
        _quotaTracker = quotaTracker;
    }

    /// <summary>
    /// Analyzes the PDF with the prebuilt-invoice model and returns a JSON string
    /// matching the application's internal schema (same shape as NuDelta / OpenAI output).
    /// </summary>
    public async Task<string?> ProcessPdfAsync(string filePath)
    {
        try
        {
            int? documentPageCount = null;
            if (_quotaTracker?.IsEnabled == true)
            {
                documentPageCount = PdfPageCounter.CountPages(filePath);
                _quotaTracker.EnsureCanSubmit(filePath, documentPageCount.Value);
            }

            DocumentIntelligenceClient client = new(
                new Uri(_endpoint),
                new AzureKeyCredential(_apiKey),
                new DocumentIntelligenceClientOptions
                {
                    Retry =
                    {
                        MaxRetries = 0
                    }
                });

            BinaryData pdfData = BinaryData.FromBytes(await File.ReadAllBytesAsync(filePath));

            Log.Information("Submitting PDF to Azure Document Intelligence. File: {file}", filePath);

            Operation<AnalyzeResult> operation = await client.AnalyzeDocumentAsync(
                WaitUntil.Started,
                "prebuilt-invoice",
                pdfData);

            if (documentPageCount.HasValue)
            {
                _quotaTracker?.RecordSuccessfulSubmission(filePath, documentPageCount.Value);
            }

            using CancellationTokenSource pollCts = new(TimeSpan.FromMinutes(5));
            await operation.WaitForCompletionAsync(pollCts.Token);

            AnalyzeResult result = operation.Value;

            if (result.Documents is not { Count: > 0 })
            {
                Log.Error("Azure Document Intelligence returned no documents. File: {file}", filePath);
                return null;
            }

            AnalyzedDocument invoice = result.Documents[0];
            Log.Information(
                "Azure Document Intelligence analysis complete. Confidence: {Confidence}. File: {file}",
                invoice.Confidence, filePath);

            return BuildJsonSchema(invoice);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Azure Document Intelligence processing error. File: {file}", filePath);
            throw;
        }
    }

    private static string BuildJsonSchema(AnalyzedDocument invoice)
    {
        // ── Top-level scalar fields ──────────────────────────────────────────
        string invn      = GetString(invoice, "InvoiceId");
        string reference = GetString(invoice, "PurchaseOrder");
        string issue     = GetDate(invoice, "InvoiceDate");
        string sale      = GetDate(invoice, "ServiceStartDate");
        string payment   = GetString(invoice, "PaymentTerm");
        string maturity  = GetDate(invoice, "DueDate");
        string total     = GetCurrencyAmount(invoice, "InvoiceTotal");
        string left      = GetCurrencyAmount(invoice, "AmountDue");

        // Currency code: prefer explicit CurrencyCode field, fall back to InvoiceTotal currency.
        string currency = GetString(invoice, "CurrencyCode");
        if (string.IsNullOrEmpty(currency) &&
            invoice.Fields.TryGetValue("InvoiceTotal", out DocumentField? totalForCcy) &&
            totalForCcy.ValueCurrency != null)
        {
            currency = totalForCcy.ValueCurrency.CurrencyCode ?? string.Empty;
        }

        // Paid = InvoiceTotal − AmountDue (computed when both fields are present).
        string paid = string.Empty;
        if (invoice.Fields.TryGetValue("InvoiceTotal", out DocumentField? invoiceTotalField) &&
            invoice.Fields.TryGetValue("AmountDue", out DocumentField? amountDueField) &&
            invoiceTotalField.ValueCurrency != null &&
            amountDueField.ValueCurrency != null)
        {
            double paidAmount =
                invoiceTotalField.ValueCurrency.Amount - amountDueField.ValueCurrency.Amount;
            if (paidAmount != 0)
                paid = paidAmount.ToString(CultureInfo.InvariantCulture);
        }

        // IBAN — check inside the PaymentDetails array.
        string iban = string.Empty;
        if (invoice.Fields.TryGetValue("PaymentDetails", out DocumentField? paymentDetailsField) &&
            paymentDetailsField.ValueList is { Count: > 0 })
        {
            foreach (DocumentField detail in paymentDetailsField.ValueList)
            {
                if (detail.ValueDictionary != null &&
                    detail.ValueDictionary.TryGetValue("PaymentIban", out DocumentField? ibanField))
                {
                    iban = ibanField.Content ?? string.Empty;
                    if (!string.IsNullOrEmpty(iban)) break;
                }
            }
        }

        // ── Seller (vendor) ──────────────────────────────────────────────────
        string sellerNip    = GetString(invoice, "VendorTaxId");
        string sellerName   = GetString(invoice, "VendorName");
        string sellerCity   = GetAddressPart(invoice, "VendorAddress", a => a.City);
        string sellerStreet = GetAddressPart(invoice, "VendorAddress", a => a.StreetAddress);
        string sellerZip    = GetAddressPart(invoice, "VendorAddress", a => a.PostalCode);

        // ── Buyer (customer) ─────────────────────────────────────────────────
        string buyerNip    = GetString(invoice, "CustomerTaxId");
        string buyerName   = GetString(invoice, "CustomerName");
        string buyerCity   = GetAddressPart(invoice, "CustomerAddress", a => a.City);
        string buyerStreet = GetAddressPart(invoice, "CustomerAddress", a => a.StreetAddress);
        string buyerZip    = GetAddressPart(invoice, "CustomerAddress", a => a.PostalCode);

        // ── Line items ───────────────────────────────────────────────────────
        List<object> rows = BuildLineItems(invoice);

        // ── Table totals ─────────────────────────────────────────────────────
        string subTotal   = GetCurrencyAmount(invoice, "SubTotal");
        string totalTax   = GetCurrencyAmount(invoice, "TotalTax");
        string totalGross = GetCurrencyAmount(invoice, "InvoiceTotal");

        // ── Assemble schema ──────────────────────────────────────────────────
        var schema = new
        {
            data = new
            {
                invn,
                reference,
                issue,
                sale,
                payment,
                maturity,
                currency,
                total,
                paid,
                left,
                iban,
                seller = new
                {
                    nip     = sellerNip,
                    name    = sellerName,
                    city    = sellerCity,
                    street  = sellerStreet,
                    zipcode = sellerZip
                },
                buyer = new
                {
                    nip     = buyerNip,
                    name    = buyerName,
                    city    = buyerCity,
                    street  = buyerStreet,
                    zipcode = buyerZip
                },
                tables = new
                {
                    rows,
                    total = new[]
                    {
                        new
                        {
                            valNetto   = subTotal,
                            valVat     = totalTax,
                            valBrutto  = totalGross
                        }
                    }
                }
            }
        };

        return JsonSerializer.Serialize(schema);
    }

    // ── Field extraction helpers ─────────────────────────────────────────────

    private static string GetString(AnalyzedDocument doc, string key) =>
        doc.Fields.TryGetValue(key, out DocumentField? f) ? f.Content ?? string.Empty : string.Empty;

    private static string GetDate(AnalyzedDocument doc, string key)
    {
        if (!doc.Fields.TryGetValue(key, out DocumentField? f)) return string.Empty;
        if (f.ValueDate.HasValue) return f.ValueDate.Value.ToString("yyyy-MM-dd");
        return f.Content ?? string.Empty;
    }

    private static string GetCurrencyAmount(AnalyzedDocument doc, string key)
    {
        if (!doc.Fields.TryGetValue(key, out DocumentField? f)) return string.Empty;
        if (f.ValueCurrency != null)
            return f.ValueCurrency.Amount.ToString(CultureInfo.InvariantCulture);
        return f.Content ?? string.Empty;
    }

    private static string GetAddressPart(
        AnalyzedDocument doc,
        string key,
        Func<AddressValue, string?> selector)
    {
        if (!doc.Fields.TryGetValue(key, out DocumentField? f)) return string.Empty;
        return f.ValueAddress != null ? selector(f.ValueAddress) ?? string.Empty : string.Empty;
    }

    private static List<object> BuildLineItems(AnalyzedDocument invoice)
    {
        List<object> rows = [];

        if (!invoice.Fields.TryGetValue("Items", out DocumentField? itemsField) ||
            itemsField.ValueList is null)
        {
            return rows;
        }

        int index = 1;
        foreach (DocumentField item in itemsField.ValueList)
        {
            if (item.ValueDictionary is null) continue;

            string itemName = GetItemString(item, "Description");
            string itemQty  = GetItemNumber(item, "Quantity");
            string itemUnit = GetItemString(item, "Unit");
            string itemUnitPrice = GetItemCurrency(item, "UnitPrice");
            string itemTaxRate   = GetItemString(item, "TaxRate");
            // "Amount" on a line item is the line total (gross or net depending on locale).
            string itemAmount = GetItemCurrency(item, "Amount");

            rows.Add(new
            {
                no         = index.ToString(),
                name       = itemName,
                amount     = itemQty,
                unit       = itemUnit,
                priceNetto = itemUnitPrice,
                vat        = itemTaxRate,
                valNetto   = string.Empty,
                valVat     = string.Empty,
                valBrutto  = itemAmount
            });

            index++;
        }

        return rows;
    }

    private static string GetItemString(DocumentField item, string key) =>
        item.ValueDictionary != null && item.ValueDictionary.TryGetValue(key, out DocumentField? f)
            ? f.Content ?? string.Empty
            : string.Empty;

    private static string GetItemNumber(DocumentField item, string key)
    {
        if (item.ValueDictionary == null || !item.ValueDictionary.TryGetValue(key, out DocumentField? f))
            return string.Empty;
        if (f.ValueDouble.HasValue)
            return f.ValueDouble.Value.ToString(CultureInfo.InvariantCulture);
        return f.Content ?? string.Empty;
    }

    private static string GetItemCurrency(DocumentField item, string key)
    {
        if (item.ValueDictionary == null || !item.ValueDictionary.TryGetValue(key, out DocumentField? f))
            return string.Empty;
        if (f.ValueCurrency != null)
            return f.ValueCurrency.Amount.ToString(CultureInfo.InvariantCulture);
        return f.Content ?? string.Empty;
    }
}
