using System.Text.Json.Nodes;
using PDF2XLS.Helpers;

namespace PDF2XLS;

public static class InvoiceDataMapper
{
    /// <summary>
    /// Maps a JSON result node (in the application's internal schema) to the flat
    /// string dictionary used by GSheets.AppendRowWithBatchUpdate.
    /// </summary>
    public static Dictionary<string, string?> Map(JsonNode? root, Guid runId, string? documentLink)
    {
        JsonNode? dataNode = root?["data"];

        // Top-level scalar fields
        string invNumber       = GetValFromNode(dataNode?["invn"]);
        string refNumber       = GetValFromNode(dataNode?["reference"]);
        string issueDateString = GetValFromNode(dataNode?["issue"]);
        DateTime.TryParse(issueDateString, out DateTime issueDate);
        issueDateString = issueDate.ToString("yyyy-MM-dd");
        string saleDateString  = GetValFromNode(dataNode?["sale"]);
        DateTime.TryParse(saleDateString, out DateTime saleDate);
        saleDateString  = saleDate.ToString("yyyy-MM-dd");
        string  paymentMethod  = GetValFromNode(dataNode?["payment"]);
        string  maturity       = GetValFromNode(dataNode?["maturity"]);
        string? currency       = CurrencyResolver.GetIsoCurrencyCode(GetValFromNode(dataNode?["currency"]));
        string? totalAmount    = StringHelper.RemoveLetters(GetValFromNode(dataNode?["total"]));
        string? paidAmount     = StringHelper.RemoveLetters(GetValFromNode(dataNode?["paid"]));
        string? leftToPay      = StringHelper.RemoveLetters(GetValFromNode(dataNode?["left"]));
        string  iban           = GetValFromNode(dataNode?["iban"]);

        // Seller
        JsonNode? seller     = dataNode?["seller"];
        string  sellerNip    = GetValFromNode(seller?["nip"]);
        string? sellerName   = StringHelper.AbbreviateCompanyType(GetValFromNode(seller?["name"]));
        string  sellerCity   = GetValFromNode(seller?["city"]);
        string  sellerStreet = GetValFromNode(seller?["street"]);
        string  sellerZip    = GetValFromNode(seller?["zipcode"]);

        // Buyer
        JsonNode? buyer     = dataNode?["buyer"];
        string  buyerNip    = GetValFromNode(buyer?["nip"]);
        string? buyerName   = StringHelper.AbbreviateCompanyType(GetValFromNode(buyer?["name"]));
        string  buyerCity   = GetValFromNode(buyer?["city"]);
        string  buyerStreet = GetValFromNode(buyer?["street"]);
        string  buyerZip    = GetValFromNode(buyer?["zipcode"]);

        // Table totals
        JsonNode? tablesNode = dataNode?["tables"];
        JsonArray totals     = tablesNode?["total"]?.AsArray() ?? [];
        JsonNode? totalNode  = totals.Count > 0 ? totals[0] : null;
        string totalNet      = GetValFromNode(totalNode?["valNetto"]);
        string totalVat      = GetValFromNode(totalNode?["valVat"]);
        string totalGross    = GetValFromNode(totalNode?["valBrutto"]);

        // Prefix invoice/reference numbers with a quote so Google Sheets treats them as text.
        if (!string.IsNullOrEmpty(invNumber))  invNumber  = string.Concat("\'", invNumber);
        if (!string.IsNullOrEmpty(refNumber))  refNumber  = string.Concat("\'", refNumber);

        return new Dictionary<string, string?>
        {
            { "InvoiceNumber",   invNumber },
            { "ReferenceNumber", refNumber },
            { "IssueDate",       issueDateString },
            { "SaleDate",        saleDateString },
            { "PaymentMethod",   paymentMethod },
            { "Maturity",        maturity },
            { "Currency",        currency },
            { "TotalAmount",     totalAmount },
            { "PaidAmount",      paidAmount },
            { "AmountLeftToPay", leftToPay },
            { "IBAN",            iban },
            { "SellerNIP",       sellerNip },
            { "SellerName",      sellerName },
            { "SellerCity",      sellerCity },
            { "SellerStreet",    sellerStreet },
            { "SellerZIP",       sellerZip },
            { "BuyerNIP",        buyerNip },
            { "BuyerName",       buyerName },
            { "BuyerCity",       buyerCity },
            { "BuyerStreet",     buyerStreet },
            { "BuyerZIP",        buyerZip },
            { "DocumentLink",    documentLink },
            { "TotalNet",        totalNet },
            { "TotalVat",        totalVat },
            { "TotalGross",      totalGross },
            { "RunID",           runId.ToString() }
        };
    }

    /// <summary>
    /// Returns the string value of a JSON node.
    /// Handles NuDelta's wrapped {"ans": {"val": "..."}} format as well as plain values.
    /// </summary>
    public static string GetValFromNode(JsonNode? node)
    {
        switch (node)
        {
            case null:
                return "";
            case JsonValue:
                return node.ToString();
        }

        JsonNode? ansNode = node["ans"];
        if (ansNode?["val"] != null)
            return ansNode["val"]?.ToString() ?? "";

        return node.ToString();
    }
}
