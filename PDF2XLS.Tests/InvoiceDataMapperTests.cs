using System.Text.Json.Nodes;

namespace PDF2XLS.Tests;

public sealed class InvoiceDataMapperTests
{
    [Fact]
    public void Map_PrefersInvoiceIdOverTransactionIdAndLegacyValue()
    {
        JsonNode root = JsonNode.Parse("""
            {
              "data": {
                "invoiceId": "INV-42",
                "transactionId": "TX-99",
                "invn": "TX-99"
              }
            }
            """)!;

        Dictionary<string, string?> result = InvoiceDataMapper.Map(root, Guid.Empty, null);

        Assert.Equal("'INV-42", result["InvoiceNumber"]);
    }

    [Fact]
    public void Map_PrefersNoteNumberOverTransactionId()
    {
        JsonNode root = JsonNode.Parse("""
            {
              "data": {
                "noteNumber": "NO/17/2026",
                "transactionId": "TX-99",
                "invn": "TX-99"
              }
            }
            """)!;

        Dictionary<string, string?> result = InvoiceDataMapper.Map(root, Guid.Empty, null);

        Assert.Equal("'NO/17/2026", result["InvoiceNumber"]);
    }

    [Fact]
    public void Map_UsesTransactionIdWhenNoPrimaryIdentifierExists()
    {
        JsonNode root = JsonNode.Parse("""
            {
              "data": {
                "transactionId": "TX-99",
                "invn": "INCORRECT-FALLBACK"
              }
            }
            """)!;

        Dictionary<string, string?> result = InvoiceDataMapper.Map(root, Guid.Empty, null);

        Assert.Equal("'TX-99", result["InvoiceNumber"]);
    }

    [Fact]
    public void Map_PreservesLegacyProviderInvoiceNumber()
    {
        JsonNode root = JsonNode.Parse("""
            {
              "data": {
                "invn": "LEGACY-123"
              }
            }
            """)!;

        Dictionary<string, string?> result = InvoiceDataMapper.Map(root, Guid.Empty, null);

        Assert.Equal("'LEGACY-123", result["InvoiceNumber"]);
    }
}
