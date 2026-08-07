using Microsoft.Extensions.Configuration;

namespace PDF2XLS.Tests;

public sealed class PromptFileLoaderTests
{
    [Fact]
    public void LoadAndRender_ReadsConfiguredFileAndInjectsSchema()
    {
        using TemporaryPromptDirectory temporaryDirectory = new();
        temporaryDirectory.WritePrompt("Before {schema} After");
        IConfiguration config = BuildConfiguration(temporaryDirectory.RelativePromptPath);

        string rendered = PromptFileLoader.LoadAndRender(
            config,
            "{\"type\":\"object\"}",
            temporaryDirectory.Path);

        Assert.Equal("Before {\"type\":\"object\"} After", rendered);
    }

    [Fact]
    public void GetValidationError_RejectsMissingPromptFile()
    {
        using TemporaryPromptDirectory temporaryDirectory = new();
        IConfiguration config = BuildConfiguration(temporaryDirectory.RelativePromptPath);

        string? error = PromptFileLoader.GetValidationError(config, temporaryDirectory.Path);

        Assert.Contains("does not exist", error);
    }

    [Theory]
    [InlineData("")]
    [InlineData("No schema placeholder")]
    [InlineData("{schema} and {schema}")]
    public void GetValidationError_RejectsInvalidPromptTemplate(string template)
    {
        using TemporaryPromptDirectory temporaryDirectory = new();
        temporaryDirectory.WritePrompt(template);
        IConfiguration config = BuildConfiguration(temporaryDirectory.RelativePromptPath);

        string? error = PromptFileLoader.GetValidationError(config, temporaryDirectory.Path);

        Assert.NotNull(error);
    }

    [Fact]
    public void ShippedPrompt_DefinesRequiredSynonymsAndPrecedence()
    {
        string promptPath = System.IO.Path.Combine(
            AppContext.BaseDirectory,
            "prompts",
            "invoice-extraction.txt");
        string prompt = File.ReadAllText(promptPath);

        Assert.Contains("Transaction date", prompt);
        Assert.Contains("Merchant", prompt);
        Assert.Contains("Merchant name", prompt);
        Assert.Contains("Invoice ID and Transaction ID are present", prompt);
        Assert.Contains("invoiceId first, then noteNumber, then transactionId", prompt);
        Assert.Contains("NOTA KSIĘGOWO-OBCIĄŻENIOWA NR:", prompt);
        Assert.Contains("Nota księgowa:", prompt);
        Assert.Contains("Nota obciążeniowa:", prompt);
        Assert.Equal(1, CountOccurrences(prompt, PromptFileLoader.SchemaPlaceholder));
    }

    [Fact]
    public void ConfigurationValidator_AcceptsReadablePromptFileWithOneSchemaPlaceholder()
    {
        using TemporaryPromptDirectory temporaryDirectory = new();
        temporaryDirectory.WritePrompt("Extract using {schema}");
        string serviceAccountPath = System.IO.Path.Combine(
            temporaryDirectory.Path,
            "service-account.json");
        File.WriteAllText(serviceAccountPath, "{}");

        IConfiguration config = new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                ["GoogleSheets:ServiceAccountFile"] = serviceAccountPath,
                ["GoogleSheets:SpreadsheetId"] = "spreadsheet-id",
                ["GoogleSheets:SheetName"] = "Sheet1",
                ["GoogleSheets:ApplicationName"] = "PDF2XLS.Tests",
                ["OpenAI:OpenAI_APIKey"] = "test-key",
                ["OpenAI:OpenAI_Model"] = "test-model",
                [PromptFileLoader.ConfigurationKey] = System.IO.Path.Combine(
                    temporaryDirectory.Path,
                    temporaryDirectory.RelativePromptPath)
            })
            .Build();

        List<string> errors = ConfigurationValidator.Validate(
            config,
            "OpenAIResponses",
            uploadEnabled: false);

        Assert.Empty(errors);
    }

    private static IConfiguration BuildConfiguration(string promptFile) =>
        new ConfigurationBuilder()
            .AddInMemoryCollection(new Dictionary<string, string?>
            {
                [PromptFileLoader.ConfigurationKey] = promptFile
            })
            .Build();

    private static int CountOccurrences(string value, string searchValue)
    {
        int count = 0;
        int position = 0;
        while ((position = value.IndexOf(searchValue, position, StringComparison.Ordinal)) >= 0)
        {
            count++;
            position += searchValue.Length;
        }

        return count;
    }

    private sealed class TemporaryPromptDirectory : IDisposable
    {
        public TemporaryPromptDirectory()
        {
            Path = System.IO.Path.Combine(
                System.IO.Path.GetTempPath(),
                "PDF2XLS.Tests",
                Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }
        public string RelativePromptPath => System.IO.Path.Combine("prompts", "test-prompt.txt");

        public void WritePrompt(string contents)
        {
            string promptPath = System.IO.Path.Combine(Path, RelativePromptPath);
            Directory.CreateDirectory(System.IO.Path.GetDirectoryName(promptPath)!);
            File.WriteAllText(promptPath, contents);
        }

        public void Dispose()
        {
            if (Directory.Exists(Path))
                Directory.Delete(Path, recursive: true);
        }
    }
}
