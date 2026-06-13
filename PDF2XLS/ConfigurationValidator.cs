using Microsoft.Extensions.Configuration;

namespace PDF2XLS;

public static class ConfigurationValidator
{
    /// <summary>
    /// Validates that all configuration keys required by the selected workflow are present.
    /// Returns a list of human-readable error messages; an empty list means the config is valid.
    /// </summary>
    public static List<string> Validate(IConfiguration config, string preferredApi, bool uploadEnabled)
    {
        List<string> errors = [];

        if (string.IsNullOrWhiteSpace(preferredApi))
        {
            errors.Add("PreferredAPI is required");
            return errors;
        }

        // Common fields required by every workflow.
        string? serviceAccountFile = config["GoogleSheets:ServiceAccountFile"];
        if (string.IsNullOrWhiteSpace(serviceAccountFile))
            errors.Add("GoogleSheets:ServiceAccountFile is required");
        else if (!File.Exists(serviceAccountFile))
            errors.Add($"GoogleSheets:ServiceAccountFile points to a file that does not exist: {serviceAccountFile}");
        if (string.IsNullOrWhiteSpace(config["GoogleSheets:SpreadsheetId"]))
            errors.Add("GoogleSheets:SpreadsheetId is required");
        if (string.IsNullOrWhiteSpace(config["GoogleSheets:SheetName"]))
            errors.Add("GoogleSheets:SheetName is required");
        if (string.IsNullOrWhiteSpace(config["GoogleSheets:ApplicationName"]))
            errors.Add("GoogleSheets:ApplicationName is required");

        if (uploadEnabled && string.IsNullOrWhiteSpace(config["UploadPDF:PDF2URLPath"]))
            errors.Add("UploadPDF:PDF2URLPath is required when UploadPDF:Enabled is true");

        if (!string.IsNullOrWhiteSpace(config["Seq:ServerAddress"]) &&
            string.IsNullOrWhiteSpace(config["Seq:ApiKey"]))
            errors.Add("Seq:ApiKey is required when Seq:ServerAddress is configured");

        // Workflow-specific required fields.
        switch (preferredApi)
        {
            case "NuDelta":
                if (string.IsNullOrWhiteSpace(config["NuDeltaCredentials:Username"]))
                    errors.Add("NuDeltaCredentials:Username is required for the NuDelta workflow");
                if (string.IsNullOrWhiteSpace(config["NuDeltaCredentials:Password"]))
                    errors.Add("NuDeltaCredentials:Password is required for the NuDelta workflow");
                break;

            case "OpenAIResponses":
                if (string.IsNullOrWhiteSpace(config["OpenAI:OpenAI_APIKey"]))
                    errors.Add("OpenAI:OpenAI_APIKey is required for the OpenAIResponses workflow");
                if (string.IsNullOrWhiteSpace(config["OpenAI:OpenAI_Model"]))
                    errors.Add("OpenAI:OpenAI_Model is required for the OpenAIResponses workflow");
                if (string.IsNullOrWhiteSpace(config["OpenAI:Prompt"]))
                    errors.Add("OpenAI:Prompt is required for the OpenAIResponses workflow");
                break;

            case "AzureDocumentIntelligence":
                if (string.IsNullOrWhiteSpace(config["AzureDocumentIntelligence:Endpoint"]))
                    errors.Add("AzureDocumentIntelligence:Endpoint is required for the AzureDocumentIntelligence workflow");
                if (string.IsNullOrWhiteSpace(config["AzureDocumentIntelligence:ApiKey"]))
                    errors.Add("AzureDocumentIntelligence:ApiKey is required for the AzureDocumentIntelligence workflow");
                string? monthlyPageLimit = config["AzureDocumentIntelligence:MonthlyPageLimit"];
                if (!string.IsNullOrWhiteSpace(monthlyPageLimit) &&
                    (!int.TryParse(monthlyPageLimit, out int parsedLimit) || parsedLimit < 0))
                {
                    errors.Add("AzureDocumentIntelligence:MonthlyPageLimit must be a non-negative integer. Use 0 or leave empty to disable the internal quota guard.");
                }
                break;

            default:
                errors.Add(
                    $"PreferredAPI '{preferredApi}' is not valid. " +
                    "Supported values: NuDelta, OpenAIResponses, AzureDocumentIntelligence");
                break;
        }

        return errors;
    }
}
