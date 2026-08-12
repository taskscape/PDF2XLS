using Microsoft.Extensions.Configuration;

namespace PDF2XLS;

/// <summary>
/// Loads and validates the OpenAI extraction prompt from a file next to the application
/// (or from an explicitly configured absolute path).
/// </summary>
public static class PromptFileLoader
{
    public const string ConfigurationKey = "OpenAI:PromptFile";
    public const string SchemaPlaceholder = "{schema}";

    public static string LoadAndRender(
        IConfiguration config,
        string responseSchema,
        string? baseDirectory = null)
    {
        string template = LoadTemplate(config, baseDirectory);
        return template.Replace(SchemaPlaceholder, responseSchema, StringComparison.Ordinal);
    }

    public static string? GetValidationError(
        IConfiguration config,
        string? baseDirectory = null)
    {
        try
        {
            _ = LoadTemplate(config, baseDirectory);
            return null;
        }
        catch (InvalidOperationException ex)
        {
            return ex.Message;
        }
    }

    private static string LoadTemplate(IConfiguration config, string? baseDirectory)
    {
        string? configuredPath = config[ConfigurationKey];
        if (string.IsNullOrWhiteSpace(configuredPath))
            throw new InvalidOperationException($"{ConfigurationKey} is required for the OpenAIResponses workflow");

        string promptPath;
        try
        {
            string root = baseDirectory ?? AppContext.BaseDirectory;
            if (!Path.IsPathFullyQualified(root))
            {
                throw new InvalidOperationException(
                    $"Prompt base directory must be fully qualified: {root}");
            }

            bool isRooted = Path.IsPathRooted(configuredPath);
            bool isFullyQualified = Path.IsPathFullyQualified(configuredPath);
            if (isRooted && !isFullyQualified)
            {
                throw new InvalidOperationException(
                    $"{ConfigurationKey} must be either a relative path or a fully qualified path: {configuredPath}");
            }

            promptPath = isFullyQualified
                ? Path.GetFullPath(configuredPath)
                : Path.GetFullPath(configuredPath, root);
        }
        catch (Exception ex) when (ex is ArgumentException or NotSupportedException or PathTooLongException)
        {
            throw new InvalidOperationException(
                $"{ConfigurationKey} is not a valid file path: {configuredPath}", ex);
        }

        if (!File.Exists(promptPath))
            throw new InvalidOperationException($"{ConfigurationKey} points to a file that does not exist: {promptPath}");

        string template;
        try
        {
            template = File.ReadAllText(promptPath);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new InvalidOperationException($"{ConfigurationKey} could not be read: {promptPath}", ex);
        }

        if (string.IsNullOrWhiteSpace(template))
            throw new InvalidOperationException($"{ConfigurationKey} points to an empty prompt file: {promptPath}");

        int firstPlaceholder = template.IndexOf(SchemaPlaceholder, StringComparison.Ordinal);
        if (firstPlaceholder < 0)
            throw new InvalidOperationException(
                $"{ConfigurationKey} prompt must contain the {SchemaPlaceholder} placeholder exactly once: {promptPath}");

        int secondPlaceholder = template.IndexOf(
            SchemaPlaceholder,
            firstPlaceholder + SchemaPlaceholder.Length,
            StringComparison.Ordinal);
        if (secondPlaceholder >= 0)
            throw new InvalidOperationException(
                $"{ConfigurationKey} prompt must contain the {SchemaPlaceholder} placeholder exactly once: {promptPath}");

        return template;
    }
}
