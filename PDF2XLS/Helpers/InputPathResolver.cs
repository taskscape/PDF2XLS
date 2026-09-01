namespace PDF2XLS.Helpers;

public static class InputPathResolver
{
    /// <summary>
    /// Resolves CLI input paths to one or more PDF files.
    /// Accepts a single .pdf file, multiple .pdf files, or a directory (top-level .pdf files only).
    /// When multiple paths are supplied, directories are expanded inline and duplicate files are skipped.
    /// A missing .pdf whose sibling was already renamed to .skp is reported as already skipped.
    /// </summary>
    public static InputPathResult Resolve(IEnumerable<string?> inputPaths)
    {
        List<string> normalizedPaths = inputPaths
            .Where(p => !string.IsNullOrWhiteSpace(p))
            .Select(p => p!.Trim().Trim('"'))
            .ToList();

        if (normalizedPaths.Count == 0)
            return InputPathResult.Failure("No input path was provided.", InputPathFailureKind.InvalidPath);

        if (normalizedPaths.Count == 1 && Directory.Exists(normalizedPaths[0]))
            return ResolveDirectory(normalizedPaths[0]);

        List<string> pdfFiles = [];
        List<string> skippedFiles = [];
        HashSet<string> seenFiles = new(StringComparer.OrdinalIgnoreCase);
        HashSet<string> seenSkippedFiles = new(StringComparer.OrdinalIgnoreCase);

        foreach (string path in normalizedPaths)
        {
            if (Directory.Exists(path))
            {
                List<string> directoryFiles = GetPdfFilesFromDirectory(path);
                if (directoryFiles.Count == 0)
                {
                    return InputPathResult.Failure(
                        $"No PDF files found in folder: {path}",
                        InputPathFailureKind.EmptyDirectory);
                }

                AddUniqueFiles(pdfFiles, seenFiles, directoryFiles);
                continue;
            }

            if (File.Exists(path))
            {
                if (!IsPdfFile(path))
                {
                    return InputPathResult.Failure(
                        $"File {path} is not a PDF file.",
                        InputPathFailureKind.NotPdf);
                }

                AddUniqueFiles(pdfFiles, seenFiles, [path]);
                continue;
            }

            string? skippedPath = FindSkippedSibling(path);
            if (skippedPath != null)
            {
                if (seenSkippedFiles.Add(skippedPath))
                    skippedFiles.Add(skippedPath);
                continue;
            }

            return InputPathResult.Failure($"Path does not exist: {path}", InputPathFailureKind.InvalidPath);
        }

        if (pdfFiles.Count == 0)
        {
            return InputPathResult.Failure(
                FormatAlreadySkippedMessage(skippedFiles),
                InputPathFailureKind.AlreadySkipped);
        }

        return InputPathResult.FromFiles(
            pdfFiles,
            summaryPath: string.Join("; ", normalizedPaths),
            isDirectory: false,
            isMultipleInputs: normalizedPaths.Count > 1);
    }

    private static InputPathResult ResolveDirectory(string directoryPath)
    {
        List<string> pdfFiles = GetPdfFilesFromDirectory(directoryPath);
        if (pdfFiles.Count == 0)
        {
            return InputPathResult.Failure(
                $"No PDF files found in folder: {directoryPath}",
                InputPathFailureKind.EmptyDirectory);
        }

        return InputPathResult.FromFiles(
            pdfFiles,
            summaryPath: directoryPath,
            isDirectory: true,
            isMultipleInputs: false);
    }

    private static List<string> GetPdfFilesFromDirectory(string directoryPath) =>
        Directory
            .GetFiles(directoryPath, "*.*", SearchOption.TopDirectoryOnly)
            .Where(IsPdfFile)
            .OrderBy(f => f, StringComparer.OrdinalIgnoreCase)
            .ToList();

    private static void AddUniqueFiles(List<string> target, HashSet<string> seen, IEnumerable<string> candidates)
    {
        foreach (string candidate in candidates)
        {
            if (seen.Add(candidate))
                target.Add(candidate);
        }
    }

    private static bool IsPdfFile(string filePath) =>
        string.Equals(Path.GetExtension(filePath), ".pdf", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Finds a sibling created by <c>TryMarkFileAsSkipped</c>: <c>{path}.skp</c>, then
    /// <c>{filename} (1).skp</c> … <c>{filename} (999).skp</c>.
    /// </summary>
    private static string? FindSkippedSibling(string path)
    {
        string preferred = $"{path}.skp";
        if (File.Exists(preferred))
            return preferred;

        string? directory = Path.GetDirectoryName(path);
        string fileName = Path.GetFileName(path);
        if (string.IsNullOrEmpty(fileName))
            return null;

        for (int i = 1; i <= 999; i++)
        {
            string numberedName = $"{fileName} ({i}).skp";
            string candidate = string.IsNullOrEmpty(directory)
                ? numberedName
                : Path.Combine(directory, numberedName);
            if (File.Exists(candidate))
                return candidate;
        }

        return null;
    }

    private static string FormatAlreadySkippedMessage(IReadOnlyList<string> skippedFiles) =>
        skippedFiles.Count == 1
            ? $"File was already marked as skipped: {skippedFiles[0]}"
            : $"Files were already marked as skipped: {string.Join("; ", skippedFiles)}";
}

public enum InputPathFailureKind
{
    InvalidPath,
    NotPdf,
    EmptyDirectory,
    AlreadySkipped
}

public sealed class InputPathResult
{
    public bool IsSuccess { get; private init; }
    public IReadOnlyList<string> Files { get; private init; } = [];
    public string InputPath { get; private init; } = string.Empty;
    public bool IsDirectory { get; private init; }
    public bool IsMultipleInputs { get; private init; }
    public string? ErrorMessage { get; private init; }
    public InputPathFailureKind? FailureKind { get; private init; }

    public static InputPathResult FromFiles(
        IReadOnlyList<string> pdfFiles,
        string summaryPath,
        bool isDirectory,
        bool isMultipleInputs) => new()
    {
        IsSuccess = true,
        Files = pdfFiles,
        InputPath = summaryPath,
        IsDirectory = isDirectory,
        IsMultipleInputs = isMultipleInputs
    };

    public static InputPathResult Failure(string errorMessage, InputPathFailureKind kind) => new()
    {
        IsSuccess = false,
        ErrorMessage = errorMessage,
        FailureKind = kind
    };
}
