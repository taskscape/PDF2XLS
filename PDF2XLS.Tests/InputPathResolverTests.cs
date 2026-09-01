using PDF2XLS.Helpers;

namespace PDF2XLS.Tests;

public sealed class InputPathResolverTests
{
    [Fact]
    public void Resolve_ReportsMissingPathWhenNoSkippedSiblingExists()
    {
        using TemporaryPdfDirectory directory = new();
        string missingPdf = directory.File("missing.pdf");

        InputPathResult result = InputPathResolver.Resolve([missingPdf]);

        Assert.False(result.IsSuccess);
        Assert.Equal(InputPathFailureKind.InvalidPath, result.FailureKind);
        Assert.Equal($"Path does not exist: {missingPdf}", result.ErrorMessage);
    }

    [Fact]
    public void Resolve_ReportsAlreadySkippedWhenPreferredSkpSiblingExists()
    {
        using TemporaryPdfDirectory directory = new();
        string pdfPath = directory.File("invoice.pdf");
        string skippedPath = directory.Write("invoice.pdf.skp", "skipped");

        InputPathResult result = InputPathResolver.Resolve([pdfPath]);

        Assert.False(result.IsSuccess);
        Assert.Equal(InputPathFailureKind.AlreadySkipped, result.FailureKind);
        Assert.Equal($"File was already marked as skipped: {skippedPath}", result.ErrorMessage);
    }

    [Fact]
    public void Resolve_ReportsAlreadySkippedWhenNumberedSkpSiblingExists()
    {
        using TemporaryPdfDirectory directory = new();
        string pdfPath = directory.File("invoice.pdf");
        string skippedPath = directory.Write("invoice.pdf (1).skp", "skipped");

        InputPathResult result = InputPathResolver.Resolve([pdfPath]);

        Assert.False(result.IsSuccess);
        Assert.Equal(InputPathFailureKind.AlreadySkipped, result.FailureKind);
        Assert.Equal($"File was already marked as skipped: {skippedPath}", result.ErrorMessage);
    }

    [Fact]
    public void Resolve_ProcessesExistingPdfEvenWhenSkpSiblingExists()
    {
        using TemporaryPdfDirectory directory = new();
        string pdfPath = directory.Write("invoice.pdf", "pdf");
        directory.Write("invoice.pdf.skp", "skipped");

        InputPathResult result = InputPathResolver.Resolve([pdfPath]);

        Assert.True(result.IsSuccess);
        Assert.Equal([pdfPath], result.Files);
    }

    [Fact]
    public void Resolve_IgnoresAlreadySkippedFileWhenAnotherPdfRemains()
    {
        using TemporaryPdfDirectory directory = new();
        string remainingPdf = directory.Write("keep.pdf", "pdf");
        string skippedPdf = directory.File("skip.pdf");
        directory.Write("skip.pdf.skp", "skipped");

        InputPathResult result = InputPathResolver.Resolve([skippedPdf, remainingPdf]);

        Assert.True(result.IsSuccess);
        Assert.Equal([remainingPdf], result.Files);
    }

    [Fact]
    public void Resolve_ReportsAllAlreadySkippedFilesWhenNoneRemain()
    {
        using TemporaryPdfDirectory directory = new();
        string firstPdf = directory.File("one.pdf");
        string secondPdf = directory.File("two.pdf");
        string firstSkipped = directory.Write("one.pdf.skp", "skipped");
        string secondSkipped = directory.Write("two.pdf.skp", "skipped");

        InputPathResult result = InputPathResolver.Resolve([firstPdf, secondPdf]);

        Assert.False(result.IsSuccess);
        Assert.Equal(InputPathFailureKind.AlreadySkipped, result.FailureKind);
        Assert.Equal(
            $"Files were already marked as skipped: {firstSkipped}; {secondSkipped}",
            result.ErrorMessage);
    }

    private sealed class TemporaryPdfDirectory : IDisposable
    {
        public TemporaryPdfDirectory()
        {
            Path = System.IO.Path.Join(
                System.IO.Path.GetTempPath(),
                "PDF2XLS.Tests",
                Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(Path);
        }

        public string Path { get; }

        public string File(string fileName) => System.IO.Path.Join(Path, fileName);

        public string Write(string fileName, string contents)
        {
            string path = File(fileName);
            System.IO.File.WriteAllText(path, contents);
            return path;
        }

        public void Dispose()
        {
            if (Directory.Exists(Path))
                Directory.Delete(Path, recursive: true);
        }
    }
}
