using System.Text;
using System.Text.RegularExpressions;

namespace PDF2XLS;

public static class PdfPageCounter
{
    private static readonly Regex PageObjectRegex = new(
        @"/Type\s*/Page(?!s)\b",
        RegexOptions.CultureInvariant | RegexOptions.Compiled);

    public static int CountPages(string filePath)
    {
        byte[] pdfBytes = File.ReadAllBytes(filePath);
        string pdfText = Encoding.Latin1.GetString(pdfBytes);

        int pageCount = PageObjectRegex.Matches(pdfText).Count;
        if (pageCount <= 0)
        {
            throw new InvalidOperationException(
                $"Could not determine PDF page count safely. Azure submission skipped. File: {filePath}");
        }

        return pageCount;
    }
}
