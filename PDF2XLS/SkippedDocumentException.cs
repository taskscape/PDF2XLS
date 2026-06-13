namespace PDF2XLS;

public sealed class SkippedDocumentException : Exception
{
    public SkippedDocumentException(string message)
        : base(message)
    {
    }

    public SkippedDocumentException(string message, Exception innerException)
        : base(message, innerException)
    {
    }
}
