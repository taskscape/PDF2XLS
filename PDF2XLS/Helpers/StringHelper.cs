namespace PDF2XLS.Helpers;

public static class StringHelper
{
    public static string RemoveLetters(string input)
    {
        string allowedChars = "0123456789.,";
        string cleaned = new string(input.Where(c => allowedChars.Contains(c)).ToArray());
        return cleaned;
    }

    public static bool IsValidHttpUrl(string url)
    {
        if (!Uri.TryCreate(url, UriKind.Absolute, out Uri uriResult))
        {
            return false;
        }
        
        return uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps;
    }
}