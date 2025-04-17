using System.Text.RegularExpressions;

namespace PDF2XLS.Helpers;

public static class StringHelper
{
    public static string RemoveLetters(string input)
    {
        string allowedChars = "0123456789.,";
        string cleaned = new(input.Where(c => allowedChars.Contains(c)).ToArray());
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

    public static string AbbreviateCompanyType(string companyName)
    {
        if (string.IsNullOrWhiteSpace(companyName))
            return companyName;
        
        string result = companyName.Trim();

        result = Regex.Replace(
            result,
            @"\bsp\.\s*z\.?\s*o\.?o\.?(?!\w)",
            "sp. z o.o.",
            RegexOptions.IgnoreCase
        );
        
        Dictionary<string, string> polishReplacements = new(StringComparer.OrdinalIgnoreCase)
        {
            { @"\bsp[oó][łl]ka akcyjna\b", "S.A." },
            { @"\bsp[oó][łl]ka z ograniczon[ąa] odpowiedzialno[śs]ci[ąa]\b", "sp. z o.o." },
            { @"\bsp[oó][łl]ka komandytowa\b", "sp. k." },
            { @"\bsp[oó][łl]ka jawna\b", "sp. j." },
            { @"\bsp[oó][łl]ka partnerska\b", "sp. p." },
            { @"\bsp[oó][łl]ka komandytowo[- ]akcyjna\b", "sp.k.a." },
            { @"\bsp[oó][łl]ka cywilna\b", "s.c." }
        };

        result = polishReplacements.Aggregate(result, (current, rule) => Regex.Replace(current, rule.Key, rule.Value, RegexOptions.IgnoreCase));

        Dictionary<string, string> otherReplacements = new(StringComparer.OrdinalIgnoreCase)
        {
            { @"\blimited\b", "Ltd" },
            { @"\bincorporated\b", "Inc" },
            { @"\bcorporation\b", "Corp" },
            { @"\bcompany\b", "Co." },
            { @"\bpublic limited company\b", "PLC" },
            { @"\bgesellschaft mit beschr[aä]nkter haftung\b", "GmbH" },
            { @"\baktiengesellschaft\b", "AG" }
        };

        foreach (KeyValuePair<string, string> rule in otherReplacements)
        {
            string pattern = rule.Key + @"(?=\s*$)";
            result = Regex.Replace(result, pattern, rule.Value, RegexOptions.IgnoreCase);
        }

        return result;
    }
}