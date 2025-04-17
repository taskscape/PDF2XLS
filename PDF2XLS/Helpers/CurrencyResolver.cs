using System.Globalization;

namespace PDF2XLS.Helpers;

public static class CurrencyResolver
{
    private static readonly Dictionary<string?, string> AmbiguousSymbolFallback =
        new(StringComparer.OrdinalIgnoreCase)
    {
        { "$",  "USD" },
        { "¥",  "JPY" },
        { "kr", "SEK" },
        { "£",  "GBP" },
        { "₨",  "INR" },
        { "Rs", "INR" },
        { "lei","RON" },
        { "₩",  "KRW" },
        { "R$", "BRL" },
        { "KSh","KES" },
        { "Sh", "SOS" },
        { "NT$", "TWD" }
    };
    
    private static readonly Dictionary<string?, string> SymbolToIso = new(StringComparer.OrdinalIgnoreCase);

    static CurrencyResolver()
    {
        List<RegionInfo> allRegions = CultureInfo
            .GetCultures(CultureTypes.SpecificCultures)
            .Select(c => new RegionInfo(c.Name))
            .GroupBy(r => r.Name)
            .Select(g => g.First())
            .ToList();
        
        Dictionary<string?, HashSet<string>> symbolToIsoCandidates = new Dictionary<string?, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (RegionInfo region in allRegions)
        {
            string? symbol  = region.CurrencySymbol;
            string isoCode = region.ISOCurrencySymbol;

            if (!symbolToIsoCandidates.ContainsKey(symbol))
            {
                symbolToIsoCandidates[symbol] = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            }
            symbolToIsoCandidates[symbol].Add(isoCode);
        }
        
        foreach (KeyValuePair<string?, HashSet<string>> kvp in symbolToIsoCandidates)
        {
            string? symbol = kvp.Key;
            HashSet<string> isoCodes = kvp.Value;

            if (isoCodes.Count == 1)
            {
                SymbolToIso[symbol] = isoCodes.First();
            }
            else
            {
                if (AmbiguousSymbolFallback.TryGetValue(symbol, out string fallbackIso))
                {
                    SymbolToIso[symbol] = fallbackIso;
                }
            }
        }
    }
    
    public static string? GetIsoCurrencyCode(string? currencySymbol)
    {
        return string.IsNullOrWhiteSpace(currencySymbol) ? "" : SymbolToIso.GetValueOrDefault(currencySymbol, currencySymbol);
    }
}


