using System;
using System.IO;
using System.Text;

internal static class UiLanguage
{
    private static readonly string FilePath = Path.Combine(AppContext.BaseDirectory, "ui-language.txt");
    private static string? current;
    public static string Current
    {
        get
        {
            if (current != null) return current;
            try { current = File.Exists(FilePath) && File.ReadAllText(FilePath, Encoding.UTF8).Trim().Equals("ja", StringComparison.OrdinalIgnoreCase) ? "ja" : "vi"; }
            catch { current = "vi"; }
            return current;
        }
    }
    public static string T(string vi, string ja) => Current == "ja" ? ja : vi;
    public static bool Set(string language)
    {
        if (language != "vi" && language != "ja") return false;
        try
        {
            File.WriteAllText(FilePath, language + Environment.NewLine, new UTF8Encoding(false));
            current = language;
            return true;
        }
        catch { return false; }
    }
}
