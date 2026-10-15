using PdfSharp.Fonts;

namespace ConverterReport.Core.Pdf;

/// <summary>
/// Поиск шрифтов для PDF. На Windows используются системные Times New Roman / Arial,
/// на других ОС — метрически совместимые Liberation (или DejaVu). Шрифты встраиваются в PDF.
/// </summary>
public sealed class SystemFontResolver : IFontResolver
{
    // Для каждого семейства: файлы для обычного, полужирного, курсива, полужирного курсива (в порядке предпочтения).
    private static readonly Dictionary<string, string[][]> Families = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Times New Roman"] = new[]
        {
            new[] { "times.ttf", "LiberationSerif-Regular.ttf", "DejaVuSerif.ttf" },
            new[] { "timesbd.ttf", "LiberationSerif-Bold.ttf", "DejaVuSerif-Bold.ttf" },
            new[] { "timesi.ttf", "LiberationSerif-Italic.ttf", "DejaVuSerif-Italic.ttf" },
            new[] { "timesbi.ttf", "LiberationSerif-BoldItalic.ttf", "DejaVuSerif-BoldItalic.ttf" },
        },
        ["Arial"] = new[]
        {
            new[] { "arial.ttf", "LiberationSans-Regular.ttf", "DejaVuSans.ttf" },
            new[] { "arialbd.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf" },
            new[] { "ariali.ttf", "LiberationSans-Italic.ttf", "DejaVuSans-Oblique.ttf" },
            new[] { "arialbi.ttf", "LiberationSans-BoldItalic.ttf", "DejaVuSans-BoldOblique.ttf" },
        },
        ["Calibri"] = new[]
        {
            new[] { "calibri.ttf", "Carlito-Regular.ttf", "arial.ttf", "LiberationSans-Regular.ttf", "DejaVuSans.ttf" },
            new[] { "calibrib.ttf", "Carlito-Bold.ttf", "arialbd.ttf", "LiberationSans-Bold.ttf", "DejaVuSans-Bold.ttf" },
            new[] { "calibrii.ttf", "Carlito-Italic.ttf", "ariali.ttf", "LiberationSans-Italic.ttf", "DejaVuSans-Oblique.ttf" },
            new[] { "calibriz.ttf", "Carlito-BoldItalic.ttf", "arialbi.ttf", "LiberationSans-BoldItalic.ttf", "DejaVuSans-BoldOblique.ttf" },
        },
    };

    private const string FallbackFamily = "Arial";

    private static readonly Lazy<Dictionary<string, string>> FontFiles = new(IndexFontFiles);
    private readonly Dictionary<string, byte[]> _cache = new(StringComparer.OrdinalIgnoreCase);
    private readonly object _lock = new();

    public static void Register()
    {
        if (GlobalFontSettings.FontResolver is not SystemFontResolver)
            GlobalFontSettings.FontResolver = new SystemFontResolver();
    }

    public FontResolverInfo? ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        var style = (isBold ? 1 : 0) + (isItalic ? 2 : 0);
        var faces = Families.TryGetValue(familyName.Trim(), out var f) ? f : Families[FallbackFamily];

        var file = Find(faces[style]);
        if (file is not null)
            return new FontResolverInfo(file);

        // Нет нужного начертания — имитируем на основе обычного.
        var regular = Find(faces[0]) ?? Find(Families[FallbackFamily][0]);
        if (regular is not null)
            return new FontResolverInfo(regular, isBold, isItalic);

        throw new InvalidOperationException(
            $"Не найден шрифт \"{familyName}\". Установите Times New Roman/Arial (Windows) или Liberation/DejaVu (Linux) " +
            "либо положите файлы .ttf в папку fonts рядом с программой.");
    }

    public byte[]? GetFont(string faceName)
    {
        lock (_lock)
        {
            if (!_cache.TryGetValue(faceName, out var data))
            {
                data = File.ReadAllBytes(FontFiles.Value[faceName]);
                _cache[faceName] = data;
            }

            return data;
        }
    }

    private static string? Find(IEnumerable<string> candidates) =>
        candidates.FirstOrDefault(c => FontFiles.Value.ContainsKey(c));

    private static Dictionary<string, string> IndexFontFiles()
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var dir in FontDirectories().Where(Directory.Exists))
        {
            try
            {
                foreach (var file in Directory.EnumerateFiles(dir, "*.tt?", SearchOption.AllDirectories))
                    result.TryAdd(Path.GetFileName(file), file);
            }
            catch (Exception)
            {
                // Нет доступа к каталогу — пропускаем.
            }
        }

        return result;
    }

    private static IEnumerable<string> FontDirectories()
    {
        yield return Path.Combine(AppContext.BaseDirectory, "fonts");
        yield return Environment.GetFolderPath(Environment.SpecialFolder.Fonts);
        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        if (!string.IsNullOrEmpty(local))
            yield return Path.Combine(local, "Microsoft", "Windows", "Fonts");
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        if (!string.IsNullOrEmpty(home))
        {
            yield return Path.Combine(home, ".fonts");
            yield return Path.Combine(home, ".local", "share", "fonts");
        }

        yield return "/usr/share/fonts";
        yield return "/usr/local/share/fonts";
        yield return "/Library/Fonts";
        yield return "/System/Library/Fonts";
    }
}
