using ClosedXML.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ConverterReport.Core.Archive;

namespace ConverterReport.Core.Workbook;

/// <summary>
/// Старая книга отчета (которую вела версия 1.2.2): перенос листов прошлых дней в общую книгу и сроки этапов
/// из последнего листа. Старая книга только читается (даже если открыта в Excel).
/// </summary>
public static class LegacyWorkbook
{
    /// <summary>Имена листов «ДД.ММ» без загрузки книги (быстро — только список листов).</summary>
    public static List<string> DailySheetNames(string path)
    {
        using var stream = OpenRead(path);
        using var doc = SpreadsheetDocument.Open(stream, false);
        return doc.WorkbookPart?.Workbook.Descendants<Sheet>()
                   .Select(s => s.Name?.Value ?? "")
                   .Where(SummaryWorkbookBuilder.IsDailySheet)
                   .ToList()
               ?? new List<string>();
    }

    public static XLWorkbook Load(string path)
    {
        using var stream = OpenRead(path);
        return new XLWorkbook(stream);
    }

    /// <summary>
    /// Копирует листы <paramref name="names"/> из старой книги в общую: значения (формулы старой книги ссылаются
    /// на соседние листы — переносятся вычисленные значения), оформление, скрытые строки, область печати.
    /// Цвета темы переводятся в явные: в книгах разные темы, иначе цвета на листах изменились бы.
    /// </summary>
    public static int CopySheets(XLWorkbook legacy, XLWorkbook master, IEnumerable<string> names)
    {
        var theme = legacy.Theme;
        var copied = 0;
        foreach (var name in names)
        {
            if (!legacy.TryGetWorksheet(name, out var source) || master.TryGetWorksheet(name, out _))
                continue;
            var copy = source.CopyTo(master, name);
            foreach (var cell in copy.CellsUsed(XLCellsUsedOptions.All))
            {
                if (cell.HasFormula)
                {
                    var value = Computed(source.Cell(cell.Address.RowNumber, cell.Address.ColumnNumber));
                    cell.FormulaA1 = null;
                    cell.Value = value;
                }

                ResolveThemeColors(cell.Style, theme);
            }

            copy.Visibility = XLWorksheetVisibility.Visible;
            copy.TabSelected = false;
            copy.TabActive = false;
            copied++;
        }

        return copied;
    }

    /// <summary>
    /// Сроки этапов из раздела «4. Сроки» листа дня: номер этапа → стадия → [начало, окончание] или [текст].
    /// </summary>
    internal static Dictionary<string, Dictionary<string, List<string>>> ReadDeadlines(XlsxSheet ws)
    {
        var result = new Dictionary<string, Dictionary<string, List<string>>>(StringComparer.OrdinalIgnoreCase);
        var last = ws.Values.LastRow;
        var title = Enumerable.Range(1, last).FirstOrDefault(r => ws.Text(r, 1).StartsWith("4. Сроки", StringComparison.OrdinalIgnoreCase));
        if (title == 0)
            return result;
        var header = Enumerable.Range(title + 1, 4).FirstOrDefault(r => ws.Text(r, 1).Equals("Этап", StringComparison.OrdinalIgnoreCase));
        if (header == 0)
            return result;

        // Стадии — объединенные ячейки строки заголовка; «Начало»/«Окончание» — в следующей строке.
        var lastColumn = Enumerable.Range(1, ws.Values.LastColumn).LastOrDefault(c => ws.Text(header, c).Length > 0);
        var stages = new List<(string Name, int Start, int End)>();
        for (var c = 3; c <= lastColumn; c++)
        {
            var name = ws.Text(header, c);
            if (name.Length == 0)
                continue;
            var span = ws.MergeEnd(header, c);
            int start = 0, end = 0;
            for (var k = c; k <= span; k++)
            {
                var label = ws.Text(header + 1, k);
                if (start == 0 && label.StartsWith("Начало", StringComparison.OrdinalIgnoreCase))
                    start = k;
                else if (end == 0 && label.StartsWith("Окончание", StringComparison.OrdinalIgnoreCase))
                    end = k;
            }

            if (start > 0)
                stages.Add((name, start, end));
            c = span;
        }

        for (var r = header + 2; r <= last; r++)
        {
            var number = ws.Text(r, 1);
            if (number.Length == 0)
                break;
            var deadlines = new Dictionary<string, List<string>>();
            foreach (var (name, start, end) in stages)
            {
                var a = Value(ws, r, start);
                var b = end > 0 ? Value(ws, r, end) : "";
                if (a.Length == 0 && b.Length == 0)
                    continue;
                deadlines[name] = b.Length == 0 && Formats.ParseDate(a) is null ? new List<string> { a } : new List<string> { a, b };
            }

            if (deadlines.Count > 0)
                result[number] = deadlines;
        }

        return result;
    }

    private static string Value(XlsxSheet ws, int row, int column) => ws.Values.Get(row, column) switch
    {
        DateTime d => d.ToString("dd.MM.yyyy", Formats.Ru),
        double n when n is > 30000 and < 80000 => DateTime.FromOADate(n).ToString("dd.MM.yyyy", Formats.Ru),
        _ => ws.Text(row, column),
    };

    /// <summary>
    /// Значение формулы, сохраненное Excel. Если его нет (книга сохранена не Excel) — вычисленное ClosedXML,
    /// а если и это не удалось — пусто.
    /// </summary>
    private static XLCellValue Computed(IXLCell cell)
    {
        var cached = cell.CachedValue;
        if (!cached.IsBlank)
            return cached;
        try
        {
            return cell.Value;
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            return Blank.Value;
        }
    }

    private static FileStream OpenRead(string path) =>
        new(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);

    // ------------------------------------------------------------------ цвета темы → явные

    private static void ResolveThemeColors(IXLStyle style, IXLTheme theme)
    {
        if (Resolve(style.Font.FontColor, theme) is { } font)
            style.Font.FontColor = font;
        if (style.Fill.PatternType != XLFillPatternValues.None)
        {
            if (Resolve(style.Fill.BackgroundColor, theme) is { } background)
                style.Fill.BackgroundColor = background;
            if (Resolve(style.Fill.PatternColor, theme) is { } pattern)
                style.Fill.PatternColor = pattern;
        }

        var border = style.Border;
        if (Resolve(border.LeftBorderColor, theme) is { } left)
            border.LeftBorderColor = left;
        if (Resolve(border.RightBorderColor, theme) is { } right)
            border.RightBorderColor = right;
        if (Resolve(border.TopBorderColor, theme) is { } top)
            border.TopBorderColor = top;
        if (Resolve(border.BottomBorderColor, theme) is { } bottom)
            border.BottomBorderColor = bottom;
    }

    private static XLColor? Resolve(XLColor color, IXLTheme theme)
    {
        if (color.ColorType != XLColorType.Theme)
            return null;
        var baseColor = color.ThemeColor switch
        {
            XLThemeColor.Background1 => theme.Background1,
            XLThemeColor.Text1 => theme.Text1,
            XLThemeColor.Background2 => theme.Background2,
            XLThemeColor.Text2 => theme.Text2,
            XLThemeColor.Accent1 => theme.Accent1,
            XLThemeColor.Accent2 => theme.Accent2,
            XLThemeColor.Accent3 => theme.Accent3,
            XLThemeColor.Accent4 => theme.Accent4,
            XLThemeColor.Accent5 => theme.Accent5,
            XLThemeColor.Accent6 => theme.Accent6,
            XLThemeColor.Hyperlink => theme.Hyperlink,
            XLThemeColor.FollowedHyperlink => theme.FollowedHyperlink,
            _ => null,
        };
        if (baseColor is null || baseColor.ColorType != XLColorType.Color)
            return null;
        var (r, g, b) = ApplyTint(baseColor.Color.R, baseColor.Color.G, baseColor.Color.B, color.ThemeTint);
        return XLColor.FromArgb(r, g, b);
    }

    /// <summary>Оттенок цвета темы — как в Excel: изменение яркости в модели HLS.</summary>
    internal static (int R, int G, int B) ApplyTint(int r, int g, int b, double tint)
    {
        if (Math.Abs(tint) < 1e-9)
            return (r, g, b);
        var (h, l, s) = ToHls(r / 255.0, g / 255.0, b / 255.0);
        l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
        var (rr, gg, bb) = FromHls(h, l, s);
        return ((int)Math.Round(rr * 255), (int)Math.Round(gg * 255), (int)Math.Round(bb * 255));
    }

    private static (double H, double L, double S) ToHls(double r, double g, double b)
    {
        var max = Math.Max(r, Math.Max(g, b));
        var min = Math.Min(r, Math.Min(g, b));
        var l = (max + min) / 2;
        if (max - min < 1e-12)
            return (0, l, 0);
        var d = max - min;
        var s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
        double h;
        if (max == r)
            h = (g - b) / d + (g < b ? 6 : 0);
        else if (max == g)
            h = (b - r) / d + 2;
        else
            h = (r - g) / d + 4;
        return (h / 6, l, s);
    }

    private static (double R, double G, double B) FromHls(double h, double l, double s)
    {
        if (s < 1e-12)
            return (l, l, l);
        var q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        var p = 2 * l - q;
        return (Hue(p, q, h + 1.0 / 3), Hue(p, q, h), Hue(p, q, h - 1.0 / 3));

        static double Hue(double p, double q, double t)
        {
            if (t < 0)
                t += 1;
            if (t > 1)
                t -= 1;
            if (t < 1.0 / 6)
                return p + (q - p) * 6 * t;
            if (t < 0.5)
                return q;
            if (t < 2.0 / 3)
                return p + (q - p) * (2.0 / 3 - t) * 6;
            return p;
        }
    }
}

/// <summary>Старая книга: список листов читается быстро, сама книга загружается только при необходимости (и один раз).</summary>
public sealed class LegacySource : IDisposable
{
    private XLWorkbook? _book;

    public LegacySource(string path)
    {
        Path = path;
    }

    public string Path { get; }

    public XLWorkbook Book => _book ??= LegacyWorkbook.Load(Path);

    public void Dispose() => _book?.Dispose();
}
