using System.Text.RegularExpressions;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Model;
using ConverterReport.Core.Screens;
using ConverterReport.Core.Workbook;

namespace ConverterReport.Core.Archive;

/// <summary>Данные одного этапа в прошлой версии отчета: листы Д.N/З.N (или «Документы»/«Замечания»).</summary>
public sealed class ArchivePart
{
    /// <summary>Номер листов («5» — листы Д.5/З.5) или «Экран» для книги одного этапа.</summary>
    public required string Key { get; init; }

    /// <summary>Листы для показа пользователю: «Д.5, З.5».</summary>
    public required string Sheets { get; init; }

    /// <summary>Этап, который предлагается по книге (заголовок блока «Свода», номер листа, текст отчета).</summary>
    public string? SuggestedStage { get; init; }

    /// <summary>Откуда взят предлагаемый номер этапа.</summary>
    public string Note { get; init; } = "";

    public required ScreenData Data { get; init; }

    public DateTime? From => Dates().DefaultIfEmpty().Min() is var d && d != default ? d : null;

    public DateTime? To => Dates().DefaultIfEmpty().Max() is var d && d != default ? d : null;

    private IEnumerable<DateTime> Dates() => Data.Remarks.Select(s => s.Timestamp).Concat(Data.Loads.Select(l => l.End));
}

/// <summary>Прошлая версия отчета: этапы (листы данных) и листы дней.</summary>
public sealed class ArchiveBook
{
    public required string Path { get; init; }

    public List<ArchivePart> Parts { get; } = new();

    /// <summary>Листы дней «ДД.ММ».</summary>
    public List<string> DailySheets { get; } = new();
}

/// <summary>
/// Чтение прошлых версий отчета: сводных книг (листы Д.N/З.N, «Свод», листы дней) и книг одного этапа
/// с листами «Замечания»/«Документы» (как у экрана).
/// </summary>
public static class ArchiveReader
{
    public const string ScreenKey = "Экран";

    private static readonly Regex DataSheet = new(@"^\s*З\.\s*(?<key>.+?)\s*$", RegexOptions.CultureInvariant);
    private static readonly Regex SheetReference = new(@"'?З\.(?<key>[^'!]+)'?!", RegexOptions.CultureInvariant);
    private static readonly Regex StageLabel = new(@"^\s*(?<n>[0-9A-Za-zА-Яа-я.\-]+)\s+этап", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
    private static readonly Regex StageInText = new(@"по\s+(?<n>\d+)\s*(?:-?\s*му\s+)?этап", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
    private static readonly Regex StageInFileName = new(@"(?<n>\d+)\s*этап", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    public static ArchiveBook Read(string path, ScreenLayout layout)
    {
        using var xlsx = new XlsxReader(path);
        var book = new ArchiveBook { Path = path };
        book.DailySheets.AddRange(xlsx.SheetNames.Where(SummaryWorkbookBuilder.IsDailySheet));
        var reader = new ScreenReader(layout);

        var keys = xlsx.SheetNames.Select(n => DataSheet.Match(n)).Where(m => m.Success).Select(m => m.Groups["key"].Value).ToList();
        if (keys.Count > 0)
        {
            var labels = SvodLabels(xlsx);
            foreach (var key in keys)
            {
                var remarks = xlsx.ReadSheet("З." + key).Values;
                var docsName = xlsx.SheetNames.FirstOrDefault(n => n.Trim() == "Д." + key);
                var documents = docsName is null ? null : xlsx.ReadSheet(docsName).Values;
                var (stage, note) = labels.TryGetValue(key, out var label)
                    ? (label, $"в «Своде» — «{label} этап»")
                    : (key, "по номеру листа");
                book.Parts.Add(new ArchivePart
                {
                    Key = key,
                    Sheets = docsName is null ? "З." + key : $"Д.{key}, З.{key}",
                    SuggestedStage = stage,
                    Note = note,
                    Data = reader.FromRaw(remarks, documents, $"{Path.GetFileName(path)}: З.{key}"),
                });
            }
        }
        else if (xlsx.SheetNames.FirstOrDefault(n => Same(n, layout.RemarksSheet)) is { } remarksName)
        {
            var documentsName = xlsx.SheetNames.FirstOrDefault(n => Same(n, layout.DocumentsSheet));
            var (stage, note) = SingleStage(xlsx, path);
            book.Parts.Add(new ArchivePart
            {
                Key = ScreenKey,
                Sheets = documentsName is null ? remarksName : $"{remarksName}, {documentsName}",
                SuggestedStage = stage,
                Note = note,
                Data = reader.FromRaw(xlsx.ReadSheet(remarksName).Values,
                    documentsName is null ? null : xlsx.ReadSheet(documentsName).Values, Path.GetFileName(path)),
            });
        }

        return book;
    }

    /// <summary>Лист дня (значения и объединения) — для сроков этапов.</summary>
    internal static XlsxSheet? ReadDailySheet(string path, DateTime reportDate)
    {
        using var xlsx = new XlsxReader(path);
        var last = xlsx.SheetNames.Where(SummaryWorkbookBuilder.IsDailySheet)
            .Select(n => (Name: n, Date: SummaryWorkbookBuilder.DailyDate(n, reportDate)))
            .Where(x => x.Date <= reportDate)
            .OrderBy(x => x.Date)
            .LastOrDefault();
        return last.Name is null ? null : xlsx.ReadSheet(last.Name);
    }

    /// <summary>
    /// Номера этапов блоков «Свода»: заголовок «N этап» над блоком, формулы которого ссылаются на лист З.X
    /// (на месте этапа 5 в одной версии мог стоять этап 9 — номер листа и этапа не обязательно совпадают).
    /// </summary>
    private static Dictionary<string, string> SvodLabels(XlsxReader xlsx)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!xlsx.HasSheet(SvodLayout.SheetName))
            return result;
        var svod = xlsx.ReadSheet(SvodLayout.SheetName, maxRow: 40, formulas: true);
        var firstColumn = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var ((_, column), formula) in svod.Formulas)
        {
            var refs = SheetReference.Matches(formula).Select(m => m.Groups["key"].Value.Trim()).Distinct().ToList();
            if (refs.Count == 1 && (!firstColumn.TryGetValue(refs[0], out var c) || column < c))
                firstColumn[refs[0]] = column;
        }

        foreach (var (key, column) in firstColumn)
        {
            if (StageLabel.Match(svod.Text(1, column - 1)) is { Success: true } m)
                result[key] = m.Groups["n"].Value;
        }

        return result;
    }

    /// <summary>Номер этапа книги одного этапа: из текста отчета («по 3 этапу») или из имени файла.</summary>
    private static (string? Stage, string Note) SingleStage(XlsxReader xlsx, string path)
    {
        foreach (var name in xlsx.SheetNames.Where(n => n.StartsWith("Отч", StringComparison.OrdinalIgnoreCase)))
        {
            var sheet = xlsx.ReadSheet(name, maxRow: 10).Values;
            foreach (var value in sheet.Cells.Values.OfType<string>())
            {
                if (StageInText.Match(value) is { Success: true } m)
                    return (m.Groups["n"].Value, $"в листе «{name}» — «по {m.Groups["n"].Value} этапу»");
            }
        }

        return StageInFileName.Match(Path.GetFileNameWithoutExtension(path)) is { Success: true } f
            ? (f.Groups["n"].Value, "по имени файла")
            : (null, "номер этапа не найден — укажите");
    }

    private static bool Same(string a, string b) =>
        string.Equals(a.Trim().Replace('ё', 'е'), b.Trim().Replace('ё', 'е'), StringComparison.OrdinalIgnoreCase);
}
