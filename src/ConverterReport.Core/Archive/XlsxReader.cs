using System.Globalization;
using System.Text.RegularExpressions;
using ConverterReport.Core.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConverterReport.Core.Archive;

/// <summary>Лист, прочитанный <see cref="XlsxReader"/>: значения, объединенные ячейки и (по запросу) формулы.</summary>
internal sealed class XlsxSheet
{
    public required RawSheet Values { get; init; }

    public List<(int Row1, int Col1, int Row2, int Col2)> Merges { get; } = new();

    public Dictionary<(int Row, int Column), string> Formulas { get; } = new();

    public string Text(int row, int column) => Values.Get(row, column) switch
    {
        null => "",
        double d => d.ToString(Formats.Ru),
        DateTime dt => dt.ToString("dd.MM.yyyy", Formats.Ru),
        var v => v.ToString()?.Trim() ?? "",
    };

    /// <summary>Последний столбец объединенной ячейки, начинающейся в (row, column) (или сам column).</summary>
    public int MergeEnd(int row, int column) =>
        Merges.FirstOrDefault(m => m.Row1 <= row && row <= m.Row2 && m.Col1 == column) is { Col2: > 0 } m ? m.Col2 : column;
}

/// <summary>
/// Чтение значений листов книги Excel без загрузки всей книги (Open XML, значения, сохраненные Excel).
/// Большие книги отчета (сотни листов дней) так читаются за секунды. Файл открывается только на чтение
/// и с разрешением совместного доступа — его можно читать, даже если он открыт в Excel.
/// </summary>
internal sealed class XlsxReader : IDisposable
{
    private static readonly HashSet<uint> BuiltinDateFormats = new() { 14, 15, 16, 17, 18, 19, 20, 21, 22, 45, 46, 47 };

    private readonly FileStream _stream;
    private readonly SpreadsheetDocument _doc;
    private readonly List<string> _strings = new();
    private readonly HashSet<uint> _dateStyles = new();
    private readonly Dictionary<string, WorksheetPart> _parts = new(StringComparer.OrdinalIgnoreCase);

    public XlsxReader(string path)
    {
        _stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        try
        {
            _doc = SpreadsheetDocument.Open(_stream, false);
        }
        catch
        {
            _stream.Dispose();
            throw;
        }

        var wb = _doc.WorkbookPart ?? throw new InvalidDataException("Файл не является книгой Excel.");
        foreach (var sheet in wb.Workbook.Descendants<Sheet>())
        {
            if (sheet.Name?.Value is not { } name || sheet.Id?.Value is not { } id)
                continue;
            if (wb.GetPartById(id) is WorksheetPart part)
            {
                _parts[name] = part;
                SheetNames.Add(name);
            }
        }

        if (wb.SharedStringTablePart is { } sst)
        {
            using var reader = OpenXmlReader.Create(sst);
            while (reader.Read())
            {
                if (reader.ElementType == typeof(SharedStringItem) && reader.IsStartElement)
                    _strings.Add(reader.LoadCurrentElement()?.InnerText ?? "");
            }
        }

        if (wb.WorkbookStylesPart?.Stylesheet is { } styles)
        {
            var custom = styles.NumberingFormats?.Elements<NumberingFormat>()
                .Where(f => f.NumberFormatId?.Value is not null)
                .ToDictionary(f => f.NumberFormatId!.Value, f => f.FormatCode?.Value ?? "") ?? new Dictionary<uint, string>();
            var index = 0u;
            foreach (var format in styles.CellFormats?.Elements<CellFormat>() ?? Enumerable.Empty<CellFormat>())
            {
                var id = format.NumberFormatId?.Value ?? 0;
                if (BuiltinDateFormats.Contains(id) || (custom.TryGetValue(id, out var code) && IsDateFormat(code)))
                    _dateStyles.Add(index);
                index++;
            }
        }
    }

    public List<string> SheetNames { get; } = new();

    public bool HasSheet(string name) => _parts.ContainsKey(name);

    /// <summary>Значения листа (до строки <paramref name="maxRow"/>); с формулами — если <paramref name="formulas"/>.</summary>
    public XlsxSheet ReadSheet(string name, int maxRow = int.MaxValue, bool formulas = false)
    {
        var part = _parts[name];
        var result = new XlsxSheet { Values = new RawSheet { Name = name } };
        var raw = result.Values;
        using var reader = OpenXmlReader.Create(part);
        while (reader.Read())
        {
            if (!reader.IsStartElement)
                continue;
            if (reader.ElementType == typeof(Row))
            {
                var row = (Row)reader.LoadCurrentElement()!;
                var r = (int)(row.RowIndex?.Value ?? (uint)(raw.LastRow + 1));
                if (r > maxRow)
                    break;
                var c = 0;
                foreach (var cell in row.Elements<Cell>())
                {
                    if (cell.CellReference?.Value is { } reference)
                        (_, c) = Parse(reference);
                    else
                        c++;
                    if (formulas && cell.CellFormula?.Text is { Length: > 0 } f)
                        result.Formulas[(r, c)] = f;
                    if (Value(cell) is not { } value || value is string { Length: 0 })
                        continue;
                    raw.Cells[(r, c)] = value;
                    raw.LastRow = Math.Max(raw.LastRow, r);
                    raw.LastColumn = Math.Max(raw.LastColumn, c);
                }
            }
            else if (reader.ElementType == typeof(Columns))
            {
                foreach (var column in ((Columns)reader.LoadCurrentElement()!).Elements<Column>())
                {
                    if (column.Width?.Value is not { } width || column.Min?.Value is not { } min || column.Max?.Value is not { } max)
                        continue;
                    for (var i = min; i <= Math.Min(max, min + 200); i++)
                        raw.ColumnWidths[(int)i] = width;
                }
            }
            else if (reader.ElementType == typeof(MergeCell))
            {
                if (((MergeCell)reader.LoadCurrentElement()!).Reference?.Value is { } range && range.Split(':') is { Length: 2 } ends)
                {
                    var (r1, c1) = Parse(ends[0]);
                    var (r2, c2) = Parse(ends[1]);
                    result.Merges.Add((r1, c1, r2, c2));
                }
            }
        }

        return result;
    }

    private object? Value(Cell cell)
    {
        var type = cell.DataType?.Value;
        var text = cell.CellValue?.Text;
        if (type == CellValues.InlineString)
            return cell.InlineString?.InnerText;
        if (text is null)
            return null;
        if (type == CellValues.SharedString)
            return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var i) && i >= 0 && i < _strings.Count ? _strings[i] : null;
        if (type == CellValues.String)
            return text;
        if (type == CellValues.Boolean)
            return text == "1";
        if (type == CellValues.Error)
            return null;
        if (!double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out var number))
            return text;
        if (cell.StyleIndex?.Value is { } style && _dateStyles.Contains(style))
        {
            if (number >= 1 && number < 2958466)
                return DateTime.FromOADate(number);
            if (number >= 0 && number < 1)
                return TimeSpan.FromDays(number);
        }

        return number;
    }

    /// <summary>Формат даты/времени: есть d, m, y, h или s вне кавычек и квадратных скобок.</summary>
    internal static bool IsDateFormat(string code)
    {
        var cleaned = Regex.Replace(code, "\"[^\"]*\"|\\[[^\\]]*\\]|\\\\.|_.|\\*.", "");
        return cleaned.IndexOfAny(new[] { 'd', 'D', 'm', 'M', 'y', 'Y', 'h', 'H', 's', 'S' }) >= 0
               && !cleaned.Equals("General", StringComparison.OrdinalIgnoreCase);
    }

    internal static (int Row, int Column) Parse(string reference)
    {
        var col = 0;
        var i = 0;
        for (; i < reference.Length && char.IsLetter(reference[i]); i++)
            col = col * 26 + (char.ToUpperInvariant(reference[i]) - 'A' + 1);
        var row = int.TryParse(reference.AsSpan(i), NumberStyles.Integer, CultureInfo.InvariantCulture, out var r) ? r : 0;
        return (row, col);
    }

    public void Dispose()
    {
        _doc.Dispose();
        _stream.Dispose();
    }
}
