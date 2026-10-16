using ClosedXML.Excel;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Model;

namespace ConverterReport.Core.Screens;

/// <summary>
/// Чтение файла «Экран» этапа: листы «Замечания» (состояние на дату/время) и «Документы» (загрузки в АИС ГГЭ).
/// Файл открывается только на чтение и с разрешением совместного доступа — его можно читать, даже если он открыт в Excel.
/// Для ячеек с формулами используются сохраненные Excel значения.
/// </summary>
public sealed class ScreenReader
{
    private readonly ScreenLayout _layout;

    public ScreenReader(ScreenLayout layout)
    {
        _layout = layout;
    }

    public ScreenData Read(string path)
    {
        using var file = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        using var buffer = new MemoryStream();
        file.CopyTo(buffer);
        buffer.Position = 0;
        return Read(buffer, path);
    }

    public ScreenData Read(Stream stream, string sourceName)
    {
        using var workbook = new XLWorkbook(stream);
        var remarks = FindSheet(workbook, _layout.RemarksSheet);
        var documents = FindSheet(workbook, _layout.DocumentsSheet);
        var data = FromRaw(remarks is null ? null : Raw(remarks), documents is null ? null : Raw(documents), sourceName);
        if (remarks is null)
            data.Warnings.Insert(0, $"В файле нет листа \"{_layout.RemarksSheet}\".");
        if (documents is null)
            data.Warnings.Add($"В файле нет листа \"{_layout.DocumentsSheet}\".");
        return data;
    }

    /// <summary>Данные экрана по значениям листов «Замечания» и «Документы» (экран, архив, общая книга).</summary>
    /// <param name="remarksRow">Номер строки для сообщений по номеру строки листа (если строки сдвинуты при объединении).</param>
    public ScreenData FromRaw(RawSheet? remarks, RawSheet? documents, string sourceName, Func<int, int>? remarksRow = null, Func<int, int>? loadsRow = null)
    {
        var data = new ScreenData { FilePath = sourceName, RawRemarks = remarks, RawDocuments = documents };
        if (remarks is not null)
            ReadRemarks(remarks, data, remarksRow ?? (r => r));
        if (documents is not null)
            ReadLoads(documents, data, loadsRow ?? (r => r));
        return data;
    }

    /// <summary>Строки «Замечаний» с данными: номер строки → дата и время.</summary>
    internal IEnumerable<(int Row, DateTime Timestamp)> RemarkRows(RawSheet sheet)
    {
        var dateCol = Col(_layout.RemarksDate);
        var timeCol = Col(_layout.RemarksTime);
        var allCol = Col(_layout.RemarksAll);
        for (var r = _layout.RemarksFirstRow; r <= sheet.LastRow; r++)
        {
            // Строка считается заполненной, если в «Всего» стоит число.
            if (!Formats.IsNumber(sheet.Get(r, allCol)))
                continue;
            var date = Formats.ParseDate(sheet.Get(r, dateCol));
            var time = Formats.ParseTime(sheet.Get(r, timeCol));
            if (date is not null && time is not null)
                yield return (r, date.Value.Date + time.Value);
        }
    }

    /// <summary>Строки «Документов» с загрузками: номер строки → окончание (или начало, если окончания нет).</summary>
    internal IEnumerable<(int Row, DateTime End)> LoadRows(RawSheet sheet)
    {
        for (var r = _layout.DocumentsFirstRow; r <= sheet.LastRow; r++)
        {
            var start = Formats.Combine(sheet.Get(r, Col(_layout.LoadStartDate)), sheet.Get(r, Col(_layout.LoadStartTime)));
            var end = Formats.Combine(sheet.Get(r, Col(_layout.LoadEndDate)), sheet.Get(r, Col(_layout.LoadEndTime))) ?? start;
            if (end is not null)
                yield return (r, end.Value);
        }
    }

    private void ReadRemarks(RawSheet sheet, ScreenData data, Func<int, int> sourceRow)
    {
        var dateCol = Col(_layout.RemarksDate);
        var timeCol = Col(_layout.RemarksTime);
        int[] allCols = { Col(_layout.RemarksAll), Col(_layout.RemarksSurvey), Col(_layout.RemarksTech), Col(_layout.RemarksEstimate) };
        int[] critCols = { Col(_layout.CriticalAll), Col(_layout.CriticalSurvey), Col(_layout.CriticalTech), Col(_layout.CriticalEstimate) };

        var byTime = new SortedDictionary<DateTime, RemarksSnapshot>();
        for (var r = _layout.RemarksFirstRow; r <= sheet.LastRow; r++)
        {
            // Строка считается заполненной, если в «Всего» стоит число.
            if (!Formats.IsNumber(sheet.Get(r, allCols[0])))
                continue;

            var date = Formats.ParseDate(sheet.Get(r, dateCol));
            var time = Formats.ParseTime(sheet.Get(r, timeCol));
            if (date is null || time is null)
            {
                data.Warnings.Add($"Лист \"{sheet.Name}\", строка {r}: заполнены данные, но не распознаны дата/время.");
                continue;
            }

            var snapshot = new RemarksSnapshot { Timestamp = date.Value.Date + time.Value, Row = sourceRow(r) };
            for (var i = 0; i < 4; i++)
            {
                snapshot.All[i] = ReadCounts(sheet, r, allCols[i]);
                snapshot.Critical[i] = ReadCounts(sheet, r, critCols[i]);
            }

            if (byTime.ContainsKey(snapshot.Timestamp))
                data.Warnings.Add($"Лист \"{sheet.Name}\": на {Formats.DateTimeText(snapshot.Timestamp)} несколько строк, используется строка {r}.");
            byTime[snapshot.Timestamp] = snapshot;
        }

        data.Remarks.AddRange(byTime.Values);
    }

    private void ReadLoads(RawSheet sheet, ScreenData data, Func<int, int> sourceRow)
    {
        var startDate = Col(_layout.LoadStartDate);
        var startTime = Col(_layout.LoadStartTime);
        var endDate = Col(_layout.LoadEndDate);
        var endTime = Col(_layout.LoadEndTime);

        for (var r = _layout.DocumentsFirstRow; r <= sheet.LastRow; r++)
        {
            var start = Formats.Combine(sheet.Get(r, startDate), sheet.Get(r, startTime));
            var end = Formats.Combine(sheet.Get(r, endDate), sheet.Get(r, endTime)) ?? start;
            if (end is null)
                continue; // пустая строка, строка «Итого» или загрузка еще не выполнена

            var counts = new LoadCounts(
                Int(sheet, r, _layout.LoadSurvey),
                Int(sheet, r, _layout.LoadProject),
                Int(sheet, r, _layout.LoadEstimate),
                Int(sheet, r, _layout.LoadVor),
                Int(sheet, r, _layout.LoadKac),
                Int(sheet, r, _layout.LoadOther),
                Int(sheet, r, _layout.LoadAnswers),
                Int(sheet, r, _layout.LoadCancelled),
                Int(sheet, r, _layout.LoadErrors));

            if (start is not null && start > end)
                data.Warnings.Add($"Лист \"{sheet.Name}\", строка {r}: начало загрузки позже окончания.");

            data.Loads.Add(new DocumentLoad
            {
                Row = sourceRow(r),
                Number = Text(sheet.Get(r, 1)),
                Start = start,
                End = end.Value,
                Counts = counts,
                Note = Text(sheet.Get(r, Col(_layout.LoadNote))),
            });
        }

        data.Loads.Sort((a, b) => a.End.CompareTo(b.End));
    }

    /// <summary>
    /// Значения всех ячеек листа — для копии листа в сводную книгу (листы «З.N» и «Д.N»).
    /// <paramref name="maxColumn"/> — только столбцы до него (служебные столбцы листов общей книги не читаются).
    /// </summary>
    internal static RawSheet Raw(IXLWorksheet sheet, int maxColumn = int.MaxValue)
    {
        var raw = new RawSheet { Name = sheet.Name };
        foreach (var cell in sheet.CellsUsed(XLCellsUsedOptions.Contents))
        {
            if (cell.Address.ColumnNumber > maxColumn)
                continue;
            var value = Value(cell);
            if (value is null || value is string { Length: 0 })
                continue;
            var row = cell.Address.RowNumber;
            var col = cell.Address.ColumnNumber;
            raw.Cells[(row, col)] = value;
            raw.LastRow = Math.Max(raw.LastRow, row);
            raw.LastColumn = Math.Max(raw.LastColumn, col);
        }

        foreach (var column in sheet.ColumnsUsed())
        {
            if (column.ColumnNumber() <= maxColumn)
                raw.ColumnWidths[column.ColumnNumber()] = column.Width;
        }
        return raw;
    }

    private static StatusCounts ReadCounts(RawSheet sheet, int row, int firstColumn) => new(
        Formats.ParseInt(sheet.Get(row, firstColumn)),
        Formats.ParseInt(sheet.Get(row, firstColumn + 1)),
        Formats.ParseInt(sheet.Get(row, firstColumn + 2)),
        Formats.ParseInt(sheet.Get(row, firstColumn + 3)));

    private static int Int(RawSheet sheet, int row, string column) =>
        string.IsNullOrWhiteSpace(column) ? 0 : Formats.ParseInt(sheet.Get(row, Col(column)));

    private static string? Text(object? v)
    {
        var s = v switch
        {
            null => null,
            double d => d.ToString(Formats.Ru),
            _ => v.ToString(),
        };
        return string.IsNullOrWhiteSpace(s) ? null : s.Trim();
    }

    private static int Col(string letter) => XLHelper.GetColumnNumberFromLetter(letter.Trim().ToUpperInvariant());

    /// <summary>Значение ячейки как object: double, DateTime, TimeSpan, string, bool или null.</summary>
    internal static object? Value(IXLCell cell)
    {
        XLCellValue v;
        try
        {
            v = cell.HasFormula ? cell.CachedValue : cell.Value;
            if (cell.HasFormula && v.IsBlank)
                v = cell.Value; // в файле нет сохраненного значения — пробуем вычислить формулу
        }
        catch
        {
            return null;
        }

        return v.Type switch
        {
            XLDataType.Number => v.GetNumber(),
            XLDataType.DateTime => v.GetDateTime(),
            XLDataType.TimeSpan => v.GetTimeSpan(),
            XLDataType.Text => v.GetText(),
            XLDataType.Boolean => v.GetBoolean(),
            _ => null,
        };
    }

    internal static IXLWorksheet? FindSheet(XLWorkbook workbook, string name)
    {
        var wanted = Normalize(name);
        return workbook.Worksheets.FirstOrDefault(ws => Normalize(ws.Name) == wanted);

        static string Normalize(string s) => s.Trim().ToLowerInvariant().Replace('ё', 'е');
    }
}
