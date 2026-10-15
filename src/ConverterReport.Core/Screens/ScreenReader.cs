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
        var data = new ScreenData { FilePath = sourceName };

        var remarks = FindSheet(workbook, _layout.RemarksSheet);
        if (remarks is null)
        {
            data.Warnings.Add($"В файле нет листа \"{_layout.RemarksSheet}\".");
        }
        else
        {
            ReadRemarks(remarks, data);
            data.RawRemarks = Raw(remarks);
        }

        var documents = FindSheet(workbook, _layout.DocumentsSheet);
        if (documents is null)
        {
            data.Warnings.Add($"В файле нет листа \"{_layout.DocumentsSheet}\".");
        }
        else
        {
            ReadLoads(documents, data);
            data.RawDocuments = Raw(documents);
        }

        return data;
    }

    private void ReadRemarks(IXLWorksheet sheet, ScreenData data)
    {
        var dateCol = Col(_layout.RemarksDate);
        var timeCol = Col(_layout.RemarksTime);
        int[] allCols = { Col(_layout.RemarksAll), Col(_layout.RemarksSurvey), Col(_layout.RemarksTech), Col(_layout.RemarksEstimate) };
        int[] critCols = { Col(_layout.CriticalAll), Col(_layout.CriticalSurvey), Col(_layout.CriticalTech), Col(_layout.CriticalEstimate) };

        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;
        var byTime = new SortedDictionary<DateTime, RemarksSnapshot>();
        for (var r = _layout.RemarksFirstRow; r <= lastRow; r++)
        {
            var row = sheet.Row(r);
            // Строка считается заполненной, если в «Всего» стоит число.
            if (!Formats.IsNumber(Value(row.Cell(allCols[0]))))
                continue;

            var date = Formats.ParseDate(Value(row.Cell(dateCol)));
            var time = Formats.ParseTime(Value(row.Cell(timeCol)));
            if (date is null || time is null)
            {
                data.Warnings.Add($"Лист \"{sheet.Name}\", строка {r}: заполнены данные, но не распознаны дата/время.");
                continue;
            }

            var snapshot = new RemarksSnapshot { Timestamp = date.Value.Date + time.Value, Row = r };
            for (var i = 0; i < 4; i++)
            {
                snapshot.All[i] = ReadCounts(row, allCols[i]);
                snapshot.Critical[i] = ReadCounts(row, critCols[i]);
            }

            if (byTime.ContainsKey(snapshot.Timestamp))
                data.Warnings.Add($"Лист \"{sheet.Name}\": на {Formats.DateTimeText(snapshot.Timestamp)} несколько строк, используется строка {r}.");
            byTime[snapshot.Timestamp] = snapshot;
        }

        data.Remarks.AddRange(byTime.Values);
    }

    private void ReadLoads(IXLWorksheet sheet, ScreenData data)
    {
        var startDate = Col(_layout.LoadStartDate);
        var startTime = Col(_layout.LoadStartTime);
        var endDate = Col(_layout.LoadEndDate);
        var endTime = Col(_layout.LoadEndTime);
        var lastRow = sheet.LastRowUsed()?.RowNumber() ?? 0;

        for (var r = _layout.DocumentsFirstRow; r <= lastRow; r++)
        {
            var row = sheet.Row(r);
            var start = Formats.Combine(Value(row.Cell(startDate)), Value(row.Cell(startTime)));
            var end = Formats.Combine(Value(row.Cell(endDate)), Value(row.Cell(endTime))) ?? start;
            if (end is null)
                continue; // пустая строка, строка «Итого» или загрузка еще не выполнена

            var counts = new LoadCounts(
                Int(row, _layout.LoadSurvey),
                Int(row, _layout.LoadProject),
                Int(row, _layout.LoadEstimate),
                Int(row, _layout.LoadVor),
                Int(row, _layout.LoadKac),
                Int(row, _layout.LoadOther),
                Int(row, _layout.LoadAnswers),
                Int(row, _layout.LoadCancelled),
                Int(row, _layout.LoadErrors));

            if (start is not null && start > end)
                data.Warnings.Add($"Лист \"{sheet.Name}\", строка {r}: начало загрузки позже окончания.");

            data.Loads.Add(new DocumentLoad
            {
                Row = r,
                Number = Text(row.Cell(1)),
                Start = start,
                End = end.Value,
                Counts = counts,
                Note = Text(row.Cell(Col(_layout.LoadNote))),
            });
        }

        data.Loads.Sort((a, b) => a.End.CompareTo(b.End));
    }

    /// <summary>Значения всех ячеек листа — для копии листа в сводную книгу (листы «З.N» и «Д.N»).</summary>
    private static RawSheet Raw(IXLWorksheet sheet)
    {
        var raw = new RawSheet { Name = sheet.Name };
        foreach (var cell in sheet.CellsUsed(XLCellsUsedOptions.Contents))
        {
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
            raw.ColumnWidths[column.ColumnNumber()] = column.Width;
        return raw;
    }

    private static StatusCounts ReadCounts(IXLRow row, int firstColumn) => new(
        Formats.ParseInt(Value(row.Cell(firstColumn))),
        Formats.ParseInt(Value(row.Cell(firstColumn + 1))),
        Formats.ParseInt(Value(row.Cell(firstColumn + 2))),
        Formats.ParseInt(Value(row.Cell(firstColumn + 3))));

    private static int Int(IXLRow row, string column) =>
        string.IsNullOrWhiteSpace(column) ? 0 : Formats.ParseInt(Value(row.Cell(Col(column))));

    private static string? Text(IXLCell cell)
    {
        var v = Value(cell);
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

    private static IXLWorksheet? FindSheet(XLWorkbook workbook, string name)
    {
        var wanted = Normalize(name);
        return workbook.Worksheets.FirstOrDefault(ws => Normalize(ws.Name) == wanted);

        static string Normalize(string s) => s.Trim().ToLowerInvariant().Replace('ё', 'е');
    }
}
