using ConverterReport.Core.Configuration;
using ConverterReport.Core.Model;
using ConverterReport.Core.Screens;

namespace ConverterReport.Core.Archive;

/// <summary>Источник строк этапа для объединения: экран, лист общей книги или прошлая версия отчета.</summary>
internal sealed record StageSource(string Name, RawSheet? Remarks, RawSheet? Documents);

/// <summary>Результат объединения: данные этапа и сколько строк добавлено не из экрана.</summary>
internal sealed record MergedStage(ScreenData Data, int AddedRemarks, int AddedLoads);

/// <summary>
/// Объединение данных этапа: строки экрана берутся целиком (как есть), а строки общей книги и прошлых версий
/// отчета — только за даты раньше первой строки экрана (то, чего в экране уже нет: прошлые циклы экспертизы,
/// этапы без экрана). Совпадающие строки (то же время) берутся из первого источника. Строки идут по времени.
/// </summary>
internal sealed class StageMerger
{
    private readonly ScreenLayout _layout;
    private readonly ScreenReader _reader;

    public StageMerger(ScreenLayout layout)
    {
        _layout = layout;
        _reader = new ScreenReader(layout);
    }

    /// <param name="live">Экран этапа (null — архивный этап).</param>
    /// <param name="others">Прочие источники по убыванию приоритета: лист общей книги, прошлые версии отчета.</param>
    public MergedStage? Merge(ScreenData? live, IReadOnlyList<StageSource> others, string sourceName)
    {
        var sources = new List<StageSource>();
        if (live is not null)
            sources.Add(new StageSource(live.FilePath, live.RawRemarks, live.RawDocuments));
        sources.AddRange(others);
        if (sources.All(s => s.Remarks is null && s.Documents is null))
            return null;

        var remarks = MergeSheets(live?.RawRemarks, sources.Select(s => s.Remarks), _layout.RemarksFirstRow,
            sheet => _reader.RemarkRows(sheet).Select(x => (x.Row, x.Timestamp, x.Timestamp.Ticks.ToString())), out var addedRemarks);
        var documents = MergeSheets(live?.RawDocuments, sources.Select(s => s.Documents), _layout.DocumentsFirstRow,
            sheet => _reader.LoadRows(sheet).Select(x => (x.Row, x.End, LoadKey(sheet, x.Row, x.End))), out var addedLoads);

        if (live is not null && addedRemarks == 0 && addedLoads == 0)
            return new MergedStage(live, 0, 0);

        // Строки экрана сдвинуты вниз на число добавленных строк: в сообщениях — номера строк экрана (у архивных — 0).
        Func<int, int> Shift(int first, int count) => r => r < first ? r : r < first + count ? 0 : r - count;
        var data = live is null
            ? _reader.FromRaw(remarks, documents, sourceName)
            : _reader.FromRaw(remarks, documents, sourceName,
                Shift(_layout.RemarksFirstRow, addedRemarks), Shift(_layout.DocumentsFirstRow, addedLoads));
        if (live is not null)
        {
            // Предупреждения — по экрану (номера строк экрана), а не по объединенным листам.
            data.Warnings.Clear();
            data.Warnings.AddRange(live.Warnings);
        }

        return new MergedStage(data, addedRemarks, addedLoads);
    }

    private string LoadKey(RawSheet sheet, int row, DateTime end)
    {
        var start = Formats.Combine(sheet.Get(row, Col(_layout.LoadStartDate)), sheet.Get(row, Col(_layout.LoadStartTime)));
        return $"{end.Ticks}/{start?.Ticks}";
    }

    /// <summary>
    /// Лист экрана (<paramref name="live"/>) целиком + строки прочих источников раньше первой строки экрана.
    /// Без экрана — все строки всех источников (без повторов) по времени, шапка — из первого источника.
    /// </summary>
    private static RawSheet? MergeSheets(RawSheet? live, IEnumerable<RawSheet?> all, int firstRow,
        Func<RawSheet, IEnumerable<(int Row, DateTime Time, string Key)>> rows, out int added)
    {
        added = 0;
        var sheets = all.Where(s => s is not null).Cast<RawSheet>().ToList();
        if (sheets.Count == 0)
            return null;

        var head = live ?? sheets[0];
        var liveRows = live is null ? new List<(int Row, DateTime Time, string Key)>() : rows(live).ToList();
        var cutoff = liveRows.Count > 0 ? liveRows.Min(r => r.Time) : DateTime.MaxValue;
        var seen = new HashSet<string>(liveRows.Select(r => r.Key));
        var earlier = new List<(DateTime Time, RawSheet Sheet, int Row)>();
        foreach (var sheet in sheets.Where(s => !ReferenceEquals(s, live)))
        {
            foreach (var (row, time, key) in rows(sheet))
            {
                if (time < cutoff && seen.Add(key))
                    earlier.Add((time, sheet, row));
            }
        }

        if (live is not null && earlier.Count == 0)
            return live;

        var result = new RawSheet { Name = head.Name };
        foreach (var (col, width) in head.ColumnWidths)
            result.ColumnWidths[col] = width;
        var byRow = sheets.ToDictionary(s => s, ByRow);
        var headRows = byRow[head];
        foreach (var (row, cells) in headRows.Where(kv => kv.Key < firstRow))
            Put(result, row, cells);

        var r = firstRow;
        foreach (var (_, sheet, row) in earlier.OrderBy(e => e.Time))
            Put(result, r++, byRow[sheet].GetValueOrDefault(row) ?? new List<(int, object)>());
        added = earlier.Count;

        if (live is not null)
        {
            foreach (var (row, cells) in headRows.Where(kv => kv.Key >= firstRow).OrderBy(kv => kv.Key))
                Put(result, r + row - firstRow, cells);
        }

        return result;
    }

    private static Dictionary<int, List<(int Column, object Value)>> ByRow(RawSheet sheet)
    {
        var rows = new Dictionary<int, List<(int, object)>>();
        foreach (var ((row, col), value) in sheet.Cells)
        {
            if (!rows.TryGetValue(row, out var list))
                rows[row] = list = new List<(int, object)>();
            list.Add((col, value));
        }

        return rows;
    }

    private static void Put(RawSheet sheet, int row, List<(int Column, object Value)> cells)
    {
        foreach (var (col, value) in cells)
        {
            sheet.Cells[(row, col)] = value;
            sheet.LastColumn = Math.Max(sheet.LastColumn, col);
        }

        sheet.LastRow = Math.Max(sheet.LastRow, row);
    }

    private static int Col(string letter) => ClosedXML.Excel.XLHelper.GetColumnNumberFromLetter(letter.Trim().ToUpperInvariant());
}
