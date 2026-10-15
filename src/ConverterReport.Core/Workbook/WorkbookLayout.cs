using ClosedXML.Excel;

namespace ConverterReport.Core.Workbook;

/// <summary>Колонки блока листа «Свод»: замечания (5 колонок) и загрузки (10 колонок).</summary>
internal readonly record struct SvodBlock(int Label, int Remarks, int Loads)
{
    public int Total => Remarks;
    public int Fixed => Remarks + 1;
    public int Fixing => Remarks + 2;
    public int Review => Remarks + 3;
    public int FixedPerDay => Remarks + 4;

    /// <summary>ИИ, ПД, СМ, ВОР, КАЦ, Иные, Всего, Ответы, Аннулировано, Ошибки.</summary>
    public int Load(int index) => Loads + index;

    public int LoadTotal => Loads + 6;
}

/// <summary>
/// Раскладка листа «Свод» — как в книге отчета: B — дата, C — день недели, D — подпись для диаграмм, E — время;
/// F–J и L–U — итог по всем этапам; далее блоки по 17 колонок на каждый этап.
/// </summary>
internal sealed class SvodLayout
{
    public const string SheetName = "Свод";
    public const int HeaderRow = 1;
    public const int FirstRow = 2;
    public const int DateColumn = 2;
    public const int WeekdayColumn = 3;
    public const int LabelColumn = 4;
    public const int TimeColumn = 5;
    private const int FirstProjectColumn = 23; // W
    private const int BlockWidth = 17;

    public static readonly SvodBlock Overall = new(0, 6, 12);

    public SvodLayout(DateTime start, DateTime end)
    {
        Start = start.Date;
        End = end.Date < start.Date ? start.Date : end.Date;
    }

    public DateTime Start { get; }

    public DateTime End { get; }

    public int LastRow => Row(End);

    public int Row(DateTime date) => FirstRow + (date.Date - Start).Days;

    public static SvodBlock Project(int index)
    {
        var label = FirstProjectColumn + index * BlockWidth;
        return new SvodBlock(label, label + 1, label + 7);
    }

    public static string Ref(int column, int row1, int row2) =>
        $"{SheetName}!${Col(column)}${row1}:${Col(column)}${row2}";

    public static string Ref(int column, int row) => $"{SheetName}!${Col(column)}${row}";

    public static string Col(int column) => XLHelper.GetColumnLetterFromNumber(column);
}

internal enum ChartKind
{
    Remarks,
    Loads,
}

/// <summary>Ряд диаграммы со ссылками на лист «Свод».</summary>
internal sealed record ChartSeriesRef(string Name, string NameRef, string ValuesRef, string Color, bool IsLine, string? TrendColor = null, bool LabelsBelow = false);

/// <summary>Диаграмма на листе «Диаграммы»: положение (строки/колонки, с 0) и ряды.</summary>
internal sealed record ChartPlacement(
    ChartKind Kind,
    string Title,
    int FromRow,
    int ToRow,
    int FromColumn,
    int ToColumn,
    string CategoriesRef,
    IReadOnlyList<ChartSeriesRef> Series);

/// <summary>Результат построения сводной книги.</summary>
public sealed class SummaryWorkbook
{
    public const string DetailedSheet = "Отчет_дет";
    public const string ChartsSheet = "Диаграммы";
    public const string DataSheet = "Отчет";

    public required string Path { get; init; }

    /// <summary>Лист суточного отчета («ДД.ММ»).</summary>
    public required string DailySheet { get; init; }

    public int ChartPages { get; init; }

    internal List<ChartPlacement> Charts { get; } = new();
}
