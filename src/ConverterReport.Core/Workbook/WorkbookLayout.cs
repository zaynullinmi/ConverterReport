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

    /// <summary>Окончание экспертизы: 1 на эту дату (вертикальная линия на диаграмме), иначе #Н/Д.</summary>
    public int Marker => Remarks + 5;

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

/// <summary>Ряд диаграммы: имя — ссылка на заголовок «Свода», значения — именованный диапазон периода.</summary>
internal sealed record ChartSeriesRef(string Name, string NameRef, string ValuesRef, string Color, bool IsLine, string? TrendColor = null, bool LabelsBelow = false)
{
    /// <summary>Значения периода по умолчанию — кэш диаграммы (видно до пересчета книги).</summary>
    public IReadOnlyList<double?>? Cache { get; init; }
}

/// <summary>Диаграмма на листе «Диаграммы»: положение (строки/колонки, с 0) и ряды.</summary>
internal sealed record ChartPlacement(
    ChartKind Kind,
    string Title,
    int FromRow,
    int ToRow,
    int FromColumn,
    int ToColumn,
    string CategoriesRef,
    IReadOnlyList<ChartSeriesRef> Series)
{
    /// <summary>Сводная диаграмма (по всем этапам).</summary>
    public bool Summary { get; init; }

    /// <summary>Количество точек (дней) по оси категорий.</summary>
    public int CategoryCount { get; init; }

    /// <summary>Верх шкалы (null — автоматически).</summary>
    public double? AxisMax { get; init; }

    /// <summary>Номер точки с датой окончания экспертизы (для кэша).</summary>
    public int? MarkerIndex { get; init; }

    /// <summary>Ряд «окончание экспертизы» (1 на дату окончания) — вертикальная пунктирная линия; null — линии нет.</summary>
    public string? MarkerRef { get; init; }

    /// <summary>Заголовок — ссылка на ячейку (меняется вместе с периодом); <see cref="Title"/> — его кэш.</summary>
    public string? TitleRef { get; init; }

    /// <summary>Подписи оси дат периода по умолчанию (кэш).</summary>
    public IReadOnlyList<string>? CategoryCache { get; init; }
}

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

    /// <summary>Сколько листов прошлых дней перенесено из старой книги.</summary>
    public int ImportedSheets { get; set; }

    /// <summary>Не удалось прочитать старую книгу (лист дня в общую книгу при этом записан).</summary>
    public string? LegacyWarning { get; set; }

    internal List<ChartPlacement> Charts { get; } = new();
}
