namespace ConverterReport.Core.Model;

/// <summary>Счетчики замечаний: Всего, Устранено, Устраняется, Рассматривается (ГГЭ).</summary>
public readonly record struct StatusCounts(int Total, int Fixed, int Fixing, int Review)
{
    public static readonly StatusCounts Zero = default;

    /// <summary>Не устранено = устраняется + на рассмотрении.</summary>
    public int Open => Fixing + Review;

    /// <summary>Всего = Устранено + Устраняется + Рассматривается.</summary>
    public bool IsConsistent => Total == Fixed + Fixing + Review;

    public static StatusCounts operator +(StatusCounts a, StatusCounts b) =>
        new(a.Total + b.Total, a.Fixed + b.Fixed, a.Fixing + b.Fixing, a.Review + b.Review);

    public static StatusCounts operator -(StatusCounts a, StatusCounts b) =>
        new(a.Total - b.Total, a.Fixed - b.Fixed, a.Fixing - b.Fixing, a.Review - b.Review);
}

/// <summary>Разделы замечаний. All — все замечания, остальные — «в т.ч.».</summary>
public enum RemarkCategory
{
    All = 0,
    Survey = 1,
    Tech = 2,
    Estimate = 3,
}

public static class RemarkCategories
{
    public static readonly RemarkCategory[] Values = { RemarkCategory.All, RemarkCategory.Survey, RemarkCategory.Tech, RemarkCategory.Estimate };

    public static string Title(RemarkCategory c) => c switch
    {
        RemarkCategory.All => "Все замечания",
        RemarkCategory.Survey => "Инженерные изыскания",
        RemarkCategory.Tech => "Тех. часть",
        RemarkCategory.Estimate => "Сметная часть",
        _ => c.ToString(),
    };
}

/// <summary>Строка листа «Замечания» экрана: состояние замечаний на дату и время.</summary>
public sealed class RemarksSnapshot
{
    public DateTime Timestamp { get; init; }

    /// <summary>Строка листа (для сообщений пользователю). 0 — расчетное значение.</summary>
    public int Row { get; init; }

    /// <summary>Все замечания по разделам (индекс — <see cref="RemarkCategory"/>).</summary>
    public StatusCounts[] All { get; init; } = new StatusCounts[4];

    /// <summary>Критические замечания по разделам.</summary>
    public StatusCounts[] Critical { get; init; } = new StatusCounts[4];

    public StatusCounts Get(RemarkCategory category, bool critical) => (critical ? Critical : All)[(int)category];

    public static RemarksSnapshot Empty(DateTime timestamp) => new() { Timestamp = timestamp };

    public static RemarksSnapshot operator +(RemarksSnapshot a, RemarksSnapshot b) => Combine(a, b, a.Timestamp, (x, y) => x + y);

    public static RemarksSnapshot operator -(RemarksSnapshot a, RemarksSnapshot b) => Combine(a, b, a.Timestamp, (x, y) => x - y);

    private static RemarksSnapshot Combine(RemarksSnapshot a, RemarksSnapshot b, DateTime ts, Func<StatusCounts, StatusCounts, StatusCounts> op)
    {
        var result = new RemarksSnapshot { Timestamp = ts };
        for (var i = 0; i < 4; i++)
        {
            result.All[i] = op(a.All[i], b.All[i]);
            result.Critical[i] = op(a.Critical[i], b.Critical[i]);
        }

        return result;
    }

    /// <summary>Сумма разделов (ИИ + Тех.часть + Сметная) совпадает с итогом — как столбцы «Сверка» в экране.</summary>
    public bool SectionsMatch(bool critical)
    {
        var arr = critical ? Critical : All;
        var sum = arr[1] + arr[2] + arr[3];
        return sum == arr[0];
    }
}

/// <summary>Количество документов в загрузке по видам.</summary>
public readonly record struct LoadCounts(
    int Survey,
    int Project,
    int Estimate,
    int Vor,
    int Kac,
    int Other,
    int Answers = 0,
    int Cancelled = 0,
    int Errors = 0)
{
    public static readonly LoadCounts Zero = default;

    /// <summary>Всего = ИИ + ПД + СМ + ВОР + КАЦ + Иные (как столбец «Всего» экрана).</summary>
    public int Total => Survey + Project + Estimate + Vor + Kac + Other;

    public static LoadCounts operator +(LoadCounts a, LoadCounts b) => new(
        a.Survey + b.Survey, a.Project + b.Project, a.Estimate + b.Estimate, a.Vor + b.Vor, a.Kac + b.Kac,
        a.Other + b.Other, a.Answers + b.Answers, a.Cancelled + b.Cancelled, a.Errors + b.Errors);
}

/// <summary>Строка листа «Документы» экрана — одна загрузка в АИС ГГЭ.</summary>
public sealed class DocumentLoad
{
    public int Row { get; init; }

    public string? Number { get; init; }

    public DateTime? Start { get; init; }

    public DateTime End { get; init; }

    public LoadCounts Counts { get; init; }

    public string? Note { get; init; }
}

/// <summary>Данные, прочитанные из одного файла «Экран».</summary>
public sealed class ScreenData
{
    public string FilePath { get; init; } = "";

    /// <summary>Заполненные строки «Замечаний», по возрастанию времени.</summary>
    public List<RemarksSnapshot> Remarks { get; } = new();

    /// <summary>Загрузки, по возрастанию времени окончания.</summary>
    public List<DocumentLoad> Loads { get; } = new();

    public List<string> Warnings { get; } = new();

    /// <summary>Все значения листа «Замечания» (для листа «З.N» сводной книги).</summary>
    public RawSheet? RawRemarks { get; set; }

    /// <summary>Все значения листа «Документы» (для листа «Д.N» сводной книги).</summary>
    public RawSheet? RawDocuments { get; set; }
}

/// <summary>Значения ячеек листа (без формул и оформления): ключ — (строка, столбец), нумерация с 1.</summary>
public sealed class RawSheet
{
    public string Name { get; init; } = "";

    public Dictionary<(int Row, int Column), object> Cells { get; } = new();

    public Dictionary<int, double> ColumnWidths { get; } = new();

    public int LastRow { get; set; }

    public int LastColumn { get; set; }

    public object? Get(int row, int column) => Cells.TryGetValue((row, column), out var v) ? v : null;
}
