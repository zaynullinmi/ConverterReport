using ConverterReport.Core.Configuration;
using ConverterReport.Core.Model;

namespace ConverterReport.Core.Reporting;

/// <summary>Точка диаграммы: состояние замечаний на конец дня и загрузки за день.</summary>
public sealed class DailyPoint
{
    public DateTime Date { get; init; }

    /// <summary>Все замечания по последней строке экрана за день (не позже времени отчета). null — данных нет.</summary>
    public StatusCounts? Remarks { get; set; }

    /// <summary>Устранено за сутки (разница «Устранено» с предыдущей точкой).</summary>
    public int? FixedPerDay { get; set; }

    /// <summary>Загружено документов за день (по дате окончания загрузки).</summary>
    public int Loaded { get; set; }
}

/// <summary>Рассчитанные данные по одному этапу.</summary>
public sealed class ProjectReport
{
    public required ProjectConfig Project { get; init; }

    public string Number => Project.Number;

    public string Name => Project.DisplayName;

    public ScreenData? Screen { get; init; }

    /// <summary>Ошибка чтения экрана (файл не найден и т.п.).</summary>
    public string? Error { get; set; }

    /// <summary>Состояние замечаний на время отчета (последняя заполненная строка не позже времени отчета).</summary>
    public RemarksSnapshot? Current { get; set; }

    /// <summary>Состояние на начало отчетного периода.</summary>
    public RemarksSnapshot? Previous { get; set; }

    public RemarksSnapshot? Delta => Current is null || Previous is null ? null : Current - Previous;

    /// <summary>Загрузки для раздела «1. Загрузка документов» (см. <see cref="ReportModel.LoadsFrom"/>).</summary>
    public List<DocumentLoad> PeriodLoads { get; } = new();

    public LoadCounts PeriodLoadsSum => PeriodLoads.Aggregate(LoadCounts.Zero, (acc, l) => acc + l.Counts);

    public DateTime? PeriodLoadsStart => PeriodLoads.Count == 0 ? null : PeriodLoads.Min(l => l.Start ?? l.End);

    public DateTime? PeriodLoadsEnd => PeriodLoads.Count == 0 ? null : PeriodLoads.Max(l => l.End);

    /// <summary>Ряды для диаграмм (по дням).</summary>
    public List<DailyPoint> Daily { get; } = new();

    public DateTime? ExpertiseEnd { get; set; }

    public List<string> Warnings { get; } = new();

    public bool HasRemarksSeries => Daily.Any(p => p.Remarks is not null);

    public bool HasLoadsSeries => Daily.Any(p => p.Loaded != 0);

    public int LoadedInCharts => Daily.Sum(p => p.Loaded);

    /// <summary>Данные экрана есть, но не на отчетное время (использована более ранняя строка).</summary>
    public bool IsStale(DateTime reportTime) => Current is not null && Current.Timestamp != reportTime;
}

/// <summary>Полностью рассчитанный отчет — источник для PDF и выгрузки в Excel.</summary>
public sealed class ReportModel
{
    public required ReportConfig Config { get; init; }

    public DateTime ReportTime { get; init; }

    public DateTime PeriodStart { get; init; }

    /// <summary>Раздел «1. Загрузка документов» — загрузки за день (с 00:00), иначе — с начала отчетного периода.</summary>
    public bool LoadsForDay { get; init; }

    /// <summary>Начало выборки загрузок раздела 1 (по окончанию загрузки).</summary>
    public DateTime LoadsFrom => LoadsForDay ? ReportTime.Date : PeriodStart;

    /// <summary>
    /// Загрузка попадает в раздел 1: окончание за день (с 00:00 включительно — как на диаграммах) или после начала
    /// отчетного периода, и не позже времени отчета.
    /// </summary>
    public bool IsReportLoad(DateTime end) => (LoadsForDay ? end >= LoadsFrom : end > LoadsFrom) && end <= ReportTime;

    public string NoLoadsText => LoadsForDay ? "За день загрузок документов не было." : "За отчетный период загрузок документов не было.";

    public List<ProjectReport> Projects { get; init; } = new();

    /// <summary>Сводные ряды по всем этапам (для общей диаграммы). null — если этап один.</summary>
    public ProjectReport? Summary { get; set; }

    public List<string> Notes { get; init; } = new();

    /// <summary>Этапы, которые выводятся в отчете (у остальных — только сводная диаграмма).</summary>
    public List<ProjectReport> ShownProjects => Projects.Where(p => p.Project.ShowInReport).ToList();

    /// <summary>Все этапы с экраном, «2-4, 9» — для имен файлов и сводной диаграммы (архивные этапы не входят).</summary>
    public string ProjectsList => Formats.CompressNumbers(Projects.Where(p => !p.Project.IsArchive).Select(p => p.Number));

    /// <summary>Этапы отчета, «3-4, 9».</summary>
    public string ShownProjectsList => Formats.CompressNumbers(ShownProjects.Select(p => p.Number));

    /// <summary>«по 3-4, 9 этапам строительства ВСМ-1».</summary>
    public string Caption => CaptionFor(ShownProjectsList);

    /// <summary>Подпись детального отчета — список без пробела: «по 3-4,9 этапам строительства ВСМ-1».</summary>
    public string DetailedCaption => CaptionFor(ShownProjectsList.Replace(", ", ","));

    private string CaptionFor(string list)
    {
        var template = ShownProjects.Count == 1 ? Config.Report.ProjectsCaptionSingle : Config.Report.ProjectsCaption;
        return Formats.ApplyTemplate(template, values: new Dictionary<string, string> { ["projects"] = list });
    }

    public RemarksSnapshot TotalCurrent =>
        ShownProjects.Where(p => p.Current is not null).Aggregate(RemarksSnapshot.Empty(ReportTime), (acc, p) => acc + p.Current!);

    public RemarksSnapshot TotalDelta =>
        ShownProjects.Where(p => p.Delta is not null).Aggregate(RemarksSnapshot.Empty(ReportTime), (acc, p) => acc + p.Delta!);

    public LoadCounts TotalPeriodLoads => ShownProjects.Aggregate(LoadCounts.Zero, (acc, p) => acc + p.PeriodLoadsSum);

    public IEnumerable<string> AllWarnings =>
        Projects.SelectMany(p => p.Warnings.Select(w => $"Этап {p.Number}: {w}"));
}
