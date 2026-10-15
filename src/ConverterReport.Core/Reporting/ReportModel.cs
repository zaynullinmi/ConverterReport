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

    /// <summary>Загрузки, завершенные в отчетном периоде.</summary>
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

    public List<ProjectReport> Projects { get; init; } = new();

    /// <summary>Сводные ряды по всем этапам (для общей диаграммы). null — если этап один.</summary>
    public ProjectReport? Summary { get; set; }

    public List<string> Notes { get; init; } = new();

    /// <summary>«2-4, 9».</summary>
    public string ProjectsList => Formats.CompressNumbers(Projects.Select(p => p.Number));

    /// <summary>«по 2-4, 9 этапам строительства ВСМ-1».</summary>
    public string Caption
    {
        get
        {
            var template = Projects.Count == 1 ? Config.Report.ProjectsCaptionSingle : Config.Report.ProjectsCaption;
            return Formats.ApplyTemplate(template, values: new Dictionary<string, string> { ["projects"] = ProjectsList });
        }
    }

    public RemarksSnapshot TotalCurrent =>
        Projects.Where(p => p.Current is not null).Aggregate(RemarksSnapshot.Empty(ReportTime), (acc, p) => acc + p.Current!);

    public RemarksSnapshot TotalDelta =>
        Projects.Where(p => p.Delta is not null).Aggregate(RemarksSnapshot.Empty(ReportTime), (acc, p) => acc + p.Delta!);

    public LoadCounts TotalPeriodLoads => Projects.Aggregate(LoadCounts.Zero, (acc, p) => acc + p.PeriodLoadsSum);

    public IEnumerable<string> AllWarnings =>
        Projects.SelectMany(p => p.Warnings.Select(w => $"Этап {p.Number}: {w}"));
}
