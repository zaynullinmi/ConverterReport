using ConverterReport.Core.Configuration;
using ConverterReport.Core.Model;

namespace ConverterReport.Core.Reporting;

/// <summary>Исходные данные этапа: настройки и прочитанный экран (или ошибка чтения).</summary>
public sealed record ProjectInput(ProjectConfig Project, ScreenData? Screen, string? Error = null);

/// <summary>
/// Расчет отчета по данным экранов: состояние замечаний на время отчета, изменения за период,
/// сумма всех загрузок за период (а не только последняя), ряды для диаграмм.
/// </summary>
public static class ReportBuilder
{
    /// <summary>Насколько далеко (в днях) ось диаграммы может уходить вперед до окончания экспертизы.</summary>
    private const int MaxFutureDays = 180;

    public static ReportModel Build(
        ReportConfig config,
        IReadOnlyList<ProjectInput> inputs,
        DateTime reportTime,
        DateTime periodStart,
        IReadOnlyList<string>? notes = null)
    {
        var model = new ReportModel
        {
            Config = config,
            ReportTime = reportTime,
            PeriodStart = periodStart,
            Notes = notes?.ToList() ?? new List<string>(),
        };

        var globalChartStart = Formats.ParseDate(config.Charts.StartDate)?.Date;
        foreach (var input in inputs)
            model.Projects.Add(BuildProject(config, input, reportTime, periodStart, globalChartStart));

        if (model.Projects.Count > 1)
            model.Summary = BuildSummary(model);

        return model;
    }

    private static ProjectReport BuildProject(ReportConfig config, ProjectInput input, DateTime reportTime, DateTime periodStart, DateTime? globalChartStart)
    {
        var report = new ProjectReport { Project = input.Project, Screen = input.Screen, Error = input.Error };
        report.ExpertiseEnd = ResolveExpertiseEnd(input.Project);

        if (input.Error is not null)
            report.Warnings.Add(input.Error);
        if (input.Screen is null)
            return report;

        report.Warnings.AddRange(input.Screen.Warnings);

        var snapshots = input.Screen.Remarks.Where(s => s.Timestamp <= reportTime).ToList();
        report.Current = snapshots.LastOrDefault();
        report.Previous = snapshots.LastOrDefault(s => s.Timestamp <= periodStart);

        if (report.Current is null)
        {
            report.Warnings.Add($"в листе \"Замечания\" нет заполненных строк на {Formats.DateTimeText(reportTime)} и ранее.");
        }
        else
        {
            if (report.Current.Timestamp != reportTime)
                report.Warnings.Add($"нет строки на {Formats.DateTimeText(reportTime)} в листе \"Замечания\" — использованы данные на {Formats.DateTimeText(report.Current.Timestamp)} (строка {report.Current.Row}).");
            CheckConsistency(report);
        }

        if (report.Current is not null && report.Previous is null)
            report.Warnings.Add($"нет данных на начало периода ({Formats.DateTimeText(periodStart)}) — изменения за период не рассчитаны.");
        else if (report.Previous is not null && report.Previous.Timestamp != periodStart)
            report.Warnings.Add($"на начало периода использованы данные на {Formats.DateTimeText(report.Previous.Timestamp)}.");

        report.PeriodLoads.AddRange(input.Screen.Loads.Where(l => l.End > periodStart && l.End <= reportTime));

        BuildDaily(config, report, input.Screen, reportTime, globalChartStart);
        return report;
    }

    private static void CheckConsistency(ProjectReport report)
    {
        var s = report.Current!;
        var row = s.Row > 0 ? $" (строка {s.Row})" : "";
        if (!s.All[0].IsConsistent)
            report.Warnings.Add($"всего замечаний {s.All[0].Total} не равно сумме устранено+устраняется+рассматривается ({s.All[0].Fixed + s.All[0].Fixing + s.All[0].Review}){row}.");
        if (!s.Critical[0].IsConsistent)
            report.Warnings.Add($"всего критических {s.Critical[0].Total} не равно сумме по статусам ({s.Critical[0].Fixed + s.Critical[0].Fixing + s.Critical[0].Review}){row}.");
        if (!s.SectionsMatch(critical: false))
            report.Warnings.Add($"сумма по разделам (ИИ + Тех.часть + Сметная) не совпадает с общим числом замечаний{row}.");
        if (!s.SectionsMatch(critical: true))
            report.Warnings.Add($"сумма критических по разделам не совпадает с общим числом критических{row}.");
    }

    private static void BuildDaily(ReportConfig config, ProjectReport report, ScreenData screen, DateTime reportTime, DateTime? globalChartStart)
    {
        var reportDate = reportTime.Date;
        var snapshots = screen.Remarks.Where(s => s.Timestamp <= reportTime).ToList();
        var loads = screen.Loads.Where(l => l.End <= reportTime).ToList();

        // Начало оси: явно заданное → начало сроков экспертизы (входной контроль) → первая заполненная строка экрана.
        var from = Formats.ParseDate(report.Project.ChartStartDate)?.Date
                   ?? globalChartStart
                   ?? EarliestDeadline(report.Project)
                   ?? snapshots.FirstOrDefault()?.Timestamp.Date
                   ?? loads.FirstOrDefault()?.End.Date;
        if (from is null || from > reportDate)
            return;

        var to = AxisEnd(config, reportDate, report.ExpertiseEnd);
        var lastByDay = snapshots.GroupBy(s => s.Timestamp.Date).ToDictionary(g => g.Key, g => g.Last());
        var loadsByDay = loads.GroupBy(l => l.End.Date).ToDictionary(g => g.Key, g => g.Sum(l => l.Counts.Total));

        // До начала оси «устранено за сутки» не считаем: там может быть предыдущий цикл экспертизы.
        StatusCounts? previous = null;
        for (var d = from.Value; d <= to; d = d.AddDays(1))
        {
            var point = new DailyPoint { Date = d };
            if (d <= reportDate)
            {
                if (lastByDay.TryGetValue(d, out var snap))
                {
                    point.Remarks = snap.All[0];
                    point.FixedPerDay = previous is null ? null : snap.All[0].Fixed - previous.Value.Fixed;
                    previous = snap.All[0];
                }

                point.Loaded = loadsByDay.GetValueOrDefault(d);
            }

            report.Daily.Add(point);
        }
    }

    private static ProjectReport BuildSummary(ReportModel model)
    {
        var config = model.Config;
        var name = Formats.ApplyTemplate(config.Charts.SummaryName, values: new Dictionary<string, string> { ["projects"] = model.ProjectsList });
        var summary = new ProjectReport { Project = new ProjectConfig { Number = model.ProjectsList, Name = name } };

        var withData = model.Projects.Where(p => p.Daily.Count > 0).ToList();
        if (withData.Count == 0)
            return summary;

        var reportDate = model.ReportTime.Date;
        var from = Formats.ParseDate(config.Charts.StartDate)?.Date ?? withData.Min(p => p.Daily[0].Date);
        var ends = model.Projects.Where(p => p.ExpertiseEnd is not null).Select(p => p.ExpertiseEnd!.Value).ToList();
        summary.ExpertiseEnd = ends.Count > 0 ? ends.Max() : null;
        var to = AxisEnd(config, reportDate, summary.ExpertiseEnd);

        // Для суммы по этапам берется последнее известное значение каждого этапа (если в какой-то день строка не заполнена).
        var lastKnown = new Dictionary<ProjectReport, StatusCounts>();
        var byDate = withData.ToDictionary(p => p, p => p.Daily.ToDictionary(x => x.Date));
        for (var d = from; d <= to; d = d.AddDays(1))
        {
            var point = new DailyPoint { Date = d };
            if (d <= reportDate)
            {
                int? fixedPerDay = null;
                foreach (var p in withData)
                {
                    if (!byDate[p].TryGetValue(d, out var pp))
                        continue;
                    point.Loaded += pp.Loaded;
                    if (pp.Remarks is not null)
                        lastKnown[p] = pp.Remarks.Value;
                    // «Устранено за сутки» суммируется по этапам, чтобы подключение нового этапа не давало скачка.
                    if (pp.FixedPerDay is not null)
                        fixedPerDay = (fixedPerDay ?? 0) + pp.FixedPerDay.Value;
                }

                if (lastKnown.Count > 0)
                {
                    point.Remarks = lastKnown.Values.Aggregate(StatusCounts.Zero, (a, b) => a + b);
                    point.FixedPerDay = fixedPerDay;
                }
            }

            summary.Daily.Add(point);
        }

        return summary;
    }

    private static DateTime AxisEnd(ReportConfig config, DateTime reportDate, DateTime? expertiseEnd)
    {
        if (!config.Charts.ExtendToExpertiseEnd || expertiseEnd is null || expertiseEnd.Value.Date <= reportDate)
            return reportDate;
        var limit = reportDate.AddDays(MaxFutureDays);
        return expertiseEnd.Value.Date < limit ? expertiseEnd.Value.Date : limit;
    }

    /// <summary>Дата окончания экспертизы: явно заданная или наибольшая дата из сроков.</summary>
    public static DateTime? ResolveExpertiseEnd(ProjectConfig project)
    {
        var explicitEnd = Formats.ParseDate(project.ExpertiseEndDate);
        if (explicitEnd is not null)
            return explicitEnd.Value.Date;

        var dates = DeadlineDates(project);
        return dates.Count > 0 ? dates.Max() : null;
    }

    private static DateTime? EarliestDeadline(ProjectConfig project)
    {
        var dates = DeadlineDates(project);
        return dates.Count > 0 ? dates.Min() : null;
    }

    private static List<DateTime> DeadlineDates(ProjectConfig project) =>
        project.Deadlines.Values
            .SelectMany(v => v)
            .Select(v => Formats.ParseDate(v))
            .Where(d => d is not null)
            .Select(d => d!.Value.Date)
            .ToList();
}
