using ConverterReport.Core.Configuration;
using ConverterReport.Core.Model;
using ConverterReport.Core.Reporting;

namespace ConverterReport.Tests;

public class ReportBuilderTests
{
    private static readonly DateTime ReportAt = new(2026, 10, 15, 8, 0, 0);
    private static readonly DateTime PeriodFrom = new(2026, 10, 14, 18, 0, 0);

    [Fact]
    public void SumsAllLoadsInPeriod_NotOnlyTheLastOne()
    {
        var screen = Screen(
            loads: new[]
            {
                Load(new DateTime(2026, 10, 14, 17, 50, 0), survey: 100), // до начала периода
                Load(new DateTime(2026, 10, 14, 18, 0, 0), survey: 50),   // ровно на начало — в прошлый период
                Load(new DateTime(2026, 10, 14, 20, 45, 0), project: 6),
                Load(new DateTime(2026, 10, 14, 22, 27, 0), survey: 2, project: 1, estimate: 1),
                Load(new DateTime(2026, 10, 15, 8, 0, 0), other: 3),      // ровно на время отчета — входит
                Load(new DateTime(2026, 10, 15, 9, 0, 0), other: 1000),   // после отчета
            });

        var model = Build(new ProjectInput(Project("3"), screen));
        var p = model.Projects.Single();

        Assert.Equal(3, p.PeriodLoads.Count);
        Assert.Equal(2, p.PeriodLoadsSum.Survey);
        Assert.Equal(7, p.PeriodLoadsSum.Project);
        Assert.Equal(1, p.PeriodLoadsSum.Estimate);
        Assert.Equal(3, p.PeriodLoadsSum.Other);
        Assert.Equal(13, p.PeriodLoadsSum.Total);
        Assert.Equal(new DateTime(2026, 10, 14, 20, 40, 0), p.PeriodLoadsStart);
        Assert.Equal(new DateTime(2026, 10, 15, 8, 0, 0), p.PeriodLoadsEnd);
    }

    [Fact]
    public void TakesSnapshotAtReportTimeAndDeltaFromPeriodStart()
    {
        var screen = Screen(remarks: new[]
        {
            Snap(new DateTime(2026, 10, 14, 8, 0, 0), 398, 350, 47, 1),
            Snap(new DateTime(2026, 10, 14, 18, 0, 0), 398, 352, 45, 1),
            Snap(new DateTime(2026, 10, 15, 8, 0, 0), 400, 360, 37, 3),
            Snap(new DateTime(2026, 10, 15, 18, 0, 0), 400, 380, 17, 3), // после отчета — не учитывается
        });

        var model = Build(new ProjectInput(Project("3"), screen));
        var p = model.Projects.Single();

        Assert.Equal(new DateTime(2026, 10, 15, 8, 0, 0), p.Current!.Timestamp);
        Assert.Equal(new DateTime(2026, 10, 14, 18, 0, 0), p.Previous!.Timestamp);
        Assert.False(p.IsStale(ReportAt));
        Assert.Equal(new StatusCounts(2, 8, -8, 2), p.Delta!.All[0]);
        Assert.Equal(8, model.TotalDelta.All[0].Fixed);
        Assert.Empty(p.Warnings);
    }

    [Fact]
    public void FallsBackToLatestRowAndWarnsWhenReportRowIsEmpty()
    {
        var screen = Screen(remarks: new[]
        {
            Snap(new DateTime(2026, 10, 14, 18, 0, 0), 398, 352, 45, 1),
        });

        var p = Build(new ProjectInput(Project("4"), screen)).Projects.Single();

        Assert.True(p.IsStale(ReportAt));
        Assert.Contains(p.Warnings, w => w.Contains("нет строки на 15.10.2026 08:00"));
    }

    [Fact]
    public void WarnsAboutInconsistentTotals()
    {
        var bad = Snap(new DateTime(2026, 10, 15, 8, 0, 0), 400, 360, 37, 1); // 360 + 37 + 1 ≠ 400
        var p = Build(new ProjectInput(Project("9"), Screen(remarks: new[] { bad }))).Projects.Single();
        Assert.Contains(p.Warnings, w => w.Contains("не равно сумме"));
    }

    [Fact]
    public void SupportsAnyNumberOfProjects()
    {
        var inputs = Enumerable.Range(1, 12)
            .Select(i => new ProjectInput(Project(i.ToString()), Screen(remarks: new[] { Snap(ReportAt, 100 * i, 50 * i, 40 * i, 10 * i) })))
            .ToArray();

        var model = Build(inputs);

        Assert.Equal(12, model.Projects.Count);
        Assert.Equal("1-12", model.ProjectsList);
        Assert.Equal("по 1-12 этапам строительства ВСМ-1", model.Caption);
        Assert.Equal(7800, model.TotalCurrent.All[0].Total);
        Assert.NotNull(model.Summary);
    }

    [Fact]
    public void MissingScreenIsReportedButOthersAreBuilt()
    {
        var ok = new ProjectInput(Project("3"), Screen(remarks: new[] { Snap(ReportAt, 10, 5, 5, 0) }));
        var missing = new ProjectInput(Project("2"), null, "файл экрана не найден");

        var model = Build(missing, ok);

        Assert.Null(model.Projects[0].Current);
        Assert.Contains("файл экрана не найден", model.Projects[0].Warnings);
        Assert.Equal(10, model.TotalCurrent.All[0].Total);
        Assert.Equal("по 2-3 этапам строительства ВСМ-1", model.Caption);
    }

    [Fact]
    public void DailySeriesUsesLastRowOfDayAndFixedPerDay()
    {
        var project = Project("3");
        project.ChartStartDate = "13.10.2026";
        project.ExpertiseEndDate = "20.10.2026";
        var screen = Screen(
            remarks: new[]
            {
                Snap(new DateTime(2026, 10, 12, 18, 0, 0), 100, 10, 90, 0),
                Snap(new DateTime(2026, 10, 13, 8, 0, 0), 100, 15, 85, 0),
                Snap(new DateTime(2026, 10, 13, 18, 0, 0), 100, 20, 80, 0),
                Snap(new DateTime(2026, 10, 14, 18, 0, 0), 100, 26, 74, 0),
                Snap(new DateTime(2026, 10, 15, 8, 0, 0), 100, 30, 70, 0),
            },
            loads: new[]
            {
                Load(new DateTime(2026, 10, 14, 10, 0, 0), survey: 5),
                Load(new DateTime(2026, 10, 14, 20, 0, 0), survey: 7),
            });

        var p = Build(new ProjectInput(project, screen)).Projects.Single();

        // Ось: с 13.10 до окончания экспертизы 20.10.
        Assert.Equal(new DateTime(2026, 10, 13), p.Daily[0].Date);
        Assert.Equal(new DateTime(2026, 10, 20), p.Daily[^1].Date);
        Assert.Equal(20, p.Daily[0].Remarks!.Value.Fixed);
        Assert.Null(p.Daily[0].FixedPerDay); // до начала оси не считаем
        Assert.Equal(6, p.Daily[1].FixedPerDay);
        Assert.Equal(12, p.Daily[1].Loaded);
        Assert.Equal(4, p.Daily[2].FixedPerDay);
        Assert.Null(p.Daily[3].Remarks); // будущие даты пустые
        Assert.Equal(new DateTime(2026, 10, 20), p.ExpertiseEnd);
    }

    [Fact]
    public void ExpertiseEndIsLatestDeadlineDate()
    {
        var project = Project("9");
        project.Deadlines["Экспертиза"] = new List<string> { "12.08.2026", "09.10.2026" };
        project.Deadlines["Продление 1"] = new List<string> { "10.10.2026", "09.11.2026" };
        project.Deadlines["Продление 2"] = new List<string> { "Заявка до 23.10.2026" };

        Assert.Equal(new DateTime(2026, 11, 9), ReportBuilder.ResolveExpertiseEnd(project));
    }

    [Fact]
    public void SummaryAddsProjectsAndCarriesForwardMissingDays()
    {
        var a = Project("1");
        var b = Project("2");
        a.ChartStartDate = b.ChartStartDate = "14.10.2026";
        var sa = Screen(remarks: new[]
        {
            Snap(new DateTime(2026, 10, 14, 18, 0, 0), 100, 10, 90, 0),
            Snap(new DateTime(2026, 10, 15, 8, 0, 0), 100, 15, 85, 0),
        });
        var sb = Screen(remarks: new[]
        {
            Snap(new DateTime(2026, 10, 14, 18, 0, 0), 50, 5, 45, 0),
        });

        var model = Build(new ProjectInput(a, sa), new ProjectInput(b, sb));
        var summary = model.Summary!;

        Assert.Equal(150, summary.Daily[0].Remarks!.Value.Total);
        Assert.Equal(150, summary.Daily[1].Remarks!.Value.Total); // этап 2 на 15.10 не заполнен — берется последнее значение
        Assert.Equal(20, summary.Daily[1].Remarks!.Value.Fixed);
        Assert.Equal(5, summary.Daily[1].FixedPerDay);
    }

    // ------------------------------------------------------------------

    private static ReportModel Build(params ProjectInput[] inputs) =>
        ReportBuilder.Build(new ReportConfig(), inputs, ReportAt, PeriodFrom);

    private static ProjectConfig Project(string number) => new() { Number = number, ScreenFile = number + ".xlsx" };

    private static ScreenData Screen(RemarksSnapshot[]? remarks = null, DocumentLoad[]? loads = null)
    {
        var data = new ScreenData { FilePath = "test.xlsx" };
        if (remarks is not null)
            data.Remarks.AddRange(remarks.OrderBy(r => r.Timestamp));
        if (loads is not null)
            data.Loads.AddRange(loads.OrderBy(l => l.End));
        return data;
    }

    private static RemarksSnapshot Snap(DateTime at, int total, int fix, int fixing, int review)
    {
        var s = new RemarksSnapshot { Timestamp = at, Row = 1 };
        s.All[0] = new StatusCounts(total, fix, fixing, review);
        s.All[2] = s.All[0];
        return s;
    }

    private static DocumentLoad Load(DateTime end, int survey = 0, int project = 0, int estimate = 0, int other = 0) => new()
    {
        Start = end.AddMinutes(-5),
        End = end,
        Counts = new LoadCounts(survey, project, estimate, 0, 0, other),
    };
}
