using System.Text.RegularExpressions;
using ClosedXML.Excel;
using ConverterReport.Core;
using ConverterReport.Core.Configuration;
using DocumentFormat.OpenXml.Packaging;

namespace ConverterReport.Tests;

/// <summary>Режим «общая книга»: лист дня дописывается в одну книгу, прежние листы дней сохраняются.</summary>
public sealed class MasterWorkbookTests : IDisposable
{
    private readonly string _dir = Directory.CreateTempSubdirectory("converter-report-master-").FullName;

    public void Dispose()
    {
        try
        {
            Directory.Delete(_dir, recursive: true);
        }
        catch (IOException)
        {
        }
    }

    private string Master => Path.Combine(_dir, "out", "Сводный отчет ГГЭ ВСМ-1.xlsx");

    [Fact]
    public void AppendsDailySheetsAndKeepsPreviousOnes()
    {
        var config = Config("3", "4");
        config.Report.VisibleDailySheets = 1;

        Run(config, 13, 18);
        Run(config, 14, 18);
        var result = Run(config, 15, 8);

        Assert.Equal(Master, result.Workbook);
        Assert.True(File.Exists(Master + ".bak"));
        Assert.False(File.Exists(Path.Combine(_dir, "out", "Сводный отчет ГГЭ ВСМ-1.tmp.xlsx")));
        Assert.DoesNotContain(result.Warnings, w => w.Contains("общая книга", StringComparison.OrdinalIgnoreCase));

        using (var wb = new XLWorkbook(Master))
        {
            Assert.Equal(
                new[] { "Описание", "Д.3", "З.3", "Д.4", "З.4", "Свод", "Диаграммы", "13.10", "14.10", "15.10", "Отчет_дет", "Отчет" },
                wb.Worksheets.Select(w => w.Name));

            // Видим только последний лист дня; открыт и выделен только он.
            Assert.Equal(XLWorksheetVisibility.Hidden, wb.Worksheet("13.10").Visibility);
            Assert.Equal(XLWorksheetVisibility.Hidden, wb.Worksheet("14.10").Visibility);
            Assert.Equal(XLWorksheetVisibility.Visible, wb.Worksheet("15.10").Visibility);
            Assert.Equal(new[] { "15.10" }, wb.Worksheets.Where(w => w.TabSelected).Select(w => w.Name));
            Assert.True(wb.Worksheet("15.10").TabActive);

            // Прежний лист дня не изменился: его дата и ссылка на строку экрана на 14.10 18:00.
            var old = wb.Worksheet("14.10");
            Assert.Equal(new DateTime(2026, 10, 14), old.Cell("AE4").GetDateTime());
            Assert.Equal(0.75, Days(old.Cell("AF4")), 6);
            var reference = old.CellsUsed().Select(c => c.FormulaA1).First(f => f.Contains("'З.3'!E"));
            var row = int.Parse(Regex.Match(reference, @"'З\.3'!E(\d+)").Groups[1].Value);
            Assert.Equal(new DateTime(2026, 10, 14), wb.Worksheet("З.3").Cell(row, 2).GetDateTime());
            Assert.Equal(0.75, Days(wb.Worksheet("З.3").Cell(row, 4)), 6);
        }

        // Диаграммы пересозданы один раз: сводная + 2 этапа, по 2 диаграммы.
        using var doc = SpreadsheetDocument.Open(Master, false);
        var drawings = doc.WorkbookPart!.WorksheetParts.Where(p => p.DrawingsPart is not null).ToList();
        Assert.Single(drawings);
        Assert.Equal(6, drawings[0].DrawingsPart!.ChartParts.Count());
    }

    [Fact]
    public void ReportForSameDateReplacesSheet()
    {
        var config = Config("3");
        config.Report.VisibleDailySheets = 0;
        Run(config, 14, 18);
        Run(config, 15, 8);
        Run(config, 15, 18);

        using var wb = new XLWorkbook(Master);
        Assert.Equal(new[] { "14.10", "15.10" }, wb.Worksheets.Select(w => w.Name).Where(n => n.Contains('.') && char.IsDigit(n[0])));
        Assert.Equal(0.75, Days(wb.Worksheet("15.10").Cell("AF4")), 6);
        Assert.All(wb.Worksheets, w => Assert.Equal(XLWorksheetVisibility.Visible, w.Visibility));
    }

    [Fact]
    public void ForeignWorkbookIsNotChanged()
    {
        var config = Config("3");
        Directory.CreateDirectory(Path.GetDirectoryName(Master)!);
        using (var foreign = new XLWorkbook())
        {
            foreign.AddWorksheet("15.10").Cell("A1").Value = "книга прежнего формата";
            foreign.SaveAs(Master);
        }

        var before = File.ReadAllBytes(Master);
        var result = Run(config, 15, 8);

        Assert.NotEqual(Master, result.Workbook);
        Assert.True(File.Exists(result.Workbook));
        Assert.Contains(result.Warnings, w => w.Contains("создана не программой"));
        Assert.Equal(before, File.ReadAllBytes(Master));
    }

    [Fact]
    public void BusyWorkbookFallsBackToSeparateBook()
    {
        var config = Config("3");
        Run(config, 14, 18);
        var before = File.ReadAllBytes(Master);

        RunResult result;
        using (new FileStream(Master, FileMode.Open, FileAccess.Read, FileShare.None))
            result = Run(config, 15, 8);

        Assert.NotEqual(Master, result.Workbook);
        Assert.Contains("2026.10.15 8.00 Сводный отчет", result.Workbook);
        Assert.Contains(result.Warnings, w => w.Contains("открыта в другой программе"));
        Assert.Equal(before, File.ReadAllBytes(Master));
    }

    [Fact]
    public void MasterPathIsRelativeToReportsFolder()
    {
        var folder = Path.Combine(_dir, "Отчеты");
        var at = new DateTime(2026, 10, 15, 8, 0, 0);
        var names = new Dictionary<string, string> { ["projectsshort"] = "2-4,9" };

        Assert.Equal(Path.Combine(folder, "Общая.xlsx"), ReportGenerator.MasterPath("Общая", folder, at, names));
        Assert.Equal(Path.Combine(folder, "Общая 2026.10 2-4,9.xlsx"), ReportGenerator.MasterPath("Общая {date:yyyy.MM} {projectsShort}.xlsx", folder, at, names));
        var rooted = Path.Combine(_dir, "Книги", "Общая.xlsx");
        Assert.Equal(rooted, ReportGenerator.MasterPath(rooted, folder, at, names));
    }

    private ReportConfig Config(params string[] numbers)
    {
        var config = new ReportConfig { BaseDirectory = _dir };
        config.Report.OutputFolder = "out";
        config.Report.WorkbookMode = WorkbookMode.Append;
        config.Report.PdfEngine = PdfEngine.Builtin;
        config.Report.IncludeDaily = false;
        config.Report.IncludeDetailed = false;
        config.Report.IncludeCharts = false;
        foreach (var n in numbers)
        {
            CreateScreen($"Экран {n}.xlsx");
            config.Projects.Add(new ProjectConfig
            {
                Number = n,
                ScreenFile = $"Экран {n}.xlsx",
                Deadlines = { ["Экспертиза"] = new List<string> { "01.10.2026", "30.10.2026" } },
            });
        }

        return config;
    }

    private static double Days(IXLCell cell) =>
        cell.Value.IsTimeSpan ? cell.Value.GetTimeSpan().TotalDays : cell.GetDouble();

    private static RunResult Run(ReportConfig config, int day, int hour) =>
        new ReportGenerator().Run(config, new RunOptions { Date = new DateTime(2026, 10, day), Time = TimeSpan.FromHours(hour) });

    private void CreateScreen(string name)
    {
        if (File.Exists(Path.Combine(_dir, name)))
            return;
        var screen = new TestScreen();
        var start = new DateTime(2026, 10, 1);
        for (var d = 0; d < 15; d++)
        {
            var day = start.AddDays(d);
            var fixedCount = 10 + d * 3;
            screen.Remarks(day, TimeSpan.FromHours(8), (100, fixedCount - 1, 100 - fixedCount, 1));
            screen.Remarks(day, TimeSpan.FromHours(18), (100, fixedCount, 100 - fixedCount - 1, 1));
            screen.Load(day, new TimeSpan(20, 0, 0), day, new TimeSpan(21, 0, 0), 1, 2, 0, 0, 0, 1);
        }

        screen.SaveTo(_dir, name);
    }
}
