using ClosedXML.Excel;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Reporting;
using ConverterReport.Core.Screens;
using ConverterReport.Core.Workbook;
using DocumentFormat.OpenXml.Packaging;

namespace ConverterReport.Tests;

public sealed class WorkbookTests : IDisposable
{
    private static readonly DateTime ReportAt = new(2026, 10, 15, 8, 0, 0);
    private readonly string _dir = Directory.CreateTempSubdirectory("converter-report-wb-").FullName;

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

    [Fact]
    public void WorkbookHasSameStructureAsReportBook()
    {
        var path = Build(out var result, "3", "4", "9");

        using var wb = new XLWorkbook(path);
        var names = wb.Worksheets.Select(w => w.Name).ToList();
        Assert.Equal(new[] { "Описание", "Д.3", "З.3", "Д.4", "З.4", "Д.9", "З.9", "Свод", "Диаграммы", "15.10", "Отчет_дет", "Отчет" }, names);
        Assert.Equal("15.10", result.DailySheet);

        // Свод — формулы СУММЕСЛИМН по листам этапов, имена листов в кавычках.
        var svod = wb.Worksheet("Свод");
        Assert.Contains("SUMIFS('З.3'!$E$3:", svod.Cell(2, 24).FormulaA1);
        Assert.Contains("SUMIFS('Д.3'!$L$3:", svod.Cell(2, 30).FormulaA1);

        // Лист отчета: отчетная дата/время и ссылки на блок «Конвертер».
        var daily = wb.Worksheet("15.10");
        Assert.Equal("AE4+AF4", daily.Cell("A4").FormulaA1);
        Assert.Equal("AB9", daily.Cell("C9").FormulaA1);
        Assert.Contains("SUMPRODUCT", daily.Cell("AB9").FormulaA1);
        Assert.Contains(daily.CellsUsed(), c => c.HasFormula && c.FormulaA1.Contains("'З.3'!$E$"));

        // Отчет_дет ссылается на «Отчет», как в книге.
        Assert.StartsWith("Отчет!", wb.Worksheet("Отчет_дет").Cell(9, 2).FormulaA1);

        // Диаграммы: на каждый лист — устранение замечаний и загрузка документов; сводный лист + 3 этапа.
        using var doc = SpreadsheetDocument.Open(path, false);
        var chartsPart = doc.WorkbookPart!.WorksheetParts.Single(p => p.DrawingsPart is not null);
        Assert.Equal(8, chartsPart.DrawingsPart!.ChartParts.Count());
        Assert.Equal(4, result.ChartPages);
    }

    [Fact]
    public void SvodFormulasCalculateStatusPerDay()
    {
        var path = Build(out _, "3");

        using var wb = new XLWorkbook(path);
        var svod = wb.Worksheet("Свод");
        // Строка 2 — 01.10.2026 (первая строка экрана после начала сроков), 18:00.
        Assert.Equal(new DateTime(2026, 10, 1), svod.Cell(2, 2).GetDateTime());
        var block = 24; // X — «Всего» первого этапа
        Assert.Equal(100.0, svod.Cell(2, block).Value.GetNumber());
        Assert.Equal(10.0, svod.Cell(2, block + 1).Value.GetNumber());
        // 15.10 в 08:00 (время отчета)
        var reportRow = 2 + 14;
        Assert.Equal(60.0, svod.Cell(reportRow, block + 1).Value.GetNumber());
        // Загрузки за 14.10: 3 (ИИ 1 + ПД 2) + 6 (вечерняя загрузка)
        Assert.Equal(9.0, svod.Cell(reportRow - 1, 30 + 6).Value.GetNumber());
    }

    private string Build(out SummaryWorkbook result, params string[] numbers)
    {
        var config = new ReportConfig { BaseDirectory = _dir };
        foreach (var n in numbers)
        {
            var file = $"Экран {n}.xlsx";
            CreateScreen(file);
            config.Projects.Add(new ProjectConfig
            {
                Number = n,
                ScreenFile = file,
                Deadlines = { ["Экспертиза"] = new List<string> { "01.10.2026", "30.10.2026" } },
            });
        }

        var reader = new ScreenReader(config.ScreenLayout);
        var inputs = config.Projects.Select(p => new ProjectInput(p, reader.Read(Path.Combine(_dir, p.ScreenFile)))).ToList();
        var model = ReportBuilder.Build(config, inputs, ReportAt, ReportTime.PeriodStart(config.Report, ReportAt));
        var path = Path.Combine(_dir, "Свод.xlsx");
        result = SummaryWorkbookWriter.Write(model, path);
        return path;
    }

    private void CreateScreen(string name)
    {
        var screen = new TestScreen();
        var start = new DateTime(2026, 10, 1);
        for (var d = 0; d < 14; d++)
        {
            var day = start.AddDays(d);
            var fixedCount = 10 + d * 3;
            screen.Remarks(day, TimeSpan.FromHours(8), (100, fixedCount - 1, 100 - fixedCount, 1));
            screen.Remarks(day, TimeSpan.FromHours(18), (100, fixedCount, 100 - fixedCount - 1, 1));
            screen.Load(day, new TimeSpan(10, 0, 0), day, new TimeSpan(11, 0, 0), 1, 2, 0, 0, 0, 0);
        }

        screen.Load(new DateTime(2026, 10, 14), new TimeSpan(20, 0, 0), new DateTime(2026, 10, 14), new TimeSpan(20, 30, 0), 1, 1, 1, 1, 1, 1);
        screen.Remarks(new DateTime(2026, 10, 15), TimeSpan.FromHours(8), (100, 60, 39, 1));
        screen.SaveTo(_dir, name);
    }
}
