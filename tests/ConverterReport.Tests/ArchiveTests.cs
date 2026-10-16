using ClosedXML.Excel;
using ConverterReport.Core;
using ConverterReport.Core.Archive;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Workbook;
using DocumentFormat.OpenXml.Packaging;

namespace ConverterReport.Tests;

/// <summary>
/// Прошлые версии отчета (архив), лист отчета на любую дату, построчные загрузки, настройки диаграмм на листе.
/// </summary>
public sealed class ArchiveTests : IDisposable
{
    private static readonly DateTime D14 = new(2026, 10, 14);
    private static readonly DateTime D15 = new(2026, 10, 15);
    private readonly string _dir = Directory.CreateTempSubdirectory("converter-report-archive-").FullName;

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
    public void ArchiveSheetsAreReadWithSvodLabels()
    {
        var path = OldBook();
        var book = ArchiveReader.Read(path, new ScreenLayout());

        Assert.Equal(new[] { "1", "5" }, book.Parts.Select(p => p.Key).OrderBy(k => k));
        var five = book.Parts.Single(p => p.Key == "5");
        Assert.Equal("5", five.SuggestedStage); // в «Своде» блок подписан «5 этап» — на самом деле этап 9 указывает пользователь
        Assert.Contains("5 этап", five.Note);
        Assert.Equal(3, five.Data.Remarks.Count);
        Assert.Equal(new DateTime(2026, 2, 1), five.From!.Value.Date);
        Assert.Equal(new[] { "01.02" }, book.DailySheets);
    }

    [Fact]
    public void SelectedArchiveStagesAreMergedAndAddedAsArchive()
    {
        var config = Config();
        config.Report.ImportLegacyDays = true;
        config.Report.ArchiveWorkbooks.Add(new ArchiveSource { Path = OldBook(), Stages = { ["5"] = "9", ["1"] = "1" } });
        var result = Run(config, 15, 18);

        // Этап 1 — архивный (без экрана): в настройках, в «Своде», но не в отчете.
        var added = Assert.Single(result.AddedStages);
        Assert.Equal("1", added.Number);
        Assert.True(added.IsArchive);
        Assert.False(added.ShowInReport);
        Assert.Equal(new[] { "1", "9" }, config.Projects.Select(p => p.Number));

        // Этап 9: строки прошлого цикла (листы Д.5/З.5) — раньше строк экрана.
        var nine = result.Model.Projects.Single(p => p.Number == "9");
        Assert.Equal(new DateTime(2026, 2, 1, 18, 0, 0), nine.Screen!.Remarks[0].Timestamp);
        Assert.Equal(D15.AddHours(18), nine.Screen.Remarks[^1].Timestamp);
        Assert.Equal("9", result.Model.ProjectsList); // архивный этап в имена файлов не входит

        using var wb = new XLWorkbook(result.Workbook!);
        Assert.True(wb.TryGetWorksheet("З.1", out _));
        var svod = wb.Worksheet("Свод");
        Assert.Equal("1 этап", svod.Cell(1, SvodLayout.Project(0).Label).GetString());
        Assert.Equal("9 этап", svod.Cell(1, SvodLayout.Project(1).Label).GetString());
        var start = svod.Cell(2, 2).GetDateTime();
        Assert.Equal(new DateTime(2026, 2, 1), start);
        var feb2 = 2 + (new DateTime(2026, 2, 2) - start).Days;
        Assert.Equal(510.0, svod.Cell(feb2, SvodLayout.Project(1).Total).Value.GetNumber()); // этап 9 из архива
        Assert.Equal(1010.0, svod.Cell(feb2, SvodLayout.Overall.Total).Value.GetNumber()); // итог 1 + 9
        // На дату отчета закончившийся в феврале этап 1 в «Своде» не участвует.
        var report = 2 + (D15 - start).Days;
        Assert.True(svod.Cell(report, SvodLayout.Project(0).Total).Value.IsError);
        Assert.Equal(200.0, svod.Cell(report, SvodLayout.Overall.Total).Value.GetNumber());

        // В листе отчета — только этап 9; листы прошлых дней в отдельной книге не переносятся.
        var daily = wb.Worksheet("15.10");
        Assert.Equal(9.0, daily.Cell("AA9").Value.GetNumber());
        Assert.Equal("Всего:", daily.Cell("J10").GetString());
        Assert.False(wb.TryGetWorksheet("01.02", out _));
    }

    [Fact]
    public void NothingIsTakenWithoutSelectedStages()
    {
        var config = Config();
        config.Report.ImportLegacyDays = true;
        config.Report.WorkbookMode = WorkbookMode.Append;
        config.Report.ArchiveWorkbooks.Add(new ArchiveSource { Path = OldBook() });
        var result = Run(config, 15, 18);

        Assert.Empty(result.AddedStages);
        Assert.Equal(D14.AddHours(8), result.Model.Projects.Single().Screen!.Remarks[0].Timestamp);
        using var wb = new XLWorkbook(result.Workbook!);
        Assert.True(wb.TryGetWorksheet("01.02", out _)); // листы прошлых дней — всегда
        Assert.False(wb.TryGetWorksheet("З.1", out _));
    }

    [Fact]
    public void ArchiveDataStaysInMasterWithoutArchiveBook()
    {
        var config = Config();
        config.Report.WorkbookMode = WorkbookMode.Append;
        config.Report.ImportLegacyDays = true;
        var old = OldBook();
        config.Report.ArchiveWorkbooks.Add(new ArchiveSource { Path = old, Stages = { ["5"] = "9" } });
        Run(config, 14, 18);

        // Прошлую версию отчета убрали — данные прошлого цикла этапа 9 остаются в общей книге.
        File.Delete(old);
        config.Report.ImportLegacyDays = false;
        var result = Run(config, 15, 18);
        Assert.Equal(new DateTime(2026, 2, 1, 18, 0, 0), result.Model.Projects.Single().Screen!.Remarks[0].Timestamp);
    }

    [Fact]
    public void DailySheetIsRecalculatedForAnotherDate()
    {
        var config = Config();
        var result = Run(config, 15, 18);
        using var wb = new XLWorkbook(result.Workbook!);
        var daily = wb.Worksheet("15.10");
        var converter = daily.Column(27).CellsUsed().First(c => c.GetString() == "III.").Address.RowNumber + 3;
        Assert.Equal(200.0, daily.Cell(converter, 29).Value.GetNumber()); // «Всего» на 15.10 18:00
        Assert.Equal(160.0, daily.Cell(converter, 30).Value.GetNumber()); // «Устранено»

        Assert.Equal(150.0, daily.Cell(converter, 40).Value.GetNumber()); // на начало периода — 14.10 18:00

        // В книге меняют отчетную дату — лист берет строку экрана на 14.10 18:00, начало периода — 13.10 18:00 (строк нет).
        daily.Cell("AE4").Value = D14;
        Assert.Equal("AE4-1", daily.Cell("AE3").FormulaA1);
        Assert.Equal(150.0, daily.Cell(converter, 30).Value.GetNumber());
        Assert.Equal(0.0, daily.Cell(converter, 40).Value.GetNumber());
        Assert.Equal("18:00 13.10.2026", daily.Cell("AH3").Value.GetText());
    }

    [Fact]
    public void LoadsCanBeListedOneByOne()
    {
        var config = Config();
        config.Report.LoadsMode = LoadsMode.Each;
        var result = Run(config, 15, 18);
        using var wb = new XLWorkbook(result.Workbook!);
        var daily = wb.Worksheet("15.10");

        // За день у этапа три загрузки — три строки: этап | 1, 2, 3.
        for (var k = 1; k <= 3; k++)
        {
            var r = 8 + k;
            Assert.Equal(9.0, daily.Cell(r, 1).Value.GetNumber());
            Assert.Equal($"IF($AQ{r}=0,\"\",{k})", daily.Cell(r, 28).FormulaA1);
            Assert.Contains($"AGGREGATE(15,6,ROW(", daily.Cell(r, 43).FormulaA1);
        }

        Assert.Equal("Всего:", daily.Cell(12, 10).GetString());
    }

    [Fact]
    public void StaleMarkCanBeTurnedOff()
    {
        var config = Config();
        config.Report.MarkStaleData = false;
        var result = Run(config, 15, 12); // строки на 12:00 нет — взята 08:00
        Assert.True(result.Model.Projects.Single().IsStale(D15.AddHours(12)));
        using var wb = new XLWorkbook(result.Workbook!);
        var daily = wb.Worksheet("15.10");
        Assert.DoesNotContain(daily.CellsUsed(), c => c.HasFormula && c.FormulaA1.Contains("\"9*\""));
        Assert.DoesNotContain(daily.CellsUsed(), c => c.HasFormula && c.FormulaA1.Contains("данные экрана на"));
    }

    [Fact]
    public void ChartsUseNamedRangesAndKeepManualAxisMax()
    {
        var config = Config();
        config.Report.WorkbookMode = WorkbookMode.Append;
        config.Charts.LoadsAxisMaxProject = 250;
        var result = Run(config, 14, 18);

        using (var wb = new XLWorkbook(result.Workbook!))
        {
            var charts = wb.Worksheet(SummaryWorkbook.ChartsSheet);
            Assert.Contains(charts.DefinedNames, n => n.Name == "p1_cat");
            Assert.Contains(charts.DefinedNames, n => n.Name == "p1_end");
            var axis = charts.Column(27).CellsUsed().First(c => c.GetString().StartsWith("Верх шкалы загрузок", StringComparison.Ordinal));
            Assert.Equal(250.0, axis.CellRight().Value.GetNumber());
            axis.CellRight().Value = 120; // изменено вручную в общей книге
            wb.Save();
        }

        result = Run(config, 15, 18);
        using (var doc = SpreadsheetDocument.Open(result.Workbook!, false))
        {
            var xml = doc.WorkbookPart!.WorksheetParts.Single(p => p.DrawingsPart is not null).DrawingsPart!.ChartParts
                .Select(p => new StreamReader(p.GetStream()).ReadToEnd()).ToList();
            Assert.Contains(xml, x => x.Contains("<c:f>Диаграммы!p1_cat</c:f>"));
            Assert.Contains(xml, x => x.Contains("<c:max val=\"120\""));
            Assert.Contains(xml, x => x.Contains("<c:f>Диаграммы!p1_end</c:f>") && x.Contains("dropLines"));
            Assert.Contains(xml, x => System.Text.RegularExpressions.Regex.IsMatch(x, @"<c:title>\s*<c:tx>\s*<c:strRef>\s*<c:f>Диаграммы!\$AB\$"));
        }

        using var master = new XLWorkbook(result.Workbook!);
        var sheet = master.Worksheet(SummaryWorkbook.ChartsSheet);
        var cell = sheet.Column(27).CellsUsed().First(c => c.GetString().StartsWith("Верх шкалы загрузок", StringComparison.Ordinal));
        Assert.Equal(120.0, cell.CellRight().Value.GetNumber());
        Assert.Equal(250.0, cell.CellRight().CellRight().Value.GetNumber());
    }

    // ------------------------------------------------------------------ данные

    /// <summary>Экран этапа 9: 14.10 08:00 и 18:00, 15.10 08:00 и 18:00; загрузки 15.10.</summary>
    private ReportConfig Config()
    {
        new TestScreen()
            .Remarks(D14, TimeSpan.FromHours(8), (200, 140, 59, 1))
            .Remarks(D14, TimeSpan.FromHours(18), (200, 150, 49, 1))
            .Remarks(D15, TimeSpan.FromHours(8), (200, 155, 44, 1))
            .Remarks(D15, TimeSpan.FromHours(18), (200, 160, 39, 1))
            .Load(D15, new TimeSpan(9, 0, 0), D15, new TimeSpan(9, 30, 0), 1)
            .Load(D15, new TimeSpan(11, 0, 0), D15, new TimeSpan(11, 30, 0), 0, 2)
            .Load(D15, new TimeSpan(16, 0, 0), D15, new TimeSpan(16, 30, 0), 0, 0, 3)
            .SaveTo(_dir, "Экран 9.xlsx");
        var config = new ReportConfig { BaseDirectory = _dir };
        config.Report.OutputFolder = "out";
        config.Report.PdfEngine = PdfEngine.Builtin;
        config.Report.IncludeDaily = false;
        config.Report.IncludeDetailed = false;
        config.Report.IncludeCharts = false;
        config.Projects.Add(new ProjectConfig
        {
            Number = "9",
            ScreenFile = "Экран 9.xlsx",
            Deadlines = { ["Экспертиза"] = new List<string> { "01.10.2026", "30.10.2026" } },
        });
        return config;
    }

    /// <summary>
    /// Прошлая версия отчета: листы Д.1/З.1 (этап 1) и Д.5/З.5 (прошлый цикл этапа 9, в «Своде» подписан «5 этап»),
    /// лист дня «01.02».
    /// </summary>
    private string OldBook()
    {
        var path = Path.Combine(_dir, "Старая версия 1.13.xlsx");
        var feb = new DateTime(2026, 2, 1);
        var stage1 = Path.Combine(_dir, "s1.xlsx");
        var stage5 = Path.Combine(_dir, "s5.xlsx");
        new TestScreen()
            .Remarks(feb, TimeSpan.FromHours(18), (500, 100, 390, 10))
            .Remarks(feb.AddDays(1), TimeSpan.FromHours(18), (500, 110, 380, 10))
            .Load(feb, new TimeSpan(10, 0, 0), feb, new TimeSpan(11, 0, 0), 7)
            .SaveTo(_dir, "s1.xlsx");
        new TestScreen()
            .Remarks(feb, TimeSpan.FromHours(18), (510, 0, 510, 0))
            .Remarks(feb.AddDays(1), TimeSpan.FromHours(18), (510, 20, 480, 10))
            .Remarks(feb.AddDays(2), TimeSpan.FromHours(18), (510, 40, 460, 10))
            .Load(feb, new TimeSpan(12, 0, 0), feb, new TimeSpan(13, 0, 0), 0, 9)
            .SaveTo(_dir, "s5.xlsx");

        using var wb = new XLWorkbook();
        foreach (var (file, key) in new[] { (stage1, "1"), (stage5, "5") })
        {
            using var src = new XLWorkbook(file);
            src.Worksheet("Замечания").CopyTo(wb, "З." + key);
            src.Worksheet("Документы").CopyTo(wb, "Д." + key);
        }

        var svod = wb.AddWorksheet("Свод");
        svod.Cell(1, 2).Value = "Дата";
        svod.Cell(1, 41).Value = "5 этап";
        svod.Cell(2, 2).Value = feb;
        svod.Cell(2, 24).FormulaA1 = "SUMIFS('З.1'!$E$2:$E$100,'З.1'!$B$2:$B$100,B2)";
        svod.Cell(2, 42).FormulaA1 = "SUMIFS('З.5'!$E$2:$E$100,'З.5'!$B$2:$B$100,B2)";
        var day = wb.AddWorksheet("01.02");
        day.Cell("A1").Value = "ОТЧЕТ";
        wb.SaveAs(path);
        return path;
    }

    private static RunResult Run(ReportConfig config, int day, int hour) =>
        new ReportGenerator().Run(config, new RunOptions { Date = new DateTime(2026, 10, day), Time = TimeSpan.FromHours(hour) });
}
