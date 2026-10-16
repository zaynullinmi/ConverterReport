using ClosedXML.Excel;
using ConverterReport.Core;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Workbook;
using DocumentFormat.OpenXml.Packaging;

namespace ConverterReport.Tests;

/// <summary>Вывод этапов в отчет, верх шкалы загрузок, оформление диаграмм и перенос листов из старой книги.</summary>
public sealed class OutputSettingsTests : IDisposable
{
    private readonly string _dir = Directory.CreateTempSubdirectory("converter-report-output-").FullName;

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
    public void ProjectNotShownIsOnlyInSummaryChart()
    {
        var config = Config("2", "3", "4");
        config.Projects[0].ShowInReport = false;
        var result = Run(config, 15, 8);

        // Имена файлов — по всем загруженным этапам, подписи отчета — по выводимым.
        Assert.EndsWith("Сводный отчет ГГЭ 2-4 этапы ВСМ-1 1.0.xlsx", result.Workbook);
        using (var wb = new XLWorkbook(result.Workbook!))
        {
            var daily = wb.Worksheet("15.10");
            Assert.Equal("по 3-4 этапам строительства ВСМ-1", daily.Cell("A3").GetString());
            Assert.Equal(new[] { 3.0, 4.0 }, new[] { daily.Cell("A9").GetDouble(), daily.Cell("A10").GetDouble() });
            Assert.Equal("по 3-4 этапам строительства ВСМ-1", wb.Worksheet("Отчет_дет").Cell("A3").GetString());
            Assert.True(wb.TryGetWorksheet("Д.2", out _));
            Assert.Contains(wb.Worksheet("Свод").Row(1).CellsUsed(), c => c.GetString() == "2 этап");
        }

        // Диаграммы: сводная + этапы 3 и 4 (у этапа 2 своей страницы нет), по 2 на странице.
        using var doc = SpreadsheetDocument.Open(result.Workbook!, false);
        var charts = doc.WorkbookPart!.WorksheetParts.Single(p => p.DrawingsPart is not null).DrawingsPart!.ChartParts.ToList();
        Assert.Equal(6, charts.Count);
    }

    [Fact]
    public void LoadsAxisMaxAndChartStyleAsInReportBook()
    {
        var config = Config("3", "4");
        config.Charts.LoadsAxisMaxSummary = 500;
        config.Charts.LoadsAxisMaxProject = 250;
        config.Projects[1].LoadsAxisMax = 120;
        config.Projects[0].ExpertiseEndDate = "20.10.2026";
        var result = Run(config, 15, 8);

        using var doc = SpreadsheetDocument.Open(result.Workbook!, false);
        var xml = doc.WorkbookPart!.WorksheetParts.Single(p => p.DrawingsPart is not null).DrawingsPart!.ChartParts
            .Select(p => new StreamReader(p.GetStream()).ReadToEnd())
            .ToList();
        var loads = xml.Where(x => x.Contains("clustered")).ToList();
        var remarks = xml.Where(x => x.Contains("stacked")).ToList();
        Assert.Equal(3, loads.Count);
        Assert.Contains("<c:max val=\"500\"", loads[0]);
        Assert.Contains("<c:max val=\"250\"", loads[1]);
        Assert.Contains("<c:max val=\"120\"", loads[2]);
        Assert.All(loads, x => Assert.Contains("<c:gapWidth val=\"150\"", x));

        // Сводная — без линий тренда, этапы — с ними; шкала замечаний от 0; подписи 9 пт.
        Assert.DoesNotContain("trendline", remarks[0]);
        Assert.Contains("trendline", remarks[1]);
        Assert.All(remarks, x => Assert.Contains("<c:min val=\"0\"", x));
        Assert.All(remarks, x => Assert.Contains("sz=\"900\"", x));
        // Окончание экспертизы этапа 3 на оси — вертикальная линия; у сводной ее нет.
        Assert.Contains("dropLines", remarks[1]);
        Assert.DoesNotContain("dropLines", remarks[0]);

        using var wb = new XLWorkbook(result.Workbook!);
        var setup = wb.Worksheet("Диаграммы").PageSetup;
        Assert.Equal(1, setup.PagesWide);
        Assert.Equal(0, setup.PagesTall);
    }

    [Fact]
    public void ColumnWidthsAreAsInReportBook()
    {
        var result = Run(Config("3"), 15, 8);
        using var doc = SpreadsheetDocument.Open(result.Workbook!, false);
        var wbPart = doc.WorkbookPart!;
        var sheet = wbPart.Workbook.Descendants<DocumentFormat.OpenXml.Spreadsheet.Sheet>().Single(s => s.Name == "15.10");
        var part = (WorksheetPart)wbPart.GetPartById(sheet.Id!);
        var first = part.Worksheet.Descendants<DocumentFormat.OpenXml.Spreadsheet.Column>().First(c => c.Min!.Value == 1);
        Assert.Equal(4.66, first.Width!.Value, 2);
    }

    [Fact]
    public void SvodUsesLastFilledRowWhenReportTimeRowIsMissing()
    {
        // В экране последняя строка — 15.10 08:00, отчет на 18:00: «Свод» на 15.10 не должен «проваливаться».
        var result = Run(Config("3", "4"), 15, 18);
        using var wb = new XLWorkbook(result.Workbook!);
        var svod = wb.Worksheet("Свод");
        var row = 2 + (new DateTime(2026, 10, 15) - svod.Cell(2, 2).GetDateTime()).Days;
        Assert.Equal(200.0, svod.Cell(row, 6).Value.GetNumber()); // «Всего» по двум этапам
        Assert.Equal(120.0, svod.Cell(row, 7).Value.GetNumber()); // «Устранено» — строка 08:00
    }

    [Theory]
    [InlineData(18, ReportLoadsMode.Auto, 20)]   // вечер: все загрузки за день (с 00:00) — как на диаграмме
    [InlineData(8, ReportLoadsMode.Auto, 14)]    // утро: с 18:00 прошлого дня
    [InlineData(18, ReportLoadsMode.Period, 25)] // всегда с начала периода
    [InlineData(8, ReportLoadsMode.Day, 9)]      // всегда за день
    public void ReportLoadsAreForDayInEveningAndSincePeriodStartInMorning(int hour, ReportLoadsMode mode, int expected)
    {
        var d14 = new DateTime(2026, 10, 14);
        var d15 = new DateTime(2026, 10, 15);
        new TestScreen()
            .Remarks(d14, TimeSpan.FromHours(18), (100, 50, 49, 1))
            .Remarks(d15, TimeSpan.FromHours(8), (100, 51, 48, 1))
            .Load(d14, new TimeSpan(10, 0, 0), d14, new TimeSpan(11, 0, 0), 300)          // до начала периода
            .Load(d14, new TimeSpan(20, 0, 0), d14, new TimeSpan(22, 0, 0), 5)            // прошлый день после 18:00
            .Load(d14, new TimeSpan(23, 30, 0), d15, TimeSpan.Zero, 2)                    // окончание ровно в 00:00
            .Load(d15, new TimeSpan(7, 0, 0), d15, new TimeSpan(7, 30, 0), 0, 7)
            .Load(d15, new TimeSpan(12, 0, 0), d15, new TimeSpan(13, 0, 0), 0, 0, 11)
            .Load(d15, new TimeSpan(18, 30, 0), d15, new TimeSpan(19, 0, 0), 1000)        // после отчета
            .SaveTo(_dir, "Экран 7.xlsx");
        var config = Config();
        config.Projects.Add(new ProjectConfig { Number = "7", ScreenFile = "Экран 7.xlsx" });
        config.Report.ReportLoads = mode;
        var result = Run(config, 15, hour);

        Assert.Equal(expected, result.Model.Projects.Single().PeriodLoadsSum.Total);

        using var wb = new XLWorkbook(result.Workbook!);
        var daily = wb.Worksheet("15.10");
        var forDay = expected is 20 or 9;
        // Начало выборки (AE5) и признак «за день» (AG5) — формулы от дат отчета AE3:AF4.
        Assert.Equal(forDay ? 1.0 : 0.0, daily.Cell(5, 33).Value.GetNumber());
        Assert.Equal((forDay ? d15 : d14).ToOADate(), daily.Cell(5, 31).Value.GetNumber());
        Assert.Contains("$AG$5", daily.Cell(9, 28).FormulaA1);
        Assert.Equal(expected, daily.Cell(10, 24).Value.GetNumber()); // «Всего:» раздела 1

        // Вечером сумма раздела 1 совпадает с загрузками дня на диаграмме («Свод»).
        if (hour == 18 && mode == ReportLoadsMode.Auto)
        {
            var svod = wb.Worksheet("Свод");
            var row = 2 + (d15 - svod.Cell(2, 2).GetDateTime()).Days;
            Assert.Equal(expected, svod.Cell(row, SvodLayout.Project(0).LoadTotal).Value.GetNumber());
        }
    }

    [Fact]
    public void ChartsShowNaAsEmptyCellLikeReportBook()
    {
        var result = Run(Config("3", "4"), 15, 8);
        using var doc = SpreadsheetDocument.Open(result.Workbook!, false);
        var xml = doc.WorkbookPart!.WorksheetParts.Single(p => p.DrawingsPart is not null).DrawingsPart!.ChartParts
            .Select(p => new StreamReader(p.GetStream()).ReadToEnd())
            .ToList();
        Assert.NotEmpty(xml);
        Assert.All(xml, x => Assert.Contains("<c16r3:dispNaAsBlank val=\"1\"", x));
    }

    [Fact]
    public void OldDefaultFileNamesAreUpgraded()
    {
        var path = Path.Combine(_dir, "old.json");
        File.WriteAllText(path, """
            { "Report": { "FileName": "{date:yyyy.MM.dd} {time:H.mm} Суточный {projectsShort} этапы {version}",
                          "DetailedFileName": "{date:yyyy.MM.dd} {time:H.mm} Суточный {projectsShort} этапы {version} (дет)",
                          "ChartsFileName": "Мои диаграммы" } }
            """);
        var config = ConfigLoader.Load(path);
        Assert.Equal("{date:yyyy.MM.dd} {time:H.mm} Свод ГГЭ {projectsShort} этапы {version}", config.Report.FileName);
        Assert.Equal("{date:yyyy.MM.dd} {time:H.mm} Свод ГГЭ {projectsShort} этапы {version} Отчет", config.Report.DetailedFileName);
        Assert.Equal("Мои диаграммы", config.Report.ChartsFileName);
    }

    [Theory]
    [InlineData(0xFF, 0xC0, 0x00, -0.5, 0x7F, 0x60, 0x00)] // «Золотистый, темнее 50%» — цвет «Всего»
    [InlineData(0x70, 0xAD, 0x47, 0.8, 0xE2, 0xEF, 0xDA)]  // «Зеленый, светлее 80%» — текущая стадия
    public void ThemeTintAsInExcel(int r, int g, int b, double tint, int er, int eg, int eb)
    {
        // Excel округляет оттенки по-своему — расхождение не больше единицы на канал (на глаз не видно).
        var (ar, ag, ab) = LegacyWorkbook.ApplyTint(r, g, b, tint);
        Assert.InRange(ar, er - 1, er + 1);
        Assert.InRange(ag, eg - 1, eg + 1);
        Assert.InRange(ab, eb - 1, eb + 1);
    }

    [Fact]
    public void PastDaysAndDeadlinesAreTakenFromOldBook()
    {
        var old = Path.Combine(_dir, "Старая книга.xlsx");
        using (var wb = new XLWorkbook())
        {
            foreach (var (name, day) in new[] { ("13.10", 13), ("14.10", 14) })
            {
                var ws = wb.AddWorksheet(name);
                ws.Cell("AE4").Value = new DateTime(2026, 10, day);
                ws.Cell("A1").Value = "ОТЧЕТ";
                ws.Cell("C10").Value = 398;
                ws.Cell("C10").Style.Font.FontColor = XLColor.FromTheme(XLThemeColor.Accent4, -0.5);
                ws.Cell("D10").FormulaA1 = "C10/2";
            }

            var last = wb.Worksheet("14.10");
            last.Cell("A40").Value = "4. Сроки.";
            last.Range("A42:B43").Merge().Value = "Этап";
            last.Range("C42:H42").Merge().Value = "Экспертиза";
            last.Cell("C43").Value = "Начало";
            last.Cell("F43").Value = "Окончание";
            last.Cell("A44").Value = 4;
            last.Cell("C44").Value = new DateTime(2026, 10, 1);
            last.Cell("F44").Value = new DateTime(2026, 10, 30);
            last.Cell("A45").Value = 3;
            last.Range("C45:H45").Merge().Value = "Заявка до 23.10.2026";
            wb.SaveAs(old);
        }

        var config = Config("3", "4");
        foreach (var p in config.Projects)
            p.Deadlines.Clear();
        config.Report.WorkbookMode = WorkbookMode.Append;
        config.Report.ImportLegacyDays = true;
        config.Report.ArchiveWorkbooks.Add(new ArchiveSource { Path = old });
        var result = Run(config, 15, 8);

        Assert.Equal(new[] { "01.10.2026", "30.10.2026" }, result.FilledDeadlines["4"]["Экспертиза"]);
        Assert.Equal(new[] { "Заявка до 23.10.2026" }, result.FilledDeadlines["3"]["Экспертиза"]);
        Assert.Equal(new[] { "01.10.2026", "30.10.2026" }, config.Projects[1].Deadlines["Экспертиза"]);

        using (var master = new XLWorkbook(result.Workbook!))
        {
            Assert.Equal(new[] { "13.10", "14.10", "15.10" }, master.Worksheets.Select(w => w.Name).Where(SummaryWorkbookBuilder.IsDailySheet));
            var copy = master.Worksheet("14.10");
            Assert.False(copy.Cell("D10").HasFormula);
            Assert.Equal(199.0, copy.Cell("D10").GetDouble());
            Assert.Equal(XLColorType.Color, copy.Cell("C10").Style.Font.FontColor.ColorType);
            // Цвет темы старой книги (тема ClosedXML: Accent4 = 8064A2), затемненный на 50%, — явным цветом.
            var color = copy.Cell("C10").Style.Font.FontColor.Color;
            var expected = LegacyWorkbook.ApplyTint(0x80, 0x64, 0xA2, -0.5);
            Assert.Equal(expected, (color.R, color.G, color.B));
        }

        // Повторно ничего не переносится, и листы не дублируются.
        var again = Run(config, 15, 18);
        using var after = new XLWorkbook(again.Workbook!);
        Assert.Equal(3, after.Worksheets.Count(w => SummaryWorkbookBuilder.IsDailySheet(w.Name)));
    }

    private ReportConfig Config(params string[] numbers)
    {
        var config = new ReportConfig { BaseDirectory = _dir };
        config.Report.OutputFolder = "out";
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

    private static RunResult Run(ReportConfig config, int day, int hour) =>
        new ReportGenerator().Run(config, new RunOptions { Date = new DateTime(2026, 10, day), Time = TimeSpan.FromHours(hour) });

    private void CreateScreen(string name)
    {
        if (File.Exists(Path.Combine(_dir, name)))
            return;
        var screen = new TestScreen();
        var start = new DateTime(2026, 10, 1);
        for (var d = 0; d < 14; d++)
        {
            var day = start.AddDays(d);
            var fixedCount = 10 + d * 3;
            screen.Remarks(day, TimeSpan.FromHours(8), (100, fixedCount - 1, 100 - fixedCount, 1));
            screen.Remarks(day, TimeSpan.FromHours(18), (100, fixedCount, 100 - fixedCount - 1, 1));
            screen.Load(day, new TimeSpan(10, 0, 0), day, new TimeSpan(11, 0, 0), 100, 200, 0, 0, 0, 1);
        }

        screen.Remarks(new DateTime(2026, 10, 15), TimeSpan.FromHours(8), (100, 60, 39, 1));
        screen.SaveTo(_dir, name);
    }
}
