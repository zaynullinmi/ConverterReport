using ConverterReport.Core;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace ConverterReport.Tests;

public sealed class PdfTests : IDisposable
{
    private readonly string _dir = Directory.CreateTempSubdirectory("converter-report-tests-").FullName;

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
    public void BuildsSinglePdfWithDailyDetailedAndCharts_SkippingProjectsWithoutData()
    {
        var config = new ReportConfig { BaseDirectory = _dir };
        config.Report.OutputFolder = "out";
        // Шесть этапов с данными + один с пустым экраном (диаграмм по нему быть не должно) + один отсутствующий файл.
        for (var i = 1; i <= 6; i++)
        {
            CreateScreen($"Экран {i}.xlsx", scale: i);
            config.Projects.Add(new ProjectConfig
            {
                Number = i.ToString(),
                ScreenFile = $"Экран {i}.xlsx",
                Deadlines = { ["Экспертиза"] = new List<string> { "01.10.2026", "30.10.2026" } },
            });
        }

        new TestScreen().SaveTo(_dir, "Пустой.xlsx");
        config.Projects.Add(new ProjectConfig { Number = "7", ScreenFile = "Пустой.xlsx" });
        config.Projects.Add(new ProjectConfig { Number = "8", ScreenFile = "Нет такого файла.xlsx" });

        var result = new ReportGenerator().Run(config, new RunOptions
        {
            Date = new DateTime(2026, 10, 15),
            Time = TimeSpan.FromHours(8),
            PerProject = true,
            ExportExcel = true,
        });

        var main = result.Files[0];
        Assert.EndsWith("2026.10.15 08.00 Суточный отчет ГГЭ.pdf", main);
        Assert.True(File.Exists(main));

        using var pdf = PdfSharp.Pdf.IO.PdfReader.Open(main, PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
        // Суточный (может занять 1–2 страницы) + детальный + сводная диаграмма + 6 этапов с данными.
        var chartPages = pdf.Pages.Cast<PdfPage>().Count(p => p.Width.Point > p.Height.Point);
        Assert.Equal(7, chartPages);
        Assert.InRange(pdf.PageCount, 9, 11);

        // Отдельные PDF по этапам: только по этапам с данными (6 шт.), плюс Excel.
        Assert.Equal(1 + 6 + 1, result.Files.Count);
        Assert.Contains(result.Files, f => f.EndsWith("Суточный отчет ГГЭ 3 этап.pdf"));
        Assert.Contains(result.Files, f => f.EndsWith(".xlsx"));
        Assert.Contains(result.Warnings, w => w.StartsWith("Этап 8:"));
    }

    [Fact]
    public void ChartsOnlyReportHasNoTextPages()
    {
        CreateScreen("Экран.xlsx", scale: 1);
        var config = new ReportConfig { BaseDirectory = _dir };
        config.Report.IncludeDaily = false;
        config.Report.IncludeDetailed = false;
        config.Projects.Add(new ProjectConfig { Number = "3", ScreenFile = "Экран.xlsx" });

        var result = new ReportGenerator().Run(config, new RunOptions { Date = new DateTime(2026, 10, 15), Time = TimeSpan.FromHours(18) });

        using var pdf = PdfSharp.Pdf.IO.PdfReader.Open(result.Files[0], PdfSharp.Pdf.IO.PdfDocumentOpenMode.Import);
        Assert.Equal(1, pdf.PageCount);
    }

    [Fact]
    public void BlankPagesAreRemoved()
    {
        SystemFontResolver.Register();
        using var doc = new PdfDocument();
        doc.AddPage(); // пустая
        var page = doc.AddPage();
        using (var gfx = XGraphics.FromPdfPage(page))
            gfx.DrawString("Текст", new XFont("Arial", 10), XBrushes.Black, 50, 50);
        doc.AddPage(); // пустая

        Assert.True(BlankPageRemover.IsBlank(doc.Pages[0]));
        Assert.False(BlankPageRemover.IsBlank(doc.Pages[1]));
        Assert.Equal(2, BlankPageRemover.RemoveBlankPages(doc));
        Assert.Equal(1, doc.PageCount);
    }

    [Fact]
    public void NotesFileIsReadAndCommentsSkipped()
    {
        File.WriteAllText(Path.Combine(_dir, "Справочно.txt"), "// комментарий\nПервая строка\n\nВторая строка\n");
        var config = new ReportConfig { BaseDirectory = _dir };

        var notes = ReportGenerator.ReadNotes(config);

        Assert.Equal(new[] { "Первая строка", "Вторая строка" }, notes);
    }

    private void CreateScreen(string name, int scale)
    {
        var screen = new TestScreen();
        var start = new DateTime(2026, 10, 1);
        for (var d = 0; d < 14; d++)
        {
            var day = start.AddDays(d);
            var fixedCount = 10 * scale + d * 3 * scale;
            screen.Remarks(day, TimeSpan.FromHours(8), (100 * scale, fixedCount - scale, 100 * scale - fixedCount, scale));
            screen.Remarks(day, TimeSpan.FromHours(18), (100 * scale, fixedCount, 100 * scale - fixedCount - scale, scale));
            screen.Load(day, new TimeSpan(10, 0, 0), day, new TimeSpan(11, 0, 0), scale, 2 * scale, 0, 0, 0, 1);
        }

        screen.Load(new DateTime(2026, 10, 14), new TimeSpan(20, 0, 0), new DateTime(2026, 10, 14), new TimeSpan(20, 30, 0), 1, 1, 1, 1, 1, 1);
        screen.Remarks(new DateTime(2026, 10, 15), TimeSpan.FromHours(8), (100 * scale, 60 * scale, 39 * scale, scale));
        screen.SaveTo(_dir, name);
    }
}
