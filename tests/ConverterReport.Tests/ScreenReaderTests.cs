using ConverterReport.Core.Configuration;
using ConverterReport.Core.Screens;

namespace ConverterReport.Tests;

public class ScreenReaderTests
{
    private static readonly TimeSpan T08 = TimeSpan.FromHours(8);
    private static readonly TimeSpan T18 = TimeSpan.FromHours(18);

    [Fact]
    public void ReadsFilledRemarksRowsOnly()
    {
        var screen = new TestScreen()
            .Remarks(new DateTime(2026, 10, 14), T08, (398, 350, 47, 1))
            .Remarks(new DateTime(2026, 10, 14), T18, (398, 360, 37, 1))
            .EmptyRemarksRow(new DateTime(2026, 10, 15), T08)
            .EmptyRemarksRow(new DateTime(2026, 10, 15), T18);

        var data = new ScreenReader(new ScreenLayout()).Read(screen.ToStream(), "test.xlsx");

        Assert.Equal(2, data.Remarks.Count);
        var last = data.Remarks[^1];
        Assert.Equal(new DateTime(2026, 10, 14, 18, 0, 0), last.Timestamp);
        Assert.Equal(398, last.All[0].Total);
        Assert.Equal(360, last.All[0].Fixed);
        Assert.Equal(37, last.All[0].Fixing);
        Assert.Equal(1, last.All[0].Review);
        Assert.True(last.SectionsMatch(critical: false));
        Assert.True(last.SectionsMatch(critical: true));
        Assert.Empty(data.Warnings);
    }

    [Fact]
    public void ReadsLoadsWithMixedDateFormats()
    {
        var screen = new TestScreen()
            .Load(new DateTime(2026, 1, 26), new TimeSpan(21, 34, 0), new DateTime(2026, 1, 27), new TimeSpan(2, 20, 0), 407, 1104, 2462, 1924, 233, 285)
            // Время записано датой со временем (так бывает после копирования).
            .Load(new DateTime(2026, 2, 20), new DateTime(2026, 2, 20, 10, 53, 0), new DateTime(2026, 2, 20), new DateTime(2026, 2, 20, 11, 16, 0), 0, 15, 2, 0, 0, 0, 13)
            // Дата и время строками.
            .Load("10.10.2026", "18:56", "10.10.2026", "19:00", 0, 3, 1, 0, 0, 2)
            // Загрузка еще не выполнена — пропускается.
            .Load("", "", "", "", 0, 0, 0, 0, 0, 0)
            .TotalsRow();

        var data = new ScreenReader(new ScreenLayout()).Read(screen.ToStream(), "test.xlsx");

        Assert.Equal(3, data.Loads.Count);
        Assert.Equal(new DateTime(2026, 1, 27, 2, 20, 0), data.Loads[0].End);
        Assert.Equal(6415, data.Loads[0].Counts.Total);
        Assert.Equal(new DateTime(2026, 2, 20, 11, 16, 0), data.Loads[1].End);
        Assert.Equal(13, data.Loads[1].Counts.Answers);
        Assert.Equal(new DateTime(2026, 10, 10, 18, 56, 0), data.Loads[2].Start);
        Assert.Equal(6, data.Loads[2].Counts.Total);
    }

    [Fact]
    public void MissingSheetsProduceWarnings()
    {
        using var wb = new ClosedXML.Excel.XLWorkbook();
        wb.AddWorksheet("Лист1");
        using var ms = new MemoryStream();
        wb.SaveAs(ms);
        ms.Position = 0;

        var data = new ScreenReader(new ScreenLayout()).Read(ms, "empty.xlsx");

        Assert.Empty(data.Remarks);
        Assert.Empty(data.Loads);
        Assert.Equal(2, data.Warnings.Count);
    }
}
