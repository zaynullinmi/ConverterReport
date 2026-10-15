using ConverterReport.Core;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Reporting;

namespace ConverterReport.Tests;

public class FormatsTests
{
    [Theory]
    [InlineData("10.10.2026", 2026, 10, 10)]
    [InlineData(" 1.2.2026 ", 2026, 2, 1)]
    [InlineData("18.02.2026", 2026, 2, 18)]
    [InlineData("2026-10-12", 2026, 10, 12)]
    public void ParseDate_Strings(string text, int y, int m, int d)
    {
        Assert.Equal(new DateTime(y, m, d), Formats.ParseDate(text)!.Value.Date);
    }

    [Fact]
    public void ParseDate_ExcelNumberAndDateTime()
    {
        Assert.Equal(new DateTime(2026, 2, 20), Formats.ParseDate(46073.0));
        Assert.Equal(new DateTime(2026, 2, 20, 10, 53, 0), Formats.ParseDate(new DateTime(2026, 2, 20, 10, 53, 0)));
        Assert.Null(Formats.ParseDate("-"));
        Assert.Null(Formats.ParseDate(""));
        Assert.Null(Formats.ParseDate(new DateTime(1899, 12, 30, 8, 0, 0))); // в ячейке только время
    }

    [Theory]
    [InlineData("18:56", 18, 56)]
    [InlineData("9:05", 9, 5)]
    [InlineData("22.36", 22, 36)]
    public void ParseTime_Strings(string text, int h, int m)
    {
        Assert.Equal(new TimeSpan(h, m, 0), Formats.ParseTime(text));
    }

    [Fact]
    public void ParseTime_NumbersDatesAndSpans()
    {
        // 0.33333333333333326 — так Excel хранит 08:00 (с погрешностью).
        Assert.Equal(new TimeSpan(8, 0, 0), Formats.ParseTime(0.33333333333333326));
        // Дата со временем в колонке «Время» — берется только время.
        Assert.Equal(new TimeSpan(10, 53, 0), Formats.ParseTime(new DateTime(2026, 2, 20, 10, 53, 0)));
        Assert.Equal(new TimeSpan(11, 16, 0), Formats.ParseTime(46073.46944444444));
        Assert.Equal(new TimeSpan(18, 0, 0), Formats.ParseTime(new TimeSpan(18, 0, 0)));
        Assert.Null(Formats.ParseTime("-"));
    }

    [Fact]
    public void Combine_UsesTimeFromDateWhenTimeIsEmpty()
    {
        Assert.Equal(new DateTime(2026, 10, 10, 18, 56, 0), Formats.Combine("10.10.2026", "18:56"));
        Assert.Equal(new DateTime(2026, 2, 20, 10, 53, 0), Formats.Combine(new DateTime(2026, 2, 20), new DateTime(2026, 2, 20, 10, 53, 0)));
        Assert.Equal(new DateTime(2026, 10, 10, 7, 0, 0), Formats.Combine("10.10.2026 07:00", null));
        Assert.Null(Formats.Combine(null, "18:00"));
    }

    [Theory]
    [InlineData(new[] { "2", "3", "4", "9" }, "2-4, 9")]
    [InlineData(new[] { "3", "4", "9" }, "3-4, 9")]
    [InlineData(new[] { "1", "2", "3", "4", "5" }, "1-5")]
    [InlineData(new[] { "3" }, "3")]
    [InlineData(new[] { "1", "3", "5" }, "1, 3, 5")]
    [InlineData(new[] { "2", "2А", "3" }, "2, 2А, 3")]
    public void CompressNumbers(string[] numbers, string expected)
    {
        Assert.Equal(expected, Formats.CompressNumbers(numbers));
    }

    [Fact]
    public void PercentAndNumbers()
    {
        Assert.Equal("90,5", Formats.Percent(360, 398));
        Assert.Equal("100", Formats.Percent(5, 5));
        Assert.Equal("0", Formats.Percent(0, 5));
        Assert.Equal("-", Formats.Percent(1, 0));
        Assert.Equal("1 367", Formats.Number(1367));
    }

    [Fact]
    public void ApplyTemplate_FileName()
    {
        var name = Formats.ApplyTemplate("{date:yyyy.MM.dd} {time:HH.mm} Суточный отчет ГГЭ {number} этап",
            new DateTime(2026, 1, 6, 21, 0, 0), new Dictionary<string, string> { ["number"] = "5" });
        Assert.Equal("2026.01.06 21.00 Суточный отчет ГГЭ 5 этап", name);
    }

    [Theory]
    [InlineData(7, 50, 8)]
    [InlineData(8, 30, 8)]
    [InlineData(17, 40, 18)]
    [InlineData(23, 30, 18)]
    [InlineData(14, 0, 18)]
    public void ReportTime_NearestSlot(int hour, int minute, int expectedHour)
    {
        var now = new DateTime(2026, 10, 15, hour, minute, 0);
        var result = ReportTime.Nearest(now, new[] { TimeSpan.FromHours(8), TimeSpan.FromHours(18) });
        Assert.Equal(new DateTime(2026, 10, 15, expectedHour, 0, 0), result);
    }

    [Fact]
    public void ReportTime_NearestSlot_AfterMidnightGoesToPreviousEvening()
    {
        var result = ReportTime.Nearest(new DateTime(2026, 10, 16, 0, 30, 0), new[] { TimeSpan.FromHours(8), TimeSpan.FromHours(18) });
        Assert.Equal(new DateTime(2026, 10, 15, 18, 0, 0), result);
    }

    [Fact]
    public void ReportTime_PeriodStartsAtPreviousEvening()
    {
        var settings = new ReportSettings();
        Assert.Equal(new DateTime(2026, 10, 14, 18, 0, 0), ReportTime.PeriodStart(settings, new DateTime(2026, 10, 15, 8, 0, 0)));
        Assert.Equal(new DateTime(2026, 10, 14, 18, 0, 0), ReportTime.PeriodStart(settings, new DateTime(2026, 10, 15, 18, 0, 0)));
    }

    [Fact]
    public void ReportTime_ExplicitDateUsesLastSlot()
    {
        var settings = new ReportSettings();
        var now = new DateTime(2026, 10, 20, 7, 0, 0);
        Assert.Equal(new DateTime(2026, 10, 15, 18, 0, 0), ReportTime.Resolve(settings, now, new DateTime(2026, 10, 15), null));
        Assert.Equal(new DateTime(2026, 10, 20, 8, 0, 0), ReportTime.Resolve(settings, now, null, TimeSpan.FromHours(8)));
    }
}
