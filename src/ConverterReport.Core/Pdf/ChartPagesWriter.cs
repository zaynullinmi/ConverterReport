using ConverterReport.Core.Reporting;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace ConverterReport.Core.Pdf;

/// <summary>Листы с диаграммами: сводный по всем этапам и по каждому этапу (устранение замечаний + загрузка документов).</summary>
internal sealed class ChartPagesWriter
{
    private const double Margin = 24;

    private readonly ReportModel _m;
    private readonly ChartRenderer _renderer;

    public ChartPagesWriter(ReportModel model)
    {
        _m = model;
        _renderer = new ChartRenderer(model.Config.Charts.FontFamily);
    }

    /// <summary>Добавляет листы; этапы без данных пропускаются (пустые листы не выводятся).</summary>
    public int AddPages(PdfDocument pdf)
    {
        var items = new List<ProjectReport>();
        if (_m.Summary is not null && _m.Config.Charts.SummaryChart)
            items.Add(_m.Summary);
        items.AddRange(_m.Projects);

        var added = 0;
        foreach (var item in items)
        {
            var remarks = item.HasRemarksSeries ? RemarksSpec(item) : null;
            var loads = item.HasLoadsSeries ? LoadsSpec(item) : null;
            if (remarks is null && loads is null)
                continue;

            var page = pdf.AddPage();
            var a3 = string.Equals(_m.Config.Charts.PageSize?.Trim(), "A3", StringComparison.OrdinalIgnoreCase);
            page.Width = XUnit.FromPoint(a3 ? 1190.55 : 841.89);
            page.Height = XUnit.FromPoint(a3 ? 841.89 : 595.28);
            using (var gfx = XGraphics.FromPdfPage(page))
                DrawPage(gfx, page.Width.Point, page.Height.Point, item, remarks, loads);
            added++;
        }

        return added;
    }

    private void DrawPage(XGraphics gfx, double width, double height, ProjectReport item, ChartSpec? remarks, ChartSpec? loads)
    {
        var family = _m.Config.Charts.FontFamily;
        var font = new XFont(family, 9);
        var bold = new XFont(family, 9, XFontStyleEx.Bold);
        var title = new XFont(family, 11, XFontStyleEx.Bold);
        var text = new XSolidBrush(Palette.ChartText);
        var accent = new XSolidBrush(Palette.ChartReview);

        // Как на листе «Диаграммы»: «По состоянию на» + дата; между диаграммами — «Всего:», заголовок загрузок,
        // дата окончания экспертизы и сколько дней осталось.
        var top = Margin;
        ChartRenderer.Text(gfx, "По состоянию на ", font, text, Margin, top, 0, 0);
        var w = gfx.MeasureString("По состоянию на ", font).Width;
        ChartRenderer.Text(gfx, Formats.Date(_m.ReportTime), bold, text, Margin + w, top, 0, 0);

        var chartsTop = top + 16;
        var available = height - chartsTop - Margin;
        const double infoHeight = 22;
        var remarksHeight = loads is null ? available : remarks is null ? 0 : (available - infoHeight) * 0.6;
        if (remarks is not null)
            _renderer.Draw(gfx, new XRect(Margin, chartsTop, width - 2 * Margin, remarksHeight), remarks);
        if (loads is null)
            return;

        var infoY = chartsTop + remarksHeight + infoHeight / 2;
        var loadsTitle = Formats.ApplyTemplate(_m.Config.Charts.LoadsTitle, values: Names(item));
        ChartRenderer.Text(gfx, "Всего:  ", bold, text, Margin + 120, infoY, 1, 0.5);
        ChartRenderer.Text(gfx, Formats.Number(item.LoadedInCharts), bold, text, Margin + 120, infoY, 0, 0.5);
        ChartRenderer.Text(gfx, loadsTitle, title, text, width / 2, infoY, 0.5, 0.5);
        if (item.ExpertiseEnd is { } end && item != _m.Summary)
        {
            var days = (end.Date - _m.ReportTime.Date).Days;
            var right = width - Margin;
            var daysText = days >= 0 ? days.ToString(Formats.Ru) : $"срок истек {-days} дн. назад";
            ChartRenderer.Text(gfx, daysText, bold, text, right, infoY, 1, 0.5);
            right -= gfx.MeasureString(daysText, bold).Width + 6;
            ChartRenderer.Text(gfx, "Осталось дней:", bold, text, right, infoY, 1, 0.5);
            right -= gfx.MeasureString("Осталось дней:", bold).Width + 30;
            ChartRenderer.Text(gfx, Formats.Date(end), bold, accent, right, infoY, 1, 0.5);
            right -= gfx.MeasureString(Formats.Date(end), bold).Width + 6;
            ChartRenderer.Text(gfx, "Дата окончания экспертизы:", bold, accent, right, infoY, 1, 0.5);
        }

        var loadsTop = chartsTop + remarksHeight + infoHeight;
        _renderer.Draw(gfx, new XRect(Margin, loadsTop, width - 2 * Margin, height - Margin - loadsTop), loads);
    }

    private ChartSpec RemarksSpec(ProjectReport p)
    {
        var cs = _m.Config.Charts;
        var daily = p.Daily;
        var spec = new ChartSpec
        {
            Title = Formats.ApplyTemplate(cs.RemarksTitle, values: Names(p)),
            Dates = daily.Select(d => d.Date).ToList(),
            ReportIndex = ReportIndex(daily),
            Deadline = p.ExpertiseEnd,
            Labels = cs.DataLabels,
            DateFormat = "dd.MM.yyyy ddd",
        };
        spec.Bars.Add(new BarSeries("Устраняется", Palette.ChartFixing, daily.Select(d => (double?)d.Remarks?.Fixing).ToArray()));
        spec.Bars.Add(new BarSeries("Рассматривается", Palette.ChartReview, daily.Select(d => (double?)d.Remarks?.Review).ToArray()));
        spec.Bars.Add(new BarSeries("Устранено за сутки", Palette.ChartFixedPerDay, daily.Select(d => (double?)d.FixedPerDay).ToArray()));
        spec.Lines.Add(new LineSeries("Всего", Palette.ChartTotal, daily.Select(d => (double?)d.Remarks?.Total).ToArray(),
            Trend: cs.TrendLine, TrendColor: XColor.FromArgb(0x44, 0x72, 0xC4)));
        spec.Lines.Add(new LineSeries("Устранено", Palette.ChartFixed, daily.Select(d => (double?)d.Remarks?.Fixed).ToArray(),
            Trend: cs.TrendLine, Labels: LabelPosition.Below, TrendColor: XColor.FromArgb(0xED, 0x7D, 0x31)));
        return spec;
    }

    private ChartSpec LoadsSpec(ProjectReport p)
    {
        var cs = _m.Config.Charts;
        var reportDate = _m.ReportTime.Date;
        var spec = new ChartSpec
        {
            Title = "",
            Dates = p.Daily.Select(d => d.Date).ToList(),
            ReportIndex = ReportIndex(p.Daily),
            Labels = cs.DataLabels,
            DateFormat = "dd.MM.yyyy ddd",
        };
        spec.Bars.Add(new BarSeries("Загружено документов", Palette.ChartLoaded,
            p.Daily.Select(d => d.Date <= reportDate ? (double?)d.Loaded : null).ToArray()));
        return spec;
    }

    private int ReportIndex(List<DailyPoint> daily)
    {
        var reportDate = _m.ReportTime.Date;
        var index = daily.FindLastIndex(d => d.Date <= reportDate);
        return Math.Max(0, index);
    }

    private static Dictionary<string, string> Names(ProjectReport p) => new()
    {
        ["name"] = p.Name,
        ["number"] = p.Number,
    };
}
