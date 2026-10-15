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
        var text = new XSolidBrush(Palette.ChartText);
        var accent = new XSolidBrush(Palette.ChartReview);

        // Шапка: дата состояния слева, справа — загружено всего / окончание экспертизы / осталось дней.
        var top = Margin;
        ChartRenderer.Text(gfx, "По состоянию на ", font, text, Margin, top, 0, 0);
        var w = gfx.MeasureString("По состоянию на ", font).Width;
        ChartRenderer.Text(gfx, Formats.DateTimeText(_m.ReportTime), bold, text, Margin + w, top, 0, 0);

        var info = new List<(string Label, string Value)>();
        if (loads is not null && item.Daily.Count > 0)
            info.Add(($"Загружено документов с {Formats.Date(item.Daily[0].Date)}: ", Formats.Number(item.LoadedInCharts)));
        if (item.ExpertiseEnd is { } end)
        {
            info.Add(("Дата окончания экспертизы: ", Formats.Date(end)));
            var days = (end.Date - _m.ReportTime.Date).Days;
            info.Add(days >= 0 ? ("Осталось дней: ", days.ToString(Formats.Ru)) : ("Срок истек ", $"{-days} дн. назад"));
        }

        var x = width - Margin;
        for (var i = info.Count - 1; i >= 0; i--)
        {
            var (label, value) = info[i];
            var vw = gfx.MeasureString(value, bold).Width;
            ChartRenderer.Text(gfx, value, bold, accent, x, top, 1, 0);
            x -= vw;
            ChartRenderer.Text(gfx, label, font, text, x, top, 1, 0);
            x -= gfx.MeasureString(label, font).Width + 18;
        }

        var chartsTop = top + 16;
        var available = height - chartsTop - Margin;
        if (remarks is not null && loads is not null)
        {
            var remarksHeight = available * 0.6;
            _renderer.Draw(gfx, new XRect(Margin, chartsTop, width - 2 * Margin, remarksHeight), remarks);
            _renderer.Draw(gfx, new XRect(Margin, chartsTop + remarksHeight + 10, width - 2 * Margin, available - remarksHeight - 10), loads);
        }
        else
        {
            _renderer.Draw(gfx, new XRect(Margin, chartsTop, width - 2 * Margin, available), (remarks ?? loads)!);
        }
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
        };
        spec.Bars.Add(new BarSeries("Устраняется", Palette.ChartFixing, daily.Select(d => (double?)d.Remarks?.Fixing).ToArray()));
        spec.Bars.Add(new BarSeries("Рассматривается", Palette.ChartReview, daily.Select(d => (double?)d.Remarks?.Review).ToArray()));
        spec.Bars.Add(new BarSeries("Устранено за сутки", Palette.ChartFixedPerDay, daily.Select(d => (double?)d.FixedPerDay).ToArray()));
        spec.Lines.Add(new LineSeries("Всего", Palette.ChartTotal, daily.Select(d => (double?)d.Remarks?.Total).ToArray()));
        spec.Lines.Add(new LineSeries("Устранено", Palette.ChartFixed, daily.Select(d => (double?)d.Remarks?.Fixed).ToArray(),
            Trend: cs.TrendLine, Labels: LabelPosition.Below));
        return spec;
    }

    private ChartSpec LoadsSpec(ProjectReport p)
    {
        var cs = _m.Config.Charts;
        var reportDate = _m.ReportTime.Date;
        var spec = new ChartSpec
        {
            Title = Formats.ApplyTemplate(cs.LoadsTitle, values: Names(p)),
            Dates = p.Daily.Select(d => d.Date).ToList(),
            ReportIndex = ReportIndex(p.Daily),
            Labels = cs.DataLabels,
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
