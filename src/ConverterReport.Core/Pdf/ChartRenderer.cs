using ConverterReport.Core.Configuration;
using PdfSharp.Drawing;

namespace ConverterReport.Core.Pdf;

/// <summary>Ряд гистограммы (столбцы складываются в стопку).</summary>
internal sealed record BarSeries(string Name, XColor Color, double?[] Values);

/// <summary>Ряд-линия. <see cref="Trend"/> — дополнительно нарисовать линейный тренд (прогноз) до конца оси.</summary>
internal sealed record LineSeries(string Name, XColor Color, double?[] Values, bool Trend = false, LabelPosition Labels = LabelPosition.Above, XColor? TrendColor = null);

internal enum LabelPosition
{
    Above,
    Below,
}

internal sealed class ChartSpec
{
    public string Title { get; init; } = "";
    public IReadOnlyList<DateTime> Dates { get; init; } = Array.Empty<DateTime>();
    public List<BarSeries> Bars { get; } = new();
    public List<LineSeries> Lines { get; } = new();

    /// <summary>Индекс отчетной даты на оси (последняя точка с данными).</summary>
    public int ReportIndex { get; init; }

    /// <summary>Дата окончания экспертизы (вертикальная линия), если попадает на ось.</summary>
    public DateTime? Deadline { get; init; }

    public DataLabelsMode Labels { get; init; } = DataLabelsMode.Auto;

    /// <summary>Формат подписей дат; «ddd» — день недели.</summary>
    public string DateFormat { get; init; } = "dd.MM";
}

/// <summary>
/// Векторные диаграммы для PDF (без растровых картинок и без Excel):
/// столбцы с накоплением + линии + линия тренда + отметка окончания экспертизы.
/// </summary>
internal sealed class ChartRenderer
{
    private readonly string _family;

    public ChartRenderer(string fontFamily)
    {
        _family = fontFamily;
    }

    private XFont Font(double size, bool bold = false) =>
        new(_family, size, bold ? XFontStyleEx.Bold : XFontStyleEx.Regular);

    public void Draw(XGraphics gfx, XRect area, ChartSpec spec)
    {
        var n = spec.Dates.Count;
        gfx.DrawRectangle(new XPen(Palette.ChartFrame, 0.75), area);

        var titleFont = Font(12, bold: true);
        var titleHeight = 8.0;
        if (!string.IsNullOrEmpty(spec.Title))
        {
            Text(gfx, spec.Title, titleFont, Brush(Palette.ChartText), area.X + area.Width / 2, area.Y + 6, 0.5, 0);
            titleHeight = gfx.MeasureString(spec.Title, titleFont).Height + 10;
        }

        var legendHeight = DrawLegend(gfx, area, spec);
        if (n == 0)
            return;

        // Масштаб по оси значений.
        double max = 0, min = 0;
        for (var i = 0; i < n; i++)
        {
            double pos = 0, neg = 0;
            foreach (var bar in spec.Bars)
            {
                var v = bar.Values[i] ?? 0;
                if (v > 0) pos += v; else neg += v;
            }

            max = Math.Max(max, pos);
            min = Math.Min(min, neg);
            foreach (var line in spec.Lines)
            {
                if (line.Values[i] is { } lv)
                {
                    max = Math.Max(max, lv);
                    min = Math.Min(min, lv);
                }
            }
        }

        if (max <= min)
            max = min + 1;
        var step = NiceStep((max - min) / 6);
        var yMax = Math.Ceiling(max * 1.08 / step) * step;
        var yMin = Math.Floor(min / step) * step;
        if (yMax <= yMin)
            yMax = yMin + step;

        var axisFont = Font(7);
        var yLabelWidth = Enumerable.Range(0, (int)Math.Round((yMax - yMin) / step) + 1)
            .Select(k => gfx.MeasureString(Formats.Number((int)Math.Round(yMin + k * step)), axisFont).Width)
            .DefaultIfEmpty(10).Max();

        var sameYear = spec.Dates.Select(d => d.Year).Distinct().Count() == 1;
        var dateFormat = spec.DateFormat == "dd.MM" && !sameYear ? "dd.MM.yy" : spec.DateFormat;
        string DateLabel(DateTime d)
        {
            var text = d.ToString(dateFormat, Formats.Ru);
            var space = text.LastIndexOf(' ');
            return space >= 0 && space + 1 < text.Length ? text[..(space + 1)] + char.ToUpper(text[space + 1]) + text[(space + 2)..] : text;
        }

        var xLabelHeight = gfx.MeasureString(DateLabel(spec.Dates[0]), axisFont).Width + 6;

        var plot = new XRect(
            area.X + 10 + yLabelWidth,
            area.Y + titleHeight,
            area.Width - 10 - yLabelWidth - 12,
            area.Height - titleHeight - xLabelHeight - legendHeight - 6);
        if (plot.Width <= 10 || plot.Height <= 10)
            return;

        double Y(double v) => plot.Bottom - (v - yMin) / (yMax - yMin) * plot.Height;
        var slot = plot.Width / n;
        double X(int i) => plot.X + slot * (i + 0.5);

        // Сетка и подписи оси значений.
        var gridPen = new XPen(Palette.ChartGrid, 0.5);
        for (var v = yMin; v <= yMax + step / 2; v += step)
        {
            var y = Y(v);
            gfx.DrawLine(gridPen, plot.X, y, plot.Right, y);
            Text(gfx, Formats.Number((int)Math.Round(v)), axisFont, Brush(Palette.ChartText), plot.X - 3, y, 1, 0.5);
        }

        var axisPen = new XPen(Palette.ChartAxis, 0.75);
        gfx.DrawLine(axisPen, plot.X, Y(0), plot.Right, Y(0));
        gfx.DrawLine(axisPen, plot.X, plot.Y, plot.X, plot.Bottom);

        // Подписи дат: шаг подбирается так, чтобы подписи не налезали; отчетная дата подписывается всегда.
        var labelStep = Math.Max(1, (int)Math.Ceiling((axisFont.Height + 1.5) / slot));
        for (var i = 0; i < n; i++)
        {
            if ((i - spec.ReportIndex) % labelStep != 0)
                continue;
            var bold = i == spec.ReportIndex;
            TextRotated(gfx, DateLabel(spec.Dates[i]), bold ? Font(7, true) : axisFont,
                Brush(Palette.ChartText), X(i), plot.Bottom + 3, 1, 0.5);
        }

        // Отметка отчетной даты, если ось продолжается в будущее.
        if (spec.ReportIndex < n - 1)
        {
            var pen = new XPen(Palette.ChartAxis, 0.5) { DashStyle = XDashStyle.Dot };
            gfx.DrawLine(pen, X(spec.ReportIndex) + slot / 2, plot.Y, X(spec.ReportIndex) + slot / 2, plot.Bottom);
        }

        // Столбцы с накоплением.
        var barWidth = slot / 1.5;
        var wide = spec.Labels == DataLabelsMode.All || (spec.Labels == DataLabelsMode.Auto && slot >= 18);
        var labelFont = Font(6);
        for (var i = 0; i < n; i++)
        {
            double pos = 0, neg = 0;
            foreach (var bar in spec.Bars)
            {
                if (bar.Values[i] is not { } v || v == 0)
                    continue;
                double from, to;
                if (v > 0)
                {
                    from = pos;
                    to = pos + v;
                    pos = to;
                }
                else
                {
                    from = neg;
                    to = neg + v;
                    neg = to;
                }

                var top = Math.Min(Y(from), Y(to));
                var height = Math.Abs(Y(to) - Y(from));
                gfx.DrawRectangle(Brush(bar.Color), X(i) - barWidth / 2, top, barWidth, height);

                var showLabel = spec.Labels != DataLabelsMode.None && (wide || i == spec.ReportIndex);
                if (showLabel && height >= 8 && barWidth >= 12)
                    Text(gfx, Formats.Number((int)v), labelFont, Brush(ContrastText(bar.Color)), X(i), top + height / 2, 0.5, 0.5);
            }
        }

        // Значения одиночного ряда столбцов (загрузки) — над столбцами.
        if (spec.Bars.Count == 1 && spec.Lines.Count == 0 && spec.Labels != DataLabelsMode.None)
            DrawBarTopLabels(gfx, spec, slot, X, Y);

        // Линии тренда (под основными линиями), обрезаются по области построения.
        var clip = gfx.Save();
        gfx.IntersectClip(plot);
        foreach (var line in spec.Lines.Where(l => l.Trend))
            DrawTrend(gfx, line, spec.ReportIndex, n, X, Y);
        gfx.Restore(clip);

        // Линии.
        foreach (var line in spec.Lines)
        {
            var pen = new XPen(line.Color, 1.75) { LineJoin = XLineJoin.Round };
            var run = new List<XPoint>();
            for (var i = 0; i <= n; i++)
            {
                if (i < n && line.Values[i] is { } v)
                {
                    run.Add(new XPoint(X(i), Y(v)));
                    continue;
                }

                if (run.Count == 1)
                    gfx.DrawEllipse(Brush(line.Color), run[0].X - 1.5, run[0].Y - 1.5, 3, 3);
                else if (run.Count > 1)
                    gfx.DrawLines(pen, run.ToArray());
                run.Clear();
            }

            if (spec.Labels == DataLabelsMode.None)
                continue;
            for (var i = 0; i < n; i++)
            {
                if (line.Values[i] is not { } v || !(wide || i == spec.ReportIndex))
                    continue;
                var font = i == spec.ReportIndex ? Font(7, true) : labelFont;
                var dy = line.Labels == LabelPosition.Above ? -2 : 2;
                Text(gfx, Formats.Number((int)v), font, Brush(line.Color), X(i), Y(v) + dy, wide ? 0.5 : 0, line.Labels == LabelPosition.Above ? 1 : 0);
            }
        }

        // Окончание экспертизы.
        if (spec.Deadline is { } deadline)
        {
            var index = IndexOf(spec.Dates, deadline);
            if (index >= 0)
            {
                var x = X(index);
                var pen = new XPen(Palette.ChartDeadline, 1) { DashStyle = XDashStyle.Dash };
                gfx.DrawLine(pen, x, plot.Y, x, plot.Bottom);
                Text(gfx, $"Окончание экспертизы {Formats.Date(deadline)}", Font(7, true), Brush(Palette.ChartDeadline),
                    x - 3, plot.Y + 2, 1, 0);
            }
        }
    }

    private void DrawBarTopLabels(XGraphics gfx, ChartSpec spec, double slot, Func<int, double> x, Func<double, double> y)
    {
        var bar = spec.Bars[0];
        var font = Font(6);
        for (var i = 0; i < spec.Dates.Count; i++)
        {
            if (bar.Values[i] is not { } v || v == 0)
                continue;
            var text = Formats.Number((int)v);
            var isReport = i == spec.ReportIndex;
            if (spec.Labels == DataLabelsMode.Last && !isReport)
                continue;
            var size = gfx.MeasureString(text, font);
            if (size.Width <= slot * 0.95 || isReport)
                Text(gfx, text, font, Brush(bar.Color), x(i), y(v) - 1.5, 0.5, 1);
            else if (size.Height <= slot * 1.15 || spec.Labels == DataLabelsMode.All)
                TextRotated(gfx, text, font, Brush(bar.Color), x(i), y(v) - 2, 0, 0.5);
        }
    }

    private static void DrawTrend(XGraphics gfx, LineSeries line, int lastIndex, int n, Func<int, double> x, Func<double, double> y)
    {
        var pts = new List<(double X, double Y)>();
        for (var i = 0; i <= Math.Min(lastIndex, n - 1); i++)
        {
            if (line.Values[i] is { } v)
                pts.Add((i, v));
        }

        if (pts.Count < 3)
            return;
        var mx = pts.Average(p => p.X);
        var my = pts.Average(p => p.Y);
        var sxx = pts.Sum(p => (p.X - mx) * (p.X - mx));
        if (sxx == 0)
            return;
        var b = pts.Sum(p => (p.X - mx) * (p.Y - my)) / sxx;
        var a = my - b * mx;
        var first = (int)pts[0].X;
        var pen = new XPen(line.TrendColor ?? line.Color, 1.1) { DashStyle = XDashStyle.Dot };
        gfx.DrawLine(pen, x(first), y(a + b * first), x(n - 1), y(a + b * (n - 1)));
    }

    private double DrawLegend(XGraphics gfx, XRect area, ChartSpec spec)
    {
        var font = Font(8);
        var items = new List<(string Name, XColor Color, int Kind)>(); // 0 — столбец, 1 — линия, 2 — пунктир
        items.AddRange(spec.Bars.Select(b => (b.Name, b.Color, 0)));
        items.AddRange(spec.Lines.Select(l => (l.Name, l.Color, 1)));
        items.AddRange(spec.Lines.Where(l => l.Trend).Select(l => ($"Линейная ({l.Name})", l.TrendColor ?? l.Color, 2)));
        if (spec.Deadline is { } d && IndexOf(spec.Dates, d) >= 0)
            items.Add(("Окончание экспертизы", Palette.ChartDeadline, 2));
        if (items.Count == 0)
            return 4;

        const double marker = 14, gap = 4, spacing = 14;
        var widths = items.Select(i => marker + gap + gfx.MeasureString(i.Name, font).Width).ToList();
        var total = widths.Sum() + spacing * (items.Count - 1);
        var x = area.X + Math.Max(4, (area.Width - total) / 2);
        var y = area.Bottom - 12;
        for (var k = 0; k < items.Count; k++)
        {
            var (name, color, kind) = items[k];
            switch (kind)
            {
                case 0:
                    gfx.DrawRectangle(Brush(color), x + 3, y - 4, 8, 8);
                    break;
                case 1:
                    gfx.DrawLine(new XPen(color, 1.75), x, y, x + marker, y);
                    break;
                default:
                    gfx.DrawLine(new XPen(color, 1.1) { DashStyle = XDashStyle.Dot }, x, y, x + marker, y);
                    break;
            }

            Text(gfx, name, font, Brush(Palette.ChartText), x + marker + gap, y, 0, 0.5);
            x += widths[k] + spacing;
        }

        return 20;
    }

    private static int IndexOf(IReadOnlyList<DateTime> dates, DateTime date)
    {
        for (var i = 0; i < dates.Count; i++)
        {
            if (dates[i].Date == date.Date)
                return i;
        }

        return -1;
    }

    /// <summary>Шаг сетки 1/2/2,5/5 × 10^n.</summary>
    internal static double NiceStep(double raw)
    {
        if (raw <= 0)
            return 1;
        var exp = Math.Pow(10, Math.Floor(Math.Log10(raw)));
        var f = raw / exp;
        var nice = f <= 1 ? 1 : f <= 2 ? 2 : f <= 2.5 ? 2.5 : f <= 5 ? 5 : 10;
        return Math.Max(1, nice * exp);
    }

    private static XColor ContrastText(XColor background) =>
        background.R * 0.299 + background.G * 0.587 + background.B * 0.114 > 150 ? XColors.Black : XColors.White;

    private static XSolidBrush Brush(XColor color) => new(color);

    /// <summary>Текст с привязкой: ax/ay — 0 (лево/верх), 0.5 (центр), 1 (право/низ).</summary>
    internal static void Text(XGraphics gfx, string text, XFont font, XBrush brush, double x, double y, double ax, double ay)
    {
        var size = gfx.MeasureString(text, font);
        gfx.DrawString(text, font, brush, new XRect(x - size.Width * ax, y - size.Height * ay, size.Width, size.Height), XStringFormats.TopLeft);
    }

    /// <summary>Текст, повернутый на 90° (читается снизу вверх). ax=1 — текст заканчивается в точке привязки.</summary>
    internal static void TextRotated(XGraphics gfx, string text, XFont font, XBrush brush, double x, double y, double ax, double ay)
    {
        var state = gfx.Save();
        gfx.RotateAtTransform(-90, new XPoint(x, y));
        Text(gfx, text, font, brush, x, y, ax, ay);
        gfx.Restore(state);
    }
}
