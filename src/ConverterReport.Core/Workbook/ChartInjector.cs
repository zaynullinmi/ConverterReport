using System.Text.RegularExpressions;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ConverterReport.Core.Workbook;

/// <summary>
/// Добавляет диаграммы Excel на лист «Диаграммы» (ClosedXML диаграммы не создает, поэтому они дописываются через Open XML).
/// Ряды диаграмм ссылаются на лист «Свод» — при изменении данных Excel перестраивает их сам.
/// </summary>
internal static class ChartInjector
{
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

    public static void Inject(string path, string sheetName, IReadOnlyList<ChartPlacement> charts)
    {
        using var doc = SpreadsheetDocument.Open(path, true);
        var wbPart = doc.WorkbookPart ?? throw new InvalidDataException("В книге нет WorkbookPart.");
        QuoteSheetReferences(wbPart);
        if (charts.Count == 0)
            return;
        var sheet = wbPart.Workbook.Descendants<Sheet>().First(s => s.Name == sheetName);
        var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!);
        var drawingsPart = wsPart.DrawingsPart ?? wsPart.AddNewPart<DrawingsPart>();

        var anchors = new List<XElement>();
        var id = 2;
        foreach (var chart in charts)
        {
            var chartPart = drawingsPart.AddNewPart<ChartPart>();
            using (var stream = chartPart.GetStream(FileMode.Create))
                ChartXml(chart).Save(stream);
            anchors.Add(Anchor(chart, drawingsPart.GetIdOfPart(chartPart), id, $"Диаграмма {id - 1}"));
            id++;
        }

        var drawing = new XDocument(new XElement(Xdr + "wsDr",
            new XAttribute(XNamespace.Xmlns + "xdr", Xdr), new XAttribute(XNamespace.Xmlns + "a", A), anchors));
        using (var stream = drawingsPart.GetStream(FileMode.Create))
            drawing.Save(stream);

        var worksheet = wsPart.Worksheet;
        if (!worksheet.Elements<Drawing>().Any())
        {
            var element = new Drawing { Id = wsPart.GetIdOfPart(drawingsPart) };
            // <drawing> должен стоять после pageSetup/headerFooter/rowBreaks и перед legacyDrawing/tableParts/extLst.
            var after = worksheet.ChildElements.FirstOrDefault(e =>
                e is LegacyDrawing or LegacyDrawingHeaderFooter or DrawingHeaderFooter or Picture or OleObjects or Controls
                    or WebPublishItems or TableParts or WorksheetExtensionList);
            if (after is not null)
                worksheet.InsertBefore(element, after);
            else
                worksheet.Append(element);
        }

        worksheet.Save();
    }

    /// <summary>
    /// Берет в кавычки ссылки на листы с точкой в имени («Д.3!A1» → «'Д.3'!A1»): Excel понимает оба варианта,
    /// а LibreOffice и другие программы — только в кавычках. ClosedXML при сохранении кавычки иногда убирает.
    /// </summary>
    private static void QuoteSheetReferences(WorkbookPart wbPart)
    {
        var names = wbPart.Workbook.Descendants<Sheet>()
            .Select(s => s.Name?.Value ?? "")
            .Where(n => n.Length > 0 && n.Any(ch => !char.IsLetterOrDigit(ch) && ch != '_'))
            .OrderByDescending(n => n.Length)
            .ToList();
        if (names.Count == 0)
            return;

        var patterns = names
            .Select(n => new Regex(@"(?<![\w'.])" + Regex.Escape(n) + "!", RegexOptions.CultureInvariant))
            .Zip(names, (re, n) => (re, quoted: "'" + n.Replace("'", "''") + "'!"))
            .ToList();
        foreach (var wsPart in wbPart.WorksheetParts)
        {
            var changed = false;
            foreach (var formula in wsPart.Worksheet.Descendants<CellFormula>())
            {
                var text = formula.Text;
                if (string.IsNullOrEmpty(text))
                    continue;
                var fixedText = text;
                foreach (var (re, quoted) in patterns)
                    fixedText = re.Replace(fixedText, quoted);
                if (fixedText == text)
                    continue;
                formula.Text = fixedText;
                changed = true;
            }

            if (changed)
                wsPart.Worksheet.Save();
        }
    }

    private static XElement Anchor(ChartPlacement chart, string relId, int id, string name) =>
        new(Xdr + "twoCellAnchor", new XAttribute("editAs", "oneCell"),
            new XElement(Xdr + "from", Cell(chart.FromColumn, chart.FromRow)),
            new XElement(Xdr + "to", Cell(chart.ToColumn, chart.ToRow)),
            new XElement(Xdr + "graphicFrame", new XAttribute("macro", ""),
                new XElement(Xdr + "nvGraphicFramePr",
                    new XElement(Xdr + "cNvPr", new XAttribute("id", id), new XAttribute("name", name)),
                    new XElement(Xdr + "cNvGraphicFramePr")),
                new XElement(Xdr + "xfrm",
                    new XElement(A + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                    new XElement(A + "ext", new XAttribute("cx", 0), new XAttribute("cy", 0))),
                new XElement(A + "graphic",
                    new XElement(A + "graphicData", new XAttribute("uri", C.NamespaceName),
                        new XElement(C + "chart", new XAttribute(XNamespace.Xmlns + "c", C), new XAttribute(XNamespace.Xmlns + "r", R),
                            new XAttribute(R + "id", relId))))),
            new XElement(Xdr + "clientData"));

    private static IEnumerable<XElement> Cell(int column, int row) => new[]
    {
        new XElement(Xdr + "col", column),
        new XElement(Xdr + "colOff", 0),
        new XElement(Xdr + "row", row),
        new XElement(Xdr + "rowOff", 0),
    };

    // ------------------------------------------------------------------ chartSpace
    // Оформление — как у диаграмм книги отчета: подписи 9 пт, шкала замечаний от 0, область построения
    // растянута на всю диаграмму, у сводной диаграммы нет линий тренда.

    private const int MainCategoryAxis = 50010;
    private const int MainValueAxis = 50020;
    private const int MarkerCategoryAxis = 50030;
    private const int MarkerValueAxis = 50040;
    private const int MarkerSeriesIndex = 20;

    private static XDocument ChartXml(ChartPlacement chart)
    {
        var remarks = chart.Kind == ChartKind.Remarks;
        var plot = new XElement(C + "plotArea", remarks
            ? ManualLayout(0.028, 0.075, 0.963, 0.73)
            : ManualLayout(0.028, 0.13, 0.963, 0.63));
        var bars = chart.Series.Where(s => !s.IsLine).ToList();
        var lines = chart.Series.Where(s => s.IsLine).ToList();
        var index = 0;

        if (bars.Count > 0)
        {
            var bar = new XElement(C + "barChart",
                Val("barDir", "col"),
                Val("grouping", remarks ? "stacked" : "clustered"),
                Val("varyColors", 0));
            foreach (var s in bars)
                bar.Add(BarSeries(s, chart, index++));
            bar.Add(Val("gapWidth", remarks ? 50 : 150));
            if (remarks)
                bar.Add(Val("overlap", 100));
            bar.Add(Val("axId", MainCategoryAxis), Val("axId", MainValueAxis));
            plot.Add(bar);
        }

        if (lines.Count > 0)
        {
            var line = new XElement(C + "lineChart", Val("grouping", "standard"), Val("varyColors", 0));
            foreach (var s in lines)
                line.Add(LineSeries(s, chart, index++));
            line.Add(Val("marker", 1), Val("axId", MainCategoryAxis), Val("axId", MainValueAxis));
            plot.Add(line);
        }

        var marker = chart.MarkerIndex is { } m && m >= 0 && m < chart.CategoryCount;
        if (marker)
            plot.Add(MarkerChart(chart));

        plot.Add(CategoryAxis(MainCategoryAxis, MainValueAxis, deleted: false),
            ValueAxis(MainValueAxis, MainCategoryAxis, min: 0, max: chart.AxisMax, format: remarks ? "General" : "#,##0", deleted: false));
        if (marker)
            plot.Add(CategoryAxis(MarkerCategoryAxis, MarkerValueAxis, deleted: true), ValueAxis(MarkerValueAxis, MarkerCategoryAxis, min: 0, max: 1, format: "General", deleted: true));

        var chartElement = new XElement(C + "chart");
        if (!string.IsNullOrEmpty(chart.Title))
            chartElement.Add(Title(chart.Title));
        chartElement.Add(Val("autoTitleDeleted", string.IsNullOrEmpty(chart.Title) ? 1 : 0), plot);
        if (remarks)
        {
            chartElement.Add(new XElement(C + "legend", Val("legendPos", "b"), Val("overlay", 0), TextProps(900)));
        }

        chartElement.Add(Val("plotVisOnly", 1), Val("dispBlanksAs", "gap"));

        var space = new XElement(C + "chartSpace",
            new XAttribute(XNamespace.Xmlns + "c", C), new XAttribute(XNamespace.Xmlns + "a", A), new XAttribute(XNamespace.Xmlns + "r", R),
            Val("date1904", 0),
            Val("lang", "ru-RU"),
            Val("roundedCorners", 0),
            chartElement,
            new XElement(C + "spPr",
                new XElement(A + "solidFill", Rgb("FFFFFF")),
                new XElement(A + "ln", new XAttribute("w", 9525), new XElement(A + "solidFill", Rgb("D9D9D9")))),
            TextProps(900));
        return new XDocument(new XDeclaration("1.0", "UTF-8", "yes"), space);
    }

    private static XElement ManualLayout(double x, double y, double w, double h) =>
        new(C + "layout", new XElement(C + "manualLayout",
            Val("layoutTarget", "inner"), Val("xMode", "edge"), Val("yMode", "edge"),
            Val("x", x), Val("y", y), Val("w", w), Val("h", h)));

    private static XElement BarSeries(ChartSeriesRef s, ChartPlacement chart, int index)
    {
        XElement labels;
        if (chart.Kind == ChartKind.Loads)
        {
            // Подпись колонки выше верха шкалы Excel сам переносит внутрь диаграммы, сбоку у верхнего края (как в книге отчета).
            labels = Labels(null, 900, rotate: true, bold: false, color: "0070C0");
        }
        else
        {
            labels = s.Color switch
            {
                "0070C0" => Labels("ctr", 700, rotate: true, bold: false, color: "FFFFFF"),
                "92D050" => Labels("ctr", 900, rotate: false, bold: false, color: "404040"),
                _ => Labels("ctr", 900, rotate: true, bold: false, color: chart.Summary ? "C00000" : "FF0000"),
            };
        }

        return new XElement(C + "ser",
            Val("idx", index), Val("order", index),
            SeriesName(s),
            new XElement(C + "spPr", new XElement(A + "solidFill", Rgb(s.Color))),
            Val("invertIfNegative", 0),
            labels,
            Categories(chart),
            Values(s));
    }

    private static XElement LineSeries(ChartSeriesRef s, ChartPlacement chart, int index)
    {
        var ser = new XElement(C + "ser",
            Val("idx", index), Val("order", index),
            SeriesName(s),
            new XElement(C + "spPr",
                new XElement(A + "ln", new XAttribute("w", 28575), new XAttribute("cap", "rnd"),
                    new XElement(A + "solidFill", Rgb(s.Color)), new XElement(A + "round"))),
            new XElement(C + "marker", Val("symbol", "none")),
            Labels(s.LabelsBelow ? "b" : "t", 900, rotate: true, bold: false, color: s.Color));
        if (s.TrendColor is not null && !chart.Summary)
        {
            ser.Add(new XElement(C + "trendline",
                new XElement(C + "spPr",
                    new XElement(A + "ln", new XAttribute("w", 19050), new XAttribute("cap", "rnd"),
                        new XElement(A + "solidFill", Rgb(s.TrendColor)), new XElement(A + "prstDash", new XAttribute("val", "sysDot")))),
                Val("trendlineType", "linear"),
                Val("dispRSqr", 0),
                Val("dispEq", 0)));
        }

        ser.Add(Categories(chart), Values(s), Val("smooth", 0));
        return ser;
    }

    /// <summary>
    /// Вертикальная пунктирная линия на дате окончания экспертизы: одна точка на вспомогательной шкале 0..1
    /// и линия проекции от нее до оси дат (как линия, нарисованная вручную в книге отчета).
    /// </summary>
    private static XElement MarkerChart(ChartPlacement chart)
    {
        var points = new XElement(C + "numLit", new XElement(C + "formatCode", "General"), Val("ptCount", chart.CategoryCount),
            new XElement(C + "pt", new XAttribute("idx", chart.MarkerIndex!.Value), new XElement(C + "v", 1)));
        return new XElement(C + "lineChart",
            Val("grouping", "standard"),
            Val("varyColors", 0),
            new XElement(C + "ser",
                Val("idx", MarkerSeriesIndex), Val("order", MarkerSeriesIndex),
                new XElement(C + "tx", new XElement(C + "v", "Окончание экспертизы")),
                new XElement(C + "spPr", new XElement(A + "ln", new XElement(A + "noFill"))),
                new XElement(C + "marker", Val("symbol", "none")),
                Categories(chart),
                new XElement(C + "val", points),
                Val("smooth", 0)),
            new XElement(C + "dropLines", new XElement(C + "spPr",
                new XElement(A + "ln", new XAttribute("w", 15875),
                    new XElement(A + "solidFill", Rgb("0070C0")), new XElement(A + "prstDash", new XAttribute("val", "dash"))))),
            Val("marker", 1),
            Val("axId", MarkerCategoryAxis), Val("axId", MarkerValueAxis));
    }

    private static XElement SeriesName(ChartSeriesRef s) =>
        new(C + "tx", new XElement(C + "strRef", new XElement(C + "f", s.NameRef),
            new XElement(C + "strCache", Val("ptCount", 1),
                new XElement(C + "pt", new XAttribute("idx", 0), new XElement(C + "v", s.Name)))));

    private static XElement Categories(ChartPlacement chart) =>
        new(C + "cat", new XElement(C + "strRef", new XElement(C + "f", chart.CategoriesRef)));

    private static XElement Values(ChartSeriesRef s) =>
        new(C + "val", new XElement(C + "numRef", new XElement(C + "f", s.ValuesRef)));

    private static XElement Labels(string? position, int size, bool rotate, bool bold, string color)
    {
        var labels = new XElement(C + "dLbls");
        labels.Add(
            new XElement(C + "spPr", new XElement(A + "noFill"), new XElement(A + "ln", new XElement(A + "noFill"))),
            new XElement(C + "txPr",
                new XElement(A + "bodyPr", rotate ? new XAttribute("rot", -5400000) : null, new XAttribute("vert", "horz")),
                new XElement(A + "lstStyle"),
                new XElement(A + "p",
                    new XElement(A + "pPr", new XElement(A + "defRPr", new XAttribute("sz", size), new XAttribute("b", bold ? 1 : 0),
                        new XElement(A + "solidFill", Rgb(color)))),
                    new XElement(A + "endParaRPr", new XAttribute("lang", "ru-RU")))),
            position is null ? null : Val("dLblPos", position),
            Val("showLegendKey", 0),
            Val("showVal", 1),
            Val("showCatName", 0),
            Val("showSerName", 0),
            Val("showPercent", 0),
            Val("showBubbleSize", 0));
        return labels;
    }

    private static XElement CategoryAxis(int id, int crossAx, bool deleted) =>
        new(C + "catAx",
            Val("axId", id),
            new XElement(C + "scaling", Val("orientation", "minMax")),
            Val("delete", deleted ? 1 : 0),
            Val("axPos", "b"),
            new XElement(C + "numFmt", new XAttribute("formatCode", "General"), new XAttribute("sourceLinked", 1)),
            Val("majorTickMark", deleted ? "none" : "out"),
            Val("minorTickMark", "none"),
            Val("tickLblPos", deleted ? "none" : "low"),
            new XElement(C + "spPr", new XElement(A + "ln", new XAttribute("w", 9525), new XElement(A + "solidFill", Rgb("BFBFBF")))),
            new XElement(C + "txPr",
                new XElement(A + "bodyPr", new XAttribute("rot", -5400000), new XAttribute("vert", "horz")),
                new XElement(A + "lstStyle"),
                new XElement(A + "p", new XElement(A + "pPr", new XElement(A + "defRPr", new XAttribute("sz", 900))),
                    new XElement(A + "endParaRPr", new XAttribute("lang", "ru-RU")))),
            Val("crossAx", crossAx),
            Val("crosses", "autoZero"),
            Val("auto", 1),
            Val("lblAlgn", "ctr"),
            Val("lblOffset", 100),
            Val("noMultiLvlLbl", 0));

    private static XElement ValueAxis(int id, int crossAx, double? min, double? max, string format, bool deleted)
    {
        var scaling = new XElement(C + "scaling", Val("orientation", "minMax"));
        if (max is > 0)
            scaling.Add(Val("max", max.Value));
        if (min is not null)
            scaling.Add(Val("min", min.Value));
        return new XElement(C + "valAx",
            Val("axId", id),
            scaling,
            Val("delete", deleted ? 1 : 0),
            Val("axPos", deleted ? "r" : "l"),
            deleted
                ? null
                : new XElement(C + "majorGridlines",
                    new XElement(C + "spPr", new XElement(A + "ln", new XAttribute("w", 6350), new XElement(A + "solidFill", Rgb("D9D9D9"))))),
            new XElement(C + "numFmt", new XAttribute("formatCode", format), new XAttribute("sourceLinked", 0)),
            Val("majorTickMark", "none"),
            Val("minorTickMark", "none"),
            Val("tickLblPos", deleted ? "none" : "nextTo"),
            new XElement(C + "spPr", new XElement(A + "ln", new XElement(A + "noFill"))),
            TextProps(900),
            Val("crossAx", crossAx),
            Val("crosses", deleted ? "max" : "autoZero"),
            Val("crossBetween", "between"));
    }

    private static XElement Title(string text) =>
        new(C + "title",
            new XElement(C + "tx",
                new XElement(C + "rich",
                    new XElement(A + "bodyPr"),
                    new XElement(A + "lstStyle"),
                    new XElement(A + "p",
                        new XElement(A + "pPr", new XElement(A + "defRPr", new XAttribute("sz", 1400), new XAttribute("b", 1))),
                        new XElement(A + "r",
                            new XElement(A + "rPr", new XAttribute("lang", "ru-RU"), new XAttribute("sz", 1400), new XAttribute("b", 1)),
                            new XElement(A + "t", text))))),
            Val("overlay", 0));

    private static XElement TextProps(int size) =>
        new(C + "txPr",
            new XElement(A + "bodyPr"),
            new XElement(A + "lstStyle"),
            new XElement(A + "p",
                new XElement(A + "pPr", new XElement(A + "defRPr", new XAttribute("sz", size))),
                new XElement(A + "endParaRPr", new XAttribute("lang", "ru-RU"))));

    private static XElement Rgb(string hex) => new(A + "srgbClr", new XAttribute("val", hex));

    private static XElement Val(string name, object value) => new(C + name, new XAttribute("val", value));
}
