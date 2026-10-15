using ConverterReport.Core.Configuration;
using ConverterReport.Core.Model;
using ConverterReport.Core.Reporting;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;

namespace ConverterReport.Core.Pdf;

/// <summary>
/// Текстовая часть PDF (MigraDoc): суточный отчет на дату и детальный отчет.
/// Структура повторяет листы книги: «ДД.ММ» (разделы 1–4) и «Отчет_дет».
/// </summary>
internal sealed class ReportDocumentBuilder
{
    private const double PageWidthMm = 210;
    private const double LeftMarginMm = 15;
    private const double RightMarginMm = 10;
    private const double ContentWidthMm = PageWidthMm - LeftMarginMm - RightMarginMm;

    private readonly ReportModel _m;
    private readonly ReportSettings _s;
    private readonly Document _doc = new();

    public ReportDocumentBuilder(ReportModel model)
    {
        _m = model;
        _s = model.Config.Report;
        SetupStyles();
    }

    public bool HasContent => _doc.Sections.Count > 0;

    public Document Build(bool daily, bool detailed)
    {
        _doc.Info.Title = $"{_s.Title} {_s.Subtitle} {_m.Caption}";
        _doc.Info.Subject = $"По состоянию на {Formats.DateTimeText(_m.ReportTime)}";
        _doc.Info.Author = "ConverterReport";

        if (daily)
            AddDailyReport();
        if (detailed && _m.Projects.Any(p => p.Current is not null))
            AddDetailedReport();
        return _doc;
    }

    // ------------------------------------------------------------------ стили

    private void SetupStyles()
    {
        var normal = _doc.Styles[StyleNames.Normal]!;
        normal.Font.Name = _s.FontFamily;
        normal.Font.Size = 11;
        normal.ParagraphFormat.SpaceAfter = 0;

        var title = _doc.Styles.AddStyle("ReportTitle", StyleNames.Normal);
        title.Font.Size = 14;
        title.Font.Bold = true;
        title.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        var subtitle = _doc.Styles.AddStyle("ReportSubtitle", StyleNames.Normal);
        subtitle.Font.Size = 12;
        subtitle.ParagraphFormat.Alignment = ParagraphAlignment.Center;

        var heading = _doc.Styles.AddStyle("SectionHeading", StyleNames.Normal);
        heading.Font.Size = 12;
        heading.Font.Bold = true;
        heading.ParagraphFormat.SpaceBefore = Unit.FromPoint(10);
        heading.ParagraphFormat.SpaceAfter = Unit.FromPoint(4);
        heading.ParagraphFormat.KeepWithNext = true;

        var note = _doc.Styles.AddStyle("Footnote", StyleNames.Normal);
        note.Font.Size = 8;
        note.Font.Italic = true;

        var caption = _doc.Styles.AddStyle("TableCaption", StyleNames.Normal);
        caption.Font.Size = 11;
        caption.Font.Bold = true;
        caption.ParagraphFormat.Alignment = ParagraphAlignment.Center;
        caption.ParagraphFormat.SpaceBefore = Unit.FromPoint(8);
        caption.ParagraphFormat.SpaceAfter = Unit.FromPoint(2);
        caption.ParagraphFormat.KeepWithNext = true;
    }

    private Section NewSection()
    {
        var section = _doc.AddSection();
        var ps = section.PageSetup;
        ps.PageFormat = PageFormat.A4;
        ps.Orientation = Orientation.Portrait;
        ps.LeftMargin = Unit.FromMillimeter(LeftMarginMm);
        ps.RightMargin = Unit.FromMillimeter(RightMarginMm);
        ps.TopMargin = Unit.FromMillimeter(12);
        ps.BottomMargin = Unit.FromMillimeter(12);
        ps.FooterDistance = Unit.FromMillimeter(5);
        return section;
    }

    private void AddHeader(Section section, string title)
    {
        section.AddParagraph(title, "ReportTitle");
        section.AddParagraph(_s.Subtitle, "ReportSubtitle");
        section.AddParagraph(_m.Caption, "ReportSubtitle");
        var p = section.AddParagraph(_m.ReportTime.ToString("dd.MM.yyyy H:mm", Formats.Ru), "ReportSubtitle");
        p.Format.Font.Size = 11;
        p.Format.SpaceAfter = Unit.FromPoint(4);
    }

    // ------------------------------------------------------------------ суточный отчет

    private void AddDailyReport()
    {
        var section = NewSection();
        AddHeader(section, _s.Title);
        AddLoadsSection(section);
        AddRemarksSection(section);
        AddChangesSection(section);
        if (_s.IncludeDeadlines)
            AddDeadlinesSection(section);
        AddNotes(section);
    }

    private void AddLoadsSection(Section section)
    {
        section.AddParagraph(
            "1. Загрузка документов в АИС ГГЭ.",
            "SectionHeading");

        var projects = _m.Projects.Where(p => p.PeriodLoads.Count > 0 || !_s.HideProjectsWithoutLoads).ToList();
        if (_m.Projects.All(p => p.PeriodLoads.Count == 0))
        {
            section.AddParagraph("За отчетный период загрузок документов не было.");
            return;
        }

        var each = _s.LoadsMode == LoadsMode.Each;
        var table = NewTable(section, new double[] { 11, 10, 17, 11, 17, 11, 15, 17, 13, 16, 19, 12, 14 }, 9);

        var h1 = HeaderRow(table);
        var h2 = HeaderRow(table);
        Head(h1, 0, "Этап", mergeDown: 1);
        Head(h1, 1, "№", mergeDown: 1);
        Head(h1, 2, "Начало", mergeRight: 1);
        Head(h1, 4, "Окончание", mergeRight: 1);
        string[] kinds = { "Инженерные изыскания", "Проектная документация", "Сметы", "Ведомости объемов работ", "Конъюнктурный анализ цен", "Иные" };
        for (var i = 0; i < kinds.Length; i++)
            Head(h1, 6 + i, kinds[i], mergeDown: 1, size: 7);
        Head(h1, 12, "Всего", mergeDown: 1);
        Head(h2, 2, "Дата", size: 8);
        Head(h2, 3, "Время", size: 8);
        Head(h2, 4, "Дата", size: 8);
        Head(h2, 5, "Время", size: 8);

        foreach (var p in projects)
        {
            if (each && p.PeriodLoads.Count > 0)
            {
                foreach (var load in p.PeriodLoads)
                    LoadRow(table, p.Number, load.Number ?? "", load.Start ?? load.End, load.End, load.Counts, false);
                if (p.PeriodLoads.Count > 1)
                    LoadRow(table, p.Number, "Итого", p.PeriodLoadsStart, p.PeriodLoadsEnd, p.PeriodLoadsSum, true);
            }
            else
            {
                LoadRow(table, p.Number, p.PeriodLoads.Count.ToString(Formats.Ru), p.PeriodLoadsStart, p.PeriodLoadsEnd, p.PeriodLoadsSum, false);
            }
        }

        var total = table.AddRow();
        total.Shading.Color = Palette.TotalsRow;
        Put(total, 0, "Всего:", bold: true, align: ParagraphAlignment.Right).MergeRight = 5;
        PutLoadCounts(total, _m.TotalPeriodLoads, bold: true);
    }

    private static void LoadRow(Table table, string stage, string number, DateTime? start, DateTime? end, LoadCounts counts, bool subtotal)
    {
        var row = table.AddRow();
        if (subtotal)
            row.Shading.Color = Palette.TotalsRow;
        Put(row, 0, stage, bold: true);
        Put(row, 1, number);
        Put(row, 2, start is null ? "" : Formats.Date(start.Value));
        Put(row, 3, start is null ? "" : Formats.Time(start.Value));
        Put(row, 4, end is null ? "" : Formats.Date(end.Value));
        Put(row, 5, end is null ? "" : Formats.Time(end.Value));
        PutLoadCounts(row, counts, bold: subtotal);
    }

    private static void PutLoadCounts(Row row, LoadCounts c, bool bold)
    {
        int[] values = { c.Survey, c.Project, c.Estimate, c.Vor, c.Kac, c.Other };
        for (var i = 0; i < values.Length; i++)
            Put(row, 6 + i, Formats.Number(values[i]), bold: bold, align: ParagraphAlignment.Right);
        Put(row, 12, Formats.Number(c.Total), bold: true, align: ParagraphAlignment.Right);
    }

    private void AddRemarksSection(Section section)
    {
        section.AddParagraph("2. Устранение замечаний.", "SectionHeading");
        var table = StatusTable(section, critical1: "%¹", critical2: "%²");
        foreach (var p in _m.Projects)
        {
            if (p.Current is null)
            {
                NoDataRow(table, p, 17);
                continue;
            }

            StatusRow(table, StageLabel(p), p.Current, p.Current, highlight: true);
        }

        var total = _m.TotalCurrent;
        StatusRow(table, "Всего:", total, total, highlight: false, totals: true);

        AddStaleFootnotes(section);
    }

    private void AddChangesSection(Section section)
    {
        var delta = _m.TotalDelta;
        section.AddParagraph(
            $"3. Изменение за период с {Short(_m.PeriodStart)} по {Short(_m.ReportTime)} " +
            $"(устранено всего {Formats.Number(delta.All[0].Fixed)} замечаний).",
            "SectionHeading");

        // Как в книге: таблица изменений без шапки — колонки те же, что в разделе 2.
        var table = StatusTable(section, critical1: "%¹", critical2: "%²", header: false);
        foreach (var p in _m.Projects)
        {
            if (p.Delta is null || p.Current is null)
            {
                NoDataRow(table, p, 17);
                continue;
            }

            // Проценты изменений — от текущего количества замечаний этапа (как в книге).
            StatusRow(table, StageLabel(p), p.Delta, p.Current, highlight: false);
        }

        StatusRow(table, "Всего:", delta, _m.TotalCurrent, highlight: false, totals: true);
        Footnote(section, "1. Процент общего числа критических замечаний приведен к общему количеству всех замечаний.");
        Footnote(section, "2. Процент критических замечаний приведен к общему количеству критических замечаний.");

        var cur = _m.TotalCurrent;
        var open = cur.All[0].Open;
        var openCritical = cur.Critical[0].Open;
        var summary = section.AddParagraph();
        summary.Format.SpaceBefore = Unit.FromPoint(8);
        summary.Format.Font.Bold = true;
        summary.AddText(
            $"На текущую дату не устранено {Formats.Number(open)} ({Formats.Percent(open, cur.All[0].Total)}%) замечаний, " +
            $"в том числе {Formats.Number(openCritical)} ({Formats.Percent(openCritical, cur.Critical[0].Total)}%) критических.");
    }

    /// <summary>Таблица «Все замечания / Критические замечания» (17 колонок).</summary>
    private static string Short(DateTime value) => value.ToString("H:mm dd.MM.yyyy", Formats.Ru);

    private Table StatusTable(Section section, string critical1, string critical2, bool header = true)
    {
        var widths = new List<double> { 11 };
        for (var i = 0; i < 8; i++)
        {
            widths.Add(13);
            widths.Add(8.75);
        }

        var table = NewTable(section, widths.ToArray(), 9);
        if (!header)
            return table;
        var h1 = HeaderRow(table);
        var h2 = HeaderRow(table);
        var h3 = HeaderRow(table);
        Head(h1, 0, "Этап", mergeDown: 2);
        Head(h1, 1, "Все замечания", mergeRight: 7);
        Head(h1, 9, "Критические замечания", mergeRight: 7);
        string[] names = { "Всего", "Устранено", "Устраняется", "Рассм. ГГЭ" };
        Color[] colors = { Palette.Total, Palette.Fixed, Palette.Fixing, Palette.Review };
        for (var block = 0; block < 2; block++)
        {
            for (var i = 0; i < 4; i++)
            {
                var col = 1 + block * 8 + i * 2;
                Head(h2, col, names[i], mergeRight: 1, color: colors[i]);
                Head(h3, col, "Шт.", color: colors[i], size: 8);
                var pct = block == 0 ? "%" : (i == 0 ? critical1 : critical2);
                Head(h3, col + 1, pct, color: colors[i], size: 8);
            }
        }

        return table;
    }

    /// <summary>Строка таблицы статусов. Проценты считаются от <paramref name="basis"/>.</summary>
    private static void StatusRow(Table table, string stage, RemarksSnapshot values, RemarksSnapshot basis, bool highlight, bool totals = false)
    {
        var row = table.AddRow();
        if (totals)
            row.Shading.Color = Palette.TotalsRow;
        Put(row, 0, stage, bold: true);

        var all = values.All[0];
        var crit = values.Critical[0];
        var allBasis = basis.All[0].Total;
        var critBasis = basis.Critical[0].Total;

        var allTotal = Put(row, 1, Formats.Number(all.Total), Palette.Total, bold: true, align: ParagraphAlignment.Right);
        Put(row, 2, ReferenceEquals(values, basis) ? Formats.Percent(all.Total, all.Total) : Formats.Percent(all.Total, allBasis), Palette.Total, size: 7);
        Put(row, 3, Formats.Number(all.Fixed), Palette.Fixed, bold: totals, align: ParagraphAlignment.Right);
        Put(row, 4, Formats.Percent(all.Fixed, allBasis), Palette.Fixed, size: 7);
        Put(row, 5, Formats.Number(all.Fixing), Palette.Fixing, bold: totals, align: ParagraphAlignment.Right);
        Put(row, 6, Formats.Percent(all.Fixing, allBasis), Palette.Fixing, size: 7);
        Put(row, 7, Formats.Number(all.Review), Palette.Review, bold: totals, align: ParagraphAlignment.Right);
        Put(row, 8, Formats.Percent(all.Review, allBasis), Palette.Review, size: 7);

        var critTotal = Put(row, 9, Formats.Number(crit.Total), Palette.Total, bold: true, align: ParagraphAlignment.Right);
        Put(row, 10, Formats.Percent(crit.Total, allBasis), Palette.Total, size: 7);
        Put(row, 11, Formats.Number(crit.Fixed), Palette.Fixed, bold: totals, align: ParagraphAlignment.Right);
        Put(row, 12, Formats.Percent(crit.Fixed, critBasis), Palette.Fixed, size: 7);
        Put(row, 13, Formats.Number(crit.Fixing), Palette.Fixing, bold: totals, align: ParagraphAlignment.Right);
        Put(row, 14, Formats.Percent(crit.Fixing, critBasis), Palette.Fixing, size: 7);
        Put(row, 15, Formats.Number(crit.Review), Palette.Review, bold: totals, align: ParagraphAlignment.Right);
        Put(row, 16, Formats.Percent(crit.Review, critBasis), Palette.Review, size: 7);

        // Несовпадение сумм подсвечивается красным фоном (как в книге).
        if (highlight)
        {
            if (!all.IsConsistent || !values.SectionsMatch(critical: false))
                allTotal.Shading.Color = Palette.Error;
            if (!crit.IsConsistent || !values.SectionsMatch(critical: true))
                critTotal.Shading.Color = Palette.Error;
        }
    }

    private void NoDataRow(Table table, ProjectReport p, int columns)
    {
        var row = table.AddRow();
        Put(row, 0, p.Number, bold: true);
        var cell = Put(row, 1, p.Error is not null ? "нет данных (ошибка чтения экрана)" : "нет данных", Palette.Muted);
        cell.MergeRight = columns - 2;
    }

    private string StageLabel(ProjectReport p) => p.IsStale(_m.ReportTime) ? p.Number + "*" : p.Number;

    private void AddStaleFootnotes(Section section)
    {
        foreach (var p in _m.Projects.Where(p => p.IsStale(_m.ReportTime)))
            Footnote(section, $"* Этап {p.Number}: данные экрана на {Formats.DateTimeText(p.Current!.Timestamp)} (на отчетное время строка не заполнена).");
    }

    private void AddDeadlinesSection(Section section)
    {
        var withDeadlines = _m.Projects.Where(p => p.Project.Deadlines.Count > 0).ToList();
        if (withDeadlines.Count == 0)
            return;

        var stages = _s.DeadlineStages
            .Concat(withDeadlines.SelectMany(p => p.Project.Deadlines.Keys))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(stage => withDeadlines.Any(p => FindDeadline(p.Project, stage) is not null))
            .ToList();
        if (stages.Count == 0)
            return;

        section.AddParagraph("4. Сроки.", "SectionHeading");
        var widths = new List<double> { 11 };
        foreach (var _ in stages)
        {
            widths.Add(10);
            widths.Add(10);
        }

        var table = NewTable(section, widths.ToArray(), 9);
        var h1 = HeaderRow(table);
        var h2 = HeaderRow(table);
        Head(h1, 0, "Этап", mergeDown: 1);
        for (var i = 0; i < stages.Count; i++)
        {
            Head(h1, 1 + i * 2, stages[i], mergeRight: 1);
            Head(h2, 1 + i * 2, "Начало", size: 8);
            Head(h2, 2 + i * 2, "Окончание", size: 8);
        }

        var today = _m.ReportTime.Date;
        foreach (var p in withDeadlines)
        {
            var row = table.AddRow();
            Put(row, 0, p.Number, bold: true);
            for (var i = 0; i < stages.Count; i++)
            {
                var values = FindDeadline(p.Project, stages[i]) ?? new List<string>();
                var col = 1 + i * 2;
                if (values.Count == 1 && Formats.ParseDate(values[0]) is null)
                {
                    // Произвольный текст (например, «Заявка до 23.10.2026») — на обе колонки.
                    var cell = Put(row, col, values[0], size: 8);
                    cell.MergeRight = 1;
                    cell.Format.Font.Italic = true;
                    continue;
                }

                var start = values.Count > 0 ? Formats.ParseDate(values[0]) : null;
                var end = values.Count > 1 ? Formats.ParseDate(values[1]) : null;
                var c1 = Put(row, col, start is null ? (values.Count > 0 ? values[0] : "") : Formats.Date(start.Value));
                var c2 = Put(row, col + 1, end is null ? (values.Count > 1 ? values[1] : "") : Formats.Date(end.Value));
                if (start is not null && end is not null && start.Value.Date <= today && today <= end.Value.Date)
                {
                    c1.Shading.Color = Palette.CurrentStage;
                    c2.Shading.Color = Palette.CurrentStage;
                }
            }
        }
    }

    private static List<string>? FindDeadline(ProjectConfig project, string stage) =>
        project.Deadlines.FirstOrDefault(kv => string.Equals(kv.Key.Trim(), stage.Trim(), StringComparison.OrdinalIgnoreCase)).Value;

    private void AddNotes(Section section)
    {
        var lines = _m.Notes.Where(l => !string.IsNullOrWhiteSpace(l)).ToList();
        if (lines.Count == 0)
            return;

        section.AddParagraph().Format.SpaceAfter = Unit.FromPoint(6);
        foreach (var line in lines)
        {
            var p = section.AddParagraph(line.Trim());
            p.Format.Font.Italic = true;
            p.Format.Alignment = ParagraphAlignment.Justify;
            p.Format.SpaceAfter = Unit.FromPoint(2);
        }
    }

    // ------------------------------------------------------------------ детальный отчет

    private void AddDetailedReport()
    {
        var section = NewSection();
        AddHeader(section, _s.DetailedTitle);
        foreach (var category in RemarkCategories.Values)
        {
            var title = category == RemarkCategory.All ? null : "в т.ч. " + RemarkCategories.Title(category);
            if (title is not null)
                section.AddParagraph(title, "TableCaption");
            else
                section.AddParagraph().Format.SpaceAfter = Unit.FromPoint(6);
            DetailedTable(section, category);
        }

        AddStaleFootnotes(section);
    }

    private void DetailedTable(Section section, RemarkCategory category)
    {
        var table = NewTable(section, new double[] { 13, 14, 13, 10, 13, 10, 13, 10, 14, 12, 10, 12, 10, 12, 10 }, 9);
        var h1 = HeaderRow(table);
        var h2 = HeaderRow(table);
        var h3 = HeaderRow(table);
        Head(h1, 0, "Этап", mergeDown: 2);
        Head(h1, 1, "Все замечания", mergeRight: 6);
        Head(h1, 8, "Критические замечания", mergeRight: 6);
        string[] names = { "Устранено", "Устраняется", "Рассм. ГГЭ" };
        Color[] colors = { Palette.Fixed, Palette.Fixing, Palette.Review };
        for (var block = 0; block < 2; block++)
        {
            var baseCol = 1 + block * 7;
            Head(h2, baseCol, "Всего", mergeDown: 1, color: Palette.Total);
            for (var i = 0; i < 3; i++)
            {
                var col = baseCol + 1 + i * 2;
                Head(h2, col, names[i], mergeRight: 1, color: colors[i]);
                Head(h3, col, "Шт.", color: colors[i], size: 8);
                Head(h3, col + 1, "%", color: colors[i], size: 8);
            }
        }

        // Держим таблицу на одной странице.
        h1.KeepWith = 2;

        var withData = _m.Projects.Where(p => p.Current is not null).ToList();
        foreach (var p in _m.Projects)
        {
            if (p.Current is null)
            {
                NoDataRow(table, p, 15);
                continue;
            }

            DetailedRow(table, StageLabel(p), p.Current.Get(category, false), p.Current.Get(category, true), false,
                highlight: category == RemarkCategory.All ? !p.Current.SectionsMatch(false) : false);
        }

        var all = withData.Aggregate(StatusCounts.Zero, (a, p) => a + p.Current!.Get(category, false));
        var crit = withData.Aggregate(StatusCounts.Zero, (a, p) => a + p.Current!.Get(category, true));
        DetailedRow(table, "Всего", all, crit, true, highlight: false);
    }

    private static void DetailedRow(Table table, string stage, StatusCounts all, StatusCounts crit, bool totals, bool highlight)
    {
        var row = table.AddRow();
        if (totals)
            row.Shading.Color = Palette.TotalsRow;
        Put(row, 0, stage, bold: true);
        StatusCounts[] blocks = { all, crit };
        for (var block = 0; block < 2; block++)
        {
            var c = blocks[block];
            var baseCol = 1 + block * 7;
            var totalCell = Put(row, baseCol, Formats.Number(c.Total), Palette.Total, bold: true, align: ParagraphAlignment.Right);
            if ((highlight && block == 0) || !c.IsConsistent)
                totalCell.Shading.Color = Palette.Error;
            int[] values = { c.Fixed, c.Fixing, c.Review };
            Color[] colors = { Palette.Fixed, Palette.Fixing, Palette.Review };
            for (var i = 0; i < 3; i++)
            {
                Put(row, baseCol + 1 + i * 2, Formats.Number(values[i]), colors[i], bold: totals, align: ParagraphAlignment.Right);
                Put(row, baseCol + 2 + i * 2, Formats.Percent(values[i], c.Total), colors[i], size: 7);
            }
        }
    }

    // ------------------------------------------------------------------ вспомогательные

    private static Table NewTable(Section section, double[] relativeWidths, double fontSize)
    {
        var table = section.AddTable();
        table.Borders.Width = 0.5;
        table.Borders.Color = Colors.Black;
        table.Format.Font.Size = fontSize;
        table.Rows.VerticalAlignment = VerticalAlignment.Center;
        table.TopPadding = Unit.FromPoint(1);
        table.BottomPadding = Unit.FromPoint(1);
        table.LeftPadding = Unit.FromPoint(1.5);
        table.RightPadding = Unit.FromPoint(1.5);
        var sum = relativeWidths.Sum();
        foreach (var w in relativeWidths)
            table.AddColumn(Unit.FromMillimeter(w / sum * ContentWidthMm));
        return table;
    }

    private static Row HeaderRow(Table table)
    {
        var row = table.AddRow();
        row.HeadingFormat = true;
        row.Format.Font.Bold = true;
        return row;
    }

    private static Cell Head(Row row, int column, string text, int mergeRight = 0, int mergeDown = 0, Color? color = null, double? size = null)
    {
        var cell = Put(row, column, text, color, bold: true, size: size);
        if (mergeRight > 0)
            cell.MergeRight = mergeRight;
        if (mergeDown > 0)
            cell.MergeDown = mergeDown;
        return cell;
    }

    private static Cell Put(Row row, int column, string text, Color? color = null, bool bold = false,
        ParagraphAlignment align = ParagraphAlignment.Center, double? size = null)
    {
        var cell = row.Cells[column];
        var p = cell.AddParagraph(text);
        p.Format.Alignment = align;
        if (color is { } c)
            p.Format.Font.Color = c;
        if (bold)
            p.Format.Font.Bold = true;
        if (size is { } s)
            p.Format.Font.Size = s;
        return cell;
    }

    private static void Footnote(Section section, string text)
    {
        var p = section.AddParagraph(text, "Footnote");
        p.Format.SpaceBefore = Unit.FromPoint(1);
    }
}
