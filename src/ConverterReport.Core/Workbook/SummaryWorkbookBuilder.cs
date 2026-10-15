using ClosedXML.Excel;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Model;
using ConverterReport.Core.Reporting;
using static ConverterReport.Core.Workbook.Formula;

namespace ConverterReport.Core.Workbook;

/// <summary>
/// Сводная книга Excel со структурой книги отчета: «Описание», «Д.N»/«З.N» (копии листов экранов),
/// «Свод» (SUMIFS по листам этапов — источник диаграмм), «Диаграммы», лист отчета на дату «ДД.ММ»
/// (с блоками «Конвертер»), «Отчет_дет» и «Отчет». Все расчеты в книге — формулы со ссылками.
/// </summary>
internal sealed class SummaryWorkbookBuilder
{
    private readonly ReportModel _m;
    private readonly ReportSettings _s;
    private readonly ScreenLayout _layout;
    private readonly List<ProjectReport> _projects;
    private readonly string _dailySheet;

    // Ссылки, заполняемые по ходу построения.
    private readonly Dictionary<ProjectReport, int> _documentsLastRow = new();
    private readonly Dictionary<(int Category, ProjectReport Project), int> _dataRows = new();

    public SummaryWorkbookBuilder(ReportModel model)
    {
        _m = model;
        _s = model.Config.Report;
        _layout = model.Config.ScreenLayout;
        _projects = model.Projects.Where(p => p.Screen is not null).ToList();
        _dailySheet = model.ReportTime.ToString("dd.MM", Formats.Ru);
    }

    public SummaryWorkbook Build(string path)
    {
        using var wb = new XLWorkbook();
        wb.Style.Font.FontName = "Calibri";
        wb.Style.Font.FontSize = 11;

        Description(wb.AddWorksheet("Описание"));
        foreach (var p in _projects)
        {
            Documents(wb.AddWorksheet("Д." + SheetSuffix(p)), p);
            Remarks(wb.AddWorksheet("З." + SheetSuffix(p)), p);
        }

        var svod = new SvodLayout(SvodStart(), SvodEnd());
        Svod(wb.AddWorksheet(SvodLayout.SheetName), svod);
        var charts = wb.AddWorksheet(SummaryWorkbook.ChartsSheet);
        var daily = wb.AddWorksheet(_dailySheet);
        var detailed = wb.AddWorksheet(SummaryWorkbook.DetailedSheet);
        var data = wb.AddWorksheet(SummaryWorkbook.DataSheet);

        Daily(daily);
        DataSheet(data);
        Detailed(detailed);
        var result = new SummaryWorkbook { Path = path, DailySheet = _dailySheet, ChartPages = ChartsSheet(charts, svod, out var placements) };
        result.Charts.AddRange(placements);

        daily.SetTabActive();
        daily.Select();
        wb.CalculateMode = XLCalculateMode.Auto;
        wb.FullCalculationOnLoad = true;
        wb.ForceFullCalculation = true;
        wb.SaveAs(path);
        return result;
    }

    private static string SheetSuffix(ProjectReport p) => Formats.SafeFileName(p.Number).Replace("[", "(").Replace("]", ")");

    private string D(ProjectReport p) => XlStyle.Q("Д." + SheetSuffix(p));

    private string Z(ProjectReport p) => XlStyle.Q("З." + SheetSuffix(p));

    private static string C(int column) => XLHelper.GetColumnLetterFromNumber(column);

    private static int Col(string letter) => XLHelper.GetColumnNumberFromLetter(letter.Trim().ToUpperInvariant());

    // ================================================================== Описание

    private void Description(IXLWorksheet ws)
    {
        string[] lines =
        {
            "Сводная книга сформирована программой ConverterReport автоматически — вручную заполнять ничего не нужно.",
            $"Отчетная дата и время: {Formats.DateTimeText(_m.ReportTime)}; отчетный период с {Formats.DateTimeText(_m.PeriodStart)}.",
            "Д.N, З.N — копии листов «Документы» и «Замечания» экранов этапов (данные на момент формирования).",
            "Свод — состояние замечаний и загрузки по дням (формулы СУММЕСЛИМН по листам Д.N/З.N); источник диаграмм.",
            $"{_dailySheet} — суточный отчет. Блоки «Конвертер» справа от области печати: II — загрузки за период " +
            "(сумма всех загрузок, окончание которых попадает в период), III — замечания на отчетное время, IV — на начало периода.",
            "Отчет — данные по разделам (ссылки на строки листов З.N); Отчет_дет — детальный отчет по листу «Отчет».",
            "Диаграммы — диаграммы по листу «Свод».",
            "Ячейки «Всего» подсвечиваются красным, если не совпадает сумма по статусам.",
        };
        ws.Cell(1, 2).Value = "Сводная книга отчета ГГЭ";
        ws.Cell(1, 2).Style.Font.Bold = true;
        for (var i = 0; i < lines.Length; i++)
        {
            ws.Cell(i + 2, 1).Value = $"{i + 1}.";
            ws.Cell(i + 2, 2).Value = lines[i];
        }

        ws.Column(1).Width = 4;
        ws.Column(2).Width = 150;
        ws.Range(2, 2, lines.Length + 1, 2).Style.Alignment.WrapText = true;
    }

    // ================================================================== Д.N и З.N

    /// <summary>Копия листа «Документы»: значения, даты/время приведены к датам Excel, «Всего» — формулой,
    /// W/X — начало/окончание загрузки (дата+время) для расчета загрузок за период.</summary>
    private void Documents(IXLWorksheet ws, ProjectReport p)
    {
        var raw = p.Screen!.RawDocuments;
        var first = _layout.DocumentsFirstRow;
        int[] dateCols = { Col(_layout.LoadStartDate), Col(_layout.LoadEndDate) };
        int[] timeCols = { Col(_layout.LoadStartTime), Col(_layout.LoadEndTime) };
        var countCols = new[] { _layout.LoadSurvey, _layout.LoadProject, _layout.LoadEstimate, _layout.LoadVor, _layout.LoadKac, _layout.LoadOther, _layout.LoadAnswers, _layout.LoadCancelled, _layout.LoadErrors }
            .Where(c => !string.IsNullOrWhiteSpace(c)).Select(Col).ToHashSet();
        var totalCol = Col(_layout.LoadOther) + 1;

        var last = first;
        if (raw is not null)
        {
            CopyRaw(ws, raw, (row, col, value) =>
            {
                if (row < first)
                    return value;
                if (dateCols.Contains(col))
                    return Formats.ParseDate(value)?.Date ?? value;
                if (timeCols.Contains(col))
                    return (object?)Formats.ParseTime(value) ?? value;
                if (countCols.Contains(col) && value is string s)
                    return Formats.IsNumber(s) ? Formats.ParseInt(s) : null;
                return value;
            });

            // Последняя строка таблицы загрузок — до строки «Итого:» (в ней суммы формулами).
            last = Math.Max(first, raw.LastRow);
            for (var r = first; r <= raw.LastRow; r++)
            {
                if (raw.Get(r, Col(_layout.LoadEndTime)) is string t && t.Contains("Итого", StringComparison.OrdinalIgnoreCase))
                {
                    last = r - 1;
                    var sumCols = countCols.Append(totalCol).ToList();
                    foreach (var c in sumCols)
                        ws.Cell(r, c).FormulaA1 = $"SUM({C(c)}{first}:{C(c)}{last})";
                    break;
                }
            }
        }

        var sd = Col(_layout.LoadStartDate);
        var st = Col(_layout.LoadStartTime);
        var ed = Col(_layout.LoadEndDate);
        var et = Col(_layout.LoadEndTime);
        var startCol = 23; // W
        var endCol = 24;   // X
        ws.Cell(2, startCol).Value = "Начало (дата и время)";
        ws.Cell(2, endCol).Value = "Окончание (дата и время)";
        ws.Range(2, startCol, 2, endCol).Style.Font.Bold = true;
        for (var r = first; r <= last; r++)
        {
            ws.Cell(r, totalCol).FormulaA1 = $"SUM({C(Col(_layout.LoadSurvey))}{r}:{C(Col(_layout.LoadOther))}{r})";
            ws.Cell(r, endCol).FormulaA1 =
                $"IF(ISNUMBER({C(ed)}{r}),{C(ed)}{r}+N({C(et)}{r}),IF(ISNUMBER({C(sd)}{r}),{C(sd)}{r}+N({C(st)}{r}),0))";
            ws.Cell(r, startCol).FormulaA1 = $"IF(ISNUMBER({C(sd)}{r}),{C(sd)}{r}+N({C(st)}{r}),{C(endCol)}{r})";
        }

        ws.Range(first, startCol, Math.Max(first, last), endCol).Style.NumberFormat.Format = XlStyle.DateTimeFormat;
        foreach (var c in dateCols)
            ws.Column(c).Style.NumberFormat.Format = XlStyle.DateFormat;
        foreach (var c in timeCols)
            ws.Column(c).Style.NumberFormat.Format = XlStyle.TimeFormat;
        ws.Range(1, 1, 2, totalCol + 3).Style.Font.Bold = true;
        ws.Column(startCol).Width = 18;
        ws.Column(endCol).Width = 18;
        ws.SheetView.FreezeRows(2);
        _documentsLastRow[p] = Math.Max(first, last);
    }

    /// <summary>Копия листа «Замечания»: значения, дата и время приведены к датам Excel, день недели и «устранено за сутки» — формулами.</summary>
    private void Remarks(IXLWorksheet ws, ProjectReport p)
    {
        var raw = p.Screen!.RawRemarks;
        var first = _layout.RemarksFirstRow;
        var dateCol = Col(_layout.RemarksDate);
        var timeCol = Col(_layout.RemarksTime);
        var totalCol = Col(_layout.RemarksAll);
        if (raw is null)
            return;

        CopyRaw(ws, raw, (row, col, value) =>
        {
            if (row < first)
                return value;
            if (col == dateCol)
                return Formats.ParseDate(value)?.Date ?? value;
            if (col == timeCol)
                return (object?)Formats.ParseTime(value) ?? value;
            if (col == dateCol + 1)
                return null; // день недели — формулой
            return value;
        });

        for (var r = first; r <= raw.LastRow; r++)
        {
            if (raw.Get(r, dateCol) is null)
                continue;
            ws.Cell(r, dateCol + 1).FormulaA1 = $"IF(ISNUMBER({C(dateCol)}{r}),WEEKDAY({C(dateCol)}{r},2),\"\")";
        }

        ws.Column(dateCol).Style.NumberFormat.Format = XlStyle.DateFormat;
        ws.Column(timeCol).Style.NumberFormat.Format = XlStyle.TimeFormat;
        ws.Range(1, 1, 2, Math.Max(totalCol + 37, raw.LastColumn)).Style.Font.Bold = true;
        ws.SheetView.FreezeRows(2);
    }

    private static void CopyRaw(IXLWorksheet ws, RawSheet raw, Func<int, int, object?, object?> transform)
    {
        foreach (var ((row, col), value) in raw.Cells)
        {
            var v = transform(row, col, value);
            var cell = ws.Cell(row, col);
            switch (v)
            {
                case null:
                    break;
                case DateTime dt:
                    cell.Value = dt;
                    break;
                case TimeSpan ts:
                    cell.Value = ts.TotalDays;
                    break;
                case double d:
                    cell.Value = d;
                    break;
                case int i:
                    cell.Value = i;
                    break;
                case bool b:
                    cell.Value = b;
                    break;
                default:
                    cell.Value = v.ToString();
                    break;
            }
        }

        foreach (var (col, width) in raw.ColumnWidths)
            ws.Column(col).Width = Math.Clamp(width, 3, 60);
    }

    // ================================================================== Свод

    private DateTime SvodStart()
    {
        var dates = _m.Projects.Where(p => p.Daily.Count > 0).Select(p => p.Daily[0].Date).ToList();
        if (_m.Summary is { Daily.Count: > 0 } s)
            dates.Add(s.Daily[0].Date);
        return dates.Count > 0 ? dates.Min() : _m.ReportTime.Date;
    }

    private DateTime SvodEnd()
    {
        var dates = _m.Projects.Where(p => p.Daily.Count > 0).Select(p => p.Daily[^1].Date).ToList();
        if (_m.Summary is { Daily.Count: > 0 } s)
            dates.Add(s.Daily[^1].Date);
        dates.Add(_m.ReportTime.Date);
        return dates.Max();
    }

    private void Svod(IXLWorksheet ws, SvodLayout svod)
    {
        var lastSlot = _s.ReportTimes.Select(t => Formats.ParseTime(t)).Where(t => t is not null).Select(t => t!.Value)
            .DefaultIfEmpty(new TimeSpan(18, 0, 0)).Max();

        ws.Cell(1, SvodLayout.DateColumn).Value = "Дата";
        ws.Cell(1, SvodLayout.WeekdayColumn).Value = "День";
        ws.Cell(1, SvodLayout.LabelColumn).Value = "Дата";
        ws.Cell(1, SvodLayout.TimeColumn).Value = "Время";
        BlockHeader(ws, SvodLayout.Overall, null);
        for (var i = 0; i < _projects.Count; i++)
            BlockHeader(ws, SvodLayout.Project(i), _projects[i].Name);

        for (var r = SvodLayout.FirstRow; r <= svod.LastRow; r++)
        {
            var date = svod.Start.AddDays(r - SvodLayout.FirstRow);
            if (r == SvodLayout.FirstRow)
                ws.Cell(r, SvodLayout.DateColumn).Value = date;
            else
                ws.Cell(r, SvodLayout.DateColumn).FormulaA1 = $"B{r - 1}+1";
            ws.Cell(r, SvodLayout.WeekdayColumn).FormulaA1 = $"WEEKDAY(B{r},2)";
            var weekday = date.ToString("ddd", Formats.Ru);
            ws.Cell(r, SvodLayout.LabelColumn).Value = date.ToString("dd.MM.yyyy", Formats.Ru) + " " + char.ToUpper(weekday[0]) + weekday[1..];
            // Время среза: на отчетную дату — время отчета, в остальные дни — последнее стандартное (18:00).
            var time = date == _m.ReportTime.Date ? _m.ReportTime.TimeOfDay : lastSlot;
            ws.Cell(r, SvodLayout.TimeColumn).Value = time.TotalDays;

            for (var i = 0; i < _projects.Count; i++)
                BlockRow(ws, SvodLayout.Project(i), r, _projects[i]);
            OverallRow(ws, r);
        }

        ws.Column(SvodLayout.DateColumn).Style.NumberFormat.Format = XlStyle.DateFormat;
        ws.Column(SvodLayout.TimeColumn).Style.NumberFormat.Format = XlStyle.TimeFormat;
        ws.Row(1).Style.Font.Bold = true;
        ws.Row(1).Style.Alignment.WrapText = true;
        ws.Column(1).Width = 3;
        ws.Column(SvodLayout.DateColumn).Width = 11;
        ws.Column(SvodLayout.WeekdayColumn).Width = 5;
        ws.Column(SvodLayout.LabelColumn).Width = 15;
        ws.Column(SvodLayout.TimeColumn).Width = 6;
        ws.SheetView.FreezeRows(1);
        ws.SheetView.FreezeColumns(SvodLayout.TimeColumn);
    }

    private static void BlockHeader(IXLWorksheet ws, SvodBlock b, string? name)
    {
        if (name is not null)
        {
            ws.Cell(1, b.Label).Value = name;
            ws.Cell(1, b.Label).Style.Fill.BackgroundColor = XlStyle.HeaderFill;
        }

        string[] remarks = { "Всего", "Устранено", "Устраняется", "Рассматривается", "Устранено за сутки" };
        string[] loads = { "ИИ", "ПД", "СМ", "ВОР", "КАЦ", "Иные", "Всего", "Ответы", "Аннулировано", "Ошибки" };
        for (var i = 0; i < remarks.Length; i++)
            ws.Cell(1, b.Remarks + i).Value = remarks[i];
        for (var i = 0; i < loads.Length; i++)
            ws.Cell(1, b.Loads + i).Value = loads[i];
    }

    private void BlockRow(IXLWorksheet ws, SvodBlock b, int r, ProjectReport p)
    {
        var z = Z(p);
        // Диапазоны ограничены строками данных (а не целыми столбцами) — так формулы быстрее и понятнее.
        var zFirst = _layout.RemarksFirstRow;
        var zLast = Math.Max(zFirst, p.Screen!.RawRemarks?.LastRow ?? zFirst);
        string ZR(string column) => $"{z}!${column}${zFirst}:${column}${zLast}";
        string Sumifs(string column) =>
            $"SUMIFS({ZR(column)},{ZR(_layout.RemarksDate)},$B{r},{ZR(_layout.RemarksTime)},$E{r})";

        var all = Col(_layout.RemarksAll);
        var total = C(b.Total) + r;
        ws.Cell(r, b.Total).FormulaA1 = $"IF({Sumifs(C(all))}=0,NA(),{Sumifs(C(all))})";
        for (var k = 1; k <= 3; k++)
            ws.Cell(r, b.Total + k).FormulaA1 = $"IF(ISNA({total}),NA(),{Sumifs(C(all + k))})";
        ws.Cell(r, b.FixedPerDay).FormulaA1 = r == SvodLayout.FirstRow ? "NA()" : $"IFERROR({C(b.Fixed)}{r}-{C(b.Fixed)}{r - 1},NA())";

        var d = D(p);
        var dFirst = _layout.DocumentsFirstRow;
        var dLast = _documentsLastRow.GetValueOrDefault(p, dFirst);
        string DR(string column) => $"{d}!${column}${dFirst}:${column}${dLast}";
        var endDate = _layout.LoadEndDate;
        var loadCols = new[] { _layout.LoadSurvey, _layout.LoadProject, _layout.LoadEstimate, _layout.LoadVor, _layout.LoadKac, _layout.LoadOther };
        for (var i = 0; i < loadCols.Length; i++)
            ws.Cell(r, b.Load(i)).FormulaA1 = $"SUMIFS({DR(loadCols[i])},{DR(endDate)},$B{r})";
        ws.Cell(r, b.LoadTotal).FormulaA1 = $"SUM({C(b.Load(0))}{r}:{C(b.Load(5))}{r})";
        var extra = new[] { _layout.LoadAnswers, _layout.LoadCancelled, _layout.LoadErrors };
        for (var i = 0; i < extra.Length; i++)
        {
            if (string.IsNullOrWhiteSpace(extra[i]))
                continue;
            ws.Cell(r, b.Load(7 + i)).FormulaA1 = $"SUMIFS({DR(extra[i])},{DR(endDate)},$B{r})";
        }
    }

    private void OverallRow(IXLWorksheet ws, int r)
    {
        var o = SvodLayout.Overall;
        if (_projects.Count == 0)
            return;
        var blocks = Enumerable.Range(0, _projects.Count).Select(SvodLayout.Project).ToList();

        // Итог по всем этапам: сумма по этапам, у которых на эту дату заполнена строка.
        for (var k = 0; k < 4; k++)
        {
            var sum = string.Join("+", blocks.Select(b => $"IF(ISNA({C(b.Total)}{r}),0,{C(b.Total + k)}{r})"));
            var any = string.Join("+", blocks.Select(b => $"ISNUMBER({C(b.Total)}{r})"));
            ws.Cell(r, o.Total + k).FormulaA1 = $"IF(({any})=0,NA(),{sum})";
        }

        ws.Cell(r, o.FixedPerDay).FormulaA1 = r == SvodLayout.FirstRow ? "NA()" : $"IFERROR({C(o.Fixed)}{r}-{C(o.Fixed)}{r - 1},NA())";
        for (var i = 0; i < 10; i++)
        {
            if (i == 6)
            {
                ws.Cell(r, o.LoadTotal).FormulaA1 = $"SUM({C(o.Load(0))}{r}:{C(o.Load(5))}{r})";
                continue;
            }

            ws.Cell(r, o.Load(i)).FormulaA1 = string.Join("+", blocks.Select(b => C(b.Load(i)) + r));
        }
    }

    // ================================================================== Лист отчета «ДД.ММ»

    private void Daily(IXLWorksheet ws)
    {
        var n = _projects.Count;
        ws.Style.Font.FontName = XlStyle.Font;
        ws.Style.Font.FontSize = 14;
        ws.Style.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        for (var c = 1; c <= 26; c++)
            ws.Column(c).Width = 4.66;
        ws.Column(4).Width = 5.11;
        ws.Column(5).Width = 5.66;
        ws.Column(9).Width = 6;
        ws.Column(27).Width = 9.11;
        ws.Column(28).Width = 7;
        for (var c = 29; c <= 46; c++)
            ws.Column(c).Width = 11;

        // --- Шапка и отчетные дата/время (AE3:AF4 — как в книге).
        ws.Put(1, 1, 26, _s.Title).Set(14, bold: true);
        ws.Put(2, 1, 26, _s.Subtitle).Set(14);
        ws.Put(3, 1, 26, _m.Caption).Set(14);
        ws.Put(4, 1, 26, F("AE4+AF4")).Set(12, format: XlStyle.DateTimeFormat);
        ws.Row(1).Height = 17.4;
        ws.Row(4).Height = 18.75;

        Converter(ws, 3, "0.", "Начало периода (предыдущий отчет):");
        Converter(ws, 4, "I.", "Отчетная дата и время:");
        ws.Cell(3, 31).Value = _m.PeriodStart.Date;
        ws.Cell(3, 32).Value = _m.PeriodStart.TimeOfDay.TotalDays;
        ws.Cell(4, 31).Value = _m.ReportTime.Date;
        ws.Cell(4, 32).Value = _m.ReportTime.TimeOfDay.TotalDays;
        ws.Range(3, 31, 4, 31).Style.NumberFormat.Format = XlStyle.DateFormat;
        ws.Range(3, 32, 4, 32).Style.NumberFormat.Format = XlStyle.TimeFormat;
        // Тексты для заголовка раздела 3 (без функции ТЕКСТ — она зависит от языка Excel).
        ws.Cell(3, 34).Value = PeriodText(_m.PeriodStart);
        ws.Cell(4, 34).Value = PeriodText(_m.ReportTime);
        ws.Range(3, 27, 4, 34).Style.Font.FontSize = 11;

        // --- 1. Загрузка документов.
        ws.Cell(5, 1).Value = "1. Загрузка документов в АИС ГГЭ.";
        ws.Cell(5, 1).Style.Font.Bold = true;
        Converter(ws, 5, "II.", "Конвертер - Документы (сумма загрузок за отчетный период)");
        ws.Row(6).Height = 9;

        const int h = 7;
        ws.Put(h, 1, h + 1, 2, "Этап").Set(14, bold: true);
        ws.Put(h, 3, h + 1, 3, "№").Set(14, bold: true);
        ws.Put(h, 4, h, 7, "Начало").Set(14, bold: true);
        ws.Put(h, 8, h, 11, "Окончание").Set(14, bold: true);
        string[] kinds = { "Инженерные изыскания", "Проектная документация", "Сметы", "Ведомости объемов работ", "Конъюнктурный анализ цен", "Иные" };
        for (var i = 0; i < kinds.Length; i++)
            ws.Put(h, 12 + i * 2, h + 1, 13 + i * 2, kinds[i]).Set(i == 2 || i == 5 ? 9 : 7, bold: true, wrap: true);
        ws.Put(h, 24, h + 1, 26, "Всего").Set(14, bold: true);
        ws.Put(h + 1, 4, h + 1, 5, "Дата").Set(11, bold: true);
        ws.Put(h + 1, 6, h + 1, 7, "Время").Set(11, bold: true);
        ws.Put(h + 1, 8, h + 1, 9, "Дата").Set(11, bold: true);
        ws.Put(h + 1, 10, h + 1, 11, "Время").Set(11, bold: true);
        ws.Range(h, 1, h + 1, 26).Box();
        ws.Row(h).Height = 17.4;
        ws.Row(h + 1).Height = 18.75;
        ConverterLoadsHeader(ws, h);

        var first1 = h + 2;
        for (var i = 0; i < n; i++)
        {
            var r = first1 + i;
            var p = _projects[i];
            ws.Put(r, 1, r, 2, XlStyle.Stage(p.Number)).Set(14, bold: true);
            ws.Put(r, 3, r, 3, F($"AB{r}")).Set(14);
            ws.Put(r, 4, r, 5, F($"IF($AB{r}=0,\"\",AC{r})")).Set(10, format: XlStyle.DateFormat);
            ws.Put(r, 6, r, 7, F($"IF($AB{r}=0,\"\",AD{r})")).Set(10, format: XlStyle.TimeFormat);
            ws.Put(r, 8, r, 9, F($"IF($AB{r}=0,\"\",AE{r})")).Set(10, format: XlStyle.DateFormat);
            ws.Put(r, 10, r, 11, F($"IF($AB{r}=0,\"\",AF{r})")).Set(10, format: XlStyle.TimeFormat);
            for (var k = 0; k < 6; k++)
                ws.Put(r, 12 + k * 2, r, 13 + k * 2, F($"{C(33 + k)}{r}")).Set(14, align: XLAlignmentHorizontalValues.Right, format: XlStyle.IntFormat);
            ws.Put(r, 24, r, 26, F($"AM{r}")).Set(14, bold: true, align: XLAlignmentHorizontalValues.Right, format: XlStyle.IntFormat);
            ws.Row(r).Height = 18;
            ConverterLoadsRow(ws, r, p);
            // Этапы без загрузок за период скрываются (как в книге — строки не удаляются).
            if (_s.HideProjectsWithoutLoads && p.PeriodLoads.Count == 0)
                ws.Row(r).Hide();
        }

        if (n > 0)
            ws.Range(first1, 1, first1 + n - 1, 26).Box();
        var t1 = first1 + n;
        ws.Put(t1, 10, t1, 11, "Всего:").Set(14, bold: true, align: XLAlignmentHorizontalValues.Right);
        for (var k = 0; k < 7; k++)
        {
            var c1 = 12 + k * 2;
            var c2 = k == 6 ? 26 : c1 + 1;
            ws.Put(t1, c1, t1, c2, F(n == 0 ? "0" : $"SUM({C(c1)}{first1}:{C(c1)}{t1 - 1})"))
                .Set(14, bold: true, align: XLAlignmentHorizontalValues.Right, format: XlStyle.IntFormat);
        }

        ws.Range(t1, 10, t1, 26).Box();
        ws.Row(t1).Height = 18;
        ConverterLoadsTotal(ws, t1, first1, n);
        if (_projects.All(p => p.PeriodLoads.Count == 0))
        {
            ws.Put(t1 + 1, 1, t1 + 1, 26, "За отчетный период загрузок документов не было.").Set(12, italic: true, align: XLAlignmentHorizontalValues.Left);
        }

        // --- 2. Устранение замечаний.
        var s2 = t1 + 2;
        ws.Cell(s2, 1).Value = "2. Устранение замечаний.";
        ws.Cell(s2, 1).Style.Font.Bold = true;
        ws.Row(s2 + 1).Height = 9;
        var h2 = s2 + 2;
        StatusHeader(ws, h2);
        Converter(ws, h2, "III.", "Конвертер - Замечания (строка экрана на отчетное время)");
        ConverterRemarksHeader(ws, h2, 29);
        ws.Cell(h2, 38).Value = "IV. Конвертер - Замечания на начало периода";
        ws.Cell(h2, 38).Style.Font.Bold = true;
        ws.Cell(h2, 38).Style.Font.FontSize = 11;
        ConverterRemarksHeader(ws, h2, 39);

        var first2 = h2 + 3;
        for (var i = 0; i < n; i++)
        {
            var r = first2 + i;
            var p = _projects[i];
            var label = p.IsStale(_m.ReportTime) ? p.Number + "*" : p.Number;
            ws.Put(r, 1, r, 2, p.IsStale(_m.ReportTime) ? label : XlStyle.Stage(p.Number)).Set(14, bold: true);
            StatusRow(ws, r, r, r, false);
            ConverterRemarksRow(ws, r, 29, p, p.Current);
            ConverterRemarksRow(ws, r, 39, p, p.Previous ?? p.Current);
        }

        var t2 = first2 + n;
        ws.Put(t2, 1, t2, 2, "Всего:").Set(14, bold: true, align: XLAlignmentHorizontalValues.Right);
        StatusRow(ws, t2, t2, t2, true);
        ConverterRemarksTotal(ws, t2, first2, n, 29);
        ConverterRemarksTotal(ws, t2, first2, n, 39);
        ws.Range(h2, 1, t2, 26).Box();
        StaleNotes(ws, t2 + 1);

        // --- 3. Изменение за период.
        var s3 = t2 + 2;
        var t3Converter = s3 + 2 + n;
        ws.Put(s3, 1, s3, 26, F($"\"3. Изменение за период с \"&$AH$3&\" по \"&$AH$4&\" (устранено всего \"&FIXED(AD{t3Converter},0)&\" замечаний).\""))
            .Set(14, bold: true, align: XLAlignmentHorizontalValues.Left);
        ws.Row(s3).Height = 18.75;
        ws.Row(s3 + 1).Height = 9;
        Converter(ws, s3, "V.", "Изменение за период (III − IV)");
        var first3 = s3 + 2;
        for (var i = 0; i < n; i++)
        {
            var r = first3 + i;
            var basisRow = first2 + i;
            var p = _projects[i];
            ws.Put(r, 1, r, 2, XlStyle.Stage(p.Number)).Set(14, bold: true);
            StatusRow(ws, r, r, basisRow, false);
            ws.Cell(r, 28).Value = XlStyle.StageCell(p.Number);
            for (var k = 0; k < 8; k++)
                ws.Cell(r, 29 + k).FormulaA1 = $"{C(29 + k)}{basisRow}-{C(39 + k)}{basisRow}";
        }

        var t3 = first3 + n;
        ws.Put(t3, 1, t3, 2, "Всего:").Set(14, bold: true, align: XLAlignmentHorizontalValues.Right);
        StatusRow(ws, t3, t3, t2, true);
        ConverterRemarksTotal(ws, t3, first3, n, 29);
        ws.Range(first3, 1, t3, 26).Box();

        var fn = t3 + 1;
        ws.Put(fn, 1, fn, 26, "1. Процент общего числа критических замечаний приведен к общему количеству всех замечаний.")
            .Set(10, italic: true, align: XLAlignmentHorizontalValues.Left);
        ws.Put(fn + 1, 1, fn + 1, 26, "2. Процент критических замечаний приведен к общему количеству критических замечаний.")
            .Set(10, italic: true, align: XLAlignmentHorizontalValues.Left);
        ws.Row(fn).Height = 12;
        ws.Row(fn + 1).Height = 12;

        var sum = fn + 3;
        ws.Put(sum, 1, sum, 26, F(
                $"\"На текущую дату не устранено \"&FIXED(I{t2}+L{t2},0)&\" (\"&FIXED(IF(C{t2}=0,0,(I{t2}+L{t2})/C{t2}*100),1)" +
                $"&\"%) замечаний, в том числе \"&FIXED(U{t2}+X{t2},0)&\" (\"&FIXED(IF(O{t2}=0,0,(U{t2}+X{t2})/O{t2}*100),1)&\"%) критических.\""))
            .Set(14, bold: true, align: XLAlignmentHorizontalValues.Left);
        ws.Row(sum).Height = 18.75;

        // --- 4. Сроки.
        var last = sum;
        if (_s.IncludeDeadlines)
            last = Deadlines(ws, sum + 3);

        // --- Справочно.
        var notes = _m.Notes.Where(l => !string.IsNullOrWhiteSpace(l)).Select(l => l.Trim()).ToList();
        if (notes.Count > 0)
        {
            var r = last + 2;
            var text = string.Join("\n", notes);
            ws.Put(r, 1, r, 26, text).Set(14, italic: true, align: XLAlignmentHorizontalValues.Left, wrap: true);
            var lines = notes.Sum(l => Math.Max(1, (int)Math.Ceiling(l.Length / 95.0)));
            ws.Row(r).Height = Math.Min(400, lines * 18.75 + 4);
            last = r;
        }

        // --- Печать: A4, книжная, по ширине страницы.
        var ps = ws.PageSetup;
        ps.PrintAreas.Clear();
        ps.PrintAreas.Add(1, 1, last, 26);
        ps.PaperSize = XLPaperSize.A4Paper;
        ps.PageOrientation = XLPageOrientation.Portrait;
        ps.FitToPages(1, 0);
        ps.CenterHorizontally = true;
        Margins(ps);
        ws.SheetView.ZoomScale = 70;
    }

    private static string PeriodText(DateTime value) =>
        value.ToString("H:mm", Formats.Ru) + " " + value.ToString("dd.MM.yyyy", Formats.Ru);

    private static void Converter(IXLWorksheet ws, int row, string number, string title)
    {
        ws.Cell(row, 27).Value = number;
        ws.Cell(row, 28).Value = title;
        ws.Range(row, 27, row, 28).Style.Font.Bold = true;
        ws.Range(row, 27, row, 28).Style.Font.FontSize = 11;
    }

    private static void ConverterLoadsHeader(IXLWorksheet ws, int h)
    {
        string[] top = { "Этап", "Кол-во загрузок", "Начало загрузки", "", "Окончание загрузки", "", "Загрузка" };
        string[] sub = { "", "", "Дата", "Время", "Дата", "Время", "ИИ", "ПД", "СМ", "ВОР", "КАЦ", "Иные", "Всего", "Ответы", "Аннулировано", "Ошибки", "", "Начало (дата и время)", "Окончание (дата и время)" };
        for (var i = 0; i < top.Length; i++)
            ws.Cell(h, 27 + i).Value = top[i];
        for (var i = 0; i < sub.Length; i++)
            ws.Cell(h + 1, 27 + i).Value = sub[i];
        var range = ws.Range(h, 27, h + 1, 27 + sub.Length - 1);
        range.Style.Font.FontSize = 10;
        range.Style.Font.Bold = true;
        range.Style.Fill.BackgroundColor = XlStyle.ConverterFill;
        range.Style.Alignment.WrapText = true;
    }

    /// <summary>Блок «Конвертер - Документы»: загрузки этапа, окончание которых попадает в отчетный период (AE3+AF3; AE4+AF4].</summary>
    private void ConverterLoadsRow(IXLWorksheet ws, int r, ProjectReport p)
    {
        var d = D(p);
        var last = _documentsLastRow.GetValueOrDefault(p, _layout.DocumentsFirstRow);
        var first = _layout.DocumentsFirstRow;
        string Rng(string col) => $"{d}!${col}${first}:${col}${last}";
        var cond = $"(({Rng("X")}>$AE$3+$AF$3)*({Rng("X")}<=$AE$4+$AF$4))";

        ws.Cell(r, 27).Value = XlStyle.StageCell(p.Number);
        ws.Cell(r, 28).FormulaA1 = $"SUMPRODUCT({cond}*1)";
        // AR/AS — начало первой и окончание последней загрузки периода (дата и время).
        // Округление до минуты, чтобы время отображалось без погрешности (22:20, а не 22:19).
        ws.Cell(r, 44).FormulaA1 = $"IF($AB{r}=0,\"\",ROUND(SUMPRODUCT(MIN({Rng("W")}+(1-{cond})*1E+9))*1440,0)/1440)";
        ws.Cell(r, 45).FormulaA1 = $"IF($AB{r}=0,\"\",ROUND(SUMPRODUCT(MAX({Rng("X")}*{cond}))*1440,0)/1440)";
        ws.Cell(r, 29).FormulaA1 = $"IF($AR{r}=\"\",\"\",INT($AR{r}))";
        ws.Cell(r, 30).FormulaA1 = $"IF($AR{r}=\"\",\"\",TIME(INT(ROUND(MOD($AR{r},1)*1440,0)/60),MOD(ROUND(MOD($AR{r},1)*1440,0),60),0))";
        ws.Cell(r, 31).FormulaA1 = $"IF($AS{r}=\"\",\"\",INT($AS{r}))";
        ws.Cell(r, 32).FormulaA1 = $"IF($AS{r}=\"\",\"\",TIME(INT(ROUND(MOD($AS{r},1)*1440,0)/60),MOD(ROUND(MOD($AS{r},1)*1440,0),60),0))";
        var cols = new[] { _layout.LoadSurvey, _layout.LoadProject, _layout.LoadEstimate, _layout.LoadVor, _layout.LoadKac, _layout.LoadOther };
        for (var k = 0; k < cols.Length; k++)
            ws.Cell(r, 33 + k).FormulaA1 = $"SUMPRODUCT({cond}*{Rng(cols[k])})";
        ws.Cell(r, 39).FormulaA1 = $"SUM(AG{r}:AL{r})";
        var extra = new[] { _layout.LoadAnswers, _layout.LoadCancelled, _layout.LoadErrors };
        for (var k = 0; k < extra.Length; k++)
        {
            if (!string.IsNullOrWhiteSpace(extra[k]))
                ws.Cell(r, 40 + k).FormulaA1 = $"SUMPRODUCT({cond}*{Rng(extra[k])})";
        }

        ws.Range(r, 29, r, 29).Style.NumberFormat.Format = XlStyle.DateFormat;
        ws.Range(r, 31, r, 31).Style.NumberFormat.Format = XlStyle.DateFormat;
        ws.Range(r, 30, r, 30).Style.NumberFormat.Format = XlStyle.TimeFormat;
        ws.Range(r, 32, r, 32).Style.NumberFormat.Format = XlStyle.TimeFormat;
        ws.Range(r, 44, r, 45).Style.NumberFormat.Format = XlStyle.DateTimeFormat;
        ws.Range(r, 27, r, 45).Style.Font.FontSize = 11;
    }

    private static void ConverterLoadsTotal(IXLWorksheet ws, int t, int first, int n)
    {
        ws.Cell(t, 32).Value = "Итого:";
        for (var c = 33; c <= 42; c++)
            ws.Cell(t, c).FormulaA1 = n == 0 ? "0" : $"SUM({C(c)}{first}:{C(c)}{t - 1})";
        ws.Range(t, 27, t, 42).Style.Font.Bold = true;
        ws.Range(t, 27, t, 42).Style.Font.FontSize = 11;
    }

    private static void ConverterRemarksHeader(IXLWorksheet ws, int h2, int col)
    {
        ws.Range(h2 + 1, col, h2 + 1, col + 3).Merge().Value = "Замечаний";
        ws.Range(h2 + 1, col + 4, h2 + 1, col + 7).Merge().Value = "Критических";
        ws.Cell(h2 + 2, col - 1).Value = "Этап";
        string[] names = { "Всего", "Устранено", "Устраняется", "Рассматривается" };
        for (var k = 0; k < 8; k++)
            ws.Cell(h2 + 2, col + k).Value = names[k % 4];
        var range = ws.Range(h2 + 1, col - 1, h2 + 2, col + 7);
        range.Style.Font.FontSize = 10;
        range.Style.Font.Bold = true;
        range.Style.Fill.BackgroundColor = XlStyle.ConverterFill;
        range.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
    }

    /// <summary>Строка блока «Конвертер - Замечания»: ссылки на строку листа «З.N».</summary>
    private void ConverterRemarksRow(IXLWorksheet ws, int r, int col, ProjectReport p, RemarksSnapshot? snapshot)
    {
        ws.Cell(r, col - 1).Value = XlStyle.StageCell(p.Number);
        string[] src = { _layout.RemarksAll, _layout.CriticalAll };
        for (var block = 0; block < 2; block++)
        {
            var start = Col(src[block]);
            for (var k = 0; k < 4; k++)
            {
                var cell = ws.Cell(r, col + block * 4 + k);
                if (snapshot is null || snapshot.Row <= 0)
                    cell.Value = 0;
                else
                    cell.FormulaA1 = $"{Z(p)}!{C(start + k)}{snapshot.Row}";
            }
        }

        ws.Range(r, col - 1, r, col + 7).Style.Font.FontSize = 11;
    }

    private static void ConverterRemarksTotal(IXLWorksheet ws, int t, int first, int n, int col)
    {
        ws.Cell(t, col - 1).Value = "Итого:";
        for (var k = 0; k < 8; k++)
            ws.Cell(t, col + k).FormulaA1 = n == 0 ? "0" : $"SUM({C(col + k)}{first}:{C(col + k)}{t - 1})";
        ws.Range(t, col - 1, t, col + 7).Style.Font.Bold = true;
        ws.Range(t, col - 1, t, col + 7).Style.Font.FontSize = 11;
    }

    private static void StatusHeader(IXLWorksheet ws, int h)
    {
        ws.Put(h, 1, h + 2, 2, "Этап").Set(14, bold: true);
        ws.Put(h, 3, h, 14, "Все замечания").Set(11, bold: true);
        ws.Put(h, 15, h, 26, "Критические замечания").Set(11, bold: true);
        string[] names = { "Всего", "Устранено", "Устраняется", "Рассм. ГГЭ" };
        for (var block = 0; block < 2; block++)
        {
            for (var k = 0; k < 4; k++)
            {
                var c = 3 + block * 12 + k * 3;
                var color = XlStyle.Status[k];
                ws.Put(h + 1, c, h + 1, c + 2, names[k]).Set(12, bold: true, color: color);
                ws.Put(h + 2, c, h + 2, c + 1, "Шт.").Set(12, bold: k == 0, color: color);
                var pct = block == 0 ? "%" : k == 0 ? "%¹" : "%²";
                ws.Put(h + 2, c + 2, h + 2, c + 2, pct).Set(12, color: color).Style.Fill.BackgroundColor = XlStyle.PercentFill;
            }
        }

        ws.Row(h).Height = 14.4;
        ws.Row(h + 1).Height = 15.6;
        ws.Row(h + 2).Height = 15.6;
    }

    /// <summary>
    /// Строка таблицы «Все/Критические замечания» (A–Z): значения из блока «Конвертер» (AC–AJ) той же строки,
    /// проценты — от строки <paramref name="basisRow"/> (для изменений за период — от текущего количества).
    /// </summary>
    private static void StatusRow(IXLWorksheet ws, int r, int valueRow, int basisRow, bool totals)
    {
        for (var block = 0; block < 2; block++)
        {
            var totalCol = 3 + block * 12;
            for (var k = 0; k < 4; k++)
            {
                var c = totalCol + k * 3;
                var color = XlStyle.Status[k];
                ws.Put(r, c, r, c + 1, F($"{C(29 + block * 4 + k)}{valueRow}"))
                    .Set(14, bold: totals || k == 0, color: color, align: XLAlignmentHorizontalValues.Right, format: XlStyle.IntFormat);

                // Процент: «Всего» все — 100; «Всего» критические — от всех замечаний (¹); остальные — от «Всего» своего блока (²).
                string pct;
                if (block == 0 && k == 0)
                    pct = valueRow == basisRow ? "100" : $"IF($C${basisRow}=0,0,{C(c)}{r}/$C${basisRow}*100)";
                else if (block == 1 && k == 0)
                    pct = $"IF($C${basisRow}=0,0,{C(c)}{r}/$C${basisRow}*100)";
                else
                    pct = $"IF(${C(totalCol)}${basisRow}=0,0,{C(c)}{r}/${C(totalCol)}${basisRow}*100)";
                if (totals && valueRow == basisRow && !(block == 0 && k == 0))
                    pct = $"ROUND({pct},0)";
                var pc = ws.Put(r, c + 2, r, c + 2, pct == "100" ? 100 : F(pct))
                    .Set(totals ? 11 : 9, bold: totals, color: color, format: totals && valueRow == basisRow ? "0" : XlStyle.PercentFormat);
                pc.Style.Fill.BackgroundColor = XlStyle.PercentFill;
            }

            // Подсветка: «Всего» не равно сумме по статусам.
            var tc = ws.Range(r, totalCol, r, totalCol + 1);
            tc.AddConditionalFormat()
                .WhenIsTrue($"={C(totalCol)}{r}<>{C(totalCol + 3)}{r}+{C(totalCol + 6)}{r}+{C(totalCol + 9)}{r}")
                .Fill.SetBackgroundColor(XlStyle.ErrorFill);
        }

        ws.Row(r).Height = 18;
    }

    private void StaleNotes(IXLWorksheet ws, int row)
    {
        var stale = _projects.Where(p => p.IsStale(_m.ReportTime)).ToList();
        if (stale.Count == 0)
            return;
        var text = string.Join(" ", stale.Select(p =>
            $"* Этап {p.Number}: данные экрана на {Formats.DateTimeText(p.Current!.Timestamp)} (на отчетное время строка не заполнена)."));
        ws.Put(row, 1, row, 26, text).Set(9, italic: true, align: XLAlignmentHorizontalValues.Left, wrap: true);
    }

    private int Deadlines(IXLWorksheet ws, int start)
    {
        var withDeadlines = _projects.Where(p => p.Project.Deadlines.Count > 0).ToList();
        var stages = _s.DeadlineStages
            .Concat(withDeadlines.SelectMany(p => p.Project.Deadlines.Keys))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Where(stage => withDeadlines.Any(p => FindDeadline(p.Project, stage) is not null))
            .Take(12)
            .ToList();
        if (withDeadlines.Count == 0 || stages.Count == 0)
            return start - 3;

        ws.Cell(start, 1).Value = "4. Сроки.";
        ws.Cell(start, 1).Style.Font.Bold = true;
        ws.Row(start + 1).Height = 9;
        var h = start + 2;
        ws.Put(h, 1, h + 1, 2, "Этап").Set(14, bold: true);

        // 24 колонки (C–Z) делятся между стадиями поровну.
        var widths = new int[stages.Count];
        for (var i = 0; i < stages.Count; i++)
            widths[i] = 24 / stages.Count + (i < 24 % stages.Count ? 1 : 0);
        var cols = new int[stages.Count];
        for (int i = 0, c = 3; i < stages.Count; c += widths[i], i++)
            cols[i] = c;

        for (var i = 0; i < stages.Count; i++)
        {
            var half = widths[i] / 2;
            ws.Put(h, cols[i], h, cols[i] + widths[i] - 1, stages[i]).Set(12, bold: true);
            ws.Put(h + 1, cols[i], h + 1, cols[i] + half - 1, "Начало").Set(12, bold: true);
            ws.Put(h + 1, cols[i] + half, h + 1, cols[i] + widths[i] - 1, "Окончание").Set(12, bold: true);
        }

        var today = _m.ReportTime.Date;
        var r = h + 2;
        foreach (var p in withDeadlines)
        {
            ws.Put(r, 1, r, 2, XlStyle.Stage(p.Number)).Set(14, bold: true);
            for (var i = 0; i < stages.Count; i++)
            {
                var values = FindDeadline(p.Project, stages[i]) ?? new List<string>();
                var half = widths[i] / 2;
                var c1 = cols[i];
                var c2 = cols[i] + half;
                var cEnd = cols[i] + widths[i] - 1;
                if (values.Count == 1 && Formats.ParseDate(values[0]) is null)
                {
                    ws.Put(r, c1, r, cEnd, values[0]).Set(14, italic: true);
                    continue;
                }

                var startDate = values.Count > 0 ? Formats.ParseDate(values[0]) : null;
                var endDate = values.Count > 1 ? Formats.ParseDate(values[1]) : null;
                var a = ws.Put(r, c1, r, c2 - 1, (object?)startDate?.Date ?? (values.Count > 0 ? values[0] : null)).Set(14, format: XlStyle.DateFormat);
                var b = ws.Put(r, c2, r, cEnd, (object?)endDate?.Date ?? (values.Count > 1 ? values[1] : null)).Set(14, format: XlStyle.DateFormat);
                if (startDate is not null && endDate is not null && startDate.Value.Date <= today && today <= endDate.Value.Date)
                {
                    a.Style.Fill.BackgroundColor = XlStyle.CurrentStageFill;
                    b.Style.Fill.BackgroundColor = XlStyle.CurrentStageFill;
                }
            }

            ws.Row(r).Height = 18.75;
            r++;
        }

        ws.Range(h, 1, r - 1, 26).Box();
        ws.Row(h).Height = 18.75;
        ws.Row(h + 1).Height = 18.75;
        return r - 1;
    }

    private static List<string>? FindDeadline(ProjectConfig project, string stage) =>
        project.Deadlines.FirstOrDefault(kv => string.Equals(kv.Key.Trim(), stage.Trim(), StringComparison.OrdinalIgnoreCase)).Value;

    private static void Margins(IXLPageSetup ps)
    {
        ps.Margins.Left = 0.25;
        ps.Margins.Right = 0.25;
        ps.Margins.Top = 0.75;
        ps.Margins.Bottom = 0.75;
        ps.Margins.Header = 0.3;
        ps.Margins.Footer = 0.3;
    }

    // ================================================================== «Отчет» (данные по разделам)

    private void DataSheet(IXLWorksheet ws)
    {
        ws.Style.Font.FontName = XlStyle.Font;
        ws.Style.Font.FontSize = 9;
        ws.Column(1).Width = 12;
        for (var c = 2; c <= 15; c++)
            ws.Column(c).Width = 8;

        ws.Put(1, 1, 15, _s.Title).Set(14, bold: true);
        ws.Put(2, 1, 15, _s.Subtitle).Set(9);
        ws.Put(3, 1, 15, _m.Caption).Set(9);
        ws.Put(4, 1, 15, $"по состоянию на {Formats.DateTimeText(_m.ReportTime)}").Set(9);

        string[] top = { "Всего", "Устранено", "%", "Устраняется", "%", "Рассм. ГГЭ", "%" };
        string[] allCols = { _layout.RemarksAll, _layout.RemarksSurvey, _layout.RemarksTech, _layout.RemarksEstimate };
        string[] critCols = { _layout.CriticalAll, _layout.CriticalSurvey, _layout.CriticalTech, _layout.CriticalEstimate };
        var r = 6;
        foreach (var category in RemarkCategories.Values)
        {
            if (category == RemarkCategory.All)
            {
                ws.Put(r, 2, r, 8, "Все замечания").Set(9, bold: true);
                ws.Put(r, 9, r, 15, "Критические замечания").Set(9, bold: true);
            }
            else
            {
                ws.Put(r, 1, r, 15, "в т.ч. " + RemarkCategories.Title(category)).Set(9, bold: true, align: XLAlignmentHorizontalValues.Left)
                    .Style.Fill.BackgroundColor = XLColor.FromHtml("#D9D9D9");
            }

            r++;
            for (var i = 0; i < 7; i++)
            {
                ws.Cell(r, 2 + i).Value = top[i];
                ws.Cell(r, 9 + i).Value = top[i];
            }

            ws.Range(r, 1, r, 15).Set(9, bold: true);
            r++;
            var first = r;
            foreach (var p in _projects)
            {
                ws.Cell(r, 1).Value = p.Name;
                ws.Cell(r, 1).Style.Font.Bold = true;
                ws.Cell(r, 1).Style.Fill.BackgroundColor = XLColor.FromHtml("#D9D9D9");
                DataRow(ws, r, p, Col(allCols[(int)category]), Col(critCols[(int)category]));
                _dataRows[((int)category, p)] = r;
                r++;
            }

            ws.Cell(r, 1).Value = "Всего";
            for (var block = 0; block < 2; block++)
            {
                var b = 2 + block * 7;
                foreach (var k in new[] { 0, 1, 3, 5 })
                    ws.Cell(r, b + k).FormulaA1 = _projects.Count == 0 ? "0" : $"SUM({C(b + k)}{first}:{C(b + k)}{r - 1})";
                Percents(ws, r, b);
            }

            ws.Range(r, 1, r, 15).Style.Font.Bold = true;
            ws.Range(first, 2, r, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
            r += 1;
        }

        var ps = ws.PageSetup;
        ps.PaperSize = XLPaperSize.A4Paper;
        ps.FitToPages(1, 0);
        Margins(ps);
    }

    private void DataRow(IXLWorksheet ws, int r, ProjectReport p, int allCol, int critCol)
    {
        var row = p.Current?.Row ?? 0;
        int[] src = { allCol, critCol };
        for (var block = 0; block < 2; block++)
        {
            var b = 2 + block * 7;
            int[] offsets = { 0, 1, 3, 5 };
            for (var k = 0; k < 4; k++)
            {
                var cell = ws.Cell(r, b + offsets[k]);
                if (row <= 0)
                    cell.Value = 0;
                else
                    cell.FormulaA1 = $"{Z(p)}!{C(src[block] + k)}{row}";
            }

            Percents(ws, r, b);
        }
    }

    private static void Percents(IXLWorksheet ws, int r, int b)
    {
        foreach (var k in new[] { 2, 4, 6 })
        {
            ws.Cell(r, b + k).FormulaA1 = $"IF({C(b)}{r}=0,0,{C(b + k - 1)}{r}/{C(b)}{r}*100)";
            ws.Cell(r, b + k).Style.NumberFormat.Format = XlStyle.PercentFormat;
        }
    }

    // ================================================================== «Отчет_дет»

    private void Detailed(IXLWorksheet ws)
    {
        ws.Style.Font.FontName = XlStyle.Font;
        ws.Style.Font.FontSize = 11;
        double[] widths = { 8.66, 8, 8.66, 5.66, 9, 5.66, 8.66, 5.66, 8, 8.11, 5.66, 8.66, 5.66, 9.66, 5.66 };
        for (var c = 0; c < widths.Length; c++)
            ws.Column(c + 1).Width = widths[c];

        ws.Put(1, 1, 15, _s.DetailedTitle).Set(14, bold: true);
        ws.Put(2, 1, 15, _s.Subtitle).Set(14);
        ws.Put(3, 1, 15, _m.Caption).Set(14);
        ws.Put(4, 1, 15, F($"{XlStyle.Q(_dailySheet)}!A4")).Set(12, format: XlStyle.DateTimeFormat);

        var r = 6;
        var tableTop = r;
        foreach (var category in RemarkCategories.Values)
        {
            int h;
            if (category == RemarkCategory.All)
            {
                ws.Put(r, 1, r + 2, 1, "Этап").Set(11, bold: true);
                ws.Put(r, 2, r, 8, "Все замечания").Set(11, bold: true);
                ws.Put(r, 9, r, 15, "Критические замечания").Set(11, bold: true);
                h = r + 1;
                r += 3;
            }
            else
            {
                ws.Put(r, 1, r, 15, "в т.ч. " + RemarkCategories.Title(category)).Set(11, bold: true);
                ws.Put(r + 1, 1, r + 2, 1, "Этап").Set(11, bold: true);
                h = r + 1;
                r += 3;
            }

            for (var block = 0; block < 2; block++)
            {
                var b = 2 + block * 7;
                ws.Put(h, b, h + 1, b, "Всего").Set(11, bold: true, color: XlStyle.Total);
                string[] names = { "Устранено", "Устраняется", "Рассм. ГГЭ" };
                for (var k = 0; k < 3; k++)
                {
                    var color = XlStyle.Status[k + 1];
                    ws.Put(h, b + 1 + k * 2, h, b + 2 + k * 2, names[k]).Set(11, bold: true, color: color);
                    ws.Put(h + 1, b + 1 + k * 2, h + 1, b + 1 + k * 2, "Шт.").Set(11, bold: true, color: color);
                    ws.Put(h + 1, b + 2 + k * 2, h + 1, b + 2 + k * 2, "%").Set(11, color: color).Style.Fill.BackgroundColor = XlStyle.PercentFill;
                }
            }

            var first = r;
            foreach (var p in _projects)
            {
                ws.Put(r, 1, r, 1, XlStyle.Stage(p.Number)).Set(11, bold: true);
                DetailedRow(ws, r, $"{SummaryWorkbook.DataSheet}!", _dataRows[((int)category, p)], false);
                r++;
            }

            ws.Put(r, 1, r, 1, "Всего").Set(11, bold: true);
            DetailedRow(ws, r, "", 0, true, first);
            r++;
        }

        ws.Range(tableTop, 1, r - 1, 15).Box();
        var ps = ws.PageSetup;
        ps.PrintAreas.Clear();
        ps.PrintAreas.Add(1, 1, r - 1, 15);
        ps.PaperSize = XLPaperSize.A4Paper;
        ps.PageOrientation = XLPageOrientation.Portrait;
        ps.FitToPages(1, 0);
        ps.CenterHorizontally = true;
        Margins(ps);
    }

    /// <summary>Строка детального отчета: ссылки на лист «Отчет» или (для итога) суммы по строкам выше.</summary>
    private static void DetailedRow(IXLWorksheet ws, int r, string sheetPrefix, int sourceRow, bool totals, int first = 0)
    {
        for (var block = 0; block < 2; block++)
        {
            var b = 2 + block * 7;
            int[] offsets = { 0, 1, 3, 5 };
            for (var k = 0; k < 4; k++)
            {
                var c = b + offsets[k];
                var formula = totals ? $"SUM({C(c)}{first}:{C(c)}{r - 1})" : $"{sheetPrefix}{C(c)}{sourceRow}";
                ws.Put(r, c, r, c, F(formula)).Set(11, bold: totals || k == 0, color: XlStyle.Status[k],
                    align: XLAlignmentHorizontalValues.Right, format: XlStyle.IntFormat);
            }

            for (var k = 0; k < 3; k++)
            {
                var c = b + 2 + k * 2;
                ws.Put(r, c, r, c, F($"IF({C(b)}{r}=0,0,{C(c - 1)}{r}/{C(b)}{r}*100)"))
                    .Set(totals ? 10 : 9, bold: totals, color: XlStyle.Status[k + 1], format: XlStyle.PercentFormat)
                    .Style.Fill.BackgroundColor = XlStyle.PercentFill;
            }

            ws.Range(r, b, r, b).AddConditionalFormat()
                .WhenIsTrue($"={C(b)}{r}<>{C(b + 1)}{r}+{C(b + 3)}{r}+{C(b + 5)}{r}")
                .Fill.SetBackgroundColor(XlStyle.ErrorFill);
        }

        if (totals)
            ws.Range(r, 1, r, 15).Style.Border.TopBorder = XLBorderStyleValues.Medium;
    }

    // ================================================================== «Диаграммы»

    private const int PageRows = 52;

    private int ChartsSheet(IXLWorksheet ws, SvodLayout svod, out List<ChartPlacement> placements)
    {
        placements = new List<ChartPlacement>();
        double[] widths = { 2, 10.66, 9.1, 10.1, 9.55, 9.1, 9.1, 11.33, 9.1, 9.1, 9.1, 9.1, 9.1, 9.1, 16.44, 9.1, 9.1, 16.44, 2.55, 11.66, 9.1, 9.1, 9.1, 9 };
        for (var c = 0; c < widths.Length; c++)
            ws.Column(c + 1).Width = widths[c];

        var items = new List<(ProjectReport Report, SvodBlock Block, bool Summary)>();
        if (_m.Summary is not null && _m.Config.Charts.SummaryChart && _projects.Count > 1)
            items.Add((_m.Summary, SvodLayout.Overall, true));
        for (var i = 0; i < _projects.Count; i++)
        {
            var p = _projects[i];
            if (p.HasRemarksSeries || p.HasLoadsSeries)
                items.Add((p, SvodLayout.Project(i), false));
        }

        var cs = _m.Config.Charts;
        var page = 0;
        foreach (var (report, block, summary) in items)
        {
            if (report.Daily.Count == 0)
                continue;
            var top = 1 + page * PageRows;
            var name = summary
                ? Formats.ApplyTemplate(cs.SummaryName, values: new Dictionary<string, string> { ["projects"] = ShortList() })
                : report.Name;
            var names = new Dictionary<string, string> { ["name"] = name, ["number"] = report.Number };
            var r1 = svod.Row(report.Daily[0].Date);
            var r2 = svod.Row(report.Daily[^1].Date);
            var reportRow = Math.Clamp(svod.Row(_m.ReportTime.Date), r1, r2);
            var cats = SvodLayout.Ref(SvodLayout.LabelColumn, r1, r2);

            ws.Cell(top, 2).Value = "По состоянию на";
            ws.Cell(top, 4).FormulaA1 = $"{XlStyle.Q(_dailySheet)}!$AE$4";
            ws.Cell(top, 4).Style.NumberFormat.Format = XlStyle.DateFormat;
            ws.Cell(top, 4).Style.Font.Bold = true;

            // Устранение замечаний: столбцы с накоплением + линии + линии тренда.
            var remarksSeries = new List<ChartSeriesRef>
            {
                new("Устраняется", SvodLayout.Ref(block.Fixing, 1), SvodLayout.Ref(block.Fixing, r1, r2), "FF9999", false),
                new("Рассматривается", SvodLayout.Ref(block.Review, 1), SvodLayout.Ref(block.Review, r1, r2), "0070C0", false),
                new("Устранено за сутки", SvodLayout.Ref(block.FixedPerDay, 1), SvodLayout.Ref(block.FixedPerDay, r1, r2), "92D050", false),
                new("Всего", SvodLayout.Ref(block.Total, 1), SvodLayout.Ref(block.Total, r1, r2), "806000", true, cs.TrendLine ? "4472C4" : null),
                new("Устранено", SvodLayout.Ref(block.Fixed, 1), SvodLayout.Ref(block.Fixed, r1, r2), "00B050", true, cs.TrendLine ? "ED7D31" : null, LabelsBelow: true),
            };
            placements.Add(new ChartPlacement(ChartKind.Remarks, Formats.ApplyTemplate(cs.RemarksTitle, values: names),
                top, top + 29, 0, 24, cats, remarksSeries));

            // Строка между диаграммами: Всего загружено | заголовок | окончание экспертизы | осталось дней.
            var info = top + 30;
            ws.Cell(info, 7).Value = "Всего:";
            ws.Cell(info, 8).FormulaA1 = $"SUM({SvodLayout.Ref(block.LoadTotal, r1, reportRow)})";
            ws.Cell(info, 8).Style.NumberFormat.Format = XlStyle.IntFormat;
            ws.Range(info, 9, info, 15).Merge().Value = Formats.ApplyTemplate(cs.LoadsTitle, values: names);
            ws.Range(info, 9, info, 15).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            if (report.ExpertiseEnd is { } end && !summary)
            {
                ws.Cell(info, 16).Value = "Дата окончания экспертизы:";
                ws.Cell(info, 20).Value = end;
                ws.Cell(info, 20).Style.NumberFormat.Format = XlStyle.DateFormat;
                ws.Cell(info, 22).Value = "Осталось дней:";
                ws.Cell(info, 24).FormulaA1 = $"T{info}-{XlStyle.Q(_dailySheet)}!$AE$4";
                ws.Cell(info, 24).Style.NumberFormat.Format = "0";
                ws.Range(info, 16, info, 20).Style.Font.FontColor = XLColor.FromHtml("#0070C0");
            }

            ws.Range(info, 1, info, 24).Style.Font.Bold = true;
            ws.Row(info).Height = 18;

            var loadsSeries = new List<ChartSeriesRef>
            {
                new("Загружено документов", SvodLayout.Ref(block.LoadTotal, 1), SvodLayout.Ref(block.LoadTotal, r1, r2), "5B9BD5", false),
            };
            placements.Add(new ChartPlacement(ChartKind.Loads, "", top + 30, top + 50, 0, 24, cats, loadsSeries));
            page++;
        }

        var lastRow = Math.Max(1, page * PageRows);
        var ps = ws.PageSetup;
        ps.PrintAreas.Clear();
        ps.PrintAreas.Add(1, 1, lastRow, 24);
        ps.PaperSize = string.Equals(cs.PageSize?.Trim(), "A4", StringComparison.OrdinalIgnoreCase) ? XLPaperSize.A4Paper : XLPaperSize.A3Paper;
        ps.PageOrientation = XLPageOrientation.Landscape;
        // Масштаб с запасом: ширина колонок в разных программах считается немного по-разному.
        ps.Scale = ps.PaperSize == XLPaperSize.A4Paper ? 60 : 85;
        ps.CenterHorizontally = true;
        Margins(ps);
        for (var i = 1; i < page; i++)
            ps.AddHorizontalPageBreak(i * PageRows);
        ws.SheetView.ZoomScale = 55;
        return page;
    }

    private string ShortList() => Formats.CompressNumbers(_m.Projects.Select(p => p.Number)).Replace(", ", ",");
}
