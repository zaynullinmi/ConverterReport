using ClosedXML.Excel;
using ConverterReport.Core.Model;
using ConverterReport.Core.Reporting;

namespace ConverterReport.Core.Export;

/// <summary>
/// Выгрузка рассчитанных данных в Excel — для сверки и дальнейшего анализа:
/// «Сводка» (замечания по этапам), «Загрузки за период» (каждая загрузка), «Динамика» (ряды диаграмм).
/// </summary>
public static class ExcelExporter
{
    public static void Export(ReportModel model, string path)
    {
        using var wb = new XLWorkbook();
        Summary(wb.AddWorksheet("Сводка"), model);
        Loads(wb.AddWorksheet("Загрузки за период"), model);
        Dynamics(wb.AddWorksheet("Динамика"), model);
        wb.SaveAs(path);
    }

    private static void Summary(IXLWorksheet ws, ReportModel model)
    {
        ws.Cell(1, 1).Value = $"Состояние на {Formats.DateTimeText(model.ReportTime)}, изменения с {Formats.DateTimeText(model.PeriodStart)}";
        ws.Cell(1, 1).Style.Font.Bold = true;

        string[] header =
        {
            "Этап", "Раздел", "Данные на",
            "Всего", "Устранено", "Устраняется", "Рассматривается",
            "Крит. всего", "Крит. устранено", "Крит. устраняется", "Крит. рассматривается",
            "Δ Всего", "Δ Устранено", "Δ Устраняется", "Δ Рассматривается",
            "Δ Крит. всего", "Δ Крит. устранено", "Δ Крит. устраняется", "Δ Крит. рассматривается",
        };
        WriteHeader(ws, 3, header);

        var row = 4;
        foreach (var p in model.Projects)
        {
            if (p.Current is null)
            {
                ws.Cell(row, 1).Value = p.Number;
                ws.Cell(row, 2).Value = "нет данных";
                row++;
                continue;
            }

            foreach (var category in RemarkCategories.Values)
            {
                ws.Cell(row, 1).Value = p.Number;
                ws.Cell(row, 2).Value = RemarkCategories.Title(category);
                ws.Cell(row, 3).Value = p.Current.Timestamp;
                ws.Cell(row, 3).Style.DateFormat.Format = "dd.MM.yyyy HH:mm";
                Counts(ws, row, 4, p.Current.Get(category, false));
                Counts(ws, row, 8, p.Current.Get(category, true));
                if (p.Delta is { } delta)
                {
                    Counts(ws, row, 12, delta.Get(category, false));
                    Counts(ws, row, 16, delta.Get(category, true));
                }

                row++;
            }
        }

        Finish(ws, 3, row - 1, header.Length);
    }

    private static void Loads(IXLWorksheet ws, ReportModel model)
    {
        string[] header = { "Этап", "№", "Строка экрана", "Начало", "Окончание", "ИИ", "ПД", "СМ", "ВОР", "КАЦ", "Иные", "Всего", "Ответы", "Аннулировано", "Ошибки", "Примечание" };
        WriteHeader(ws, 1, header);
        var row = 2;
        foreach (var p in model.Projects)
        {
            foreach (var l in p.PeriodLoads)
            {
                ws.Cell(row, 1).Value = p.Number;
                ws.Cell(row, 2).Value = l.Number ?? "";
                ws.Cell(row, 3).Value = l.Row;
                if (l.Start is { } start)
                    ws.Cell(row, 4).Value = start;
                ws.Cell(row, 5).Value = l.End;
                ws.Range(row, 4, row, 5).Style.DateFormat.Format = "dd.MM.yyyy HH:mm";
                var c = l.Counts;
                int[] values = { c.Survey, c.Project, c.Estimate, c.Vor, c.Kac, c.Other, c.Total, c.Answers, c.Cancelled, c.Errors };
                for (var i = 0; i < values.Length; i++)
                    ws.Cell(row, 6 + i).Value = values[i];
                ws.Cell(row, 16).Value = l.Note ?? "";
                row++;
            }
        }

        Finish(ws, 1, row - 1, header.Length);
    }

    private static void Dynamics(IXLWorksheet ws, ReportModel model)
    {
        string[] header = { "Этап", "Дата", "Всего", "Устранено", "Устраняется", "Рассматривается", "Устранено за сутки", "Загружено документов" };
        WriteHeader(ws, 1, header);
        var row = 2;
        var items = model.Summary is null ? model.Projects : model.Projects.Append(model.Summary);
        foreach (var p in items)
        {
            foreach (var d in p.Daily.Where(d => d.Date <= model.ReportTime.Date))
            {
                ws.Cell(row, 1).Value = p == model.Summary ? "Всего" : p.Number;
                ws.Cell(row, 2).Value = d.Date;
                ws.Cell(row, 2).Style.DateFormat.Format = "dd.MM.yyyy";
                if (d.Remarks is { } r)
                {
                    ws.Cell(row, 3).Value = r.Total;
                    ws.Cell(row, 4).Value = r.Fixed;
                    ws.Cell(row, 5).Value = r.Fixing;
                    ws.Cell(row, 6).Value = r.Review;
                }

                if (d.FixedPerDay is { } f)
                    ws.Cell(row, 7).Value = f;
                ws.Cell(row, 8).Value = d.Loaded;
                row++;
            }
        }

        Finish(ws, 1, row - 1, header.Length);
    }

    private static void Counts(IXLWorksheet ws, int row, int column, StatusCounts c)
    {
        ws.Cell(row, column).Value = c.Total;
        ws.Cell(row, column + 1).Value = c.Fixed;
        ws.Cell(row, column + 2).Value = c.Fixing;
        ws.Cell(row, column + 3).Value = c.Review;
    }

    private static void WriteHeader(IXLWorksheet ws, int row, string[] header)
    {
        for (var i = 0; i < header.Length; i++)
            ws.Cell(row, i + 1).Value = header[i];
        var range = ws.Range(row, 1, row, header.Length);
        range.Style.Font.Bold = true;
        range.Style.Fill.BackgroundColor = XLColor.FromHtml("#D9E1F2");
        range.Style.Alignment.WrapText = true;
    }

    private static void Finish(IXLWorksheet ws, int headerRow, int lastRow, int columns)
    {
        if (lastRow > headerRow)
            ws.Range(headerRow, 1, lastRow, columns).SetAutoFilter();
        ws.SheetView.FreezeRows(headerRow);
        for (var c = 1; c <= columns; c++)
            ws.Column(c).Width = c <= 2 ? 12 : 14;
    }
}
