using ClosedXML.Excel;

namespace ConverterReport.Tests;

/// <summary>Синтетический файл «Экран» с той же раскладкой, что и рабочий шаблон.</summary>
internal sealed class TestScreen
{
    private readonly XLWorkbook _wb = new();
    private readonly IXLWorksheet _remarks;
    private readonly IXLWorksheet _documents;
    private int _remarksRow = 3;
    private int _documentsRow = 3;

    public TestScreen()
    {
        _remarks = _wb.AddWorksheet("Замечания");
        _remarks.Cell("A2").Value = "№";
        _remarks.Cell("B2").Value = "Дата";
        _remarks.Cell("D2").Value = "Время";
        _remarks.Cell("E1").Value = "Замечаний";

        _documents = _wb.AddWorksheet("Документы");
        _documents.Cell("A1").Value = "№";
        _documents.Cell("H1").Value = "Начало загрузки";
        _documents.Cell("J1").Value = "Окончание загрузки";
    }

    /// <summary>
    /// Строка «Замечаний»: all/crit — Всего, Устранено, Устраняется, Рассматривается. Разделы делятся как 10% / 80% / остаток.
    /// </summary>
    public TestScreen Remarks(DateTime date, TimeSpan time, (int Total, int Fixed, int Fixing, int Review) all, (int Total, int Fixed, int Fixing, int Review)? crit = null)
    {
        var r = _remarksRow++;
        _remarks.Cell(r, 1).Value = r - 2;
        _remarks.Cell(r, 2).Value = date;
        _remarks.Cell(r, 4).Value = time;
        WriteBlock(r, "E", all);
        var survey = Part(all, 0.1);
        var tech = Part(all, 0.8);
        var estimate = (all.Total - survey.Total - tech.Total, all.Fixed - survey.Fixed - tech.Fixed, all.Fixing - survey.Fixing - tech.Fixing, all.Review - survey.Review - tech.Review);
        WriteBlock(r, "I", survey);
        WriteBlock(r, "M", tech);
        WriteBlock(r, "Q", estimate);

        var c = crit ?? (all.Total / 10, all.Fixed / 10, all.Fixing / 10, all.Total / 10 - all.Fixed / 10 - all.Fixing / 10);
        WriteBlock(r, "U", c);
        WriteBlock(r, "Y", (0, 0, 0, 0));
        WriteBlock(r, "AC", c);
        WriteBlock(r, "AG", (0, 0, 0, 0));
        return this;
    }

    /// <summary>Пустая строка-заготовка (дата и время есть, данных нет) — должна пропускаться.</summary>
    public TestScreen EmptyRemarksRow(DateTime date, TimeSpan time)
    {
        var r = _remarksRow++;
        _remarks.Cell(r, 1).Value = r - 2;
        _remarks.Cell(r, 2).Value = date;
        _remarks.Cell(r, 4).Value = time;
        return this;
    }

    /// <summary>Загрузка. Дата/время можно передавать строками («10.10.2026», «18:56») — как встречается в экранах.</summary>
    public TestScreen Load(object startDate, object startTime, object endDate, object endTime, params int[] counts)
    {
        var r = _documentsRow++;
        _documents.Cell(r, 1).Value = r - 2;
        Set(_documents.Cell(r, 8), startDate);
        Set(_documents.Cell(r, 9), startTime);
        Set(_documents.Cell(r, 10), endDate);
        Set(_documents.Cell(r, 11), endTime);
        for (var i = 0; i < counts.Length; i++)
            _documents.Cell(r, 12 + (i < 6 ? i : i + 1)).Value = counts[i]; // L..Q, затем S, T, U
        return this;
    }

    public TestScreen TotalsRow()
    {
        var r = _documentsRow++;
        _documents.Cell(r, 11).Value = "Итого:";
        _documents.Cell(r, 12).FormulaA1 = $"SUM(L3:L{r - 1})";
        return this;
    }

    public MemoryStream ToStream()
    {
        var ms = new MemoryStream();
        _wb.SaveAs(ms);
        ms.Position = 0;
        return ms;
    }

    public string SaveTo(string folder, string name)
    {
        var path = Path.Combine(folder, name);
        _wb.SaveAs(path);
        return path;
    }

    private void WriteBlock(int row, string firstColumn, (int Total, int Fixed, int Fixing, int Review) v)
    {
        var col = XLHelper.GetColumnNumberFromLetter(firstColumn);
        _remarks.Cell(row, col).Value = v.Total;
        _remarks.Cell(row, col + 1).Value = v.Fixed;
        _remarks.Cell(row, col + 2).Value = v.Fixing;
        _remarks.Cell(row, col + 3).Value = v.Review;
    }

    private static (int Total, int Fixed, int Fixing, int Review) Part((int Total, int Fixed, int Fixing, int Review) v, double share) =>
        ((int)(v.Total * share), (int)(v.Fixed * share), (int)(v.Fixing * share), (int)(v.Total * share) - (int)(v.Fixed * share) - (int)(v.Fixing * share));

    private static void Set(IXLCell cell, object value)
    {
        switch (value)
        {
            case DateTime dt:
                cell.Value = dt;
                break;
            case TimeSpan ts:
                cell.Value = ts;
                break;
            case string s:
                cell.Value = s;
                break;
            case null:
                break;
            default:
                cell.Value = Convert.ToDouble(value);
                break;
        }
    }
}
