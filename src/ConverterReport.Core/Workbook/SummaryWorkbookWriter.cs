using ConverterReport.Core.Reporting;

namespace ConverterReport.Core.Workbook;

/// <summary>Сводная книга Excel: листы и формулы (ClosedXML) + диаграммы (Open XML).</summary>
public static class SummaryWorkbookWriter
{
    public static SummaryWorkbook Write(ReportModel model, string path)
    {
        var result = new SummaryWorkbookBuilder(model).Build(path);
        ChartInjector.Inject(path, SummaryWorkbook.ChartsSheet, result.Charts);
        return result;
    }
}
