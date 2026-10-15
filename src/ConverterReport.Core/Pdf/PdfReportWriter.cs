using ConverterReport.Core.Reporting;
using MigraDoc.Rendering;
using PdfSharp.Pdf;

namespace ConverterReport.Core.Pdf;

public sealed record PdfWriteResult(string Path, int Pages, int ChartPages, int RemovedBlankPages);

/// <summary>
/// Один PDF: суточный отчет на дату → детальный отчет → диаграммы. Пустые страницы не выводятся.
/// </summary>
public static class PdfReportWriter
{
    public static PdfWriteResult Write(ReportModel model, string path)
    {
        using var document = Render(model, out var chartPages, out var removed);
        var pages = document.PageCount;
        if (pages == 0)
            throw new InvalidOperationException("Нет данных для формирования отчета: все разделы пустые.");
        document.Save(path);
        return new PdfWriteResult(path, pages, chartPages, removed);
    }

    public static PdfDocument Render(ReportModel model, out int chartPages, out int removedBlankPages)
    {
        SystemFontResolver.Register();
        var settings = model.Config.Report;

        var builder = new ReportDocumentBuilder(model);
        var doc = builder.Build(settings.IncludeDaily, settings.IncludeDetailed);

        PdfDocument pdf;
        if (builder.HasContent)
        {
            var renderer = new PdfDocumentRenderer { Document = doc };
            renderer.RenderDocument();
            pdf = renderer.PdfDocument;
        }
        else
        {
            pdf = new PdfDocument();
        }

        pdf.Info.Title = doc.Info.Title;
        pdf.Info.Subject = doc.Info.Subject;
        pdf.Info.Creator = "ConverterReport";

        chartPages = settings.IncludeCharts ? new ChartPagesWriter(model).AddPages(pdf) : 0;
        removedBlankPages = pdf.PageCount > 1 ? BlankPageRemover.RemoveBlankPages(pdf) : 0;
        return pdf;
    }
}
