using ConverterReport.Core.Reporting;
using MigraDoc.Rendering;
using PdfSharp.Pdf;

namespace ConverterReport.Core.Pdf;

public sealed record PdfWriteResult(string Path, int Pages, int RemovedBlankPages);

/// <summary>
/// Встроенный генератор PDF (без Excel): суточный отчет, детальный отчет и диаграммы — отдельными файлами.
/// Используется, если на компьютере нет Excel или так указано в настройках. Пустые страницы не выводятся.
/// </summary>
public static class PdfReportWriter
{
    public static PdfWriteResult WriteDaily(ReportModel model, string path) => Save(Render(model, daily: true, detailed: false, charts: false), path);

    public static PdfWriteResult WriteDetailed(ReportModel model, string path) => Save(Render(model, daily: false, detailed: true, charts: false), path);

    public static PdfWriteResult WriteCharts(ReportModel model, string path) => Save(Render(model, daily: false, detailed: false, charts: true), path);

    private static PdfWriteResult Save((PdfDocument Document, int Removed) rendered, string path)
    {
        using var document = rendered.Document;
        var pages = document.PageCount;
        if (pages == 0)
            throw new InvalidOperationException("Нет данных для формирования PDF: раздел пустой.");
        document.Save(path);
        return new PdfWriteResult(path, pages, rendered.Removed);
    }

    public static (PdfDocument Document, int Removed) Render(ReportModel model, bool daily, bool detailed, bool charts)
    {
        SystemFontResolver.Register();

        var builder = new ReportDocumentBuilder(model);
        var doc = builder.Build(daily, detailed);

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

        if (charts)
            new ChartPagesWriter(model).AddPages(pdf);
        var removed = pdf.PageCount > 1 ? BlankPageRemover.RemoveBlankPages(pdf) : 0;
        return (pdf, removed);
    }
}
