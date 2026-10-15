using System.Runtime.InteropServices;
using ConverterReport.Core.Pdf;
using PdfSharp.Pdf.IO;

namespace ConverterReport.Core.Export;

/// <summary>Лист книги и PDF, в который он сохраняется.</summary>
public sealed record SheetPdf(string Sheet, string PdfPath);

/// <summary>
/// Сохранение листов сводной книги в PDF средствами самого Excel (как «Сохранить как PDF» вручную):
/// формулы пересчитываются, диаграммы и оформление — один в один. Работает на Windows с установленным Excel.
/// </summary>
public static class ExcelPdfExporter
{
    private const int XlTypePdf = 0;
    private const int XlQualityStandard = 0;

    public static bool IsAvailable()
    {
        if (!OperatingSystem.IsWindows())
            return false;
        try
        {
            return Type.GetTypeFromProgID("Excel.Application") is not null;
        }
        catch
        {
            return false;
        }
    }

    public static void Export(string workbookPath, IReadOnlyList<SheetPdf> sheets)
    {
        if (!OperatingSystem.IsWindows())
            throw new PlatformNotSupportedException("Сохранение через Excel доступно только в Windows.");

        var type = Type.GetTypeFromProgID("Excel.Application")
                   ?? throw new InvalidOperationException("Microsoft Excel не установлен.");
        dynamic? app = null;
        dynamic? workbook = null;
        try
        {
            app = Activator.CreateInstance(type)!;
            app.Visible = false;
            app.DisplayAlerts = false;
            app.ScreenUpdating = false;
            app.AskToUpdateLinks = false;

            // UpdateLinks = 0, ReadOnly = true
            workbook = app.Workbooks.Open(Path.GetFullPath(workbookPath), 0, true);
            app.CalculateFull();

            foreach (var job in sheets)
            {
                dynamic sheet = workbook.Worksheets[job.Sheet];
                // Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas
                sheet.ExportAsFixedFormat(XlTypePdf, Path.GetFullPath(job.PdfPath), XlQualityStandard, true, false);
                Marshal.FinalReleaseComObject(sheet);
                RemoveBlankPages(job.PdfPath);
            }
        }
        finally
        {
            if (workbook is not null)
            {
                try
                {
                    workbook.Close(false);
                }
                catch
                {
                    // Книга уже закрыта.
                }

                Marshal.FinalReleaseComObject(workbook);
            }

            if (app is not null)
            {
                try
                {
                    app.Quit();
                }
                catch
                {
                    // Excel уже завершен.
                }

                Marshal.FinalReleaseComObject(app);
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }

    /// <summary>Убирает из готового PDF пустые страницы. Возвращает число удаленных.</summary>
    public static int RemoveBlankPages(string pdfPath)
    {
        int removed;
        using (var doc = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Modify))
        {
            removed = BlankPageRemover.RemoveBlankPages(doc);
            if (removed > 0)
                doc.Save(pdfPath);
        }

        return removed;
    }
}
