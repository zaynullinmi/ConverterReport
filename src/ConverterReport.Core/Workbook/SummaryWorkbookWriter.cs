using ClosedXML.Excel;
using ConverterReport.Core.Reporting;

namespace ConverterReport.Core.Workbook;

/// <summary>Сводная книга Excel: листы и формулы (ClosedXML) + диаграммы (Open XML).</summary>
public static class SummaryWorkbookWriter
{
    /// <summary>Новая книга на один отчет.</summary>
    public static SummaryWorkbook Write(ReportModel model, string path)
    {
        var result = new SummaryWorkbookBuilder(model).Build(path);
        ChartInjector.Inject(path, SummaryWorkbook.ChartsSheet, result.Charts);
        return result;
    }

    /// <summary>
    /// Дописывает лист дня в общую книгу (если книги нет — создает ее). Книга сохраняется через временный файл,
    /// прежняя версия остается рядом как «.bak».
    /// </summary>
    /// <exception cref="MasterWorkbookException">Книга занята другой программой, повреждена или создана не этой программой.</exception>
    public static SummaryWorkbook Append(ReportModel model, string masterPath, int visibleDailySheets, LegacySource? legacy = null)
    {
        var exists = File.Exists(masterPath);
        if (exists)
            EnsureNotLocked(masterPath);

        var folder = Path.GetDirectoryName(masterPath)!;
        Directory.CreateDirectory(folder);
        var temp = Path.Combine(folder, Path.GetFileNameWithoutExtension(masterPath) + ".tmp.xlsx");
        try
        {
            SummaryWorkbook result;
            using (var wb = exists ? Load(masterPath) : SummaryWorkbookBuilder.NewWorkbook())
                result = new SummaryWorkbookBuilder(model).AppendTo(wb, masterPath, temp, visibleDailySheets,
                    legacy is null ? null : (book, built) => ImportLegacy(book, built, legacy, model.ReportTime.Date));
            ChartInjector.Inject(temp, SummaryWorkbook.ChartsSheet, result.Charts);

            try
            {
                if (exists)
                    File.Replace(temp, masterPath, masterPath + ".bak", ignoreMetadataErrors: true);
                else
                    File.Move(temp, masterPath);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                throw Busy(masterPath, ex);
            }

            return result;
        }
        finally
        {
            try
            {
                File.Delete(temp);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Временный файл не критичен.
            }
        }
    }

    /// <summary>Листы прошлых дней из старой книги, которых в общей книге еще нет.</summary>
    private static void ImportLegacy(XLWorkbook master, SummaryWorkbook built, LegacySource legacy, DateTime reportDate)
    {
        try
        {
            var existing = master.Worksheets.Select(w => w.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var missing = LegacyWorkbook.DailySheetNames(legacy.Path)
                .Where(n => !existing.Contains(n) && SummaryWorkbookBuilder.DailyDate(n, reportDate) < reportDate)
                .ToList();
            if (missing.Count > 0)
                built.ImportedSheets = LegacyWorkbook.CopySheets(legacy.Book, master, missing);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            built.LegacyWarning = $"Не удалось перенести листы из старой книги {legacy.Path}: {ex.Message}";
        }
    }

    private static XLWorkbook Load(string path)
    {
        XLWorkbook wb;
        try
        {
            wb = new XLWorkbook(path);
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            throw new MasterWorkbookException($"Не удалось открыть общую книгу {path}: {ex.Message}", ex);
        }

        if (!SummaryWorkbookBuilder.IsOwnWorkbook(wb))
        {
            wb.Dispose();
            throw new MasterWorkbookException(
                $"Книга {path} создана не программой ConverterReport (например, книга прежнего формата) — программа ее не изменяет. " +
                "Укажите для общей книги новый файл: программа создаст его и будет дописывать в него листы дней.");
        }

        return wb;
    }

    /// <summary>Книгу нельзя менять, пока она открыта в Excel (у вас или у коллеги по сети).</summary>
    private static void EnsureNotLocked(string path)
    {
        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw Busy(path, ex);
        }
    }

    private static MasterWorkbookException Busy(string path, Exception inner) =>
        new($"Общая книга {path} открыта в другой программе (возможно, у коллеги) или недоступна для записи.", inner);
}

/// <summary>Лист дня нельзя дописать в общую книгу.</summary>
public sealed class MasterWorkbookException : Exception
{
    public MasterWorkbookException(string message, Exception? inner = null)
        : base(message, inner)
    {
    }
}
