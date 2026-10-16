using ClosedXML.Excel;
using ConverterReport.Core.Archive;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Model;
using ConverterReport.Core.Reporting;
using ConverterReport.Core.Screens;

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
    /// Дописывает лист дня в общую книгу (открытую <see cref="MasterBook.Open"/>). Книга сохраняется через временный файл,
    /// прежняя версия остается рядом как «.bak». Из прошлых версий отчета переносятся листы дней, которых в книге нет.
    /// </summary>
    /// <exception cref="MasterWorkbookException">Книгу не удалось сохранить (занята другой программой).</exception>
    public static SummaryWorkbook Append(ReportModel model, MasterBook master, int visibleDailySheets, IReadOnlyList<ArchiveBook>? archives = null)
    {
        var masterPath = master.Path;
        var folder = Path.GetDirectoryName(masterPath)!;
        Directory.CreateDirectory(folder);
        var temp = Path.Combine(folder, Path.GetFileNameWithoutExtension(masterPath) + ".tmp.xlsx");
        try
        {
            var result = new SummaryWorkbookBuilder(model).AppendTo(master.Workbook, masterPath, temp, visibleDailySheets,
                archives is { Count: > 0 } ? (book, built) => ImportDailySheets(book, built, archives, model.ReportTime.Date) : null);
            ChartInjector.Inject(temp, SummaryWorkbook.ChartsSheet, result.Charts);

            try
            {
                if (master.Exists)
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

    /// <summary>Листы прошлых дней из прошлых версий отчета, которых в общей книге еще нет (первая книга в списке — главная).</summary>
    private static void ImportDailySheets(XLWorkbook master, SummaryWorkbook built, IReadOnlyList<ArchiveBook> archives, DateTime reportDate)
    {
        foreach (var archive in archives.Where(a => a.DailySheets.Count > 0))
        {
            try
            {
                var existing = master.Worksheets.Select(w => w.Name).ToHashSet(StringComparer.OrdinalIgnoreCase);
                var missing = archive.DailySheets
                    .Where(n => !existing.Contains(n) && SummaryWorkbookBuilder.DailyDate(n, reportDate) < reportDate)
                    .ToList();
                if (missing.Count == 0)
                    continue;
                using var legacy = new LegacySource(archive.Path);
                built.ImportedSheets += LegacyWorkbook.CopySheets(legacy.Book, master, missing);
            }
            catch (Exception ex) when (ex is not OutOfMemoryException)
            {
                built.LegacyWarning = $"Не удалось перенести листы дней из {archive.Path}: {ex.Message}";
            }
        }
    }

    internal static XLWorkbook Load(string path)
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
    internal static void EnsureNotLocked(string path)
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

    internal static MasterWorkbookException Busy(string path, Exception inner) =>
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

/// <summary>
/// Общая книга, открытая для дописывания: листы Д.N/З.N в ней — накопленные данные этапов (в том числе из прошлых версий
/// отчета и архивных этапов без экрана).
/// </summary>
public sealed class MasterBook : IDisposable
{
    private MasterBook(string path, XLWorkbook workbook, bool exists)
    {
        Path = path;
        Workbook = workbook;
        Exists = exists;
    }

    public string Path { get; }

    public bool Exists { get; }

    internal XLWorkbook Workbook { get; }

    /// <exception cref="MasterWorkbookException">Книга занята другой программой, повреждена или создана не этой программой.</exception>
    public static MasterBook Open(string path)
    {
        if (!File.Exists(path))
            return new MasterBook(path, SummaryWorkbookBuilder.NewWorkbook(), exists: false);
        SummaryWorkbookWriter.EnsureNotLocked(path);
        return new MasterBook(path, SummaryWorkbookWriter.Load(path), exists: true);
    }

    /// <summary>Данные этапа в книге (листы Д.N/З.N без служебных столбцов программы).</summary>
    internal StageSource? Stage(string number, ScreenLayout layout)
    {
        var suffix = SummaryWorkbookBuilder.SheetSuffix(number);
        Workbook.TryGetWorksheet("З." + suffix, out var remarks);
        Workbook.TryGetWorksheet("Д." + suffix, out var documents);
        if (remarks is null && documents is null)
            return null;
        return new StageSource(System.IO.Path.GetFileName(Path),
            remarks is null ? null : ScreenReader.Raw(remarks, SummaryWorkbookBuilder.RemarksDataColumns(remarks, layout)),
            documents is null ? null : ScreenReader.Raw(documents, SummaryWorkbookBuilder.DocumentsDataColumns));
    }

    public void Dispose() => Workbook.Dispose();
}
