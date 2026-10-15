using System.Text;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Export;
using ConverterReport.Core.Pdf;
using ConverterReport.Core.Reporting;
using ConverterReport.Core.Screens;
using ConverterReport.Core.Workbook;

namespace ConverterReport.Core;

/// <summary>Параметры запуска (перекрывают настройки файла конфигурации).</summary>
public sealed class RunOptions
{
    public DateTime Now { get; init; } = DateTime.Now;

    public DateTime? Date { get; init; }

    public TimeSpan? Time { get; init; }

    public DateTime? PeriodStart { get; init; }

    public string? OutputFolder { get; init; }

    public PdfEngine? PdfEngine { get; init; }

    public List<string> ExtraNotes { get; init; } = new();
}

public sealed class RunResult
{
    public required ReportModel Model { get; init; }

    /// <summary>Сводная книга Excel.</summary>
    public string? Workbook { get; set; }

    /// <summary>PDF: суточный отчет, детальный отчет, диаграммы (те, что включены).</summary>
    public List<string> Pdfs { get; } = new();

    public string? DailyPdf { get; set; }

    public string? DetailedPdf { get; set; }

    public string? ChartsPdf { get; set; }

    public IEnumerable<string> Files => (Workbook is null ? Enumerable.Empty<string>() : new[] { Workbook }).Concat(Pdfs);

    public List<string> Warnings { get; } = new();

    /// <summary>Сроки этапов, взятые из старой книги (в настройках их не было) — номер этапа → стадия → даты.</summary>
    public Dictionary<string, Dictionary<string, List<string>>> FilledDeadlines { get; } = new();
}

public sealed class ConfigurationException : Exception
{
    public ConfigurationException(IEnumerable<string> errors)
        : base("Ошибки в конфигурации:" + Environment.NewLine + string.Join(Environment.NewLine, errors.Select(e => " - " + e)))
    {
    }
}

/// <summary>
/// Полный цикл: прочитать экраны всех этапов → рассчитать отчет → сводная книга Excel (листы, формулы, диаграммы)
/// → три PDF: суточный отчет, детальный отчет, диаграммы.
/// </summary>
public sealed class ReportGenerator
{
    private readonly Action<string> _info;
    private readonly Action<string> _warn;

    public ReportGenerator(Action<string>? info = null, Action<string>? warn = null)
    {
        _info = info ?? (_ => { });
        _warn = warn ?? (_ => { });
    }

    public RunResult Run(ReportConfig config, RunOptions options)
    {
        var errors = ConfigLoader.Validate(config);
        if (errors.Count > 0)
            throw new ConfigurationException(errors);

        var reportTime = ReportTime.Resolve(config.Report, options.Now, options.Date, options.Time);
        var periodStart = options.PeriodStart ?? ReportTime.PeriodStart(config.Report, reportTime);
        if (periodStart >= reportTime)
            throw new ConfigurationException(new[] { $"Начало периода ({Formats.DateTimeText(periodStart)}) должно быть раньше времени отчета ({Formats.DateTimeText(reportTime)})." });

        _info($"Отчет на {Formats.DateTimeText(reportTime)}, период с {Formats.DateTimeText(periodStart)}.");

        var inputs = ReadScreens(config);
        if (inputs.All(i => i.Screen is null))
            throw new InvalidOperationException("Не удалось прочитать ни один файл экрана — отчет не сформирован.");

        var early = new List<string>();
        using var legacy = OpenLegacy(config, early);
        var filled = FillDeadlines(config, legacy, reportTime.Date, early);

        var notes = ReadNotes(config).Concat(options.ExtraNotes).ToList();
        var model = ReportBuilder.Build(config, inputs, reportTime, periodStart, notes);
        var result = new RunResult { Model = model };
        foreach (var (number, deadlines) in filled)
            result.FilledDeadlines[number] = deadlines;
        foreach (var w in early.Concat(model.AllWarnings))
        {
            result.Warnings.Add(w);
            _warn(w);
        }

        var folder = config.ResolvePath(options.OutputFolder ?? config.Report.OutputFolder);
        Directory.CreateDirectory(folder);
        var names = new Dictionary<string, string>
        {
            ["projects"] = model.ProjectsList,
            ["projectsshort"] = model.ProjectsList.Replace(", ", ","),
            ["version"] = config.Report.Version,
        };
        string FileFor(string template, string extension) =>
            Writable(folder, Formats.SafeFileName(Formats.ApplyTemplate(template, reportTime, names)) + extension);

        // 1. Сводная книга Excel: отдельная на отчет или лист дня в общей книге.
        SummaryWorkbook? workbook = null;
        if (config.Report.WorkbookMode == WorkbookMode.Append)
        {
            var master = MasterPath(config.Report.MasterWorkbook, folder, reportTime, names);
            try
            {
                workbook = SummaryWorkbookWriter.Append(model, master, config.Report.VisibleDailySheets, legacy);
                _info($"Лист «{workbook.DailySheet}» записан в общую книгу: {master}");
                if (workbook.ImportedSheets > 0)
                    _info($"Из старой книги перенесено листов прошлых дней: {workbook.ImportedSheets}.");
                if (workbook.LegacyWarning is { } legacyWarning)
                {
                    result.Warnings.Add(legacyWarning);
                    _warn(legacyWarning);
                }
            }
            catch (MasterWorkbookException ex)
            {
                var message = ex.Message + " Сформирована отдельная книга.";
                result.Warnings.Add(message);
                _warn(message);
            }
        }

        if (workbook is null)
        {
            var path = FileFor(config.Report.WorkbookFileName, ".xlsx");
            workbook = SummaryWorkbookWriter.Write(model, path);
            _info($"Сводная книга сохранена: {path}");
        }

        var workbookPath = workbook.Path;
        result.Workbook = workbookPath;

        // 2. Три PDF.
        var jobs = new List<(string Kind, SheetPdf Job)>();
        if (config.Report.IncludeDaily)
            jobs.Add(("daily", new SheetPdf(workbook.DailySheet, FileFor(config.Report.FileName, ".pdf"))));
        if (config.Report.IncludeDetailed && model.ShownProjects.Any(p => p.Current is not null))
            jobs.Add(("detailed", new SheetPdf(SummaryWorkbook.DetailedSheet, FileFor(config.Report.DetailedFileName, ".pdf"))));
        if (config.Report.IncludeCharts && workbook.ChartPages > 0)
            jobs.Add(("charts", new SheetPdf(SummaryWorkbook.ChartsSheet, FileFor(config.Report.ChartsFileName, ".pdf"))));

        var engine = options.PdfEngine ?? config.Report.PdfEngine;
        var viaExcel = false;
        if (engine != PdfEngine.Builtin && ExcelPdfExporter.IsAvailable())
        {
            try
            {
                ExcelPdfExporter.Export(workbookPath, jobs.Select(j => j.Job).ToList());
                viaExcel = true;
            }
            catch (Exception ex) when (engine == PdfEngine.Auto)
            {
                _warn($"Не удалось сохранить PDF через Excel ({ex.Message}) — используется встроенный генератор.");
            }
        }
        else if (engine == PdfEngine.Excel)
        {
            throw new InvalidOperationException("Для сохранения PDF через Excel нужен Windows с установленным Microsoft Excel (PdfEngine: \"Excel\").");
        }

        foreach (var (kind, job) in jobs)
        {
            if (!viaExcel)
            {
                switch (kind)
                {
                    case "daily":
                        PdfReportWriter.WriteDaily(model, job.PdfPath);
                        break;
                    case "detailed":
                        PdfReportWriter.WriteDetailed(model, job.PdfPath);
                        break;
                    default:
                        PdfReportWriter.WriteCharts(model, job.PdfPath);
                        break;
                }
            }

            result.Pdfs.Add(job.PdfPath);
            switch (kind)
            {
                case "daily":
                    result.DailyPdf = job.PdfPath;
                    break;
                case "detailed":
                    result.DetailedPdf = job.PdfPath;
                    break;
                default:
                    result.ChartsPdf = job.PdfPath;
                    break;
            }

            _info($"PDF сохранен: {job.PdfPath}");
        }

        _info(viaExcel ? "PDF сохранены через Microsoft Excel." : "PDF сформированы встроенным генератором.");
        return result;
    }

    private List<ProjectInput> ReadScreens(ReportConfig config)
    {
        var projects = config.EnabledProjects;
        var reader = new ScreenReader(config.ScreenLayout);
        var inputs = new ProjectInput[projects.Count];

        // Экраны обычно лежат на сетевом диске — читаем параллельно.
        Parallel.For(0, projects.Count, new ParallelOptions { MaxDegreeOfParallelism = 4 }, i =>
        {
            var project = projects[i];
            var path = config.ResolvePath(project.ScreenFile);
            try
            {
                inputs[i] = new ProjectInput(project, reader.Read(path));
            }
            catch (FileNotFoundException)
            {
                inputs[i] = new ProjectInput(project, null, $"файл экрана не найден: {path}");
            }
            catch (DirectoryNotFoundException)
            {
                inputs[i] = new ProjectInput(project, null, $"папка с файлом экрана не найдена: {path}");
            }
            catch (Exception ex)
            {
                inputs[i] = new ProjectInput(project, null, $"не удалось прочитать экран {path}: {ex.Message}");
            }
        });

        foreach (var input in inputs)
        {
            if (input.Screen is not null)
                _info($"Этап {input.Project.Number}: прочитано строк замечаний — {input.Screen.Remarks.Count}, загрузок — {input.Screen.Loads.Count}.");
        }

        return inputs.ToList();
    }

    /// <summary>Старая книга для переноса листов прошлых дней и сроков (если включено в настройках и файл есть).</summary>
    private LegacySource? OpenLegacy(ReportConfig config, List<string> warnings)
    {
        var r = config.Report;
        if (!r.ImportLegacyDays || string.IsNullOrWhiteSpace(r.LegacyWorkbook))
            return null;
        var path = config.ResolvePath(r.LegacyWorkbook);
        if (File.Exists(path))
            return new LegacySource(path);
        warnings.Add($"Старая книга не найдена: {path} — листы прошлых дней не перенесены.");
        return null;
    }

    /// <summary>Сроки этапов, не заполненные в настройках, — из последнего листа дня старой книги.</summary>
    private Dictionary<string, Dictionary<string, List<string>>> FillDeadlines(ReportConfig config, LegacySource? legacy, DateTime reportDate, List<string> warnings)
    {
        var filled = new Dictionary<string, Dictionary<string, List<string>>>();
        var empty = config.EnabledProjects.Where(p => p.Deadlines.Count == 0).ToList();
        if (legacy is null || empty.Count == 0)
            return filled;
        try
        {
            var sheet = LegacyWorkbook.LastDailySheet(legacy.Book, reportDate);
            if (sheet is null)
                return filled;
            var found = LegacyWorkbook.ReadDeadlines(sheet);
            foreach (var project in empty)
            {
                if (!found.TryGetValue(project.Number.Trim(), out var deadlines))
                    continue;
                project.Deadlines = deadlines;
                filled[project.Number] = deadlines;
                _info($"Этап {project.Number}: сроки взяты из старой книги (лист {sheet.Name}).");
            }
        }
        catch (Exception ex) when (ex is not OutOfMemoryException)
        {
            warnings.Add($"Не удалось прочитать сроки из старой книги {legacy.Path}: {ex.Message}");
        }

        return filled;
    }

    /// <summary>Файл общей книги: относительно папки отчетов (или полный путь), с подстановками, расширение .xlsx.</summary>
    internal static string MasterPath(string template, string folder, DateTime reportTime, IReadOnlyDictionary<string, string> names)
    {
        var path = Formats.ApplyTemplate(template.Trim(), reportTime, names);
        var name = Formats.SafeFileName(Path.GetFileName(path));
        if (!name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            name += ".xlsx";
        var dir = Path.GetDirectoryName(path);
        var full = string.IsNullOrEmpty(dir) ? name : Path.Combine(dir, name);
        return Path.GetFullPath(Path.IsPathRooted(full) ? full : Path.Combine(folder, full));
    }

    /// <summary>Путь для сохранения; если файл занят (открыт в Excel или просмотрщике PDF) — имя с номером.</summary>
    private string Writable(string folder, string fileName)
    {
        var path = Path.Combine(folder, fileName);
        try
        {
            if (File.Exists(path))
                File.Delete(path);
            return path;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            var stem = Path.GetFileNameWithoutExtension(fileName);
            var ext = Path.GetExtension(fileName);
            var alternative = path;
            for (var i = 2; File.Exists(alternative); i++)
                alternative = Path.Combine(folder, $"{stem} ({i}){ext}");
            _warn($"Файл {path} занят другой программой, сохраняю как {alternative}.");
            return alternative;
        }
    }

    /// <summary>Раздел «Справочно» из текстового файла: каждая непустая строка — абзац, строки с «//» — комментарии.</summary>
    public static List<string> ReadNotes(ReportConfig config)
    {
        if (string.IsNullOrWhiteSpace(config.Report.NotesFile))
            return new List<string>();
        var path = config.ResolvePath(config.Report.NotesFile);
        if (!File.Exists(path))
            return new List<string>();

        var bytes = File.ReadAllBytes(path);
        string text;
        try
        {
            text = new UTF8Encoding(false, throwOnInvalidBytes: true).GetString(bytes);
        }
        catch (DecoderFallbackException)
        {
            // Файл сохранен в кодировке Windows (ANSI).
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            text = Encoding.GetEncoding(1251).GetString(bytes);
        }

        return text.TrimStart('﻿')
            .Split('\n')
            .Select(l => l.TrimEnd('\r'))
            .Where(l => !string.IsNullOrWhiteSpace(l) && !l.TrimStart().StartsWith("//", StringComparison.Ordinal))
            .ToList();
    }
}
