using System.Text;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Export;
using ConverterReport.Core.Pdf;
using ConverterReport.Core.Reporting;
using ConverterReport.Core.Screens;

namespace ConverterReport.Core;

/// <summary>Параметры запуска (перекрывают настройки файла конфигурации).</summary>
public sealed class RunOptions
{
    public DateTime Now { get; init; } = DateTime.Now;

    public DateTime? Date { get; init; }

    public TimeSpan? Time { get; init; }

    public DateTime? PeriodStart { get; init; }

    public string? OutputFolder { get; init; }

    public bool? PerProject { get; init; }

    public bool? ExportExcel { get; init; }

    public List<string> ExtraNotes { get; init; } = new();
}

public sealed class RunResult
{
    public required ReportModel Model { get; init; }

    public List<string> Files { get; } = new();

    public List<string> Warnings { get; } = new();
}

public sealed class ConfigurationException : Exception
{
    public ConfigurationException(IEnumerable<string> errors)
        : base("Ошибки в конфигурации:" + Environment.NewLine + string.Join(Environment.NewLine, errors.Select(e => " - " + e)))
    {
    }
}

/// <summary>
/// Полный цикл: прочитать экраны всех этапов → рассчитать отчет → сохранить PDF (и при необходимости
/// отдельные PDF по этапам и выгрузку в Excel).
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

        var notes = ReadNotes(config).Concat(options.ExtraNotes).ToList();
        var model = ReportBuilder.Build(config, inputs, reportTime, periodStart, notes);
        var result = new RunResult { Model = model };
        foreach (var w in model.AllWarnings)
        {
            result.Warnings.Add(w);
            _warn(w);
        }

        var folder = config.ResolvePath(options.OutputFolder ?? config.Report.OutputFolder);
        Directory.CreateDirectory(folder);

        var mainName = Formats.ApplyTemplate(config.Report.FileName, reportTime);
        result.Files.Add(SavePdf(model, folder, mainName));

        if (options.PerProject ?? config.Report.PerProjectReports)
        {
            foreach (var input in inputs.Where(i => i.Screen is not null))
            {
                var single = ReportBuilder.Build(config, new[] { input }, reportTime, periodStart, notes: null);
                if (single.Projects[0].Current is null && single.Projects[0].PeriodLoads.Count == 0)
                {
                    _info($"Этап {input.Project.Number}: нет данных — отдельный PDF не формируется.");
                    continue;
                }

                var name = Formats.ApplyTemplate(config.Report.ProjectFileName, reportTime, new Dictionary<string, string>
                {
                    ["number"] = input.Project.Number,
                    ["name"] = input.Project.DisplayName,
                });
                result.Files.Add(SavePdf(single, folder, name));
            }
        }

        if (options.ExportExcel ?? config.Report.ExportExcel)
        {
            var xlsx = UniquePath(folder, Formats.SafeFileName(mainName) + ".xlsx");
            ExcelExporter.Export(model, xlsx);
            _info($"Данные выгружены в Excel: {xlsx}");
            result.Files.Add(xlsx);
        }

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

    private string SavePdf(ReportModel model, string folder, string name)
    {
        var path = Path.Combine(folder, Formats.SafeFileName(name) + ".pdf");
        try
        {
            if (File.Exists(path))
                File.Delete(path);
        }
        catch (IOException)
        {
            // Файл открыт в просмотрщике — сохраняем под другим именем.
            var alternative = UniquePath(folder, Path.GetFileName(path));
            _warn($"Файл {path} занят другой программой, отчет сохранен как {alternative}.");
            path = alternative;
        }

        var result = PdfReportWriter.Write(model, path);
        var blank = result.RemovedBlankPages > 0 ? $", удалено пустых: {result.RemovedBlankPages}" : "";
        _info($"PDF сохранен: {result.Path} (страниц: {result.Pages}, из них диаграмм: {result.ChartPages}{blank}).");
        return result.Path;
    }

    private static string UniquePath(string folder, string fileName)
    {
        var path = Path.Combine(folder, fileName);
        var stem = Path.GetFileNameWithoutExtension(fileName);
        var ext = Path.GetExtension(fileName);
        for (var i = 2; File.Exists(path); i++)
            path = Path.Combine(folder, $"{stem} ({i}){ext}");
        return path;
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
