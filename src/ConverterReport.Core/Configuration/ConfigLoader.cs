using System.Globalization;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

namespace ConverterReport.Core.Configuration;

public static class ConfigLoader
{
    public const string DefaultFileName = "ConverterReport.json";

    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
        WriteIndented = true,
        Encoder = JavaScriptEncoder.Create(UnicodeRanges.All),
    };

    public static ReportConfig Load(string path)
    {
        var fullPath = Path.GetFullPath(path);
        var json = File.ReadAllText(fullPath);
        var config = JsonSerializer.Deserialize<ReportConfig>(json, Options)
                     ?? throw new InvalidDataException($"Файл конфигурации пуст: {fullPath}");
        config.BaseDirectory = Path.GetDirectoryName(fullPath) ?? Directory.GetCurrentDirectory();
        Upgrade(config);
        return config;
    }

    /// <summary>
    /// Настройки прежних версий: имена файлов по умолчанию заменяются на текущие (свои имена не трогаются),
    /// старая книга переносится в список прошлых версий отчета.
    /// </summary>
    internal static void Upgrade(ReportConfig config)
    {
        var r = config.Report;
        r.FileName = ReportSettings.ObsoleteFileNames.GetValueOrDefault(r.FileName, r.FileName);
        r.DetailedFileName = ReportSettings.ObsoleteFileNames.GetValueOrDefault(r.DetailedFileName, r.DetailedFileName);
        r.ChartsFileName = ReportSettings.ObsoleteFileNames.GetValueOrDefault(r.ChartsFileName, r.ChartsFileName);
        if (!string.IsNullOrWhiteSpace(r.LegacyWorkbook))
        {
            var legacy = r.LegacyWorkbook.Trim();
            if (!r.ArchiveWorkbooks.Any(a => string.Equals(a.Path.Trim(), legacy, StringComparison.OrdinalIgnoreCase)))
                r.ArchiveWorkbooks.Insert(0, new ArchiveSource { Path = legacy });
        }

        r.LegacyWorkbook = null;
    }

    public static void Save(ReportConfig config, string path)
    {
        File.WriteAllText(path, JsonSerializer.Serialize(config, Options));
    }

    /// <summary>Независимая копия настроек (для формирования отчета в фоне, пока окно открыто).</summary>
    public static ReportConfig Clone(ReportConfig config)
    {
        var copy = JsonSerializer.Deserialize<ReportConfig>(JsonSerializer.Serialize(config, Options), Options)!;
        copy.BaseDirectory = config.BaseDirectory;
        return copy;
    }

    /// <summary>Проверка конфигурации. Возвращает список ошибок (пустой — все в порядке).</summary>
    public static List<string> Validate(ReportConfig config)
    {
        var errors = new List<string>();
        var enabled = config.EnabledProjects;
        if (enabled.Count == 0)
            errors.Add("Не задан ни один включенный проект (раздел \"Projects\").");

        foreach (var p in enabled)
        {
            if (string.IsNullOrWhiteSpace(p.Number))
                errors.Add("У проекта не заполнен номер (\"Number\").");
            if (p.ExpertiseEndDate is { Length: > 0 } end && Formats.ParseDate(end) is null)
                errors.Add($"Этап {p.Number}: не удалось разобрать дату окончания экспертизы \"{end}\".");
            if (p.ChartStartDate is { Length: > 0 } start && Formats.ParseDate(start) is null)
                errors.Add($"Этап {p.Number}: не удалось разобрать дату начала диаграмм \"{start}\".");
        }

        foreach (var dup in enabled.GroupBy(p => p.Number.Trim()).Where(g => g.Count() > 1))
            errors.Add($"Номер этапа \"{dup.Key}\" указан несколько раз.");

        foreach (var t in config.Report.ReportTimes.Append(config.Report.PeriodStartTime))
        {
            if (Formats.ParseTime(t) is null)
                errors.Add($"Не удалось разобрать время \"{t}\" (ожидается ЧЧ:ММ).");
        }

        if (config.Report.WorkbookMode == WorkbookMode.Append && string.IsNullOrWhiteSpace(config.Report.MasterWorkbook))
            errors.Add("Не указан файл общей книги (\"MasterWorkbook\").");
        if (enabled.Count > 0 && enabled.All(p => p.IsArchive))
            errors.Add("Ни у одного включенного этапа не указан файл экрана (\"ScreenFile\").");
        if (config.Report.ImportLegacyDays && config.Report.ArchiveWorkbooks.All(a => string.IsNullOrWhiteSpace(a.Path)))
            errors.Add("Не указаны прошлые версии отчета для архива (\"ArchiveWorkbooks\").");
        foreach (var p in enabled.Where(p => p.LoadsAxisMax is < 0))
            errors.Add($"Этап {p.Number}: верх шкалы загрузок не может быть отрицательным.");
        if (config.Report.VisibleDailySheets < 0)
            errors.Add("Число видимых листов дней (\"VisibleDailySheets\") не может быть отрицательным.");

        if (config.Charts.StartDate is { Length: > 0 } chartStart && Formats.ParseDate(chartStart) is null)
            errors.Add($"Не удалось разобрать дату начала диаграмм \"{chartStart}\".");

        return errors;
    }

    /// <summary>Пример конфигурации, который создается командой --init.</summary>
    public static ReportConfig CreateSample()
    {
        var config = new ReportConfig();
        config.Projects.Add(Sample("2", @"\\server\share\2 этап\Экран ГГЭ 2 этап.xlsx", "18.11.2026"));
        config.Projects.Add(Sample("3", @"\\server\share\3 этап\Экран ГГЭ 3 этап.xlsx", "18.11.2026"));
        config.Projects.Add(Sample("4", @"\\server\share\4 этап\Экран ГГЭ 4 этап.xlsx", "17.11.2026"));
        var p9 = Sample("9", @"\\server\share\9 этап\Экран ГГЭ 9 этап.xlsx", "09.11.2026");
        p9.Deadlines["Продление 2"] = new List<string> { "Заявка до 23.10.2026" };
        config.Projects.Add(p9);
        return config;

        static ProjectConfig Sample(string number, string file, string end)
        {
            var endDate = DateTime.ParseExact(end, "dd.MM.yyyy", CultureInfo.InvariantCulture);
            string D(DateTime d) => d.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
            return new ProjectConfig
            {
                Number = number,
                ScreenFile = file,
                Deadlines = new Dictionary<string, List<string>>
                {
                    ["Входной контроль"] = new() { D(endDate.AddDays(-160)), D(endDate.AddDays(-147)) },
                    ["Экспертиза"] = new() { D(endDate.AddDays(-147)), D(endDate.AddDays(-86)) },
                    ["Продление 1"] = new() { D(endDate.AddDays(-85)), D(endDate.AddDays(-58)) },
                    ["Продление 2"] = new() { D(endDate.AddDays(-55)), D(endDate) },
                },
            };
        }
    }
}
