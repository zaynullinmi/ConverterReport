using ConverterReport.Core.Configuration;

namespace ConverterReport.Core;

/// <summary>Параметры командной строки (общие для оконной и консольной программы).</summary>
public sealed class CommandLineOptions
{
    public const string Help = """
        ConverterReport — суточный отчет ГГЭ по данным файлов «Экран»: сводная книга Excel и три PDF
        (суточный отчет, детальный отчет, диаграммы).

        Без параметров открывается окно программы. Для запуска без окна (Планировщик заданий Windows)
        укажите --auto и при необходимости другие параметры.

        Параметры:
              --auto              сформировать отчет без окна и выйти
          -c, --config <файл>     файл настроек (по умолчанию ConverterReport.json рядом с программой)
          -d, --date <дд.мм.гггг> дата отчета (по умолчанию — сегодня)
          -t, --time <чч:мм>      время отчета (по умолчанию — ближайшее из стандартных, например 08:00 или 18:00)
              --from "<дд.мм.гггг чч:мм>"  начало отчетного периода (по умолчанию 18:00 предыдущего дня)
          -o, --out <папка>       папка для книги и PDF
              --engine <excel|builtin>  чем сохранять PDF: через Excel или встроенным генератором
          -n, --note "<текст>"    добавить строку в раздел «Справочно» (можно несколько раз)
              --open              открыть PDF после формирования
              --init              создать пример файла настроек и выйти
              --no-pause          то же, что --auto (в консольной версии — не ждать нажатия клавиши)
          -h, --help              эта справка

        Примеры:
          ConverterReport --auto -t 08:00
          ConverterReport --auto -d 15.10.2026 -t 18:00
        """;

    public bool Auto { get; private set; }

    public bool ShowHelp { get; private set; }

    public bool Init { get; private set; }

    public bool Open { get; private set; }

    public bool NoPause { get; private set; }

    public bool Updated { get; private set; }

    public string? ConfigPath { get; private set; }

    public string? OutputFolder { get; private set; }

    public DateTime? Date { get; private set; }

    public TimeSpan? Time { get; private set; }

    public DateTime? From { get; private set; }

    public PdfEngine? Engine { get; private set; }

    public List<string> Notes { get; } = new();

    /// <exception cref="ConfigurationException">Неверный параметр.</exception>
    public static CommandLineOptions Parse(IReadOnlyList<string> args)
    {
        var o = new CommandLineOptions();
        for (var i = 0; i < args.Count; i++)
        {
            var arg = args[i];
            string Next() => i + 1 < args.Count ? args[++i] : throw Error($"Не указано значение для параметра {arg}.");

            switch (arg.ToLowerInvariant())
            {
                case "-h" or "--help" or "/?":
                    o.ShowHelp = true;
                    break;
                case "--auto":
                    o.Auto = true;
                    break;
                case "--updated":
                    o.Updated = true;
                    break;
                case "-c" or "--config":
                    o.ConfigPath = Next();
                    break;
                case "-d" or "--date":
                    o.Date = Formats.ParseDate(Next()) ?? throw Error("Дата отчета указывается в формате дд.мм.гггг.");
                    break;
                case "-t" or "--time":
                    o.Time = Formats.ParseTime(Next()) ?? throw Error("Время отчета указывается в формате чч:мм.");
                    break;
                case "--from":
                    o.From = Formats.ParseDate(Next()) ?? throw Error("Начало периода указывается в формате \"дд.мм.гггг чч:мм\".");
                    break;
                case "-o" or "--out":
                    o.OutputFolder = Next();
                    break;
                case "--engine":
                    o.Engine = Next().ToLowerInvariant() switch
                    {
                        "excel" => PdfEngine.Excel,
                        "builtin" or "встроенный" => PdfEngine.Builtin,
                        "auto" => PdfEngine.Auto,
                        var other => throw Error($"Неизвестное значение --engine: {other} (excel, builtin или auto)."),
                    };
                    break;
                case "-n" or "--note":
                    o.Notes.Add(Next());
                    break;
                case "--open":
                    o.Open = true;
                    break;
                case "--no-pause":
                    o.NoPause = true;
                    break;
                case "--init":
                    o.Init = true;
                    break;
                default:
                    throw Error($"Неизвестный параметр: {arg}. Справка: --help");
            }
        }

        return o;
    }

    public RunOptions ToRunOptions() => new()
    {
        Date = Date,
        Time = Time,
        PeriodStart = From,
        OutputFolder = OutputFolder,
        PdfEngine = Engine,
        ExtraNotes = Notes.ToList(),
    };

    private static ConfigurationException Error(string message) => new(new[] { message });
}

/// <summary>Журнал ConverterReport.log в папке отчетов.</summary>
public sealed class RunLog
{
    private readonly object _lock = new();
    private readonly List<string> _buffer = new();
    private string? _path;

    public bool IsOpen => _path is not null;

    public void Open(string folder)
    {
        try
        {
            Directory.CreateDirectory(folder);
            lock (_lock)
            {
                _path = Path.Combine(folder, "ConverterReport.log");
                _buffer.Insert(0, $"{Environment.NewLine}===== {DateTime.Now:dd.MM.yyyy HH:mm:ss} =====");
            }
        }
        catch
        {
            _path = null;
        }
    }

    public void Write(string level, string message)
    {
        lock (_lock)
            _buffer.Add($"{DateTime.Now:HH:mm:ss} {level,-5} {message}");
    }

    public void Flush()
    {
        lock (_lock)
        {
            if (_path is null || _buffer.Count == 0)
                return;
            try
            {
                File.AppendAllLines(_path, _buffer, new System.Text.UTF8Encoding(false));
            }
            catch
            {
                // Журнал не критичен.
            }

            _buffer.Clear();
        }
    }
}
