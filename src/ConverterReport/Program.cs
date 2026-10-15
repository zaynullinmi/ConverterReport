using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using ConverterReport.Core;
using ConverterReport.Core.Configuration;

namespace ConverterReport;

internal static class Program
{
    private const string Help = """
        ConverterReport — суточный отчет ГГЭ по данным файлов «Экран»: сводная книга Excel и три PDF
        (суточный отчет, детальный отчет, диаграммы).

        Использование:
          ConverterReport [параметры]

        Параметры:
          -c, --config <файл>     файл настроек (по умолчанию ConverterReport.json рядом с программой
                                  или в текущей папке)
          -d, --date <дд.мм.гггг> дата отчета (по умолчанию — сегодня)
          -t, --time <чч:мм>      время отчета (по умолчанию — ближайшее из ReportTimes, например 08:00 или 18:00)
              --from "<дд.мм.гггг чч:мм>"  начало отчетного периода (по умолчанию 18:00 предыдущего дня)
          -o, --out <папка>       папка для книги и PDF (перекрывает OutputFolder)
              --engine <excel|builtin>  чем сохранять PDF: через Excel (как вручную) или встроенным генератором
                                  (по умолчанию — через Excel, если он установлен)
          -n, --note "<текст>"    добавить строку в раздел «Справочно» (можно несколько раз)
              --open              открыть PDF после формирования
              --init              создать пример файла настроек и выйти
              --no-pause          не ждать нажатия клавиши в конце (для планировщика задач)
          -h, --help              эта справка

        Примеры:
          ConverterReport
          ConverterReport -t 18:00
          ConverterReport -d 15.10.2026 -t 08:00 --open
        """;

    private static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        var pause = args.Length == 0 && LaunchedFromExplorer();
        try
        {
            return Run(args, ref pause);
        }
        catch (ConfigurationException ex)
        {
            Error(ex.Message);
            return 2;
        }
        catch (Exception ex)
        {
            Error(ex.Message);
            Log.Write("ERROR", ex.ToString());
            return 1;
        }
        finally
        {
            Log.Flush();
            if (pause)
            {
                Console.WriteLine();
                Console.WriteLine("Нажмите любую клавишу для выхода...");
                Console.ReadKey(true);
            }
        }
    }

    private static int Run(string[] args, ref bool pause)
    {
        string? configPath = null;
        string? outFolder = null;
        DateTime? date = null;
        TimeSpan? time = null;
        DateTime? from = null;
        PdfEngine? engine = null;
        var open = false;
        var init = false;
        var notes = new List<string>();

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            string Next() => i + 1 < args.Length ? args[++i] : throw new ConfigurationException(new[] { $"Не указано значение для параметра {arg}." });

            switch (arg.ToLowerInvariant())
            {
                case "-h" or "--help" or "/?":
                    Console.WriteLine(Help);
                    return 0;
                case "-c" or "--config":
                    configPath = Next();
                    break;
                case "-d" or "--date":
                    date = Formats.ParseDate(Next()) ?? throw new ConfigurationException(new[] { "Дата отчета указывается в формате дд.мм.гггг." });
                    break;
                case "-t" or "--time":
                    time = Formats.ParseTime(Next()) ?? throw new ConfigurationException(new[] { "Время отчета указывается в формате чч:мм." });
                    break;
                case "--from":
                    from = Formats.ParseDate(Next()) ?? throw new ConfigurationException(new[] { "Начало периода указывается в формате \"дд.мм.гггг чч:мм\"." });
                    break;
                case "-o" or "--out":
                    outFolder = Next();
                    break;
                case "--engine":
                    engine = Next().ToLowerInvariant() switch
                    {
                        "excel" => PdfEngine.Excel,
                        "builtin" or "встроенный" => PdfEngine.Builtin,
                        "auto" => PdfEngine.Auto,
                        var other => throw new ConfigurationException(new[] { $"Неизвестное значение --engine: {other} (excel, builtin или auto)." }),
                    };
                    break;
                case "-n" or "--note":
                    notes.Add(Next());
                    break;
                case "--open":
                    open = true;
                    break;
                case "--no-pause":
                    pause = false;
                    break;
                case "--init":
                    init = true;
                    break;
                default:
                    throw new ConfigurationException(new[] { $"Неизвестный параметр: {arg}. Справка: --help" });
            }
        }

        if (init)
            return Init(configPath);

        configPath ??= FindConfig();
        if (configPath is null || !File.Exists(configPath))
        {
            Error($"Не найден файл настроек {configPath ?? ConfigLoader.DefaultFileName}. Создайте его командой: ConverterReport --init");
            return 2;
        }

        var config = ConfigLoader.Load(configPath);
        Log.Open(config.ResolvePath(outFolder ?? config.Report.OutputFolder));
        Info($"Настройки: {Path.GetFullPath(configPath)}");

        var generator = new ReportGenerator(Info, Warn);
        var result = generator.Run(config, new RunOptions
        {
            Date = date,
            Time = time,
            PeriodStart = from,
            OutputFolder = outFolder,
            PdfEngine = engine,
            ExtraNotes = notes,
        });

        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Готово. Сводная книга и PDF ({result.Pdfs.Count} шт.) — в папке {Path.GetDirectoryName(result.Workbook)}.");
        Console.ResetColor();
        if (result.Warnings.Count > 0)
            Console.WriteLine($"Предупреждений: {result.Warnings.Count} (см. выше и в журнале ConverterReport.log).");

        if (open)
        {
            foreach (var pdf in result.Pdfs)
                Process.Start(new ProcessStartInfo(pdf) { UseShellExecute = true });
        }
        return 0;
    }

    private static int Init(string? path)
    {
        path ??= Path.Combine(Directory.GetCurrentDirectory(), ConfigLoader.DefaultFileName);
        if (File.Exists(path))
        {
            Error($"Файл {path} уже существует — он не будет перезаписан.");
            return 2;
        }

        ConfigLoader.Save(ConfigLoader.CreateSample(), path);
        Info($"Создан пример настроек: {path}. Укажите в нем пути к файлам экранов и сроки этапов.");
        return 0;
    }

    private static string? FindConfig()
    {
        var candidates = new[]
        {
            Path.Combine(Directory.GetCurrentDirectory(), ConfigLoader.DefaultFileName),
            Path.Combine(AppContext.BaseDirectory, ConfigLoader.DefaultFileName),
        };
        return candidates.FirstOrDefault(File.Exists);
    }

    private static void Info(string message)
    {
        Console.WriteLine(message);
        Log.Write("INFO", message);
    }

    private static void Warn(string message)
    {
        Console.ForegroundColor = ConsoleColor.Yellow;
        Console.WriteLine("Внимание: " + message);
        Console.ResetColor();
        Log.Write("WARN", message);
    }

    private static void Error(string message)
    {
        Console.ForegroundColor = ConsoleColor.Red;
        Console.Error.WriteLine("Ошибка: " + message);
        Console.ResetColor();
        Log.Write("ERROR", message);
    }

    /// <summary>Программа запущена двойным щелчком (окно консоли принадлежит только ей) — тогда ждем клавишу в конце.</summary>
    private static bool LaunchedFromExplorer()
    {
        if (!OperatingSystem.IsWindows() || Console.IsInputRedirected)
            return false;
        try
        {
            var list = new uint[2];
            return GetConsoleProcessList(list, 2) == 1;
        }
        catch
        {
            return false;
        }
    }

    [DllImport("kernel32.dll", SetLastError = true)]
    private static extern uint GetConsoleProcessList(uint[] processList, uint processCount);

    /// <summary>Журнал ConverterReport.log в папке отчетов.</summary>
    private static class Log
    {
        private static readonly StringBuilder Buffer = new();
        private static string? _path;

        public static void Open(string folder)
        {
            try
            {
                Directory.CreateDirectory(folder);
                _path = Path.Combine(folder, "ConverterReport.log");
                Buffer.Insert(0, $"{Environment.NewLine}===== {DateTime.Now:dd.MM.yyyy HH:mm:ss} ====={Environment.NewLine}");
            }
            catch
            {
                _path = null;
            }
        }

        public static void Write(string level, string message) =>
            Buffer.AppendLine($"{DateTime.Now:HH:mm:ss} {level,-5} {message}");

        public static void Flush()
        {
            if (_path is null || Buffer.Length == 0)
                return;
            try
            {
                File.AppendAllText(_path, Buffer.ToString(), Encoding.UTF8);
            }
            catch
            {
                // Журнал не критичен.
            }

            Buffer.Clear();
        }
    }
}
