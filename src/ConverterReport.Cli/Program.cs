using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using ConverterReport.Core;
using ConverterReport.Core.Configuration;

namespace ConverterReport.Cli;

/// <summary>Консольная версия (для запуска из командной строки и на Linux). Основная программа — оконная ConverterReport.exe.</summary>
internal static class Program
{
    private static readonly RunLog Log = new();

    private static int Main(string[] args)
    {
        Console.OutputEncoding = Encoding.UTF8;
        var pause = args.Length == 0 && LaunchedFromExplorer();
        try
        {
            var options = CommandLineOptions.Parse(args);
            if (options.NoPause)
                pause = false;
            return Run(options);
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

    private static int Run(CommandLineOptions options)
    {
        if (options.ShowHelp)
        {
            Console.WriteLine(CommandLineOptions.Help);
            return 0;
        }

        if (options.Init)
            return Init(options.ConfigPath);

        var configPath = options.ConfigPath ?? FindConfig();
        if (configPath is null || !File.Exists(configPath))
        {
            Error($"Не найден файл настроек {configPath ?? ConfigLoader.DefaultFileName}. Создайте его командой: ConverterReportCli --init");
            return 2;
        }

        var config = ConfigLoader.Load(configPath);
        Log.Open(config.ResolvePath(options.OutputFolder ?? config.Report.OutputFolder));
        Info($"Настройки: {Path.GetFullPath(configPath)}");

        var result = new ReportGenerator(Info, Warn).Run(config, options.ToRunOptions());
        SaveFilledDeadlines(config, configPath, result);

        Console.WriteLine();
        Console.ForegroundColor = ConsoleColor.Green;
        Console.WriteLine($"Готово. Сводная книга и PDF ({result.Pdfs.Count} шт.) — в папке {Path.GetDirectoryName(result.Workbook)}.");
        Console.ResetColor();
        if (result.Warnings.Count > 0)
            Console.WriteLine($"Предупреждений: {result.Warnings.Count} (см. выше и в журнале ConverterReport.log).");

        if (options.Open)
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

    /// <summary>Сроки, взятые из старой книги, сохраняются в настройки — в следующий раз книга для этого не нужна.</summary>
    private static void SaveFilledDeadlines(ReportConfig config, string configPath, RunResult result)
    {
        if (result.FilledDeadlines.Count == 0)
            return;
        try
        {
            ConfigLoader.Save(config, configPath);
            Info($"Сроки этапов {string.Join(", ", result.FilledDeadlines.Keys)} сохранены в настройки.");
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            Warn($"Не удалось сохранить сроки в {configPath}: {ex.Message}");
        }
    }
}
