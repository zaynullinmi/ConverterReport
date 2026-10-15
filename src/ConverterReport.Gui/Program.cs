using ConverterReport.Core;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Updates;

namespace ConverterReport.Gui;

internal static class Program
{
    [STAThread]
    private static int Main(string[] args)
    {
        Updater.CleanupOld(AppPaths.ExePath);

        CommandLineOptions options;
        try
        {
            options = CommandLineOptions.Parse(args);
        }
        catch (ConfigurationException ex)
        {
            if (!args.Any(a => a.Equals("--auto", StringComparison.OrdinalIgnoreCase)))
                MessageBox.Show(ex.Message, "ConverterReport", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return 2;
        }

        // Без окна — для Планировщика заданий Windows.
        if (options.Auto && !options.ShowHelp)
            return Headless.Run(options);

        Application.SetHighDpiMode(HighDpiMode.SystemAware);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        if (options.ShowHelp)
        {
            TextDialog.Show(null, "ConverterReport — параметры запуска", CommandLineOptions.Help);
            return 0;
        }

        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
        Application.ThreadException += (_, e) =>
            MessageBox.Show(e.Exception.Message, "ConverterReport — ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
        Application.Run(new MainForm(options));
        return 0;
    }
}

/// <summary>Формирование отчета без окна: результат и ошибки — в журнале ConverterReport.log.</summary>
internal static class Headless
{
    public static int Run(CommandLineOptions options)
    {
        var log = new RunLog();
        try
        {
            var path = options.ConfigPath ?? AppPaths.DefaultConfigPath;
            if (!File.Exists(path))
                throw new ConfigurationException(new[] { $"Не найден файл настроек {path}. Запустите программу без параметров и заполните настройки." });
            var config = ConfigLoader.Load(path);
            log.Open(config.ResolvePath(options.OutputFolder ?? config.Report.OutputFolder));
            log.Write("INFO", $"Запуск без окна, настройки: {path}");
            var result = new ReportGenerator(m => log.Write("INFO", m), m => log.Write("WARN", m)).Run(config, options.ToRunOptions());
            log.Write("INFO", $"Готово: {string.Join("; ", result.Files)}");
            return 0;
        }
        catch (ConfigurationException ex)
        {
            log.Write("ERROR", ex.Message);
            return 2;
        }
        catch (Exception ex)
        {
            log.Write("ERROR", ex.ToString());
            return 1;
        }
        finally
        {
            if (!log.IsOpen)
                log.Open(AppPaths.ExeFolder);
            log.Flush();
        }
    }
}
