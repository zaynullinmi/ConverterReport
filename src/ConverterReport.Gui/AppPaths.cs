using ConverterReport.Core.Configuration;
using ConverterReport.Core.Updates;

namespace ConverterReport.Gui;

/// <summary>Где хранятся настройки: рядом с программой, а если туда нельзя писать — в %APPDATA%\ConverterReport.</summary>
internal static class AppPaths
{
    public static string ExePath => Environment.ProcessPath ?? Path.Combine(AppContext.BaseDirectory, Updater.ExeName);

    public static string ExeFolder => Path.GetDirectoryName(ExePath) ?? AppContext.BaseDirectory;

    public static string DefaultConfigPath
    {
        get
        {
            var local = Path.Combine(ExeFolder, ConfigLoader.DefaultFileName);
            if (File.Exists(local) || Updater.CanWrite(ExeFolder))
                return local;
            var appData = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "ConverterReport");
            Directory.CreateDirectory(appData);
            return Path.Combine(appData, ConfigLoader.DefaultFileName);
        }
    }

    /// <summary>Настройки для первого запуска: этапы без путей и сроков — их заполняют на вкладке «Этапы».</summary>
    public static ReportConfig NewConfig()
    {
        var config = new ReportConfig();
        foreach (var n in new[] { "2", "3", "4", "9" })
            config.Projects.Add(new ProjectConfig { Number = n });
        return config;
    }
}
