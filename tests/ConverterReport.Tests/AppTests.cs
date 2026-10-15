using ConverterReport.Core;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Updates;

namespace ConverterReport.Tests;

public class UpdaterTests
{
    private const string Release = """
        {
          "tag_name": "v2.0.12",
          "html_url": "https://github.com/owner/repo/releases/tag/v2.0.12",
          "body": "Что нового",
          "assets": [
            { "name": "notes.txt", "browser_download_url": "https://example.com/notes.txt" },
            { "name": "ConverterReport-win-x64.zip", "browser_download_url": "https://example.com/ConverterReport-win-x64.zip" }
          ]
        }
        """;

    [Fact]
    public void ParseRelease_TakesVersionAndArchive()
    {
        var info = Updater.ParseRelease(Release)!;
        Assert.Equal(new Version(2, 0, 12), info.Version);
        Assert.Equal("v2.0.12", info.Tag);
        Assert.Equal("https://example.com/ConverterReport-win-x64.zip", info.DownloadUrl);
        Assert.Equal("https://github.com/owner/repo/releases/tag/v2.0.12", info.PageUrl);
        Assert.Equal("Что нового", info.Notes);
    }

    [Fact]
    public void ParseRelease_WithoutArchiveOrVersion_IsNull()
    {
        Assert.Null(Updater.ParseRelease("""{ "tag_name": "v2.0.5", "assets": [] }"""));
        Assert.Null(Updater.ParseRelease("""{ "tag_name": "latest", "assets": [ { "name": "a-win-x64.zip", "browser_download_url": "u" } ] }"""));
    }

    [Theory]
    [InlineData("v2.0.7", 2, 0, 7)]
    [InlineData("2.1", 2, 1, 0)]
    [InlineData("ConverterReport v10.2.300", 10, 2, 300)]
    public void ParseVersion(string tag, int major, int minor, int build)
    {
        Assert.Equal(new Version(major, minor, build), Updater.ParseVersion(tag));
    }

    [Fact]
    public void NewerVersionWins()
    {
        // Сравнение номеров, а не строк: 2.0.10 новее 2.0.9.
        Assert.True(Updater.ParseVersion("v2.0.10") > Updater.ParseVersion("v2.0.9"));
    }

    [Fact]
    public void ReplaceAndRestart_KeepsOldCopyAndCleanupRemovesIt()
    {
        var dir = Directory.CreateTempSubdirectory("cr-update-").FullName;
        try
        {
            var current = Path.Combine(dir, "app.bin");
            var fresh = Path.Combine(dir, "new.bin");
            File.WriteAllText(current, "old");
            File.WriteAllText(fresh, "new");

            // Запуск «новой версии» здесь не нужен: проверяем подмену файлов.
            Assert.ThrowsAny<Exception>(() => Updater.ReplaceAndRestart(fresh, current, "--updated"));
            Assert.Equal("new", File.ReadAllText(current));
            Assert.Equal("old", File.ReadAllText(current + ".old"));

            Updater.CleanupOld(current);
            Assert.False(File.Exists(current + ".old"));
            Assert.True(Updater.CanWrite(dir));
        }
        finally
        {
            Directory.Delete(dir, true);
        }
    }
}

public class CommandLineTests
{
    [Fact]
    public void Parse_AllOptions()
    {
        var o = CommandLineOptions.Parse(new[]
        {
            "--auto", "-c", "my.json", "-d", "15.10.2026", "-t", "8:00", "--from", "14.10.2026 18:00",
            "-o", "out", "--engine", "builtin", "-n", "строка 1", "--note", "строка 2", "--open", "--updated",
        });

        Assert.True(o.Auto);
        Assert.True(o.Open);
        Assert.True(o.Updated);
        Assert.Equal("my.json", o.ConfigPath);
        Assert.Equal(new DateTime(2026, 10, 15), o.Date);
        Assert.Equal(new TimeSpan(8, 0, 0), o.Time);
        Assert.Equal(new DateTime(2026, 10, 14, 18, 0, 0), o.From);
        Assert.Equal(PdfEngine.Builtin, o.Engine);
        Assert.Equal(new[] { "строка 1", "строка 2" }, o.Notes);

        var run = o.ToRunOptions();
        Assert.Equal("out", run.OutputFolder);
        Assert.Equal(new DateTime(2026, 10, 14, 18, 0, 0), run.PeriodStart);
        Assert.Equal(2, run.ExtraNotes.Count);
    }

    [Fact]
    public void Parse_NoArguments_OpensWindow()
    {
        var o = CommandLineOptions.Parse(Array.Empty<string>());
        Assert.False(o.Auto);
        Assert.False(o.ShowHelp);
        Assert.Null(o.ConfigPath);
    }

    [Theory]
    [InlineData("--unknown")]
    [InlineData("-t", "25:99")]
    [InlineData("-d")]
    [InlineData("--engine", "word")]
    public void Parse_Errors(params string[] args)
    {
        Assert.Throws<ConfigurationException>(() => CommandLineOptions.Parse(args));
    }
}

public class ConfigCloneTests
{
    [Fact]
    public void Clone_IsIndependentCopy()
    {
        var config = ConfigLoader.CreateSample();
        config.BaseDirectory = "/data";
        var copy = ConfigLoader.Clone(config);

        Assert.Equal("/data", copy.BaseDirectory);
        Assert.Equal(config.Projects.Count, copy.Projects.Count);
        Assert.Equal(config.Projects[3].Deadlines["Продление 2"], copy.Projects[3].Deadlines["Продление 2"]);
        Assert.Equal(config.Updates.Repository, copy.Updates.Repository);

        copy.Projects[0].Deadlines.Clear();
        copy.Report.ReportTimes.Add("12:00");
        Assert.NotEmpty(config.Projects[0].Deadlines);
        Assert.DoesNotContain("12:00", config.Report.ReportTimes);
    }
}
