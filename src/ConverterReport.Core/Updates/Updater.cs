using System.Diagnostics;
using System.IO.Compression;
using System.Net;
using System.Reflection;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace ConverterReport.Core.Updates;

/// <summary>Новая версия программы, опубликованная в релизах GitHub.</summary>
public sealed record UpdateInfo(Version Version, string Tag, string DownloadUrl, string PageUrl, string Notes);

/// <summary>
/// Автообновление: проверка последнего релиза на GitHub, загрузка архива с новой версией и замена exe.
/// Запущенный exe в Windows нельзя удалить, но можно переименовать — старый файл переименовывается в *.old
/// и удаляется при следующем запуске.
/// </summary>
public static class Updater
{
    public const string AssetSuffix = "win-x64.zip";
    public const string ExeName = "ConverterReport.exe";

    /// <summary>Версия запущенной программы (номер сборки CI: 2.0.N).</summary>
    public static Version CurrentVersion
    {
        get
        {
            var v = Assembly.GetEntryAssembly()?.GetName().Version ?? new Version(0, 0, 0);
            return new Version(v.Major, v.Minor, Math.Max(0, v.Build));
        }
    }

    /// <summary>Возвращает новую версию или null, если установлена последняя.</summary>
    public static async Task<UpdateInfo?> CheckAsync(string repository, Version current, CancellationToken cancellation = default)
    {
        using var http = CreateClient(TimeSpan.FromSeconds(15), api: true);
        var json = await http.GetStringAsync($"https://api.github.com/repos/{repository.Trim().Trim('/')}/releases/latest", cancellation);
        var info = ParseRelease(json);
        return info is not null && info.Version > current ? info : null;
    }

    /// <summary>Разбор ответа GitHub API о релизе (tag «v2.0.7», архив «...win-x64.zip»).</summary>
    public static UpdateInfo? ParseRelease(string json)
    {
        using var doc = JsonDocument.Parse(json);
        var root = doc.RootElement;
        var tag = root.TryGetProperty("tag_name", out var t) ? t.GetString() ?? "" : "";
        var version = ParseVersion(tag);
        if (version is null)
            return null;

        string? url = null;
        if (root.TryGetProperty("assets", out var assets) && assets.ValueKind == JsonValueKind.Array)
        {
            foreach (var asset in assets.EnumerateArray())
            {
                var name = asset.TryGetProperty("name", out var n) ? n.GetString() ?? "" : "";
                if (name.EndsWith(AssetSuffix, StringComparison.OrdinalIgnoreCase) && asset.TryGetProperty("browser_download_url", out var u))
                {
                    url = u.GetString();
                    break;
                }
            }
        }

        if (url is null)
            return null;
        var page = root.TryGetProperty("html_url", out var h) ? h.GetString() ?? "" : "";
        var notes = root.TryGetProperty("body", out var b) ? b.GetString() ?? "" : "";
        return new UpdateInfo(version, tag, url, page, notes);
    }

    public static Version? ParseVersion(string tag)
    {
        var m = Regex.Match(tag, @"(\d+)\.(\d+)(?:\.(\d+))?");
        if (!m.Success)
            return null;
        return new Version(int.Parse(m.Groups[1].Value), int.Parse(m.Groups[2].Value), m.Groups[3].Success ? int.Parse(m.Groups[3].Value) : 0);
    }

    /// <summary>Скачивает архив новой версии и распаковывает exe во временную папку. Возвращает путь к новому exe.</summary>
    public static async Task<string> DownloadAsync(UpdateInfo info, IProgress<int>? progress = null, CancellationToken cancellation = default)
    {
        var dir = Path.Combine(Path.GetTempPath(), "ConverterReport-update-" + info.Version);
        Directory.CreateDirectory(dir);
        var zip = Path.Combine(dir, "update.zip");

        using (var http = CreateClient(TimeSpan.FromMinutes(10), api: false))
        using (var response = await http.GetAsync(info.DownloadUrl, HttpCompletionOption.ResponseHeadersRead, cancellation))
        {
            response.EnsureSuccessStatusCode();
            var total = response.Content.Headers.ContentLength;
            await using var source = await response.Content.ReadAsStreamAsync(cancellation);
            await using var target = File.Create(zip);
            var buffer = new byte[81920];
            long done = 0;
            int read, percent = -1;
            while ((read = await source.ReadAsync(buffer, cancellation)) > 0)
            {
                await target.WriteAsync(buffer.AsMemory(0, read), cancellation);
                done += read;
                if (total > 0 && (int)(done * 100 / total.Value) is var p && p != percent)
                    progress?.Report(percent = p);
            }
        }

        var exe = Path.Combine(dir, ExeName);
        using (var archive = ZipFile.OpenRead(zip))
        {
            var entry = archive.Entries.FirstOrDefault(e => string.Equals(e.Name, ExeName, StringComparison.OrdinalIgnoreCase))
                        ?? throw new InvalidDataException($"В архиве обновления нет файла {ExeName}.");
            entry.ExtractToFile(exe, overwrite: true);
        }

        return exe;
    }

    /// <summary>Подменяет запущенный exe новым и запускает новую версию. После вызова программу нужно закрыть.</summary>
    public static void ReplaceAndRestart(string newExe, string currentExe, string arguments = "--updated")
    {
        var old = currentExe + ".old";
        if (File.Exists(old))
            File.Delete(old);
        File.Move(currentExe, old);
        try
        {
            File.Copy(newExe, currentExe, overwrite: true);
        }
        catch
        {
            File.Move(old, currentExe); // вернуть как было
            throw;
        }

        Process.Start(new ProcessStartInfo(currentExe, arguments) { UseShellExecute = false, WorkingDirectory = Path.GetDirectoryName(currentExe) });
    }

    /// <summary>Удаляет файл предыдущей версии, оставшийся после обновления.</summary>
    public static void CleanupOld(string currentExe)
    {
        try
        {
            var old = currentExe + ".old";
            if (File.Exists(old))
                File.Delete(old);
        }
        catch
        {
            // Старый процесс еще не завершился — удалим в следующий раз.
        }
    }

    /// <summary>Можно ли записывать в папку программы (иначе обновить exe нельзя).</summary>
    public static bool CanWrite(string folder)
    {
        try
        {
            var probe = Path.Combine(folder, $".write-test-{Guid.NewGuid():N}");
            File.WriteAllText(probe, "");
            File.Delete(probe);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static HttpClient CreateClient(TimeSpan timeout, bool api)
    {
        // Корпоративный прокси: учетные данные текущего пользователя Windows.
        var handler = new HttpClientHandler { DefaultProxyCredentials = CredentialCache.DefaultCredentials };
        var http = new HttpClient(handler) { Timeout = timeout };
        http.DefaultRequestHeaders.UserAgent.ParseAdd("ConverterReport/" + CurrentVersion);
        if (api)
            http.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");
        return http;
    }
}
