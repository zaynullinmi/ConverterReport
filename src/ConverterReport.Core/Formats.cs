using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace ConverterReport.Core;

/// <summary>
/// Разбор и форматирование значений. В экранах даты и время встречаются в разном виде:
/// числом Excel, датой, датой со временем, строкой «10.10.2026» / «18:56» — все варианты приводятся к DateTime/TimeSpan.
/// </summary>
public static class Formats
{
    public static readonly CultureInfo Ru = CultureInfo.GetCultureInfo("ru-RU");

    private static readonly string[] DateFormats =
    {
        "dd.MM.yyyy", "d.M.yyyy", "dd.MM.yy", "d.M.yy", "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy",
        "dd.MM.yyyy H:mm", "dd.MM.yyyy HH:mm", "d.M.yyyy H:mm", "dd.MM.yyyy H:mm:ss", "dd.MM.yyyy HH:mm:ss",
        "yyyy-MM-dd HH:mm", "yyyy-MM-dd HH:mm:ss", "yyyy-MM-ddTHH:mm:ss",
    };

    private static readonly string[] TimeFormats = { "H:mm", "HH:mm", "H:mm:ss", "HH:mm:ss", "H.mm", "HH.mm" };

    // Диапазон дат Excel (01.01.1900 .. 31.12.9999)
    private const double MinOaDate = 1;
    private const double MaxOaDate = 2958465;

    /// <summary>Дата (с временем, если оно есть в значении). null — если значение не является датой.</summary>
    public static DateTime? ParseDate(object? value)
    {
        switch (value)
        {
            case null:
                return null;
            case DateTime dt when dt.Year >= 1900:
                return RoundToMinute(dt);
            case DateTime:
                return null; // «нулевая» дата Excel — в ячейке только время
            case TimeSpan ts when ts.TotalDays >= MinOaDate && ts.TotalDays <= MaxOaDate:
                return RoundToMinute(DateTime.FromOADate(ts.TotalDays));
            case double d when d >= MinOaDate && d <= MaxOaDate:
                return RoundToMinute(DateTime.FromOADate(d));
            case int i when i >= MinOaDate && i <= MaxOaDate:
                return DateTime.FromOADate(i);
            case string s:
                return ParseDateString(s);
            default:
                return null;
        }
    }

    /// <summary>Время суток. Для значений «дата+время» берется только время.</summary>
    public static TimeSpan? ParseTime(object? value)
    {
        switch (value)
        {
            case null:
                return null;
            case TimeSpan ts:
                return TimeSpan.FromMinutes(Math.Round((ts.TotalMinutes % 1440 + 1440) % 1440));
            case DateTime dt:
                return RoundToMinute(dt).TimeOfDay;
            case double d when d >= 0 && d <= MaxOaDate:
                return TimeSpan.FromMinutes(Math.Round((d - Math.Floor(d)) * 1440));
            case string s:
                return ParseTimeString(s);
            default:
                return null;
        }
    }

    /// <summary>Объединяет ячейки «Дата» и «Время». Если время пустое — берется время из даты (если было).</summary>
    public static DateTime? Combine(object? dateValue, object? timeValue)
    {
        var date = ParseDate(dateValue);
        if (date is null)
            return null;
        var time = ParseTime(timeValue);
        return time is null ? date : date.Value.Date + time.Value;
    }

    public static bool IsNumber(object? value) => value switch
    {
        double or int or long or decimal => true,
        string s => TryParseNumber(s, out _),
        _ => false,
    };

    /// <summary>Целое значение счетчика. Пустые ячейки, «-» и текст дают 0.</summary>
    public static int ParseInt(object? value) => value switch
    {
        double d => (int)Math.Round(d),
        int i => i,
        long l => (int)l,
        decimal m => (int)Math.Round(m),
        string s when TryParseNumber(s, out var n) => (int)Math.Round(n),
        _ => 0,
    };

    private static bool TryParseNumber(string s, out double result)
    {
        var cleaned = s.Trim().Replace(" ", "").Replace(" ", "");
        if (cleaned.Length == 0)
        {
            result = 0;
            return false;
        }

        return double.TryParse(cleaned, NumberStyles.Float, Ru, out result)
               || double.TryParse(cleaned, NumberStyles.Float, CultureInfo.InvariantCulture, out result);
    }

    private static DateTime? ParseDateString(string s)
    {
        var text = NormalizeText(s);
        if (text.Length == 0)
            return null;
        text = Regex.Replace(text, @"\s*г\.?$", "");
        if (DateTime.TryParseExact(text, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var exact))
            return exact;
        // Дата, записанная числом в текстовой ячейке
        if (TryParseNumber(text, out var n) && n >= 20000 && n <= MaxOaDate)
            return RoundToMinute(DateTime.FromOADate(n));
        return null;
    }

    private static TimeSpan? ParseTimeString(string s)
    {
        var text = NormalizeText(s);
        if (text.Length == 0)
            return null;
        if (DateTime.TryParseExact(text, TimeFormats, CultureInfo.InvariantCulture, DateTimeStyles.AllowWhiteSpaces, out var t))
            return t.TimeOfDay;
        var dt = ParseDateString(text);
        if (dt is not null && text.Contains(':'))
            return dt.Value.TimeOfDay;
        if (TryParseNumber(text, out var n) && n >= 0 && n < 1)
            return TimeSpan.FromMinutes(Math.Round(n * 1440));
        return null;
    }

    private static string NormalizeText(string s) =>
        s.Replace(' ', ' ').Replace('\n', ' ').Replace('\r', ' ').Trim();

    private static DateTime RoundToMinute(DateTime dt)
    {
        var ticks = (long)Math.Round(dt.Ticks / (double)TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute;
        return new DateTime(ticks, dt.Kind);
    }

    // ---------------- Форматирование ----------------

    public static string Number(int value) => value.ToString("#,0", Ru);

    /// <summary>Процент с одним знаком после запятой («95,9», «100», «0»). При нулевом знаменателе — «-».</summary>
    public static string Percent(int part, int total) =>
        total == 0 ? "-" : (part * 100.0 / total).ToString("0.#", Ru);

    public static string Date(DateTime value) => value.ToString("dd.MM.yyyy", Ru);

    public static string Time(DateTime value) => value.ToString("HH:mm", Ru);

    public static string DateTimeText(DateTime value) => value.ToString("dd.MM.yyyy HH:mm", Ru);

    /// <summary>«18:00 14.10.2026» — как в заголовках отчета.</summary>
    public static string TimeDate(DateTime value) => value.ToString("HH:mm dd.MM.yyyy", Ru);

    /// <summary>
    /// Сжимает список номеров этапов: [2,3,4,9] → «2-4, 9». Нечисловые номера выводятся как есть.
    /// </summary>
    public static string CompressNumbers(IEnumerable<string> numbers)
    {
        var parts = new List<string>();
        var list = numbers.Select(n => n.Trim()).Where(n => n.Length > 0).ToList();
        var i = 0;
        while (i < list.Count)
        {
            if (!int.TryParse(list[i], out var start))
            {
                parts.Add(list[i]);
                i++;
                continue;
            }

            var end = start;
            var j = i + 1;
            while (j < list.Count && int.TryParse(list[j], out var next) && next == end + 1)
            {
                end = next;
                j++;
            }

            parts.Add(end > start ? $"{start}-{end}" : start.ToString(CultureInfo.InvariantCulture));
            i = j;
        }

        return string.Join(", ", parts);
    }

    /// <summary>
    /// Подстановка в шаблон: {date:yyyy.MM.dd}, {time:HH.mm}, {number}, {name}, {projects}.
    /// </summary>
    public static string ApplyTemplate(string template, DateTime? moment = null, IReadOnlyDictionary<string, string>? values = null)
    {
        return Regex.Replace(template, @"\{(?<key>[a-zA-Z]+)(:(?<fmt>[^}]+))?\}", m =>
        {
            var key = m.Groups["key"].Value.ToLowerInvariant();
            var fmt = m.Groups["fmt"].Success ? m.Groups["fmt"].Value : null;
            if ((key == "date" || key == "time") && moment is not null)
                return moment.Value.ToString(fmt ?? (key == "date" ? "dd.MM.yyyy" : "HH:mm"), Ru);
            if (values is not null && values.TryGetValue(key, out var v))
                return v;
            return m.Value;
        });
    }

    public static string SafeFileName(string name)
    {
        var invalid = Path.GetInvalidFileNameChars().Concat(new[] { ':', '*', '?', '"', '<', '>', '|', '/', '\\' }).ToHashSet();
        var sb = new StringBuilder(name.Length);
        foreach (var ch in name)
            sb.Append(invalid.Contains(ch) ? '_' : ch);
        return sb.ToString().Trim();
    }
}
