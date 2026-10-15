using ConverterReport.Core.Configuration;

namespace ConverterReport.Core.Reporting;

/// <summary>Определение отчетных даты/времени и начала периода — чтобы не править время отчета вручную.</summary>
public static class ReportTime
{
    /// <summary>
    /// Ближайшее к <paramref name="now"/> стандартное время отчета (среди вчерашних, сегодняшних и завтрашних).
    /// Например, при запуске в 07:50 будет выбран отчет на 08:00, в 17:40 — на 18:00.
    /// </summary>
    public static DateTime Nearest(DateTime now, IEnumerable<TimeSpan> slots)
    {
        var list = slots.ToList();
        if (list.Count == 0)
            return new DateTime(now.Year, now.Month, now.Day, now.Hour, now.Minute, 0);

        return Enumerable.Range(-1, 3)
            .SelectMany(offset => list.Select(slot => now.Date.AddDays(offset) + slot))
            .OrderBy(candidate => Math.Abs((candidate - now).Ticks))
            .ThenByDescending(candidate => candidate)
            .First();
    }

    /// <summary>Отчетное время с учетом явно заданных даты/времени запуска.</summary>
    public static DateTime Resolve(ReportSettings settings, DateTime now, DateTime? date, TimeSpan? time)
    {
        var slots = settings.ReportTimes.Select(t => Formats.ParseTime(t)).Where(t => t is not null).Select(t => t!.Value).ToList();
        if (date is not null && time is not null)
            return date.Value.Date + time.Value;
        if (time is not null)
            return now.Date + time.Value;
        if (date is not null)
        {
            // Дата задана, время нет: последнее стандартное время этого дня (или 18:00).
            var slot = slots.Count > 0 ? slots.Max() : new TimeSpan(18, 0, 0);
            return date.Value.Date + slot;
        }

        return Nearest(now, slots);
    }

    /// <summary>Начало отчетного периода: время предыдущего отчета (по умолчанию 18:00 предыдущего дня).</summary>
    public static DateTime PeriodStart(ReportSettings settings, DateTime reportTime)
    {
        var time = Formats.ParseTime(settings.PeriodStartTime) ?? new TimeSpan(18, 0, 0);
        var start = reportTime.Date.AddDays(settings.PeriodStartDayOffset) + time;
        return start < reportTime ? start : reportTime.Date.AddDays(-1) + time;
    }
}
