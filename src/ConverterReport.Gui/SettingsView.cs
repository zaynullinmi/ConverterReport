using System.ComponentModel;
using System.Globalization;
using ConverterReport.Core;
using ConverterReport.Core.Configuration;

namespace ConverterReport.Gui;

/// <summary>
/// Настройки для таблицы свойств на вкладке «Настройки»: русские названия, группы и подсказки.
/// Значения читаются из <see cref="ReportConfig"/> и записываются обратно методом <see cref="ApplyTo"/>.
/// </summary>
internal sealed class SettingsView
{
    private const string Titles = "1. Заголовки";
    private const string Times = "2. Время отчета";
    private const string Files = "3. Файлы";
    private const string Daily = "4. Суточный отчет";
    private const string Charts = "5. Диаграммы";
    private const string Screen = "6. Раскладка экрана";
    private const string Updates = "7. Обновление программы";

    public SettingsView(ReportConfig c)
    {
        var r = c.Report;
        Title = r.Title;
        DetailedTitle = r.DetailedTitle;
        Subtitle = r.Subtitle;
        ProjectsCaption = r.ProjectsCaption;
        ProjectsCaptionSingle = r.ProjectsCaptionSingle;
        ReportTimes = string.Join("; ", r.ReportTimes);
        PeriodStartTime = r.PeriodStartTime;
        PeriodStartDayOffset = r.PeriodStartDayOffset;
        Version = r.Version;
        FileName = r.FileName;
        DetailedFileName = r.DetailedFileName;
        ChartsFileName = r.ChartsFileName;
        WorkbookFileName = r.WorkbookFileName;
        NotesFile = r.NotesFile ?? "";
        VisibleDailySheets = r.VisibleDailySheets;
        IncludeDeadlines = r.IncludeDeadlines;
        HideProjectsWithoutLoads = r.HideProjectsWithoutLoads;
        MarkStaleData = r.MarkStaleData;
        ReportLoads = ReportLoadsConverter.ToText(r.ReportLoads);
        DeadlineStages = string.Join("; ", r.DeadlineStages);
        LoadsMode = r.LoadsMode == Core.Configuration.LoadsMode.Each ? LoadsModeConverter.Each : LoadsModeConverter.Sum;
        FontFamily = r.FontFamily;

        var ch = c.Charts;
        PageSize = string.Equals(ch.PageSize?.Trim(), "A4", StringComparison.OrdinalIgnoreCase) ? "A4" : "A3";
        SummaryChart = ch.SummaryChart;
        ChartsStartDate = ch.StartDate ?? "";
        ExtendToExpertiseEnd = ch.ExtendToExpertiseEnd;
        TrendLine = ch.TrendLine;
        RemarksTitle = ch.RemarksTitle;
        LoadsTitle = ch.LoadsTitle;
        SummaryName = ch.SummaryName;
        LoadsAxisMaxSummary = ch.LoadsAxisMaxSummary;
        LoadsAxisMaxProject = ch.LoadsAxisMaxProject;

        Layout = c.ScreenLayout;
        CheckOnStartup = c.Updates.CheckOnStartup;
        Repository = c.Updates.Repository;
    }

    // ------------------------------------------------------------------ Заголовки

    [Category(Titles), DisplayName("Заголовок суточного отчета")]
    public string Title { get; set; }

    [Category(Titles), DisplayName("Заголовок детального отчета")]
    public string DetailedTitle { get; set; }

    [Category(Titles), DisplayName("Подзаголовок")]
    public string Subtitle { get; set; }

    [Category(Titles), DisplayName("Строка с этапами")]
    [Description("{projects} заменяется на список этапов, например «2-4, 9». Можно написать текст целиком.")]
    public string ProjectsCaption { get; set; }

    [Category(Titles), DisplayName("Строка с этапами (один этап)")]
    public string ProjectsCaptionSingle { get; set; }

    // ------------------------------------------------------------------ Время

    [Category(Times), DisplayName("Стандартные времена отчета")]
    [Description("Через «;», например «08:00; 18:00». При запуске выбирается ближайшее время.")]
    public string ReportTimes { get; set; }

    [Category(Times), DisplayName("Начало периода: время")]
    [Description("Время предыдущего отчета, с которого считается период (по умолчанию 18:00).")]
    public string PeriodStartTime { get; set; }

    [Category(Times), DisplayName("Начало периода: сдвиг в днях")]
    [Description("-1 — предыдущий день относительно даты отчета.")]
    public int PeriodStartDayOffset { get; set; }

    // ------------------------------------------------------------------ Файлы

    [Category(Files), DisplayName("Версия формы")]
    [Description("Подставляется в имена файлов вместо {version}.")]
    public string Version { get; set; }

    [Category(Files), DisplayName("Имя PDF: суточный отчет")]
    [Description("{date:yyyy.MM.dd}, {time:H.mm}, {projects} («2-4, 9»), {projectsShort} («2-4,9»), {version}.")]
    public string FileName { get; set; }

    [Category(Files), DisplayName("Имя PDF: детальный отчет")]
    public string DetailedFileName { get; set; }

    [Category(Files), DisplayName("Имя PDF: диаграммы")]
    public string ChartsFileName { get; set; }

    [Category(Files), DisplayName("Имя сводной книги Excel")]
    public string WorkbookFileName { get; set; }

    [Category(Files), DisplayName("Общая книга: видимых листов дней")]
    [Description("Режим «Дописывать лист дня в общую книгу»: сколько последних листов дней оставлять видимыми, остальные скрываются (0 — показывать все).")]
    public int VisibleDailySheets { get; set; }

    [Category(Files), DisplayName("Файл с текстом «Справочно»")]
    [Description("Текст редактируется на вкладке «Отчет» и сохраняется в этот файл.")]
    public string NotesFile { get; set; }

    // ------------------------------------------------------------------ Суточный отчет

    [Category(Daily), DisplayName("Раздел «4. Сроки»"), TypeConverter(typeof(YesNoConverter))]
    public bool IncludeDeadlines { get; set; }

    [Category(Daily), DisplayName("Загрузки в разделе 1"), TypeConverter(typeof(ReportLoadsConverter))]
    [Description("Какие загрузки выводить в разделе «1. Загрузка документов». Авто: в вечернем отчете (время отчета не раньше " +
                 "начала периода, 18:00) — все загрузки за день, как на диаграммах; в утреннем — только загрузки с начала периода " +
                 "(с 18:00 прошлого дня). На диаграммах загрузки всегда за день.")]
    public string ReportLoads { get; set; }

    [Category(Daily), DisplayName("Отмечать «*» этапы без строки на отчетное время"), TypeConverter(typeof(YesNoConverter))]
    [Description("Если в экране нет строки на отчетное время (взята более ранняя): «Да» — у номера этапа «*» и примечание под таблицей " +
                 "«* Этап N: данные экрана на …»; «Нет» — без отметки и примечания (предупреждение остается в журнале).")]
    public bool MarkStaleData { get; set; }

    [Category(Daily), DisplayName("Скрывать этапы без загрузок"), TypeConverter(typeof(YesNoConverter))]
    [Description("В разделе «1. Загрузка документов» не показывать этапы, по которым загрузок не было.")]
    public bool HideProjectsWithoutLoads { get; set; }

    [Category(Daily), DisplayName("Стадии сроков")]
    [Description("Колонки раздела «4. Сроки», через «;».")]
    public string DeadlineStages { get; set; }

    [Category(Daily), DisplayName("Загрузки в разделе 1: строки"), TypeConverter(typeof(LoadsModeConverter))]
    [Description("Сумма загрузок по этапу (одна строка: этап | кол-во загрузок | …) или каждая загрузка отдельной строкой " +
                 "(этап | 1, 2, 3… | начало | окончание | документы).")]
    public string LoadsMode { get; set; }

    [Category(Daily), DisplayName("Шрифт")]
    public string FontFamily { get; set; }

    // ------------------------------------------------------------------ Диаграммы

    [Category(Charts), DisplayName("Формат листа"), TypeConverter(typeof(PaperConverter))]
    public string PageSize { get; set; }

    [Category(Charts), DisplayName("Сводная диаграмма по всем этапам"), TypeConverter(typeof(YesNoConverter))]
    public bool SummaryChart { get; set; }

    [Category(Charts), DisplayName("Начало диаграмм (дд.мм.гггг)")]
    [Description("Пусто — с первой заполненной строки экрана после начала сроков этапа.")]
    public string ChartsStartDate { get; set; }

    [Category(Charts), DisplayName("Ось до окончания экспертизы"), TypeConverter(typeof(YesNoConverter))]
    public bool ExtendToExpertiseEnd { get; set; }

    [Category(Charts), DisplayName("Линии тренда"), TypeConverter(typeof(YesNoConverter))]
    public bool TrendLine { get; set; }

    [Category(Charts), DisplayName("Заголовок: устранение замечаний")]
    [Description("{name} — имя этапа, например «3 этап».")]
    public string RemarksTitle { get; set; }

    [Category(Charts), DisplayName("Заголовок: загрузка документов")]
    public string LoadsTitle { get; set; }

    [Category(Charts), DisplayName("Верх шкалы загрузок: сводная")]
    [Description("Верх шкалы диаграммы загрузок документов на сводной странице (0 — автоматически). Колонки выше обрезаются, значение пишется сбоку.")]
    public int LoadsAxisMaxSummary { get; set; }

    [Category(Charts), DisplayName("Верх шкалы загрузок: этапы")]
    [Description("Для страниц этапов (0 — автоматически). Для отдельного этапа можно задать свой на вкладке «Этапы».")]
    public int LoadsAxisMaxProject { get; set; }

    [Category(Charts), DisplayName("Имя сводной диаграммы")]
    [Description("{projects} — список этапов, например «2-4,9».")]
    public string SummaryName { get; set; }

    // ------------------------------------------------------------------ Раскладка экрана

    [Category(Screen), DisplayName("Столбцы листов экрана")]
    [Description("Где в файле «Экран» лежат данные. Меняется, только если изменится шаблон экрана.")]
    [TypeConverter(typeof(ExpandableObjectConverter))]
    public ScreenLayout Layout { get; }

    // ------------------------------------------------------------------ Обновление

    [Category(Updates), DisplayName("Проверять обновления при запуске"), TypeConverter(typeof(YesNoConverter))]
    public bool CheckOnStartup { get; set; }

    [Category(Updates), DisplayName("Репозиторий GitHub")]
    [Description("Откуда берутся новые версии программы (владелец/репозиторий).")]
    public string Repository { get; set; }

    public void ApplyTo(ReportConfig c)
    {
        var r = c.Report;
        r.Title = Title;
        r.DetailedTitle = DetailedTitle;
        r.Subtitle = Subtitle;
        r.ProjectsCaption = ProjectsCaption;
        r.ProjectsCaptionSingle = ProjectsCaptionSingle;
        r.ReportTimes = SplitList(ReportTimes);
        r.PeriodStartTime = PeriodStartTime;
        r.PeriodStartDayOffset = PeriodStartDayOffset;
        r.Version = Version;
        r.FileName = FileName;
        r.DetailedFileName = DetailedFileName;
        r.ChartsFileName = ChartsFileName;
        r.WorkbookFileName = WorkbookFileName;
        r.NotesFile = string.IsNullOrWhiteSpace(NotesFile) ? null : NotesFile.Trim();
        r.VisibleDailySheets = Math.Max(0, VisibleDailySheets);
        r.IncludeDeadlines = IncludeDeadlines;
        r.HideProjectsWithoutLoads = HideProjectsWithoutLoads;
        r.MarkStaleData = MarkStaleData;
        r.ReportLoads = ReportLoadsConverter.FromText(ReportLoads);
        r.DeadlineStages = SplitList(DeadlineStages);
        r.LoadsMode = LoadsMode == LoadsModeConverter.Each ? Core.Configuration.LoadsMode.Each : Core.Configuration.LoadsMode.Sum;
        r.FontFamily = FontFamily;

        var ch = c.Charts;
        ch.PageSize = PageSize;
        ch.SummaryChart = SummaryChart;
        ch.StartDate = string.IsNullOrWhiteSpace(ChartsStartDate) ? null : ChartsStartDate.Trim();
        ch.ExtendToExpertiseEnd = ExtendToExpertiseEnd;
        ch.TrendLine = TrendLine;
        ch.RemarksTitle = RemarksTitle;
        ch.LoadsTitle = LoadsTitle;
        ch.SummaryName = SummaryName;
        ch.LoadsAxisMaxSummary = Math.Max(0, LoadsAxisMaxSummary);
        ch.LoadsAxisMaxProject = Math.Max(0, LoadsAxisMaxProject);

        c.ScreenLayout = Layout;
        c.Updates.CheckOnStartup = CheckOnStartup;
        c.Updates.Repository = Repository;
    }

    private static List<string> SplitList(string text) =>
        text.Split(new[] { ';', ',', '\n' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
}

/// <summary>«Да»/«Нет» вместо True/False.</summary>
internal sealed class YesNoConverter : BooleanConverter
{
    public override object? ConvertTo(ITypeDescriptorContext? context, CultureInfo? culture, object? value, Type destinationType) =>
        destinationType == typeof(string) && value is bool b ? (b ? "Да" : "Нет") : base.ConvertTo(context, culture, value, destinationType);

    public override object? ConvertFrom(ITypeDescriptorContext? context, CultureInfo? culture, object value) =>
        value is string s ? s.Trim().Equals("Да", StringComparison.OrdinalIgnoreCase) || s.Trim().Equals("True", StringComparison.OrdinalIgnoreCase)
            : base.ConvertFrom(context, culture, value);
}

/// <summary>Выбор из фиксированного списка строк.</summary>
internal abstract class ChoiceConverter : StringConverter
{
    protected abstract string[] Choices { get; }

    public override bool GetStandardValuesSupported(ITypeDescriptorContext? context) => true;

    public override bool GetStandardValuesExclusive(ITypeDescriptorContext? context) => true;

    public override StandardValuesCollection GetStandardValues(ITypeDescriptorContext? context) => new(Choices);
}

internal sealed class LoadsModeConverter : ChoiceConverter
{
    public const string Sum = "Сумма по этапу";
    public const string Each = "Каждая загрузка отдельной строкой";

    protected override string[] Choices => new[] { Sum, Each };
}

internal sealed class ReportLoadsConverter : ChoiceConverter
{
    private const string Auto = "Утром — с начала периода, вечером — за день";
    private const string Day = "Всегда за день (с 00:00)";
    private const string Period = "Всегда с начала периода";

    protected override string[] Choices => new[] { Auto, Day, Period };

    public static string ToText(ReportLoadsMode mode) => mode switch
    {
        ReportLoadsMode.Day => Day,
        ReportLoadsMode.Period => Period,
        _ => Auto,
    };

    public static ReportLoadsMode FromText(string text) => text switch
    {
        Day => ReportLoadsMode.Day,
        Period => ReportLoadsMode.Period,
        _ => ReportLoadsMode.Auto,
    };
}

internal sealed class PaperConverter : ChoiceConverter
{
    protected override string[] Choices => new[] { "A3", "A4" };
}
