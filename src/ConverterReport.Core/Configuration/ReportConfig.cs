using System.ComponentModel;
using System.Text.Json.Serialization;

namespace ConverterReport.Core.Configuration;

/// <summary>
/// Корневая конфигурация. Хранится в JSON (по умолчанию ConverterReport.json).
/// Количество проектов (этапов) не ограничено — они перечисляются в <see cref="Projects"/>.
/// </summary>
public sealed class ReportConfig
{
    public ReportSettings Report { get; set; } = new();

    public ChartSettings Charts { get; set; } = new();

    public ScreenLayout ScreenLayout { get; set; } = new();

    public UpdateSettings Updates { get; set; } = new();

    public List<ProjectConfig> Projects { get; set; } = new();

    /// <summary>Каталог, относительно которого разрешаются относительные пути (каталог файла конфигурации).</summary>
    [JsonIgnore]
    public string BaseDirectory { get; set; } = Directory.GetCurrentDirectory();

    [JsonIgnore]
    public IReadOnlyList<ProjectConfig> EnabledProjects => Projects.Where(p => p.Enabled).ToList();

    public string ResolvePath(string path)
    {
        var expanded = Environment.ExpandEnvironmentVariables(path.Trim());
        return Path.IsPathRooted(expanded) ? expanded : Path.GetFullPath(Path.Combine(BaseDirectory, expanded));
    }
}

public sealed class ReportSettings
{
    public string Title { get; set; } = "ОТЧЕТ";

    public string DetailedTitle { get; set; } = "ОТЧЕТ ДЕТАЛЬНЫЙ";

    public string Subtitle { get; set; } = "о ходе устранения замечаний ФАУ «Главгосэкспертиза России»";

    /// <summary>Подзаголовок со списком этапов; {projects} заменяется на «2-4, 9».</summary>
    public string ProjectsCaption { get; set; } = "по {projects} этапам строительства ВСМ-1";

    /// <summary>Подзаголовок, если в отчете один этап.</summary>
    public string ProjectsCaptionSingle { get; set; } = "по {projects} этапу строительства ВСМ-1";

    /// <summary>Стандартные времена отчета. Если время не указано при запуске — берется ближайшее к текущему.</summary>
    public List<string> ReportTimes { get; set; } = new() { "08:00", "18:00" };

    /// <summary>Начало отчетного периода: время предыдущего отчета.</summary>
    public string PeriodStartTime { get; set; } = "18:00";

    /// <summary>Сдвиг дня начала отчетного периода относительно даты отчета (-1 = предыдущий день).</summary>
    public int PeriodStartDayOffset { get; set; } = -1;

    public string OutputFolder { get; set; } = "Отчеты";

    /// <summary>Версия формы отчета — подставляется в имена файлов ({version}).</summary>
    public string Version { get; set; } = "1.0";

    /// <summary>
    /// Имена файлов (без расширения). Подстановки: {date:формат}, {time:формат}, {projects} («2-4, 9»),
    /// {projectsShort} («2-4,9»), {version}.
    /// </summary>
    public string FileName { get; set; } = DefaultFileName;

    public string DetailedFileName { get; set; } = DefaultDetailedFileName;

    public string ChartsFileName { get; set; } = DefaultChartsFileName;

    // Имена PDF — как при ручном сохранении: «2026.10.15 8.00 Свод ГГЭ 2-4,9 этапы 1.0 [Отчет|Диаграммы].pdf».
    internal const string DefaultFileName = "{date:yyyy.MM.dd} {time:H.mm} Свод ГГЭ {projectsShort} этапы {version}";
    internal const string DefaultDetailedFileName = DefaultFileName + " Отчет";
    internal const string DefaultChartsFileName = DefaultFileName + " Диаграммы";

    /// <summary>Имена по умолчанию версий 2.0.4–2.0.10 — при загрузке настроек заменяются на текущие.</summary>
    internal static readonly Dictionary<string, string> ObsoleteFileNames = new()
    {
        ["{date:yyyy.MM.dd} {time:H.mm} Суточный {projectsShort} этапы {version}"] = DefaultFileName,
        ["{date:yyyy.MM.dd} {time:H.mm} Суточный {projectsShort} этапы {version} (дет)"] = DefaultDetailedFileName,
        ["{date:yyyy.MM.dd} {time:H.mm} Суточный {projectsShort} этапы {version} Диаграммы"] = DefaultChartsFileName,
    };

    /// <summary>Сводная книга Excel (.xlsx) со всеми листами, формулами и диаграммами.</summary>
    public string WorkbookFileName { get; set; } = "{date:yyyy.MM.dd} {time:H.mm} Сводный отчет ГГЭ {projectsShort} этапы ВСМ-1 {version}";

    /// <summary>
    /// Separate — новая сводная книга на каждый отчет (WorkbookFileName);
    /// Append — лист дня дописывается в общую книгу MasterWorkbook, прежние листы дней сохраняются.
    /// </summary>
    public WorkbookMode WorkbookMode { get; set; } = WorkbookMode.Separate;

    /// <summary>Общая книга для режима Append: путь относительно папки отчетов (или полный), подстановки — как в именах файлов.</summary>
    public string MasterWorkbook { get; set; } = "Сводный отчет ГГЭ ВСМ-1";

    /// <summary>Сколько последних листов дней оставлять видимыми в общей книге; остальные скрываются (0 — показывать все).</summary>
    public int VisibleDailySheets { get; set; } = 2;

    /// <summary>
    /// Подгружать данные из прошлых версий отчета (<see cref="ArchiveWorkbooks"/>): строки экранов этапов (листы Д.N/З.N),
    /// которых нет в текущих экранах, — для «Свода» и диаграмм за прошлые периоды; в общую книгу — листы прошлых дней
    /// «ДД.ММ», которых в ней еще нет. Сроки этапов, не заполненные в настройках, берутся из последнего листа дня.
    /// </summary>
    public bool ImportLegacyDays { get; set; }

    /// <summary>Прошлые версии отчета (старые сводные книги или книги одного этапа с листами «Замечания»/«Документы»).</summary>
    public List<ArchiveSource> ArchiveWorkbooks { get; set; } = new();

    /// <summary>Старая книга версий 2.0.11–2.0.14 — при загрузке настроек переносится в <see cref="ArchiveWorkbooks"/>.</summary>
    [JsonIgnore(Condition = JsonIgnoreCondition.WhenWritingNull)]
    public string? LegacyWorkbook { get; set; }

    /// <summary>
    /// Чем сохранять PDF: Auto — через Excel, если он установлен (вид как при ручном сохранении), иначе встроенным генератором;
    /// Excel — только через Excel; Builtin — только встроенным генератором (Excel не нужен).
    /// </summary>
    public PdfEngine PdfEngine { get; set; } = PdfEngine.Auto;

    /// <summary>Текстовый файл с разделом «Справочно» (необязательно). Пустой файл — раздел не выводится.</summary>
    public string? NotesFile { get; set; } = "Справочно.txt";

    /// <summary>Какие PDF сохранять: суточный отчет, детальный отчет, диаграммы.</summary>
    public bool IncludeDaily { get; set; } = true;

    public bool IncludeDetailed { get; set; } = true;

    public bool IncludeCharts { get; set; } = true;

    public bool IncludeDeadlines { get; set; } = true;

    /// <summary>
    /// Раздел «1. Загрузка документов»: Sum — одна строка на этап с суммой загрузок; Each — каждая загрузка отдельной
    /// строкой (этап | 1, 2, 3… | начало | окончание | документы).
    /// </summary>
    public LoadsMode LoadsMode { get; set; } = LoadsMode.Sum;

    /// <summary>
    /// Какие загрузки выводить в разделе «1. Загрузка документов»: Auto — в вечернем отчете (время отчета не раньше
    /// начала периода, 18:00) все загрузки за день, в утреннем — загрузки с начала отчетного периода (с 18:00
    /// прошлого дня); Day — всегда за день; Period — всегда с начала отчетного периода. На диаграммах загрузки
    /// всегда по дням.
    /// </summary>
    public ReportLoadsMode ReportLoads { get; set; } = ReportLoadsMode.Auto;

    /// <summary>
    /// Если в экране нет строки на отчетное время (взята более ранняя), отмечать номер этапа «*» и выводить под таблицей
    /// примечание «* Этап N: данные экрана на …». false — без отметки и примечания (предупреждение остается в журнале).
    /// </summary>
    public bool MarkStaleData { get; set; } = true;

    /// <summary>Не выводить в таблице загрузок этапы, по которым за период загрузок не было.</summary>
    public bool HideProjectsWithoutLoads { get; set; } = true;

    /// <summary>Колонки раздела «Сроки» (порядок вывода).</summary>
    public List<string> DeadlineStages { get; set; } = new() { "Входной контроль", "Экспертиза", "Продление 1", "Продление 2" };

    public string FontFamily { get; set; } = "Times New Roman";
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum PdfEngine
{
    Auto,
    Excel,
    Builtin,
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum WorkbookMode
{
    Separate,
    Append,
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum LoadsMode
{
    Sum,
    Each,
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum ReportLoadsMode
{
    Auto,
    Day,
    Period,
}

public sealed class ChartSettings
{
    /// <summary>Формат листов с диаграммами: A4 или A3 (альбомная ориентация).</summary>
    public string PageSize { get; set; } = "A3";

    /// <summary>Сводные диаграммы по всем этапам (выводятся, если этапов больше одного).</summary>
    public bool SummaryChart { get; set; } = true;

    public string RemarksTitle { get; set; } = "Устранение замечаний ГГЭ - {name} ВСМ-1";

    public string LoadsTitle { get; set; } = "Загрузка документов в АИС ГГЭ - {name} ВСМ-1";

    /// <summary>Имя для сводных диаграмм; {projects} заменяется на «2-4, 9».</summary>
    public string SummaryName { get; set; } = "{projects} этапы";

    /// <summary>Общая дата начала диаграмм (дд.мм.гггг). Если не задана — с первой заполненной строки экрана после начала сроков этапа (входного контроля), а если сроков нет — с первой заполненной строки.</summary>
    public string? StartDate { get; set; }

    /// <summary>Продлевать ось дат до окончания экспертизы (видно, сколько осталось).</summary>
    public bool ExtendToExpertiseEnd { get; set; } = true;

    /// <summary>Линия тренда (прогноз) по устраненным замечаниям.</summary>
    public bool TrendLine { get; set; } = true;

    public DataLabelsMode DataLabels { get; set; } = DataLabelsMode.Auto;

    /// <summary>
    /// Верх шкалы диаграммы загрузок документов: сводной и по этапам (0 — автоматически). Для этапа можно задать свой
    /// (<see cref="ProjectConfig.LoadsAxisMax"/>). Колонки выше предела обрезаются, значение пишется сбоку у верхнего края.
    /// </summary>
    public int LoadsAxisMaxSummary { get; set; } = 500;

    public int LoadsAxisMaxProject { get; set; } = 250;

    public string FontFamily { get; set; } = "Arial";
}

[JsonConverter(typeof(JsonStringEnumConverter))]
public enum DataLabelsMode
{
    /// <summary>Подписи у всех точек, если хватает места, иначе — только на отчетную дату.</summary>
    Auto,
    All,
    Last,
    None,
}

/// <summary>Прошлая версия отчета для архива.</summary>
public sealed class ArchiveSource
{
    /// <summary>Файл (путь относительно файла настроек или полный).</summary>
    public string Path { get; set; } = "";

    /// <summary>
    /// Данные каких этапов брать: ключ — номер листов Д.N/З.N («5» — листы Д.5/З.5) или «Экран» для книги одного этапа
    /// (листы «Замечания»/«Документы»), значение — номер этапа («9»). Переносятся только листы из этого списка
    /// (выбираются в окне при добавлении книги); «-» или пусто — не переносить. Листы дней переносятся всегда.
    /// </summary>
    public Dictionary<string, string> Stages { get; set; } = new();

    /// <summary>Номер этапа для листов <paramref name="key"/> или null, если их данные не переносятся.</summary>
    public string? StageFor(string key) =>
        Stages.FirstOrDefault(kv => string.Equals(kv.Key.Trim(), key.Trim(), StringComparison.OrdinalIgnoreCase)).Value is { } v
        && v.Trim() is { Length: > 0 } stage && stage != "-"
            ? stage
            : null;
}

public sealed class ProjectConfig
{
    /// <summary>Номер этапа (как выводится в столбце «Этап»).</summary>
    public string Number { get; set; } = "";

    /// <summary>Наименование для заголовков диаграмм и имени файла. По умолчанию «{Number} этап».</summary>
    public string? Name { get; set; }

    /// <summary>
    /// Путь к файлу «Экран» этапа (листы «Замечания» и «Документы»). Пусто — архивный этап: данные только из прошлых
    /// версий отчета и общей книги.
    /// </summary>
    public string ScreenFile { get; set; } = "";

    /// <summary>Этап учитывается в листе «Свод» и сводной диаграмме (экран читается).</summary>
    public bool Enabled { get; set; } = true;

    /// <summary>Выводить этап в листах отчета («ДД.ММ», «Отчет_дет», «Отчет») и отдельной страницей диаграмм.</summary>
    public bool ShowInReport { get; set; } = true;

    /// <summary>Верх шкалы диаграммы загрузок этапа; пусто — общий (<see cref="ChartSettings.LoadsAxisMaxProject"/>).</summary>
    public int? LoadsAxisMax { get; set; }

    /// <summary>Дата начала диаграмм для этапа (дд.мм.гггг), перекрывает общую.</summary>
    public string? ChartStartDate { get; set; }

    /// <summary>Дата окончания экспертизы. Если не задана — наибольшая дата из <see cref="Deadlines"/>.</summary>
    public string? ExpertiseEndDate { get; set; }

    /// <summary>Сроки: стадия → [начало, окончание] или [произвольный текст].</summary>
    public Dictionary<string, List<string>> Deadlines { get; set; } = new();

    [JsonIgnore]
    public string DisplayName => string.IsNullOrWhiteSpace(Name) ? $"{Number} этап" : Name!.Trim();

    /// <summary>Архивный этап — без файла экрана.</summary>
    [JsonIgnore]
    public bool IsArchive => string.IsNullOrWhiteSpace(ScreenFile);
}

/// <summary>
/// Расположение данных в файле «Экран». Значения по умолчанию соответствуют текущему шаблону экрана.
/// Для блоков замечаний указывается первая колонка группы из четырех: Всего, Устранено, Устраняется, Рассматривается.
/// </summary>
public sealed class ScreenLayout
{
    public override string ToString() => $"{RemarksSheet} / {DocumentsSheet}";

    [DisplayName("Замечания: лист")]
    public string RemarksSheet { get; set; } = "Замечания";
    [DisplayName("Замечания: первая строка данных")]
    public int RemarksFirstRow { get; set; } = 3;
    [DisplayName("Замечания: дата")]
    public string RemarksDate { get; set; } = "B";
    [DisplayName("Замечания: время")]
    public string RemarksTime { get; set; } = "D";
    [DisplayName("Замечания: всего (4 столбца)")]
    public string RemarksAll { get; set; } = "E";
    [DisplayName("Замечания: ИИ (4 столбца)")]
    public string RemarksSurvey { get; set; } = "I";
    [DisplayName("Замечания: тех. часть (4 столбца)")]
    public string RemarksTech { get; set; } = "M";
    [DisplayName("Замечания: сметная часть (4 столбца)")]
    public string RemarksEstimate { get; set; } = "Q";
    [DisplayName("Критические: всего (4 столбца)")]
    public string CriticalAll { get; set; } = "U";
    [DisplayName("Критические: ИИ (4 столбца)")]
    public string CriticalSurvey { get; set; } = "Y";
    [DisplayName("Критические: тех. часть (4 столбца)")]
    public string CriticalTech { get; set; } = "AC";
    [DisplayName("Критические: сметная часть (4 столбца)")]
    public string CriticalEstimate { get; set; } = "AG";

    [DisplayName("Документы: лист")]
    public string DocumentsSheet { get; set; } = "Документы";
    [DisplayName("Документы: первая строка данных")]
    public int DocumentsFirstRow { get; set; } = 3;
    [DisplayName("Документы: начало, дата")]
    public string LoadStartDate { get; set; } = "H";
    [DisplayName("Документы: начало, время")]
    public string LoadStartTime { get; set; } = "I";
    [DisplayName("Документы: окончание, дата")]
    public string LoadEndDate { get; set; } = "J";
    [DisplayName("Документы: окончание, время")]
    public string LoadEndTime { get; set; } = "K";
    [DisplayName("Документы: ИИ")]
    public string LoadSurvey { get; set; } = "L";
    [DisplayName("Документы: ПД")]
    public string LoadProject { get; set; } = "M";
    [DisplayName("Документы: СМ")]
    public string LoadEstimate { get; set; } = "N";
    [DisplayName("Документы: ВОР")]
    public string LoadVor { get; set; } = "O";
    [DisplayName("Документы: КАЦ")]
    public string LoadKac { get; set; } = "P";
    [DisplayName("Документы: иные")]
    public string LoadOther { get; set; } = "Q";
    [DisplayName("Документы: ответы")]
    public string LoadAnswers { get; set; } = "S";
    [DisplayName("Документы: аннулировано")]
    public string LoadCancelled { get; set; } = "T";
    [DisplayName("Документы: ошибки")]
    public string LoadErrors { get; set; } = "U";
    [DisplayName("Документы: примечание")]
    public string LoadNote { get; set; } = "V";
}

/// <summary>Автообновление программы из релизов GitHub.</summary>
public sealed class UpdateSettings
{
    /// <summary>Проверять наличие новой версии при запуске программы.</summary>
    public bool CheckOnStartup { get; set; } = true;

    /// <summary>Репозиторий GitHub с релизами программы (владелец/репозиторий).</summary>
    public string Repository { get; set; } = "zaynullinmi/ConverterReport";
}
