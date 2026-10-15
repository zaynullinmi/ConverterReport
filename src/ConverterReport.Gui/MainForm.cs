using System.Diagnostics;
using ConverterReport.Core;
using ConverterReport.Core.Configuration;
using ConverterReport.Core.Reporting;
using ConverterReport.Core.Updates;

namespace ConverterReport.Gui;

/// <summary>
/// Главное окно: вкладки «Отчет» (запуск), «Этапы» (список этапов и сроки) и «Настройки» (все параметры).
/// Настройки хранятся в ConverterReport.json рядом с программой.
/// </summary>
internal sealed class MainForm : Form
{
    // Колонки таблицы этапов.
    private const int ColEnabled = 0;
    private const int ColNumber = 1;
    private const int ColName = 2;
    private const int ColFile = 3;
    private const int ColBrowse = 4;
    private const int ColEnd = 5;
    private const int ColChartStart = 6;

    private readonly CommandLineOptions _options;
    private string _configPath;
    private ReportConfig _config;
    private SettingsView _settings;
    private bool _dirty;
    private bool _loading;
    private bool _running;
    private bool _closingForUpdate;
    private ProjectConfig? _deadlinesProject;

    // «Отчет»
    private DateTimePicker _date = null!;
    private ComboBox _time = null!;
    private CheckBox _customStart = null!;
    private DateTimePicker _start = null!;
    private Label _period = null!;
    private TextBox _outFolder = null!;
    private CheckBox _daily = null!;
    private CheckBox _detailed = null!;
    private CheckBox _charts = null!;
    private RadioButton _engineAuto = null!;
    private RadioButton _engineExcel = null!;
    private RadioButton _engineBuiltin = null!;
    private TextBox _notes = null!;
    private Button _run = null!;
    private ProgressBar _progress = null!;
    private Label _runState = null!;
    private RichTextBox _log = null!;
    private Button _openDaily = null!;
    private Button _openDetailed = null!;
    private Button _openCharts = null!;
    private Button _openBook = null!;
    private Button _openFolder = null!;
    private RunResult? _last;
    private string _notesLoaded = "";

    // «Этапы»
    private DataGridView _projects = null!;
    private DataGridView _deadlines = null!;
    private GroupBox _deadlinesBox = null!;

    // «Настройки»
    private PropertyGrid _grid = null!;
    private Label _configLabel = null!;

    // Строка состояния
    private ToolStripStatusLabel _status = null!;
    private ToolStripProgressBar _updateProgress = null!;
    private ToolStripMenuItem _openConfigItem = null!;

    public MainForm(CommandLineOptions options)
    {
        _options = options;
        _configPath = Path.GetFullPath(options.ConfigPath ?? AppPaths.DefaultConfigPath);
        _config = LoadOrCreate(_configPath);
        _settings = new SettingsView(_config);

        BuildUi();
        LoadToUi();
        ApplyCommandLine();

        Shown += async (_, _) =>
        {
            if (_deadlinesProject is null)
                ShowDeadlines(SelectedProject());
            if (_options.Updated)
                SetStatus($"Программа обновлена до версии {Updater.CurrentVersion}.");
            else if (_config.Updates.CheckOnStartup)
                await CheckUpdatesAsync(silent: true);
        };
        FormClosing += OnFormClosing;
    }

    // ================================================================== Построение окна

    private void BuildUi()
    {
        SuspendLayout();
        AutoScaleDimensions = new SizeF(96F, 96F);
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Segoe UI", 9F);
        Text = $"ConverterReport {Updater.CurrentVersion} — суточный отчет ГГЭ";
        ClientSize = new Size(1060, 780);
        MinimumSize = new Size(900, 640);
        StartPosition = FormStartPosition.CenterScreen;

        var tabs = new TabControl { Dock = DockStyle.Fill };
        tabs.TabPages.Add(BuildReportTab());
        tabs.TabPages.Add(BuildProjectsTab());
        tabs.TabPages.Add(BuildSettingsTab());

        var status = new StatusStrip();
        _status = new ToolStripStatusLabel { Spring = true, TextAlign = ContentAlignment.MiddleLeft };
        _updateProgress = new ToolStripProgressBar { Visible = false, Size = new Size(160, 16) };
        status.Items.Add(_status);
        status.Items.Add(_updateProgress);

        var menu = BuildMenu();
        Controls.Add(tabs);
        Controls.Add(status);
        Controls.Add(menu);
        MainMenuStrip = menu;
        ResumeLayout(false);
        PerformLayout();
    }

    private MenuStrip BuildMenu()
    {
        var menu = new MenuStrip();
        var file = new ToolStripMenuItem("Файл");
        file.DropDownItems.Add(new ToolStripMenuItem("Сохранить настройки", null, (_, _) => SaveConfig(showStatus: true)) { ShortcutKeys = Keys.Control | Keys.S });
        _openConfigItem = new ToolStripMenuItem("Открыть файл настроек…", null, (_, _) => OpenConfig());
        file.DropDownItems.Add(_openConfigItem);
        file.DropDownItems.Add(new ToolStripMenuItem("Сохранить настройки как…", null, (_, _) => SaveConfigAs()));
        file.DropDownItems.Add(new ToolStripMenuItem("Открыть папку с настройками", null, (_, _) => OpenPath(Path.GetDirectoryName(_configPath)!)));
        file.DropDownItems.Add(new ToolStripSeparator());
        file.DropDownItems.Add(new ToolStripMenuItem("Выход", null, (_, _) => Close()));

        var help = new ToolStripMenuItem("Справка");
        help.DropDownItems.Add(new ToolStripMenuItem("Проверить обновления", null, async (_, _) => await CheckUpdatesAsync(silent: false)));
        help.DropDownItems.Add(new ToolStripMenuItem("Страница программы на GitHub", null, (_, _) => OpenPath($"https://github.com/{_config.Updates.Repository}")));
        help.DropDownItems.Add(new ToolStripMenuItem("Запуск по расписанию (без окна)", null, (_, _) =>
            TextDialog.Show(this, "Параметры запуска", CommandLineOptions.Help)));
        help.DropDownItems.Add(new ToolStripSeparator());
        help.DropDownItems.Add(new ToolStripMenuItem("О программе", null, (_, _) => MessageBox.Show(this,
            $"ConverterReport {Updater.CurrentVersion}\n\nСуточный отчет ГГЭ по данным файлов «Экран»: сводная книга Excel и три PDF " +
            "(суточный отчет, детальный отчет, диаграммы).\n\nФайл настроек:\n" + _configPath,
            "О программе", MessageBoxButtons.OK, MessageBoxIcon.Information)));

        menu.Items.Add(file);
        menu.Items.Add(help);
        return menu;
    }

    private TabPage BuildReportTab()
    {
        var page = new TabPage("Отчет") { Padding = new Padding(8), UseVisualStyleBackColor = true };
        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 5 };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Absolute, 120));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        // --- Параметры: дата/время, папка | что формировать, чем сохранять PDF.
        var parameters = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, AutoSize = true };
        parameters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 55));
        parameters.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 45));

        _date = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd.MM.yyyy", Width = 110 };
        _time = new ComboBox { DropDownStyle = ComboBoxStyle.DropDown, Width = 75 };
        var now = new Button { Text = "Ближайшее", AutoSize = true };
        now.Click += (_, _) => SetNearestTime();
        _customStart = new CheckBox { Text = "Начало периода вручную:", AutoSize = true, Margin = new Padding(3, 6, 3, 3) };
        _start = new DateTimePicker { Format = DateTimePickerFormat.Custom, CustomFormat = "dd.MM.yyyy HH:mm", Width = 140, Enabled = false };
        _period = new Label { AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(3, 6, 3, 3) };
        _date.ValueChanged += (_, _) => UpdatePeriod();
        _time.TextChanged += (_, _) => UpdatePeriod();
        _start.ValueChanged += (_, _) => UpdatePeriod();
        _customStart.CheckedChanged += (_, _) =>
        {
            _start.Enabled = _customStart.Checked;
            UpdatePeriod();
        };
        var dateBox = Group("Отчетная дата и время", Stack(
            Row(Caption("Дата:"), _date, Caption("Время:"), _time, now),
            Row(_customStart, _start),
            _period));

        _outFolder = new TextBox { Width = 380 };
        _outFolder.TextChanged += (_, _) => MarkDirty();
        var browse = new Button { Text = "Обзор…", AutoSize = true };
        browse.Click += (_, _) => BrowseOutputFolder();
        var openOut = new Button { Text = "Открыть", AutoSize = true };
        openOut.Click += (_, _) => OpenPath(OutputFolder());
        var folderBox = Group("Папка для сводной книги и PDF", Stack(
            Row(_outFolder, browse, openOut),
            Hint("Относительный путь считается от папки с файлом настроек.")));

        _daily = new CheckBox { Text = "Суточный отчет (PDF, A4)", AutoSize = true };
        _detailed = new CheckBox { Text = "Детальный отчет (PDF, A4)", AutoSize = true };
        _charts = new CheckBox { Text = "Диаграммы (PDF, A3)", AutoSize = true };
        foreach (var c in new[] { _daily, _detailed, _charts })
            c.CheckedChanged += (_, _) => MarkDirty();
        var whatBox = Group("Что сформировать", Stack(_daily, _detailed, _charts, Hint("Сводная книга Excel формируется всегда.")));

        _engineAuto = new RadioButton { Text = "Через Excel, если он установлен (рекомендуется)", AutoSize = true };
        _engineExcel = new RadioButton { Text = "Только через Excel", AutoSize = true };
        _engineBuiltin = new RadioButton { Text = "Встроенным генератором (без Excel)", AutoSize = true };
        foreach (var r in new[] { _engineAuto, _engineExcel, _engineBuiltin })
            r.CheckedChanged += (_, _) => MarkDirty();
        var engineBox = Group("Как сохранять PDF", Stack(_engineAuto, _engineExcel, _engineBuiltin));

        parameters.Controls.Add(dateBox, 0, 0);
        parameters.Controls.Add(folderBox, 0, 1);
        parameters.Controls.Add(whatBox, 1, 0);
        parameters.Controls.Add(engineBox, 1, 1);

        // --- Справочно.
        _notes = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, AcceptsReturn = true, Dock = DockStyle.Fill, WordWrap = true };
        _notes.TextChanged += (_, _) => MarkDirty();
        var notesBox = Group("Справочно — текст под суточным отчетом (каждая строка — отдельный абзац)", _notes);
        notesBox.AutoSize = false;

        // --- Запуск.
        _run = new Button { Text = "Сформировать отчет", AutoSize = true, Font = new Font(Font.FontFamily, 10F, FontStyle.Bold), Padding = new Padding(14, 4, 14, 4) };
        _run.Click += (_, _) => RunReport();
        _progress = new ProgressBar { Style = ProgressBarStyle.Marquee, Width = 220, Visible = false, Margin = new Padding(12, 12, 3, 3) };
        _runState = new Label { AutoSize = true, Margin = new Padding(8, 12, 3, 3) };
        var runRow = Row(_run, _progress, _runState);
        runRow.Margin = new Padding(3, 8, 3, 8);

        // --- Журнал.
        _log = new RichTextBox { ReadOnly = true, Dock = DockStyle.Fill, Font = new Font("Consolas", 9F), BackColor = Color.White, DetectUrls = false, BorderStyle = BorderStyle.None };
        var logBox = Group("Ход работы", _log);
        logBox.AutoSize = false;

        // --- Результаты.
        _openDaily = ResultButton("Суточный отчет", () => _last?.DailyPdf);
        _openDetailed = ResultButton("Детальный отчет", () => _last?.DetailedPdf);
        _openCharts = ResultButton("Диаграммы", () => _last?.ChartsPdf);
        _openBook = ResultButton("Сводная книга Excel", () => _last?.Workbook);
        _openFolder = ResultButton("Папка с отчетами", () => _last?.Workbook is { } wb ? Path.GetDirectoryName(wb) : null);
        var results = Row(Caption("Открыть:"), _openDaily, _openDetailed, _openCharts, _openBook, _openFolder);

        root.Controls.Add(parameters, 0, 0);
        root.Controls.Add(notesBox, 0, 1);
        root.Controls.Add(runRow, 0, 2);
        root.Controls.Add(logBox, 0, 3);
        root.Controls.Add(results, 0, 4);
        page.Controls.Add(root);
        return page;
    }

    private TabPage BuildProjectsTab()
    {
        var page = new TabPage("Этапы") { Padding = new Padding(8), UseVisualStyleBackColor = true };
        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4 };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 55));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 45));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var add = new Button { Text = "Добавить этап", AutoSize = true };
        add.Click += (_, _) => AddProject();
        var remove = new Button { Text = "Удалить", AutoSize = true };
        remove.Click += (_, _) => RemoveProject();
        var up = new Button { Text = "Выше", AutoSize = true };
        up.Click += (_, _) => MoveProject(-1);
        var down = new Button { Text = "Ниже", AutoSize = true };
        down.Click += (_, _) => MoveProject(1);
        var check = new Button { Text = "Проверить файлы экранов", AutoSize = true };
        check.Click += (_, _) => CheckFiles(showMessage: true);
        var toolbar = Row(add, remove, up, down, check);

        _projects = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            RowHeadersVisible = false,
            BackgroundColor = SystemColors.Window,
            EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
        };
        _projects.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Вкл.", Width = 45 });
        _projects.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Этап", Width = 60 });
        _projects.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Название на диаграммах", Width = 170 });
        _projects.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Файл экрана", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, MinimumWidth = 200 });
        _projects.Columns.Add(new DataGridViewButtonColumn { HeaderText = "", Text = "Обзор…", UseColumnTextForButtonValue = true, Width = 75 });
        _projects.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Окончание экспертизы", Width = 150 });
        _projects.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Начало диаграмм", Width = 130 });
        _projects.Columns[ColName].ToolTipText = "Если пусто — «N этап».";
        _projects.Columns[ColEnd].ToolTipText = "дд.мм.гггг. Если пусто — наибольшая дата из сроков этапа.";
        _projects.Columns[ColChartStart].ToolTipText = "дд.мм.гггг. Если пусто — с первой заполненной строки экрана после начала сроков.";
        _projects.CellContentClick += OnProjectCellClick;
        _projects.CellValueChanged += (_, e) =>
        {
            MarkDirty();
            if (e.ColumnIndex == ColNumber)
                UpdateDeadlinesCaption();
        };
        _projects.CurrentCellDirtyStateChanged += (_, _) =>
        {
            if (_projects.IsCurrentCellDirty && _projects.CurrentCell is DataGridViewCheckBoxCell)
                _projects.CommitEdit(DataGridViewDataErrorContexts.Commit);
        };
        _projects.SelectionChanged += (_, _) => ShowDeadlines(SelectedProject());

        _deadlines = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = true,
            AllowUserToDeleteRows = true,
            AllowUserToResizeRows = false,
            BackgroundColor = SystemColors.Window,
            EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2,
        };
        _deadlines.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Стадия", Width = 200 });
        _deadlines.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Начало", Width = 160 });
        _deadlines.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Окончание", Width = 160 });
        _deadlines.CellValueChanged += (_, _) =>
        {
            if (_loading)
                return;
            CommitDeadlines(endEdit: false);
            MarkDirty();
        };
        _deadlines.UserDeletedRow += (_, _) =>
        {
            CommitDeadlines(endEdit: false);
            MarkDirty();
        };
        _deadlinesBox = Group("Сроки этапа", _deadlines);
        _deadlinesBox.AutoSize = false;

        var hint = Hint("Даты — в формате дд.мм.гггг. Вместо дат можно написать текст в колонке «Начало» (например, «Заявка до 23.10.2026»). " +
                        "Текущая стадия подсвечивается в отчете. Порядок этапов здесь — порядок строк в отчете.");
        hint.MaximumSize = new Size(1000, 0);

        root.Controls.Add(toolbar, 0, 0);
        root.Controls.Add(_projects, 0, 1);
        root.Controls.Add(_deadlinesBox, 0, 2);
        root.Controls.Add(hint, 0, 3);
        page.Controls.Add(root);
        return page;
    }

    private TabPage BuildSettingsTab()
    {
        var page = new TabPage("Настройки") { Padding = new Padding(8), UseVisualStyleBackColor = true };
        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3 };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        _configLabel = new Label { AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(3, 3, 3, 6) };
        _grid = new PropertyGrid { Dock = DockStyle.Fill, PropertySort = PropertySort.Categorized, ToolbarVisible = false, HelpVisible = true };
        _grid.PropertyValueChanged += (_, _) =>
        {
            MarkDirty();
            RefreshTimes();
            UpdatePeriod();
        };

        var save = new Button { Text = "Сохранить настройки", AutoSize = true };
        save.Click += (_, _) => SaveConfig(showStatus: true);
        var reset = new Button { Text = "Значения по умолчанию", AutoSize = true };
        reset.Click += (_, _) =>
        {
            if (MessageBox.Show(this, "Вернуть все параметры этой вкладки к значениям по умолчанию? Этапы и сроки не изменятся.",
                    "Настройки", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;
            _settings = new SettingsView(new ReportConfig());
            _grid.SelectedObject = _settings;
            RefreshTimes();
            MarkDirty();
        };

        root.Controls.Add(_configLabel, 0, 0);
        root.Controls.Add(_grid, 0, 1);
        root.Controls.Add(Row(save, reset), 0, 2);
        page.Controls.Add(root);
        return page;
    }

    // ------------------------------------------------------------------ помощники разметки

    private static GroupBox Group(string text, Control content)
    {
        var box = new GroupBox { Text = text, Dock = DockStyle.Fill, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Padding = new Padding(8, 6, 8, 8) };
        content.Dock = DockStyle.Fill;
        box.Controls.Add(content);
        return box;
    }

    private static FlowLayoutPanel Row(params Control[] controls)
    {
        var row = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, FlowDirection = FlowDirection.LeftToRight, Margin = new Padding(0) };
        row.Controls.AddRange(controls);
        return row;
    }

    private static FlowLayoutPanel Stack(params Control[] controls)
    {
        var stack = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = false, FlowDirection = FlowDirection.TopDown, Margin = new Padding(0) };
        stack.Controls.AddRange(controls);
        return stack;
    }

    private static Label Caption(string text) => new() { Text = text, AutoSize = true, Margin = new Padding(3, 7, 3, 3) };

    private static Label Hint(string text) => new() { Text = text, AutoSize = true, ForeColor = Color.DimGray, Margin = new Padding(3, 4, 3, 3) };

    private Button ResultButton(string text, Func<string?> target)
    {
        var button = new Button { Text = text, AutoSize = true, Enabled = false };
        button.Click += (_, _) =>
        {
            if (target() is { } path)
                OpenPath(path);
        };
        return button;
    }

    // ================================================================== Настройки ↔ окно

    private static ReportConfig LoadOrCreate(string path)
    {
        if (File.Exists(path))
        {
            try
            {
                return ConfigLoader.Load(path);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Не удалось прочитать файл настроек {path}:\n{ex.Message}\n\nБудут использованы настройки по умолчанию.",
                    "ConverterReport", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        var config = AppPaths.NewConfig();
        config.BaseDirectory = Path.GetDirectoryName(path)!;
        return config;
    }

    private void LoadToUi()
    {
        _loading = true;
        try
        {
            var r = _config.Report;
            _settings = new SettingsView(_config);
            _grid.SelectedObject = _settings;
            RefreshTimes();
            SetNearestTime();
            _customStart.Checked = false;
            _outFolder.Text = r.OutputFolder;
            _daily.Checked = r.IncludeDaily;
            _detailed.Checked = r.IncludeDetailed;
            _charts.Checked = r.IncludeCharts;
            _engineAuto.Checked = r.PdfEngine == PdfEngine.Auto;
            _engineExcel.Checked = r.PdfEngine == PdfEngine.Excel;
            _engineBuiltin.Checked = r.PdfEngine == PdfEngine.Builtin;
            _notes.Text = _notesLoaded = ReadNotesText();

            _projects.Rows.Clear();
            foreach (var p in _config.Projects)
                AddProjectRow(p);
            if (_projects.Rows.Count > 0)
                _projects.Rows[0].Selected = true;
            ShowDeadlines(SelectedProject());

            _configLabel.Text = "Файл настроек: " + _configPath;
            SetStatus("Настройки: " + _configPath);
        }
        finally
        {
            _loading = false;
        }

        _dirty = false;
        UpdatePeriod();
    }

    private void CollectFromUi()
    {
        CommitDeadlines();
        _projects.EndEdit();
        _settings.ApplyTo(_config);

        var r = _config.Report;
        r.OutputFolder = string.IsNullOrWhiteSpace(_outFolder.Text) ? "Отчеты" : _outFolder.Text.Trim();
        r.IncludeDaily = _daily.Checked;
        r.IncludeDetailed = _detailed.Checked;
        r.IncludeCharts = _charts.Checked;
        r.PdfEngine = _engineExcel.Checked ? PdfEngine.Excel : _engineBuiltin.Checked ? PdfEngine.Builtin : PdfEngine.Auto;

        _config.Projects = _projects.Rows.Cast<DataGridViewRow>().Select(row =>
        {
            var source = (ProjectConfig)row.Tag!;
            return new ProjectConfig
            {
                Enabled = row.Cells[ColEnabled].Value is true,
                Number = Cell(row, ColNumber),
                Name = NullIfEmpty(Cell(row, ColName)),
                ScreenFile = Cell(row, ColFile),
                ExpertiseEndDate = NullIfEmpty(Cell(row, ColEnd)),
                ChartStartDate = NullIfEmpty(Cell(row, ColChartStart)),
                Deadlines = source.Deadlines,
            };
        }).ToList();
    }

    private bool SaveConfig(bool showStatus)
    {
        try
        {
            CollectFromUi();
            Directory.CreateDirectory(Path.GetDirectoryName(_configPath)!);
            ConfigLoader.Save(_config, _configPath);
            _config.BaseDirectory = Path.GetDirectoryName(_configPath)!;
            WriteNotesText();
            _dirty = false;
            if (showStatus)
                SetStatus($"Настройки сохранены: {_configPath}");
            return true;
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Не удалось сохранить настройки в {_configPath}:\n{ex.Message}\n\nСохраните их в другое место: Файл → Сохранить настройки как…",
                "Сохранение настроек", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
    }

    private void SaveConfigAs()
    {
        using var dialog = new SaveFileDialog
        {
            Filter = "Настройки ConverterReport (*.json)|*.json",
            FileName = Path.GetFileName(_configPath),
            InitialDirectory = Path.GetDirectoryName(_configPath),
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;
        _configPath = dialog.FileName;
        _configLabel.Text = "Файл настроек: " + _configPath;
        SaveConfig(showStatus: true);
    }

    private void OpenConfig()
    {
        if (!ConfirmDiscard())
            return;
        using var dialog = new OpenFileDialog
        {
            Filter = "Настройки ConverterReport (*.json)|*.json|Все файлы|*.*",
            InitialDirectory = Path.GetDirectoryName(_configPath),
        };
        if (dialog.ShowDialog(this) != DialogResult.OK)
            return;
        _configPath = dialog.FileName;
        _config = LoadOrCreate(_configPath);
        LoadToUi();
    }

    private bool ConfirmDiscard()
    {
        if (!_dirty)
            return true;
        var answer = MessageBox.Show(this, "Сохранить изменения настроек?", "ConverterReport", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
        return answer switch
        {
            DialogResult.Yes => SaveConfig(showStatus: false),
            DialogResult.No => true,
            _ => false,
        };
    }

    private void MarkDirty()
    {
        if (!_loading)
            _dirty = true;
    }

    private string? NotesPath() =>
        string.IsNullOrWhiteSpace(_config.Report.NotesFile) ? null : _config.ResolvePath(_config.Report.NotesFile);

    private string ReadNotesText()
    {
        var path = NotesPath();
        if (path is null || !File.Exists(path))
            return "";
        return string.Join(Environment.NewLine, ReportGenerator.ReadNotes(_config));
    }

    private void WriteNotesText()
    {
        var path = NotesPath();
        if (path is null || _notes.Text == _notesLoaded)
            return;
        if (_notes.Text.Trim().Length == 0 && !File.Exists(path))
            return;
        File.WriteAllText(path, _notes.Text.Trim() + Environment.NewLine, new System.Text.UTF8Encoding(false));
        _notesLoaded = _notes.Text;
    }

    // ================================================================== Дата и время отчета

    private void RefreshTimes()
    {
        var current = _time.Text;
        _time.Items.Clear();
        foreach (var t in _settings.ReportTimes.Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            _time.Items.Add(t);
        _time.Text = current;
    }

    private void SetNearestTime()
    {
        var slots = _time.Items.Cast<string>().Select(t => Formats.ParseTime(t)).Where(t => t is not null).Select(t => t!.Value);
        var nearest = ReportTime.Nearest(DateTime.Now, slots);
        _date.Value = nearest.Date;
        _time.Text = nearest.ToString("HH:mm");
    }

    /// <summary>Дата, время и начало периода из параметров запуска (-d, -t, --from).</summary>
    private void ApplyCommandLine()
    {
        if (_options.Date is { } date)
            _date.Value = date.Date;
        if (_options.Time is { } time)
            _time.Text = $"{time.Hours:00}:{time.Minutes:00}";
        if (_options.From is { } from)
        {
            _customStart.Checked = true;
            _start.Value = from;
        }

        UpdatePeriod();
    }

    private DateTime? SelectedReportTime()
    {
        var time = Formats.ParseTime(_time.Text);
        return time is null ? null : _date.Value.Date + time.Value;
    }

    private void UpdatePeriod()
    {
        var reportTime = SelectedReportTime();
        if (reportTime is null)
        {
            _period.Text = "Время отчета указывается в формате чч:мм.";
            _period.ForeColor = Color.Firebrick;
            return;
        }

        var settings = new ReportSettings
        {
            PeriodStartTime = _settings.PeriodStartTime,
            PeriodStartDayOffset = _settings.PeriodStartDayOffset,
        };
        var start = _customStart.Checked ? _start.Value : ReportTime.PeriodStart(settings, reportTime.Value);
        if (!_customStart.Checked && _start.Value != start)
            _start.Value = start;
        _period.ForeColor = start < reportTime ? Color.DimGray : Color.Firebrick;
        _period.Text = $"Отчетный период: с {Short(start)} по {Short(reportTime.Value)}";
    }

    private static string Short(DateTime value) => value.ToString("H:mm dd.MM.yyyy", Formats.Ru);

    // ================================================================== Формирование отчета

    private void RunReport()
    {
        if (_running)
            return;

        CollectFromUi();
        var errors = ConfigLoader.Validate(_config);
        var reportTime = SelectedReportTime();
        if (reportTime is null)
            errors.Add("Время отчета указывается в формате чч:мм.");
        if (errors.Count > 0)
        {
            MessageBox.Show(this, string.Join(Environment.NewLine, errors.Select(e => "• " + e)), "Проверьте настройки", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (!_daily.Checked && !_detailed.Checked && !_charts.Checked &&
            MessageBox.Show(this, "Не выбран ни один PDF. Сформировать только сводную книгу Excel?", "ConverterReport",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        SaveConfig(showStatus: false);
        var config = ConfigLoader.Clone(_config);
        var options = new RunOptions
        {
            Date = reportTime!.Value.Date,
            Time = reportTime.Value.TimeOfDay,
            PeriodStart = _customStart.Checked ? _start.Value : null,
            ExtraNotes = NotesPath() is null
                ? _notes.Lines.Where(l => !string.IsNullOrWhiteSpace(l)).ToList()
                : new List<string>(),
        };

        _log.Clear();
        _last = null;
        SetResultButtons();
        SetRunning(true);
        AppendLog($"Настройки: {_configPath}", Color.DimGray);

        var thread = new Thread(() =>
        {
            var log = new RunLog();
            log.Open(config.ResolvePath(config.Report.OutputFolder));
            try
            {
                var generator = new ReportGenerator(
                    m =>
                    {
                        log.Write("INFO", m);
                        Ui(() => AppendLog(m, Color.Black));
                    },
                    m =>
                    {
                        log.Write("WARN", m);
                        Ui(() => AppendLog("Внимание: " + m, Color.DarkOrange));
                    });
                var result = generator.Run(config, options);
                Ui(() => Finished(result, null));
            }
            catch (Exception ex)
            {
                log.Write("ERROR", ex.ToString());
                Ui(() => Finished(null, ex));
            }
            finally
            {
                log.Flush();
            }
        })
        {
            IsBackground = true,
            Name = "ConverterReport",
        };
        // Excel (COM) надежнее работает из STA-потока.
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
    }

    private void Finished(RunResult? result, Exception? error)
    {
        SetRunning(false);
        if (error is not null)
        {
            AppendLog("Ошибка: " + error.Message, Color.Firebrick);
            _runState.Text = "Ошибка — см. журнал ниже.";
            MessageBox.Show(this, error.Message, "Отчет не сформирован", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        _last = result;
        SetResultButtons();
        var warnings = result!.Warnings.Count;
        AppendLog(warnings > 0 ? $"Готово. Предупреждений: {warnings}." : "Готово.", Color.ForestGreen);
        _runState.Text = $"Готово: {DateTime.Now:HH:mm}" + (warnings > 0 ? $", предупреждений: {warnings}" : "");
    }

    private void SetRunning(bool running)
    {
        _running = running;
        _run.Enabled = !running;
        _openConfigItem.Enabled = !running;
        _progress.Visible = running;
        _runState.Text = running ? "Формирую отчет…" : _runState.Text;
        UseWaitCursor = running;
    }

    private void SetResultButtons()
    {
        _openDaily.Enabled = _last?.DailyPdf is { } a && File.Exists(a);
        _openDetailed.Enabled = _last?.DetailedPdf is { } b && File.Exists(b);
        _openCharts.Enabled = _last?.ChartsPdf is { } c && File.Exists(c);
        _openBook.Enabled = _last?.Workbook is { } d && File.Exists(d);
        _openFolder.Enabled = _last?.Workbook is not null;
    }

    private void AppendLog(string message, Color color)
    {
        _log.SelectionStart = _log.TextLength;
        _log.SelectionLength = 0;
        _log.SelectionColor = color;
        _log.AppendText($"{DateTime.Now:HH:mm:ss}  {message}{Environment.NewLine}");
        _log.SelectionColor = _log.ForeColor;
        _log.ScrollToCaret();
    }

    private void Ui(Action action)
    {
        if (IsDisposed || !IsHandleCreated)
            return;
        try
        {
            BeginInvoke(action);
        }
        catch (InvalidOperationException)
        {
            // Окно закрыто, пока формировался отчет.
        }
    }

    private string OutputFolder()
    {
        var folder = string.IsNullOrWhiteSpace(_outFolder.Text) ? "Отчеты" : _outFolder.Text.Trim();
        var full = _config.ResolvePath(folder);
        Directory.CreateDirectory(full);
        return full;
    }

    private void BrowseOutputFolder()
    {
        using var dialog = new FolderBrowserDialog { Description = "Папка для сводной книги и PDF", UseDescriptionForTitle = true };
        try
        {
            dialog.SelectedPath = OutputFolder();
        }
        catch
        {
            // Папку не удалось создать — откроется выбор с корня.
        }

        if (dialog.ShowDialog(this) == DialogResult.OK)
            _outFolder.Text = dialog.SelectedPath;
    }

    // ================================================================== Этапы

    private void AddProjectRow(ProjectConfig p)
    {
        var copy = new ProjectConfig
        {
            Deadlines = p.Deadlines.ToDictionary(kv => kv.Key, kv => kv.Value.ToList()),
        };
        var index = _projects.Rows.Add(p.Enabled, p.Number, p.Name ?? "", p.ScreenFile, null, p.ExpertiseEndDate ?? "", p.ChartStartDate ?? "");
        _projects.Rows[index].Tag = copy;
    }

    /// <summary>Выбранная строка; пока таблица не показана на экране, CurrentRow пустой — тогда первая выделенная.</summary>
    private DataGridViewRow? SelectedRow() =>
        _projects.CurrentRow
        ?? _projects.SelectedRows.Cast<DataGridViewRow>().FirstOrDefault()
        ?? (_projects.Rows.Count > 0 ? _projects.Rows[0] : null);

    private ProjectConfig? SelectedProject() => SelectedRow()?.Tag as ProjectConfig;

    private DataGridViewRow? RowOf(ProjectConfig project) =>
        _projects.Rows.Cast<DataGridViewRow>().FirstOrDefault(r => ReferenceEquals(r.Tag, project));

    private void AddProject()
    {
        var numbers = _projects.Rows.Cast<DataGridViewRow>().Select(r => Cell(r, ColNumber)).ToHashSet();
        var next = Enumerable.Range(1, 99).First(n => !numbers.Contains(n.ToString()));
        AddProjectRow(new ProjectConfig { Number = next.ToString() });
        var row = _projects.Rows[^1];
        _projects.CurrentCell = row.Cells[ColNumber];
        MarkDirty();
    }

    private void RemoveProject()
    {
        if (_projects.CurrentRow is not { } row)
            return;
        if (MessageBox.Show(this, $"Удалить этап {Cell(row, ColNumber)} из списка? Файлы не удаляются.", "Этапы",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;
        _deadlinesProject = null;
        _projects.Rows.Remove(row);
        ShowDeadlines(SelectedProject());
        MarkDirty();
    }

    private void MoveProject(int delta)
    {
        if (_projects.CurrentRow is not { } row)
            return;
        var index = row.Index;
        var target = index + delta;
        if (target < 0 || target >= _projects.Rows.Count)
            return;
        _projects.EndEdit();
        _projects.Rows.RemoveAt(index);
        _projects.Rows.Insert(target, row);
        _projects.CurrentCell = row.Cells[ColNumber];
        MarkDirty();
    }

    private void OnProjectCellClick(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex != ColBrowse)
            return;
        var row = _projects.Rows[e.RowIndex];
        var current = Cell(row, ColFile);
        using var dialog = new OpenFileDialog
        {
            Title = $"Файл «Экран» этапа {Cell(row, ColNumber)}",
            Filter = "Книги Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|Все файлы|*.*",
        };
        try
        {
            if (current.Length > 0)
            {
                var full = _config.ResolvePath(current);
                dialog.InitialDirectory = Path.GetDirectoryName(full);
                dialog.FileName = Path.GetFileName(full);
            }
        }
        catch
        {
            // Некорректный путь — диалог откроется по умолчанию.
        }

        if (dialog.ShowDialog(this) == DialogResult.OK)
        {
            row.Cells[ColFile].Value = dialog.FileName;
            row.Cells[ColFile].Style.BackColor = Color.Empty;
            MarkDirty();
        }
    }

    private void CheckFiles(bool showMessage)
    {
        int ok = 0, total = 0;
        var missing = new List<string>();
        foreach (DataGridViewRow row in _projects.Rows)
        {
            if (row.Cells[ColEnabled].Value is not true)
            {
                row.Cells[ColFile].Style.BackColor = Color.Empty;
                continue;
            }

            total++;
            var file = Cell(row, ColFile);
            bool exists;
            try
            {
                exists = file.Length > 0 && File.Exists(_config.ResolvePath(file));
            }
            catch
            {
                exists = false;
            }

            row.Cells[ColFile].Style.BackColor = exists ? Color.FromArgb(226, 239, 218) : Color.FromArgb(255, 199, 206);
            if (exists)
                ok++;
            else
                missing.Add($"этап {Cell(row, ColNumber)}: {(file.Length == 0 ? "файл не указан" : file)}");
        }

        if (showMessage)
        {
            var text = missing.Count == 0
                ? $"Все файлы экранов найдены ({ok} из {total})."
                : $"Найдено {ok} из {total}. Не найдены:\n" + string.Join("\n", missing);
            MessageBox.Show(this, text, "Проверка файлов", MessageBoxButtons.OK, missing.Count == 0 ? MessageBoxIcon.Information : MessageBoxIcon.Warning);
        }
    }

    private void ShowDeadlines(ProjectConfig? project)
    {
        if (ReferenceEquals(project, _deadlinesProject))
            return;
        CommitDeadlines();
        _deadlinesProject = project;
        var wasLoading = _loading;
        _loading = true;
        try
        {
            _deadlines.Rows.Clear();
            if (project is null)
            {
                _deadlinesBox.Text = "Сроки этапа";
                _deadlines.Enabled = false;
                return;
            }

            _deadlines.Enabled = true;
            UpdateDeadlinesCaption();
            var stages = _settings.DeadlineStages
                .Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                .Concat(project.Deadlines.Keys)
                .Distinct(StringComparer.OrdinalIgnoreCase);
            foreach (var stage in stages)
            {
                var values = project.Deadlines.FirstOrDefault(kv => string.Equals(kv.Key, stage, StringComparison.OrdinalIgnoreCase)).Value ?? new List<string>();
                _deadlines.Rows.Add(stage, values.ElementAtOrDefault(0) ?? "", values.ElementAtOrDefault(1) ?? "");
            }
        }
        finally
        {
            _loading = wasLoading;
        }
    }

    private void UpdateDeadlinesCaption()
    {
        if (_deadlinesProject is not null && RowOf(_deadlinesProject) is { } row)
            _deadlinesBox.Text = $"Сроки этапа {Cell(row, ColNumber)}";
    }

    /// <summary>Сохраняет таблицу сроков в выбранный этап.</summary>
    private void CommitDeadlines(bool endEdit = true)
    {
        if (_deadlinesProject is null)
            return;
        if (endEdit)
            _deadlines.EndEdit();
        var result = new Dictionary<string, List<string>>();
        foreach (DataGridViewRow row in _deadlines.Rows)
        {
            if (row.IsNewRow)
                continue;
            var stage = Cell(row, 0);
            var start = Cell(row, 1);
            var end = Cell(row, 2);
            if (stage.Length == 0 || (start.Length == 0 && end.Length == 0))
                continue;
            result[stage] = end.Length == 0 && Formats.ParseDate(start) is null ? new List<string> { start } : new List<string> { start, end };
        }

        _deadlinesProject.Deadlines = result;
    }

    private static string Cell(DataGridViewRow row, int column) => (row.Cells[column].Value?.ToString() ?? "").Trim();

    private static string? NullIfEmpty(string value) => value.Length == 0 ? null : value;

    // ================================================================== Обновление программы

    private async Task CheckUpdatesAsync(bool silent)
    {
        var current = Updater.CurrentVersion;
        UpdateInfo? info;
        try
        {
            SetStatus("Проверка обновлений…");
            info = await Updater.CheckAsync(_config.Updates.Repository, current);
        }
        catch (Exception ex)
        {
            SetStatus("Не удалось проверить обновления (нет доступа к GitHub?).");
            if (!silent)
                MessageBox.Show(this, $"Не удалось проверить обновления:\n{ex.Message}", "Обновление", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (info is null)
        {
            SetStatus($"Установлена последняя версия ({current}).");
            if (!silent)
                MessageBox.Show(this, $"У вас последняя версия программы ({current}).", "Обновление", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        SetStatus($"Доступна новая версия {info.Version}.");
        var notes = info.Notes.Length > 600 ? info.Notes[..600] + "…" : info.Notes;
        var question = $"Доступна новая версия {info.Version} (у вас {current}).\n\n{notes}\n\nОбновить сейчас? Программа перезапустится, настройки сохранятся.";
        if (MessageBox.Show(this, question, "Обновление программы", MessageBoxButtons.YesNo, MessageBoxIcon.Information) != DialogResult.Yes)
            return;
        if (_running)
        {
            MessageBox.Show(this, "Дождитесь окончания формирования отчета и повторите: Справка → Проверить обновления.", "Обновление",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        if (!Updater.CanWrite(AppPaths.ExeFolder))
        {
            if (MessageBox.Show(this, $"Нет прав на запись в папку программы:\n{AppPaths.ExeFolder}\n\nОткрыть страницу загрузки новой версии?",
                    "Обновление", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                OpenPath(info.PageUrl);
            return;
        }

        try
        {
            if (_dirty)
                SaveConfig(showStatus: false);
            _updateProgress.Value = 0;
            _updateProgress.Visible = true;
            SetStatus($"Загрузка версии {info.Version}…");
            var exe = await Updater.DownloadAsync(info, new Progress<int>(p => _updateProgress.Value = Math.Clamp(p, 0, 100)));
            Updater.ReplaceAndRestart(exe, AppPaths.ExePath, $"--updated -c \"{_configPath}\"");
            _closingForUpdate = true;
            Close();
        }
        catch (Exception ex)
        {
            SetStatus("Обновление не выполнено.");
            MessageBox.Show(this, $"Не удалось обновить программу:\n{ex.Message}\n\nНовую версию можно скачать вручную: {info.PageUrl}", "Обновление",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            _updateProgress.Visible = false;
        }
    }

    // ================================================================== Прочее

    private void OnFormClosing(object? sender, FormClosingEventArgs e)
    {
        if (_closingForUpdate)
            return;
        if (_running && MessageBox.Show(this, "Отчет еще формируется. Закрыть программу?", "ConverterReport",
                MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
        {
            e.Cancel = true;
            return;
        }

        if (!ConfirmDiscard())
            e.Cancel = true;
    }

    private void SetStatus(string text) => _status.Text = text;

    private void OpenPath(string path)
    {
        try
        {
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Не удалось открыть {path}:\n{ex.Message}", "ConverterReport", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
