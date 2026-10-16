using ConverterReport.Core;
using ConverterReport.Core.Archive;
using ConverterReport.Core.Configuration;

namespace ConverterReport.Gui;

/// <summary>
/// Выбор данных прошлой версии отчета: какие этапы брать и под каким номером (на месте этапа 5 в одной версии
/// мог стоять этап 9). Берутся только отмеченные — лишнее не копируется.
/// </summary>
internal sealed class ArchiveStagesDialog : Form
{
    private const int ColTake = 0;
    private const int ColSheets = 1;
    private const int ColStage = 2;
    private const int ColPeriod = 3;
    private const int ColRows = 4;
    private const int ColNote = 5;

    private readonly ArchiveBook _book;
    private readonly DataGridView _grid;

    public ArchiveStagesDialog(ArchiveBook book, ArchiveSource current, IReadOnlyCollection<string> knownStages)
    {
        _book = book;
        Text = "Данные прошлой версии отчета: " + Path.GetFileName(book.Path);
        StartPosition = FormStartPosition.CenterParent;
        Size = new Size(980, 440);
        MinimumSize = new Size(760, 320);
        Font = new Font("Segoe UI", 9F);
        ShowInTaskbar = false;
        MinimizeBox = false;

        var intro = new Label
        {
            AutoSize = true,
            MaximumSize = new Size(940, 0),
            Margin = new Padding(3, 3, 3, 8),
            Text = "Отметьте этапы, данные которых нужно брать (только они попадут в «Свод» и диаграммы). " +
                   "Номер этапа можно исправить — например, если на листах Д.5/З.5 на самом деле данные этапа 9. " +
                   "Берутся строки за даты, которых нет в экране этапа; этапы, которых нет в списке этапов, добавятся " +
                   "как архивные (без файла экрана, только в «Своде» и сводной диаграмме). " +
                   $"Листов прошлых дней в книге: {book.DailySheets.Count} — они переносятся в общую книгу всегда.",
        };

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            RowHeadersVisible = false,
            SelectionMode = DataGridViewSelectionMode.CellSelect,
            BackgroundColor = SystemColors.Window,
            EditMode = DataGridViewEditMode.EditOnEnter,
        };
        _grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "Брать", Width = 50 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Листы", Width = 110, ReadOnly = true });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Этап", Width = 60 });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Данные за период", Width = 170, ReadOnly = true });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Строк (зам. / загр.)", Width = 125, ReadOnly = true });
        _grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "Откуда номер этапа", AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill, ReadOnly = true });
        _grid.CurrentCellDirtyStateChanged += (_, _) =>
        {
            if (_grid.IsCurrentCellDirty && _grid.CurrentCell is DataGridViewCheckBoxCell)
                _grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
        };

        foreach (var part in book.Parts)
        {
            var chosen = current.Stages.Keys.Any(k => string.Equals(k.Trim(), part.Key, StringComparison.OrdinalIgnoreCase));
            var stage = current.StageFor(part.Key) ?? part.SuggestedStage ?? "";
            var take = chosen ? current.StageFor(part.Key) is not null : true;
            var period = part.From is { } from && part.To is { } to ? $"{from:dd.MM.yyyy} – {to:dd.MM.yyyy}" : "нет данных";
            var note = part.Note + (knownStages.Contains(stage) ? "" : stage.Length > 0 ? "; этапа нет в списке — будет добавлен" : "");
            var index = _grid.Rows.Add(take, part.Sheets, stage, period, $"{part.Data.Remarks.Count} / {part.Data.Loads.Count}", note);
            _grid.Rows[index].Tag = part;
        }

        if (book.Parts.Count == 0)
            intro.Text = "В книге нет листов с данными этапов (Д.N/З.N или «Замечания»/«Документы»). " +
                         $"Листов прошлых дней: {book.DailySheets.Count} — они переносятся в общую книгу.";

        var ok = new Button { Text = "OK", DialogResult = DialogResult.OK, AutoSize = true, MinimumSize = new Size(90, 0) };
        var cancel = new Button { Text = "Отмена", DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(90, 0) };
        var all = new Button { Text = "Отметить все", AutoSize = true };
        all.Click += (_, _) => SetAll(true);
        var none = new Button { Text = "Снять все", AutoSize = true };
        none.Click += (_, _) => SetAll(false);
        var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, Dock = DockStyle.Fill, AutoSize = true };
        buttons.Controls.Add(cancel);
        buttons.Controls.Add(ok);
        buttons.Controls.Add(none);
        buttons.Controls.Add(all);
        AcceptButton = ok;
        CancelButton = cancel;

        var root = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Padding = new Padding(10) };
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.Controls.Add(intro, 0, 0);
        root.Controls.Add(_grid, 0, 1);
        root.Controls.Add(buttons, 0, 2);
        Controls.Add(root);

        FormClosing += (_, e) =>
        {
            if (DialogResult != DialogResult.OK)
                return;
            _grid.EndEdit();
            var empty = _grid.Rows.Cast<DataGridViewRow>().FirstOrDefault(r => r.Cells[ColTake].Value is true && Stage(r).Length == 0);
            if (empty is not null)
            {
                MessageBox.Show(this, $"Укажите номер этапа для листов {empty.Cells[ColSheets].Value}.", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                e.Cancel = true;
            }
        };
    }

    /// <summary>Выбранное соответствие: листы → этап («-» — не брать).</summary>
    public Dictionary<string, string> Stages =>
        _grid.Rows.Cast<DataGridViewRow>().ToDictionary(
            r => ((ArchivePart)r.Tag!).Key,
            r => r.Cells[ColTake].Value is true ? Stage(r) : "-");

    /// <summary>Номера выбранных этапов.</summary>
    public IEnumerable<string> ChosenStages => Stages.Values.Where(v => v != "-").Distinct();

    private static string Stage(DataGridViewRow row) => (row.Cells[ColStage].Value?.ToString() ?? "").Trim();

    private void SetAll(bool value)
    {
        foreach (DataGridViewRow row in _grid.Rows)
            row.Cells[ColTake].Value = value;
    }

    /// <summary>Читает книгу и показывает диалог; null — отмена или книгу не удалось прочитать.</summary>
    public static async Task<(ArchiveBook Book, Dictionary<string, string> Stages)?> Choose(
        IWin32Window owner, string path, ArchiveSource current, ScreenLayout layout, IReadOnlyCollection<string> knownStages)
    {
        ArchiveBook book;
        try
        {
            book = await Task.Run(() => ArchiveReader.Read(path, layout));
        }
        catch (Exception ex)
        {
            MessageBox.Show(owner, $"Не удалось прочитать {path}:\n{ex.Message}", "Прошлая версия отчета", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return null;
        }

        using var dialog = new ArchiveStagesDialog(book, current, knownStages);
        return dialog.ShowDialog(owner) == DialogResult.OK ? (book, dialog.Stages) : null;
    }

    /// <summary>Кратко для списка: «1, 2, 3-4, 9 (Д.5 → 9)» или «не выбраны».</summary>
    public static string Describe(ArchiveSource source)
    {
        var chosen = source.Stages.Where(kv => source.StageFor(kv.Key) is not null).ToList();
        if (chosen.Count == 0)
            return "этапы не выбраны (только листы дней)";
        var stages = Formats.CompressNumbers(chosen.Select(kv => kv.Value.Trim()).Distinct()
            .OrderBy(n => int.TryParse(n, out var x) ? x : int.MaxValue));
        var renamed = chosen.Where(kv => kv.Key != ArchiveReader.ScreenKey && kv.Key.Trim() != kv.Value.Trim())
            .Select(kv => $"Д.{kv.Key}/З.{kv.Key} → {kv.Value.Trim()}").ToList();
        return "этапы " + stages + (renamed.Count > 0 ? $" ({string.Join("; ", renamed)})" : "");
    }
}
