namespace ConverterReport.Gui;

/// <summary>Окно с моноширинным текстом (справка по параметрам запуска).</summary>
internal static class TextDialog
{
    public static void Show(IWin32Window? owner, string title, string text)
    {
        using var form = new Form
        {
            Text = title,
            StartPosition = owner is null ? FormStartPosition.CenterScreen : FormStartPosition.CenterParent,
            ClientSize = new Size(820, 520),
            MinimizeBox = false,
            ShowInTaskbar = owner is null,
            Font = new Font("Segoe UI", 9F),
        };
        var box = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            WordWrap = false,
            Dock = DockStyle.Fill,
            Font = new Font("Consolas", 9.5F),
            Text = text.ReplaceLineEndings(Environment.NewLine),
            BackColor = SystemColors.Window,
        };
        var close = new Button { Text = "Закрыть", DialogResult = DialogResult.OK, AutoSize = true, Anchor = AnchorStyles.Right };
        var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Padding = new Padding(6) };
        bottom.Controls.Add(close);
        form.Controls.Add(box);
        form.Controls.Add(bottom);
        form.AcceptButton = close;
        form.CancelButton = close;
        form.Shown += (_, _) => box.Select(0, 0);
        form.ShowDialog(owner);
    }
}
