using ClosedXML.Excel;

namespace ConverterReport.Core.Workbook;

/// <summary>Оформление ячеек сводной книги (шрифт Times New Roman, цвета как в книге отчета).</summary>
internal static class XlStyle
{
    public const string Font = "Times New Roman";
    public const string DateFormat = "dd.mm.yyyy";
    public const string TimeFormat = "h:mm";
    public const string DateTimeFormat = "dd.mm.yyyy h:mm";
    public const string IntFormat = "#,##0";
    public const string PercentFormat = "0.0";

    public static readonly XLColor Total = XLColor.FromHtml("#806000");
    public static readonly XLColor Fixed = XLColor.FromHtml("#00B050");
    public static readonly XLColor Fixing = XLColor.FromHtml("#FF0000");
    public static readonly XLColor Review = XLColor.FromHtml("#0070C0");
    public static readonly XLColor PercentFill = XLColor.FromHtml("#F2F2F2");
    public static readonly XLColor ErrorFill = XLColor.FromHtml("#FFC7CE");
    public static readonly XLColor CurrentStageFill = XLColor.FromHtml("#E2EFDA");
    public static readonly XLColor HeaderFill = XLColor.FromHtml("#D9E1F2");
    public static readonly XLColor ConverterFill = XLColor.FromHtml("#FFF2CC");

    /// <summary>Цвета статусов: Всего, Устранено, Устраняется, Рассматривается.</summary>
    public static readonly XLColor[] Status = { Total, Fixed, Fixing, Review };

    public static IXLRange Set(this IXLRange range, double size, bool bold = false, bool italic = false, XLColor? color = null,
        XLAlignmentHorizontalValues align = XLAlignmentHorizontalValues.Center, bool wrap = false, string? format = null)
    {
        var st = range.Style;
        st.Font.FontName = Font;
        st.Font.FontSize = size;
        st.Font.Bold = bold;
        st.Font.Italic = italic;
        st.Font.FontColor = color ?? XLColor.Black;
        st.Alignment.Horizontal = align;
        st.Alignment.Vertical = XLAlignmentVerticalValues.Center;
        st.Alignment.WrapText = wrap;
        if (format is not null)
            st.NumberFormat.Format = format;
        return range;
    }

    /// <summary>Объединяет диапазон (если в нем больше одной ячейки) и задает значение/формулу в первой ячейке.</summary>
    public static IXLRange Put(this IXLWorksheet ws, int r1, int c1, int r2, int c2, object? value)
    {
        var range = ws.Range(r1, c1, r2, c2);
        if (r1 != r2 || c1 != c2)
            range.Merge();
        var cell = ws.Cell(r1, c1);
        switch (value)
        {
            case null:
                break;
            case Formula f:
                cell.FormulaA1 = f.Text;
                break;
            case string s:
                cell.Value = s;
                break;
            case DateTime dt:
                cell.Value = dt;
                break;
            case TimeSpan ts:
                cell.Value = ts.TotalDays;
                break;
            case int i:
                cell.Value = i;
                break;
            case double d:
                cell.Value = d;
                break;
            default:
                cell.Value = value.ToString();
                break;
        }

        return range;
    }

    public static IXLRange Put(this IXLWorksheet ws, int row, int c1, int c2, object? value) => ws.Put(row, c1, row, c2, value);

    public static void Box(this IXLRange range, XLBorderStyleValues outside = XLBorderStyleValues.Medium, XLBorderStyleValues inside = XLBorderStyleValues.Thin)
    {
        range.Style.Border.OutsideBorder = outside;
        range.Style.Border.InsideBorder = inside;
    }

    /// <summary>Номер этапа: число, если номер числовой, иначе текст.</summary>
    public static object Stage(string number) =>
        int.TryParse(number, out var n) ? n : number;

    public static XLCellValue StageCell(string number) =>
        int.TryParse(number, out var n) ? n : number;

    public static string Q(string sheet) => "'" + sheet.Replace("'", "''") + "'";
}

/// <summary>Формула для <see cref="XlStyle.Put(IXLWorksheet,int,int,int,int,object?)"/>.</summary>
internal sealed record Formula(string Text)
{
    public static Formula F(string text) => new(text.StartsWith('=') ? text[1..] : text);
}
